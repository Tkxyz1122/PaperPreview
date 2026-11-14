"""Local-only paper browser. File moves occur only through explicit UI actions."""
from __future__ import annotations

import argparse
import base64
import copy
import ctypes
from ctypes import wintypes
from datetime import datetime
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path, PurePosixPath
import mimetypes
import os
import re
import secrets
import shutil
import struct
import subprocess
import sys
import threading
import time
import unicodedata
from urllib.parse import parse_qs, quote, urlsplit
from urllib.request import urlopen
import webbrowser

READ_FOLDER = 'HaveRead'
READ_NAMES = {'haveread', '已读', '_论文（已阅）'}
SUPPORTED = {'.pdf', '.doc', '.docx', '.ppt', '.pptx', '.epub', '.djvu', '.md', '.txt', '.rtf'}
APP_ID = 'paperpreview-v3'
DEFAULT_INTERPRETATION_ROOT = ''
PDF_RENDER_LOCK = threading.Lock()
INTERPRETATION_SUFFIX = re.compile(
    r'(?:[\s_\-—:：]+)?(?:论文深度解读|论文解读|深度解读|论文精读|阅读笔记|精读笔记|解读|'
    r'[ _-](?:deep[ _-]+analysis|reading[ _-]+notes|analysis|notes))'
    r'(?:[\s_\-—]*(?:v\d+(?:\.\d+)*|修订版|更新版|中文版|英文版))?$', re.I)


class LibraryError(Exception):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.status = status


def is_read_path(relative: str) -> bool:
    return any(part.casefold() in READ_NAMES for part in PurePosixPath(relative).parts[:-1])


def logical_folder(relative: str) -> str:
    return '/'.join(part for part in PurePosixPath(relative).parts[:-1] if part.casefold() not in READ_NAMES)


def name_tokens(stem):
    text = unicodedata.normalize('NFKC', stem).casefold()
    return tuple(re.findall(r'[a-z]+|\d+|[\u3400-\u9fff]+', text))


def interpretation_identity(name):
    stem = Path(name).stem
    match = INTERPRETATION_SUFFIX.search(stem)
    return name_tokens(stem[:match.start()].rstrip(' _-—:：')) if match else None


def associate_interpretations(records):
    """Attach notes by filename identity; never choose an ambiguous original."""
    sources, notes = [], []
    for item in records:
        identity = interpretation_identity(item['name']) if item.get('ext', '').lower() == '.pdf' else None
        if identity is not None:
            notes.append((item, identity))
        else:
            item['interpretations'] = []
            sources.append(item)
    unmatched = []
    for note, identity in notes:
        candidates = [source for source in sources if source.get('ext', '').lower() == '.pdf'
            and identity and name_tokens(Path(source['name']).stem)[:len(identity)] == identity]
        exact = [source for source in candidates if name_tokens(Path(source['name']).stem) == identity]
        if exact:
            candidates = exact
        same_folder = [source for source in candidates if source['folder'] == note['folder']]
        if same_folder:
            candidates = same_folder
        if len(candidates) == 1:
            note['sourcePath'] = candidates[0]['path']
            candidates[0]['interpretations'].append(note)
        else:
            note['matchReason'] = '有多份候选原文，未自动关联' if candidates else '尚未找到同名或同简称的原文'
            unmatched.append(note)
    for source in sources:
        source['interpretations'].sort(key=lambda note: note['name'].casefold())
    return sources, unmatched


def copy_file_to_clipboard(path: Path) -> None:
    """CF_HDROP + preferred COPY, so Explorer/chat apps receive a real file."""
    if os.name != 'nt':
        raise LibraryError('复制 PDF 文件需要 Windows 桌面环境。', 501)
    user = ctypes.WinDLL('user32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    user.OpenClipboard.argtypes = [wintypes.HWND]
    user.OpenClipboard.restype = wintypes.BOOL
    user.CloseClipboard.restype = wintypes.BOOL
    user.EmptyClipboard.restype = wintypes.BOOL
    user.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user.SetClipboardData.restype = wintypes.HANDLE
    user.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    user.RegisterClipboardFormatW.restype = wintypes.UINT
    user.CreateWindowExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.HWND, wintypes.HMENU,
        wintypes.HINSTANCE, ctypes.c_void_p]
    user.CreateWindowExW.restype = wintypes.HWND
    user.DestroyWindow.argtypes = [wintypes.HWND]
    user.DestroyWindow.restype = wintypes.BOOL
    kernel.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel.GlobalLock.restype = ctypes.c_void_p
    kernel.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel.GlobalFree.argtypes = [wintypes.HGLOBAL]
    kernel.GlobalFree.restype = wintypes.HGLOBAL

    def allocate(data):
        handle = kernel.GlobalAlloc(0x0002, len(data))  # GMEM_MOVEABLE
        if not handle:
            raise ctypes.WinError(ctypes.get_last_error())
        pointer = kernel.GlobalLock(handle)
        if not pointer:
            kernel.GlobalFree(handle)
            raise ctypes.WinError(ctypes.get_last_error())
        ctypes.memmove(pointer, data, len(data))
        kernel.GlobalUnlock(handle)
        return handle

    # DROPFILES is 20 bytes: pFiles, POINT{x,y}, fNC, fWide.
    dropfiles = struct.pack('<IiiII', 20, 0, 0, 0, 1) + (str(path) + '\0\0').encode('utf-16le')
    handles = [(15, allocate(dropfiles))]
    effect_format = user.RegisterClipboardFormatW('Preferred DropEffect')
    handles.append((effect_format, allocate(struct.pack('<I', 1))))  # DROPEFFECT_COPY
    opened = False
    # EmptyClipboard needs a real owner window for subsequent SetClipboardData.
    owner = None
    try:
        owner = user.CreateWindowExW(0, 'STATIC', 'Paper PDF clipboard', 0, 0, 0, 0, 0, None, None, None, None)
        if not owner:
            raise ctypes.WinError(ctypes.get_last_error())
        for _ in range(15):
            if user.OpenClipboard(owner):
                opened = True
                break
            time.sleep(0.04)
        if not opened:
            raise LibraryError('剪贴板正被其他程序占用，请稍后再点一次。', 409)
        if not user.EmptyClipboard():
            raise ctypes.WinError(ctypes.get_last_error())
        for index, (fmt, handle) in enumerate(handles):
            if not user.SetClipboardData(fmt, handle):
                raise ctypes.WinError(ctypes.get_last_error())
            handles[index] = (fmt, None)  # Windows now owns the allocation.
    finally:
        if opened:
            user.CloseClipboard()
        if owner:
            user.DestroyWindow(owner)
        for _, handle in handles:
            if handle:
                kernel.GlobalFree(handle)


class PaperLibrary:
    def __init__(self, root: Path | None, html: Path, clipboard_writer=copy_file_to_clipboard, runtime_dir: Path | None = None):
        self.root = root.resolve(strict=True) if root is not None else None
        self.runtime = (runtime_dir or (self.root / '.run' if self.root else html.parent)).resolve()
        self.html = html.resolve(strict=True)
        self.lock = threading.RLock()
        self.clipboard_writer = clipboard_writer
        self.metadata = {}
        self.file_metadata = {}
        self.covers = {}
        self.port = 0
        self.token = secrets.token_urlsafe(32)
        self.interpretation_library = None
        self.interpretation_root = ''
        page = self.html.read_text(encoding='utf-8')
        match = re.search(r'<script id="libraryData" type="application/json">(.*?)</script>', page, re.S)
        if match:
            initial = json.loads(match.group(1))
            for record in initial.get('files', []):
                self.metadata[(record.get('name'), record.get('size'))] = record
        self.root_id = hashlib.sha256(str(self.runtime).casefold().encode()).hexdigest()[:16]
        try:
            import pymupdf
            self.pdf_renderer = pymupdf
        except ImportError:
            self.pdf_renderer = None
        self.poppler = None  # Kept as an inert compatibility field; no global tools are used.
        self.state_path = self.runtime / 'library_state.json'
        try:
            self.state = json.loads(self.state_path.read_text(encoding='utf-8'))
            if not isinstance(self.state.get('records'), dict):
                raise ValueError('Invalid state')
        except (OSError, ValueError, TypeError):
            self.state = {'schema': 1, 'records': {}}

    @staticmethod
    def file_identity(path):
        stat = path.stat()
        return f'{stat.st_dev:x}:{stat.st_ino:x}'

    def _save_config(self, updates):
        self.runtime.mkdir(parents=True, exist_ok=True)
        config_path = self.runtime / 'library_config.json'
        try:
            settings = json.loads(config_path.read_text(encoding='utf-8-sig'))
        except (OSError, ValueError):
            settings = {}
        settings.update(updates)
        temporary = self.runtime / 'library_config.json.tmp'
        temporary.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temporary, config_path)

    @staticmethod
    def _directory(value, label):
        if not isinstance(value, str):
            raise LibraryError(f'请输入{label}的完整路径。')
        if not value.strip():
            return None
        path = Path(value.strip()).expanduser()
        if not path.is_absolute():
            raise LibraryError(f'{label}需要完整的绝对路径。')
        try:
            path = path.resolve(strict=True)
            if not path.is_dir():
                raise OSError('Not a directory')
        except OSError:
            raise LibraryError(f'{label}不存在或无法读取，请检查路径。', 404)
        return path

    def configure_paths(self, paper_root, interpretation_root, persist=True):
        paper = self._directory(paper_root, '论文存放文件夹')
        reading = self._directory(interpretation_root, '解读文件夹')
        if paper and (paper == self.runtime or paper.is_relative_to(self.runtime)):
            raise LibraryError('请选择实际存放论文的目录，程序目录不作为论文目录。')
        candidate = InterpretationLibrary(reading, self.html, self.clipboard_writer) if reading else None
        if candidate:
            candidate.token = self.token
        with self.lock:
            if persist:
                self._save_config({'paperRoot': str(paper) if paper else '',
                                   'interpretationRoot': str(reading) if reading else ''})
            self.root = paper
            self.interpretation_root = str(reading) if reading else ''
            self.interpretation_library = candidate
        return {'paperRoot': str(paper) if paper else '', 'interpretationRoot': self.interpretation_root}

    def _require_root(self):
        if self.root is None:
            raise LibraryError('请先设置论文存放路径。', 409)

    def _skip_directory(self, path):
        return path.resolve() == self.runtime or path.resolve().is_relative_to(self.runtime)

    def set_interpretation_root(self, directory, persist=False):
        if not isinstance(directory, str) or not directory.strip():
            raise LibraryError('请输入论文解读文件夹的完整路径。')
        raw = Path(directory.strip()).expanduser()
        if not raw.is_absolute():
            raise LibraryError('请输入完整的绝对路径，例如 D:\\论文解读。')
        try:
            target = raw.resolve(strict=True)
            if not target.is_dir():
                raise OSError('Not a directory')
            candidate = InterpretationLibrary(target, self.html, self.clipboard_writer)
            candidate.token = self.token
        except OSError:
            raise LibraryError('该文件夹不存在或无法读取，请检查路径。', 404)
        with self.lock:
            if persist:
                self._save_config({'interpretationRoot': str(target)})
            self.interpretation_library = candidate
            self.interpretation_root = str(target)
        return {'rootWindows': str(target)}

    def pick_directory(self, initial=''):
        """Open the native Windows folder picker and return the chosen absolute path."""
        if os.name != 'nt':
            raise LibraryError('文件夹选择仅支持 Windows，请手动填写路径。')
        executable = shutil.which('powershell.exe')
        if not executable:
            candidate = (Path(os.environ.get('SystemRoot', r'C:\Windows')) / 'System32'
                         / 'WindowsPowerShell' / 'v1.0' / 'powershell.exe')
            if candidate.exists():
                executable = str(candidate)
        if not executable:
            raise LibraryError('未找到系统文件夹选择器，请手动填写路径。', 503)
        self.runtime.mkdir(parents=True, exist_ok=True)
        marker = self.runtime / f'.pick-{secrets.token_hex(8)}.txt'
        script = r'''
$ErrorActionPreference = 'Stop'
$shell = New-Object -ComObject Shell.Application
$initial = [Environment]::GetEnvironmentVariable('PAPER_PICK_INITIAL')
if ($initial -and [System.IO.Directory]::Exists($initial)) {
    $folder = $shell.BrowseForFolder(0, '选择文件夹', 0, $initial)
} else {
    $folder = $shell.BrowseForFolder(0, '选择文件夹', 0)
}
if ($folder) {
    $path = $folder.Self.Path
    if ($path -and [System.IO.Directory]::Exists($path)) {
        [System.IO.File]::WriteAllText($env:PAPER_PICK_OUTPUT, $path)
    }
}
'''
        encoded = base64.b64encode(script.encode('utf-16-le')).decode('ascii')
        environment = dict(os.environ)
        environment['PAPER_PICK_OUTPUT'] = str(marker)
        environment['PAPER_PICK_INITIAL'] = initial.strip() if isinstance(initial, str) else ''
        flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        try:
            completed = subprocess.run(
                [executable, '-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-EncodedCommand', encoded],
                env=environment, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL, creationflags=flags, timeout=300)
        except subprocess.TimeoutExpired:
            raise LibraryError('文件夹选择超时，请重试或手动填写路径。', 408)
        except OSError:
            raise LibraryError('无法打开文件夹选择窗口，请手动填写路径。', 503)
        try:
            selected = marker.read_text(encoding='utf-8-sig').strip()
        except OSError:
            selected = ''
        finally:
            try:
                marker.unlink()
            except OSError:
                pass
        if completed.returncode != 0:
            raise LibraryError('文件夹选择窗口未能打开，请手动填写路径。', 503)
        return {'path': selected}

    def interpretation_scan(self, since=None):
        if self.interpretation_library is None:
            return {'schema': 3, 'collection': 'interpretations', 'rootWindows': self.interpretation_root,
                    'rootName': '论文解读', 'folders': [], 'files': [], 'unmatchedInterpretations': [],
                    'interpretationCount': 0, 'available': False, 'generatedAt': time.strftime('%Y-%m-%dT%H:%M:%S%z')}
        return self.interpretation_library.scan(since)

    def _save_state(self):
        self.runtime.mkdir(parents=True, exist_ok=True)
        temporary = self.runtime / 'library_state.json.tmp'
        temporary.write_text(json.dumps(self.state, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temporary, self.state_path)

    def _within_root(self, path):
        self._require_root()
        resolved = path.resolve()
        if resolved == self.root or not resolved.is_relative_to(self.root):
            raise LibraryError('文件操作超出论文资料范围。', 403)
        return resolved

    def _move_checked(self, source, target):
        source = self._within_root(source)
        self._within_root(target)
        if target.exists():
            raise LibraryError('目标已有同名文件或目录，未覆盖。', 409)
        target.parent.mkdir(parents=True, exist_ok=True)
        os.rename(source, target)

    def _annual_folder(self, year):
        canonical = self.root / f'{year}-01-12'
        if canonical.exists():
            self._within_root(canonical)
            if not canonical.is_dir():
                raise LibraryError(f'{canonical.name} 已存在且不是文件夹。', 409)
            return canonical
        for candidate in self.root.iterdir():
            match = re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})', candidate.name)
            if match and int(match[1]) == year and int(match[2]) == 1 and int(match[3]) == 12 and candidate.is_dir():
                self._within_root(candidate)
                return candidate
        return canonical

    def _quarter_folder(self, base, year, start):
        canonical = base / f'{year}-{start:02}-{start+2:02}'
        if canonical.exists() or not base.exists():
            return canonical
        for candidate in base.iterdir():
            match = re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})', candidate.name)
            if match and tuple(map(int, match.groups())) == (year, start, start+2) and candidate.is_dir():
                self._within_root(candidate)
                return candidate
        return canonical

    def _merge_archive(self, source, target, report):
        self._within_root(source)
        self._within_root(target)
        if not target.exists():
            self._move_checked(source, target)
            return True
        if not source.is_dir() or not target.is_dir():
            report['conflicts'].append(source.relative_to(self.root).as_posix())
            return False
        for item in list(source.iterdir()):
            destination = target / item.name
            if item.is_symlink():
                report['errors'].append('跳过链接：' + item.relative_to(self.root).as_posix())
                continue
            try:
                if item.is_dir() and destination.is_dir():
                    self._merge_archive(item, destination, report)
                elif destination.exists():
                    report['conflicts'].append(item.relative_to(self.root).as_posix())
                else:
                    self._move_checked(item, destination)
            except (LibraryError, OSError) as exc:
                report['errors'].append(item.relative_to(self.root).as_posix() + '：' + str(exc))
        if not any(source.iterdir()):
            self._within_root(source)
            source.rmdir()  # Remove only a verified, now-empty source folder.
            return True
        return False

    def organize(self, current_year=None):
        """Explicit refresh action: archive old quarters, then ingest root-level PDFs."""
        self._require_root()
        year_now = current_year if current_year is not None else datetime.now().year
        report = {'currentYear': year_now, 'imported': [], 'archived': [], 'conflicts': [], 'errors': []}
        with self.lock:
            # Archive first, so new old-year papers join an existing archived quarter.
            for source in list(self.root.iterdir()):
                match = re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})', source.name)
                if not match or not source.is_dir() or source.is_symlink():
                    continue
                year, start, end = map(int, match.groups())
                if year == year_now or start not in (1, 4, 7, 10) or end != start + 2:
                    continue
                try:
                    target = self._annual_folder(year) / source.name
                    relative = source.relative_to(self.root).as_posix()
                    if self._merge_archive(source, target, report):
                        report['archived'].append({'from': relative, 'to': target.relative_to(self.root).as_posix()})
                except (LibraryError, OSError) as exc:
                    report['errors'].append(source.name + '：' + str(exc))
            # Only loose PDFs dropped directly into the library root are new arrivals.
            # Papers in any subfolder retain their existing classification.
            for source in list(self.root.iterdir()):
                if source.name.startswith('.') or source.suffix.lower() != '.pdf' or source.is_symlink() or not source.is_file():
                    continue
                name = source.name
                try:
                    self._within_root(source)
                    stat = source.stat()
                    identity = self.file_identity(source)
                    saved = self.state['records'].get(identity)
                    added_at = saved['addedAt'] if saved else stat.st_mtime
                    joined = datetime.fromtimestamp(added_at)
                    first_month = (joined.month - 1) // 3 * 3 + 1
                    quarter = f'{joined.year}-{first_month:02}-{first_month+2:02}'
                    base = self.root if joined.year == year_now else self._annual_folder(joined.year)
                    parent = self._quarter_folder(base, joined.year, first_month)
                    target = parent / name
                    if target.exists():
                        report['conflicts'].append(source.relative_to(self.root).as_posix())
                        continue
                    old_path = source.relative_to(self.root).as_posix()
                    self._move_checked(source, target)
                    self.state['records'][identity] = {'addedAt': added_at, 'year': joined.year,
                        'quarter': quarter, 'path': target.relative_to(self.root).as_posix()}
                    report['imported'].append({'from': old_path, 'to': target.relative_to(self.root).as_posix(), 'addedAt': added_at})
                except (LibraryError, OSError, ValueError, OverflowError) as exc:
                    report['errors'].append(name + '：' + str(exc))
            # Identity follows same-volume renames, including manual annual grouping.
            for current, directories, names in os.walk(self.root, followlinks=False):
                directories[:] = [name for name in directories if not name.startswith('.') and not (Path(current)/name).is_symlink() and (Path(current)/name).resolve().is_relative_to(self.root) and not self._skip_directory(Path(current)/name)]
                for name in names:
                    p = Path(current) / name
                    if p.suffix.lower() != '.pdf':
                        continue
                    try:
                        self._within_root(p)
                        saved = self.state['records'].get(self.file_identity(p))
                        if saved:
                            saved['path'] = p.relative_to(self.root).as_posix()
                    except OSError:
                        pass
            if report['imported'] or self.state_path.exists():
                self._save_state()
        return report

    def resolve_document(self, relative, require_pdf=False):
        self._require_root()
        if not isinstance(relative, str) or '\0' in relative:
            raise LibraryError('无效的文件路径。')
        relative = relative.replace('\\', '/')
        pure = PurePosixPath(relative)
        if pure.is_absolute() or not pure.parts or any(p in ('.', '..') or p.startswith('.') or ':' in p for p in pure.parts):
            raise LibraryError('只能访问论文资料内的文档。', 403)
        try:
            path = (self.root / Path(*pure.parts)).resolve(strict=True)
        except (FileNotFoundError, OSError):
            raise LibraryError('原文件已被移动或删除，请刷新目录。', 404)
        if not path.is_relative_to(self.root) or not path.is_file() or path.suffix.lower() not in SUPPORTED:
            raise LibraryError('只能访问论文资料内的文档。', 403)
        if require_pdf and path.suffix.lower() != '.pdf':
            raise LibraryError('此操作仅适用于 PDF 文件。')
        return path

    def scan(self, since=None):
        with self.lock:
            if self.root is None:
                signature = hashlib.sha256(('unconfigured:' + self.interpretation_root).encode()).hexdigest()
                if since == signature:
                    return {'unchanged': True, 'signature': signature}
                return {'schema': 4, 'rootWindows': '', 'rootName': '论文库', 'folders': [], 'files': [],
                    'paperConfigured': False, 'interpretationRoot': self.interpretation_root,
                    'interpretationAvailable': self.interpretation_library is not None,
                    'currentYear': datetime.now().year, 'signature': signature,
                    'generatedAt': time.strftime('%Y-%m-%dT%H:%M:%S%z')}
            records, folders = [], []
            for current, directories, names in os.walk(self.root, followlinks=False):
                current_path = Path(current)
                directories[:] = [name for name in directories if not name.startswith('.') and name not in ('$RECYCLE.BIN', 'System Volume Information') and not (current_path / name).is_symlink() and (current_path / name).resolve().is_relative_to(self.root) and not self._skip_directory(current_path/name)]
                for name in directories:
                    folders.append((current_path / name).relative_to(self.root).as_posix())
                for name in names:
                    p = current_path / name
                    if name.startswith('.') or p.suffix.lower() not in SUPPORTED:
                        continue
                    try:
                        actual = p.resolve(strict=True)
                        if not actual.is_relative_to(self.root) or not actual.is_file():
                            continue
                        stat = p.stat()
                    except OSError:
                        continue
                    relative = p.relative_to(self.root).as_posix()
                    old = self.metadata.get((name, stat.st_size), {})
                    if not old and self.pdf_renderer and p.suffix.lower() == '.pdf':
                        meta_key = (str(p), stat.st_mtime_ns, stat.st_size)
                        if meta_key not in self.file_metadata:
                            try:
                                with PDF_RENDER_LOCK, self.pdf_renderer.open(str(p)) as document:
                                    metadata = document.metadata or {}
                                    title = re.sub(r'\s+', ' ', metadata.get('title', '') or '').strip()
                                    details = {'pages': document.page_count, 'author': (metadata.get('author', '') or '')[:600]}
                                    if len(title) > 12 and not re.search(r'^(untitled|microsoft|word|latex)|\.tex$|\btemplate\b', title, re.I):
                                        details['title'] = title
                                self.file_metadata[meta_key] = details
                            except Exception:
                                self.file_metadata[meta_key] = {}
                        old = self.file_metadata[meta_key]
                    record = {key: old[key] for key in ('title', 'pages', 'author') if key in old}
                    record.update(path=relative, name=name, folder=logical_folder(relative), size=stat.st_size,
                                  modified=stat.st_mtime, ext=p.suffix.lower(), readFromFolder=is_read_path(relative))
                    saved = self.state['records'].get(self.file_identity(p))
                    record['addedAt'] = saved.get('addedAt') if saved else None
                    if 'title' not in record:
                        record['title'] = p.stem.replace('_', ' ')
                    if old.get('thumb'):
                        record['thumb'] = old['thumb']
                    elif p.suffix.lower() == '.pdf' and self.pdf_renderer:
                        record['thumb'] = f'/api/thumbnail?path={quote(relative, safe="")}&v={stat.st_mtime_ns}&token={self.token}'
                    else:
                        record['thumb'] = ''
                    records.append(record)
            folders.sort()
            records.sort(key=lambda item: item['path'])
            reading_data = self.interpretation_scan() if self.interpretation_library is not None else None
            if reading_data:
                originals = {}
                for source in reading_data['files']:
                    key = unicodedata.normalize('NFKC', source['name']).casefold()
                    originals.setdefault(key, []).append(source)
                for record in records:
                    key = unicodedata.normalize('NFKC', record['name']).casefold()
                    candidates = originals.get(key, [])
                    if len(candidates) > 1:
                        candidates = [source for source in candidates if source['size'] == record['size']]
                    record['interpretations'] = [{key: note[key] for key in ('path', 'name', 'size', 'modified', 'collection') if key in note}
                        for note in candidates[0].get('interpretations', [])] if len(candidates) == 1 else []
            fingerprint = json.dumps([str(self.root), folders, [(r['path'], r['size'], r['modified']) for r in records],
                self.interpretation_root, reading_data.get('signature') if reading_data else None], ensure_ascii=False, separators=(',', ':'))
            signature = hashlib.sha256(fingerprint.encode()).hexdigest()
            if since == signature:
                return {'unchanged': True, 'signature': signature}
            return {'schema': 2, 'rootWindows': str(self.root), 'rootName': self.root.name,
                    'generatedAt': time.strftime('%Y-%m-%dT%H:%M:%S%z'), 'folders': folders,
                    'files': records, 'signature': signature, 'currentYear': datetime.now().year,
                    'interpretationRoot': self.interpretation_root, 'paperConfigured': True,
                    'interpretationAvailable': reading_data.get('available', False) if reading_data else False}

    def mark_read(self, relative, read=True):
        if not isinstance(read, bool):
            raise LibraryError('无效的已阅状态。')
        with self.lock:
            source = self.resolve_document(relative, require_pdf=True)
            old_relative = source.relative_to(self.root).as_posix()
            already_read = is_read_path(old_relative)
            if already_read == read:
                return {'path': old_relative, 'read': read, 'changed': False}
            if read:
                parent = source.parent / READ_FOLDER
            else:
                parent = self.root / Path(logical_folder(old_relative))
            parent_resolved = parent.resolve()
            if not parent_resolved.is_relative_to(self.root):
                raise LibraryError('目标文件夹超出论文资料范围。', 403)
            if parent.exists() and (not parent.is_dir() or parent.is_symlink()):
                raise LibraryError('目标文件夹不可用。', 409)
            target = parent / source.name
            if target.exists():
                raise LibraryError('目标文件夹已有同名文件，未移动、未覆盖。请手动处理后刷新。', 409)
            parent.mkdir(parents=True, exist_ok=True)
            try:
                # Same-folder-volume rename preserves the actual file, metadata and contents.
                # On Windows this refuses an existing destination, including a raced one.
                if os.name != 'nt' and target.exists():
                    raise FileExistsError(target)
                os.rename(source, target)
            except FileExistsError:
                raise LibraryError('目标文件夹已有同名文件，未覆盖。', 409)
            except OSError:
                raise LibraryError('文件移动失败，可能被其他程序占用。请关闭 PDF 后重试。', 409)
            return {'path': target.relative_to(self.root).as_posix(), 'previousPath': old_relative,
                    'read': read, 'changed': True}

    def copy_pdf(self, relative):
        with self.lock:
            path = self.resolve_document(relative, require_pdf=True)
            self.clipboard_writer(path)
            return {'copied': True, 'name': path.name}

    def thumbnail(self, relative):
        path = self.resolve_document(relative, require_pdf=True)
        if not self.pdf_renderer:
            raise LibraryError('请重新运行启动脚本，完成本地环境初始化。', 503)
        stat = path.stat()
        key = (str(path), stat.st_mtime_ns, stat.st_size)
        with PDF_RENDER_LOCK:
            if key not in self.covers:
                try:
                    with self.pdf_renderer.open(str(path)) as document:
                        page = document.load_page(0)
                        scale = 960 / max(page.rect.width, page.rect.height)
                        pixmap = page.get_pixmap(matrix=self.pdf_renderer.Matrix(scale, scale), alpha=False)
                        rendered = pixmap.tobytes('jpeg', jpg_quality=78)
                    if len(self.covers) >= 100:
                        self.covers.pop(next(iter(self.covers)))
                    self.covers[key] = rendered
                except Exception:
                    raise LibraryError('首页缩略图暂不可用，请打开 PDF 全文。', 404)
            return self.covers[key]


class InterpretationLibrary(PaperLibrary):
    """A separate directory exposed for reading and copying, with exact folder paths."""

    def __init__(self, root, html, clipboard_writer=copy_file_to_clipboard):
        super().__init__(root, html, clipboard_writer)
        match = re.search(r'<script id="libraryData" type="application/json">(.*?)</script>', self.html.read_text(encoding='utf-8'), re.S)
        payload = json.loads(match.group(1)) if match else {}
        initial = payload.get('readingLibrary', payload)
        for source in initial.get('files', []):
            for record in [source] + source.get('interpretations', []):
                self.metadata[(record.get('name'), record.get('size'))] = record

    def scan(self, since=None):
        data = super().scan()
        signature = hashlib.sha256((str(self.root) + data['signature']).encode()).hexdigest()
        if since == signature:
            return {'unchanged': True, 'signature': signature}
        for record in data['files']:
            record['folder'] = str(PurePosixPath(record['path']).parent)
            if record['folder'] == '.':
                record['folder'] = ''
            record['readFromFolder'] = False
            record['collection'] = 'interpretations'
            if record.get('thumb', '').startswith('/api/'):
                record['thumb'] = record['thumb'].replace('/api/thumbnail?', '/api/interpretations/thumbnail?', 1)
        sources, unmatched = associate_interpretations(data['files'])
        data.update(schema=3, collection='interpretations', signature=signature, available=self.root.is_dir(),
                    files=sources, unmatchedInterpretations=unmatched,
                    interpretationCount=len(unmatched) + sum(len(source['interpretations']) for source in sources))
        return data

    def organize(self, current_year=None):
        raise LibraryError('论文解读目录只读浏览，不参与整理。', 403)

    def mark_read(self, relative, read=True):
        raise LibraryError('论文解读目录保留原始结构，不移动文件。', 403)


def handler_for(library):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt, *args):
            pass

        def _headers(self, status, mime, length=None):
            self.send_response(status)
            self.send_header('Content-Type', mime)
            if length is not None:
                self.send_header('Content-Length', str(length))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Referrer-Policy', 'no-referrer')

        def json(self, body, status=200):
            data = json.dumps(body, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
            self._headers(status, 'application/json; charset=utf-8', len(data))
            self.end_headers()
            self.wfile.write(data)

        def _host_allowed(self):
            return self.headers.get('Host') in (f'127.0.0.1:{library.port}', f'localhost:{library.port}')

        def _authorized(self, query=None, mutation=False):
            if not self._host_allowed():
                return False
            token = self.headers.get('X-Paper-Token') or ((query or {}).get('token') or [''])[0]
            if not secrets.compare_digest(str(token), library.token):
                return False
            if mutation:
                origin = self.headers.get('Origin')
                if origin and origin not in (f'http://127.0.0.1:{library.port}', f'http://localhost:{library.port}'):
                    return False
                if self.headers.get('Content-Type', '').split(';')[0] != 'application/json':
                    return False
            return True

        def do_GET(self):
            try:
                parsed = urlsplit(self.path)
                query = parse_qs(parsed.query)
                if not self._host_allowed():
                    raise LibraryError('请求来源不受支持。', 403)
                if parsed.path == '/api/ping':
                    return self.json({'app': APP_ID, 'rootId': library.root_id})
                if parsed.path in ('/', '/index.html'):
                    text = library.html.read_text(encoding='utf-8')
                    text = text.replace('</head>', f'<meta name="paper-service-token" content="{library.token}"></head>', 1)
                    data = text.encode('utf-8')
                    self._headers(200, 'text/html; charset=utf-8', len(data))
                    self.end_headers()
                    self.wfile.write(data)
                    return
                if not self._authorized(query):
                    raise LibraryError('请从 PaperPreview 中的启动快捷方式打开完整功能页面。', 403)
                if parsed.path == '/api/library':
                    return self.json(library.scan(query.get('since', [None])[0]))
                if parsed.path == '/api/interpretations':
                    return self.json(library.interpretation_scan(query.get('since', [None])[0]))
                document_library = library
                document_route = parsed.path
                if parsed.path.startswith('/api/interpretations/'):
                    document_library = library.interpretation_library
                    if document_library is None:
                        raise LibraryError('论文解读目录无法读取，请检查读取路径。', 404)
                    document_route = parsed.path.replace('/api/interpretations/', '/api/', 1)
                relative = query.get('path', [''])[0]
                if document_route == '/api/thumbnail':
                    data = document_library.thumbnail(relative)
                    self._headers(200, 'image/jpeg', len(data))
                    self.end_headers()
                    self.wfile.write(data)
                    return
                if document_route == '/api/file':
                    path = document_library.resolve_document(relative)
                    total = path.stat().st_size
                    start, end = 0, total - 1
                    status = 200
                    byte_range = self.headers.get('Range', '')
                    if byte_range:
                        match = re.fullmatch(r'bytes=(\d+)-(\d*)', byte_range)
                        if not match:
                            raise LibraryError('不支持的文件范围。', 416)
                        start = int(match[1])
                        end = min(int(match[2]) if match[2] else end, end)
                        if start > end or start >= total:
                            raise LibraryError('无效的文件范围。', 416)
                        status = 206
                    mime = mimetypes.guess_type(path.name)[0] or 'application/octet-stream'
                    self._headers(status, mime, end - start + 1)
                    self.send_header('Accept-Ranges', 'bytes')
                    self.send_header('Content-Disposition', "inline; filename*=UTF-8''" + quote(path.name))
                    if status == 206:
                        self.send_header('Content-Range', f'bytes {start}-{end}/{total}')
                    self.end_headers()
                    with path.open('rb') as source:
                        source.seek(start)
                        remaining = end - start + 1
                        while remaining:
                            block = source.read(min(262144, remaining))
                            if not block:
                                break
                            self.wfile.write(block)
                            remaining -= len(block)
                    return
                raise LibraryError('未找到此页面。', 404)
            except LibraryError as exc:
                self.json({'error': str(exc)}, exc.status)
            except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
                pass
            except Exception:
                self.json({'error': '读取失败，请刷新目录后重试。'}, 500)

        def do_POST(self):
            try:
                if not self._authorized(mutation=True):
                    raise LibraryError('请从 PaperPreview 中的启动快捷方式打开完整功能页面。', 403)
                length = int(self.headers.get('Content-Length', 0))
                if not 0 < length <= 16384:
                    raise LibraryError('无效的请求。')
                try:
                    data = json.loads(self.rfile.read(length))
                except (ValueError, UnicodeError):
                    raise LibraryError('无效的请求内容。')
                endpoint = urlsplit(self.path).path
                if endpoint == '/api/mark-read':
                    if data.get('collection') == 'interpretations':
                        raise LibraryError('论文解读目录只读浏览，不移动文件。', 403)
                    return self.json(library.mark_read(data.get('path'), data.get('read', True)))
                if endpoint == '/api/copy-pdf':
                    if data.get('collection') == 'interpretations':
                        if library.interpretation_library is None:
                            raise LibraryError('请先配置论文解读目录。', 404)
                        return self.json(library.interpretation_library.copy_pdf(data.get('path')))
                    return self.json(library.copy_pdf(data.get('path')))
                if endpoint == '/api/settings/interpretations':
                    return self.json(library.set_interpretation_root(data.get('path'), persist=True))
                if endpoint == '/api/settings/paths':
                    return self.json(library.configure_paths(data.get('paperRoot'), data.get('interpretationRoot')))
                if endpoint == '/api/pick-directory':
                    return self.json(library.pick_directory(data.get('initial')))
                if endpoint == '/api/refresh-organize':
                    if data.get('collection') == 'interpretations':
                        raise LibraryError('论文解读目录只读浏览，不参与整理。', 403)
                    return self.json(library.organize())
                raise LibraryError('不支持此操作。', 404)
            except LibraryError as exc:
                self.json({'error': str(exc)}, exc.status)
            except Exception:
                self.json({'error': '操作未完成，请稍后重试。'}, 500)

    return Handler


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', type=Path)
    parser.add_argument('--html', type=Path)
    parser.add_argument('--runtime', type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument('--port', type=int, default=8873)
    parser.add_argument('--open', action='store_true')
    args = parser.parse_args()
    if sys.prefix == sys.base_prefix:
        raise SystemExit('请使用 PaperPreview 文件夹中的启动快捷方式，以本地 .venv 运行。')
    html = args.html or args.runtime / '论文资料浏览器.html'
    try:
        settings = json.loads((args.runtime / 'library_config.json').read_text(encoding='utf-8-sig'))
    except (OSError, ValueError):
        settings = {}
    library = PaperLibrary(None, html, runtime_dir=args.runtime)
    paper_root = str(args.root) if args.root else settings.get('paperRoot', '')
    reading_root = settings.get('interpretationRoot', '')
    try:
        library.configure_paths(paper_root, '', persist=False)
    except LibraryError:
        pass
    if reading_root:
        try:
            library.set_interpretation_root(reading_root)
        except LibraryError:
            library.interpretation_root = reading_root
    for port in range(args.port, args.port + 10):
        try:
            server = ThreadingHTTPServer(('127.0.0.1', port), handler_for(library))
            library.port = port
            break
        except OSError:
            try:
                with urlopen(f'http://127.0.0.1:{port}/api/ping', timeout=0.4) as response:
                    info = json.load(response)
                if info.get('app') == APP_ID and info.get('rootId') == library.root_id:
                    if args.open:
                        webbrowser.open(f'http://127.0.0.1:{port}/')
                    return
            except Exception:
                pass
    else:
        raise SystemExit('No available localhost port.')
    if args.open:
        webbrowser.open(f'http://127.0.0.1:{library.port}/')
    server.serve_forever()


if __name__ == '__main__':
    main()
