"""Regressions for reading folders nested inside a paper library."""
from pathlib import Path
from http.server import ThreadingHTTPServer
from urllib.request import Request, urlopen
import hashlib
import importlib.util
import json
import tempfile
import threading
import unittest

import pymupdf

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location('paperpreview_server', ROOT / 'server.py')
SERVER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(SERVER)


class InterpretationVisibilityTest(unittest.TestCase):
    def setUp(self):
        scratch = ROOT / '_Temp'
        scratch.mkdir(exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(prefix='reading_visibility_', dir=scratch)
        self.base = Path(self.temp.name).resolve()
        self.scratch = scratch.resolve()
        assert self.base.is_relative_to(self.scratch)
        self.papers = self.base / 'papers'
        self.notes = self.papers / 'Reading'
        self.outside = self.base / 'Reading'
        self.runtime = self.base / 'runtime'
        for folder in (self.papers, self.notes, self.outside, self.runtime):
            folder.mkdir(parents=True, exist_ok=True)
        self.original = self.pdf(self.papers / '2026-10-12' / 'SigLIP 2.pdf')
        self.reading_original = self.pdf(self.notes / 'SigLIP 2.pdf')
        self.note = self.pdf(self.notes / 'SigLIP 2_论文解读.pdf')
        self.library = SERVER.PaperLibrary(None, ROOT / '论文资料浏览器.html', runtime_dir=self.runtime)

    def tearDown(self):
        assert self.base.is_relative_to(self.scratch) and self.base != self.scratch
        self.temp.cleanup()

    def pdf(self, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        with pymupdf.open() as document:
            document.new_page().insert_text((60, 80), 'PaperPreview fixture')
            document.save(path)
        return path

    def configure(self, notes=None):
        self.library.configure_paths(str(self.papers), str(notes or self.notes))

    def test_nested_folder_hidden_but_notes_remain_attached_and_openable(self):
        self.configure()
        data = self.library.scan()
        self.assertEqual([item['path'] for item in data['files']], ['2026-10-12/SigLIP 2.pdf'])
        self.assertNotIn('Reading', data['folders'])
        attached = data['files'][0]['interpretations']
        self.assertEqual(len(attached), 1)
        self.assertEqual(attached[0]['name'], self.note.name)
        note_path = self.library.interpretation_library.resolve_document(attached[0]['path'])
        self.assertEqual(note_path, self.note.resolve())
        with self.assertRaises(SERVER.LibraryError):
            self.library.resolve_document('Reading/' + self.note.name)

    def test_path_relationship_not_folder_name_controls_hiding(self):
        self.pdf(self.papers / 'ReadingExtra' / 'Keep.pdf')
        self.configure()
        first = self.library.scan()
        self.assertIn('ReadingExtra', first['folders'])
        self.configure(self.outside)
        second = self.library.scan(first['signature'])
        self.assertFalse(second.get('unchanged', False))
        self.assertIn('Reading', second['folders'])
        self.assertEqual(len(second['files']), 4)
        self.library.configure_paths(str(self.papers), '')
        self.assertIn('Reading', self.library.scan()['folders'])

    def test_deep_folder_and_all_descendants_hidden(self):
        deep = self.papers / 'References' / 'Notes'
        self.pdf(deep / 'subfolder' / 'Nested.pdf')
        self.configure(deep)
        data = self.library.scan()
        self.assertIn('References', data['folders'])
        self.assertFalse(any(p.startswith('References/Notes') for p in data['folders']))
        self.assertFalse(any(item['path'].startswith('References/Notes/') for item in data['files']))

    def test_same_paper_and_reading_root_does_not_hide_entire_library(self):
        self.configure(self.papers)
        self.assertTrue(self.library.scan()['files'])
        self.assertIn('2026-10-12', self.library.scan()['folders'])

    def test_annual_archive_preserves_reading_directory_and_its_parent(self):
        for reading_relative in ('2024-01-03', '2024-04-06/Notes'):
            with self.subTest(reading_relative=reading_relative):
                reading = self.papers / reading_relative
                paper = self.pdf(reading / 'Reference.pdf')
                before = hashlib.sha256(paper.read_bytes()).hexdigest()
                self.configure(reading)
                self.library.organize(current_year=2026)
                self.assertTrue(paper.exists())
                self.assertEqual(hashlib.sha256(paper.read_bytes()).hexdigest(), before)
                self.assertEqual(self.library.interpretation_root, str(reading.resolve()))

    def test_http_settings_apply_exclusion_without_restart(self):
        http = ThreadingHTTPServer(('127.0.0.1', 0), SERVER.handler_for(self.library))
        self.library.port = http.server_port
        thread = threading.Thread(target=http.serve_forever, daemon=True)
        thread.start()
        base = f'http://127.0.0.1:{http.server_port}'
        headers = {'X-Paper-Token': self.library.token, 'Content-Type': 'application/json', 'Origin': base}
        try:
            body = json.dumps({'paperRoot': str(self.papers), 'interpretationRoot': str(self.notes)}).encode()
            with urlopen(Request(base + '/api/settings/paths', data=body, headers=headers)) as response:
                self.assertEqual(response.status, 200)
            with urlopen(Request(base + '/api/library', headers=headers)) as response:
                data = json.load(response)
            self.assertEqual(len(data['files']), 1)
            self.assertNotIn('Reading', data['folders'])
        finally:
            http.shutdown()
            http.server_close()
            thread.join(timeout=2)


if __name__ == '__main__':
    unittest.main()
