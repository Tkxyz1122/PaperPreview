param([switch]$NoOpen, [int]$Port = 8873)
$ErrorActionPreference = 'Stop'
$taskApp = $PSScriptRoot
$taskVenv = Join-Path $taskApp '.venv'
$taskPython = Join-Path $taskVenv 'Scripts\python.exe'
$taskLock = $null
try {
    for ($taskAttempt = 0; $taskAttempt -lt 60; $taskAttempt++) {
        try {
            $taskLock = [System.IO.File]::Open((Join-Path $taskApp '.environment.lock'), [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            break
        } catch [System.IO.IOException] { Start-Sleep -Milliseconds 500 }
    }
    if (-not $taskLock) { throw 'The first-start setup is still busy. Please try again shortly.' }
    if (-not (Test-Path -LiteralPath $taskPython)) {
        $taskCandidates = @()
        $taskProbe = "import sys,struct,platform; print(sys.executable if sys.version_info >= (3,11) and struct.calcsize('P') == 8 and platform.machine().lower() in ('amd64','x86_64') else '')"
        $taskLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
        if ($taskLauncher) {
            try {
                $taskLaunchedPython = & $taskLauncher.Source -3 -I -c $taskProbe 2>$null
                if ($LASTEXITCODE -eq 0 -and $taskLaunchedPython) { $taskCandidates += "$taskLaunchedPython".Trim() }
            } catch { }
        }
        foreach ($taskCommand in @('python.exe', 'python3.exe')) {
            $taskFound = Get-Command $taskCommand -ErrorAction SilentlyContinue
            if ($taskFound -and $taskFound.Source -notlike '*\WindowsApps\*') { $taskCandidates += $taskFound.Source }
        }
        $taskBasePython = $null
        foreach ($taskCandidate in $taskCandidates) {
            if (-not (Test-Path -LiteralPath $taskCandidate)) { continue }
            try {
                $taskVersion = & $taskCandidate -I -c $taskProbe 2>$null
                if ($LASTEXITCODE -eq 0 -and $taskVersion) { $taskBasePython = "$taskVersion".Trim(); break }
            } catch { }
        }
        if (-not $taskBasePython) { throw 'Install 64-bit Python 3.11 or newer for Windows x64, with the Python launcher or Add Python to PATH enabled.' }
        Write-Output 'PaperPreview: creating the local .venv for first start...'
        & $taskBasePython -I -m venv --copies $taskVenv
        if ($LASTEXITCODE -ne 0) { throw 'Could not create PaperPreview/.venv.' }
    }
    $taskRequirements = Join-Path $taskApp 'requirements.txt'
    $taskMarker = Join-Path $taskVenv '.paperpreview-ready'
    $taskHash = (Get-FileHash -LiteralPath $taskRequirements -Algorithm SHA256).Hash
    $taskReady = (Test-Path -LiteralPath $taskMarker) -and ((Get-Content -LiteralPath $taskMarker -Raw).Trim() -eq $taskHash)
    if (-not $taskReady) {
        Write-Output 'PaperPreview: installing bundled components into .venv...'
        & $taskPython -I -m pip install --disable-pip-version-check --no-index --find-links (Join-Path $taskApp 'wheels') -r $taskRequirements
        if ($LASTEXITCODE -ne 0) { throw 'Local component installation failed. Keep the wheels folder with PaperPreview.' }
        & $taskPython -I -c 'import sys,pymupdf; assert sys.prefix != sys.base_prefix'
        if ($LASTEXITCODE -ne 0) { throw 'Local environment validation failed.' }
        [System.IO.File]::WriteAllText($taskMarker, $taskHash, [System.Text.UTF8Encoding]::new($false))
    }
    $taskLock.Dispose()
    $taskLock = $null
    try {
        $taskShell = New-Object -ComObject WScript.Shell
        $taskShortcut = $taskShell.CreateShortcut((Join-Path $taskApp '打开论文预览.lnk'))
        $taskShortcut.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $taskShortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $PSCommandPath + '"'
        $taskShortcut.WorkingDirectory = $taskApp
        $taskShortcut.IconLocation = (Join-Path $taskApp 'PaperPreview.ico') + ',0'
        $taskShortcut.Description = 'PaperPreview - local paper browser'
        $taskShortcut.WindowStyle = 7
        $taskShortcut.Save()
    } catch { Write-Warning 'Could not create the optional shortcut. You can keep using the .cmd launcher.' }
    $taskScript = Join-Path $taskApp 'server.py'
    $taskQuotedScript = [string][char]34 + $taskScript + [string][char]34
    $taskArguments = @('-I', $taskQuotedScript, '--port', "$Port")
    if (-not $NoOpen) { $taskArguments += '--open' }
    $taskProcess = Start-Process -FilePath $taskPython -ArgumentList $taskArguments -WorkingDirectory $taskApp -WindowStyle Hidden -PassThru
    if ($NoOpen) { Write-Output ('PaperPreview PID: ' + $taskProcess.Id) }
} catch {
    $_ | Out-String | Set-Content -LiteralPath (Join-Path $taskApp 'startup-error.log') -Encoding UTF8
    Write-Error ('PaperPreview could not start. See startup-error.log. ' + $_.Exception.Message)
    exit 1
} finally {
    if ($taskLock) { $taskLock.Dispose() }
}
