param(
    [switch]$InstallDependencies = $true
)

$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot

$PythonExe = $null
$PyCmd = Get-Command py -ErrorAction SilentlyContinue

try {
    $Candidate = (Get-Command python -ErrorAction Stop).Source
    $Resolved = & $Candidate -c "import sys; print(sys.executable)" 2>$null
    if ($LASTEXITCODE -eq 0 -and $Resolved) {
        $PythonExe = $Resolved.Trim()
    }
}
catch {}

if ((-not $PythonExe) -and $PyCmd) {
    $Resolved = & $PyCmd.Source -c "import sys; print(sys.executable)"
    if ($LASTEXITCODE -eq 0 -and $Resolved) {
        $PythonExe = $Resolved.Trim()
    }
}

if (-not $PythonExe -or -not (Test-Path $PythonExe)) {
    throw 'Python was not found. Install Python 3.10+ and rerun install.ps1.'
}
$PythonwExe = Join-Path (Split-Path -Parent $PythonExe) 'pythonw.exe'
if (-not (Test-Path $PythonwExe)) {
    $PythonwExe = $PythonExe
}

if ($InstallDependencies) {
    & $PythonExe -m pip install -r (Join-Path $ProjectRoot 'requirements.txt')
}

$LauncherScript = Join-Path $ProjectRoot 'run_launcher.py'
$TrayScript = Join-Path $ProjectRoot 'run_tray.py'

$ShellRoot = 'HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\RMRRPDFSplitter'
$CommandKey = Join-Path $ShellRoot 'command'

New-Item -Path $ShellRoot -Force | Out-Null
New-ItemProperty -Path $ShellRoot -Name '(default)' -Value 'Split PDF...' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $ShellRoot -Name 'Icon' -Value $PythonwExe -PropertyType String -Force | Out-Null

New-Item -Path $CommandKey -Force | Out-Null
$CommandValue = '"{0}" "{1}" "%1"' -f $PythonwExe, $LauncherScript
New-ItemProperty -Path $CommandKey -Name '(default)' -Value $CommandValue -PropertyType String -Force | Out-Null

$StartupFolder = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
$ShortcutPath = Join-Path $StartupFolder 'PDF Splitter Tray.lnk'

$Wsh = New-Object -ComObject WScript.Shell
$Shortcut = $Wsh.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $PythonwExe
$Shortcut.Arguments = '"{0}"' -f $TrayScript
$Shortcut.WorkingDirectory = $ProjectRoot
$Shortcut.IconLocation = $PythonwExe
$Shortcut.Save()

Start-Process -FilePath $PythonwExe -ArgumentList ('"{0}"' -f $TrayScript) -WorkingDirectory $ProjectRoot -WindowStyle Hidden

Write-Host 'Install complete.'
Write-Host 'Context menu entry registered for current user.'
Write-Host 'Tray startup shortcut created and app launched.'
