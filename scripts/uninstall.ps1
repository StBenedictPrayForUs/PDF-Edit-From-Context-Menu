$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$ShellRoot = 'HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\RMRRPDFSplitter'
if (Test-Path $ShellRoot) {
    Remove-Item -Path $ShellRoot -Recurse -Force
}

$StartupFolder = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
$ShortcutPath = Join-Path $StartupFolder 'PDF Splitter Tray.lnk'
if (Test-Path $ShortcutPath) {
    Remove-Item $ShortcutPath -Force
}

Write-Host 'Uninstall complete.'
Write-Host 'Context menu and startup shortcut removed.'
Write-Host 'If a tray instance is currently running, quit it from the tray icon.'
