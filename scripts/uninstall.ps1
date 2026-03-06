$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$SplitShellRoot = 'HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\RMRRPDFSplitter'
if (Test-Path $SplitShellRoot) {
    Remove-Item -Path $SplitShellRoot -Recurse -Force
}

$CombineExtensions = @('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp')
foreach ($Extension in $CombineExtensions) {
    $CombineShellRoot = "HKCU:\Software\Classes\SystemFileAssociations\$Extension\shell\RMRRCombineToPDF"
    if (Test-Path $CombineShellRoot) {
        Remove-Item -Path $CombineShellRoot -Recurse -Force
    }
}

$ImageExtensions = @('.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp')
foreach ($Extension in $ImageExtensions) {
    $ConvertShellRoot = "HKCU:\Software\Classes\SystemFileAssociations\$Extension\shell\RMRRConvertImageToPDF"
    if (Test-Path $ConvertShellRoot) {
        Remove-Item -Path $ConvertShellRoot -Recurse -Force
    }
}

$StartupFolder = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
$ShortcutPath = Join-Path $StartupFolder 'PDF Splitter Tray.lnk'
if (Test-Path $ShortcutPath) {
    Remove-Item $ShortcutPath -Force
}

Write-Host 'Uninstall complete.'
Write-Host 'Context menu entries and startup shortcut removed.'
Write-Host 'If a tray instance is currently running, quit it from the tray icon.'
