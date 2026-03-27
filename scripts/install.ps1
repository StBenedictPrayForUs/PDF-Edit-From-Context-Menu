param(
    [switch]$InstallDependencies = $true
)

$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot

function Get-TrayProcesses {
    Get-CimInstance Win32_Process |
        Where-Object {
            $_.Name -match 'pythonw?\.exe' -and
            ($_.CommandLine -match 'run_tray\.py' -or $_.CommandLine -match 'app\.tray_runtime')
        }
}

function Stop-TrayProcesses {
    $Processes = @(Get-TrayProcesses)
    foreach ($Process in $Processes) {
        Stop-Process -Id $Process.ProcessId -Force -ErrorAction SilentlyContinue
    }
    if ($Processes.Count -gt 0) {
        Start-Sleep -Seconds 1
    }
}

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

$SplitShellRoot = 'HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\RMRRPDFSplitter'
$SplitCommandKey = Join-Path $SplitShellRoot 'command'

New-Item -Path $SplitShellRoot -Force | Out-Null
New-ItemProperty -Path $SplitShellRoot -Name '(default)' -Value 'Split PDF...' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $SplitShellRoot -Name 'Icon' -Value $PythonwExe -PropertyType String -Force | Out-Null
New-ItemProperty -Path $SplitShellRoot -Name 'MultiSelectModel' -Value 'Single' -PropertyType String -Force | Out-Null

New-Item -Path $SplitCommandKey -Force | Out-Null
$SplitCommandValue = '"{0}" "{1}" "%1"' -f $PythonwExe, $LauncherScript
New-ItemProperty -Path $SplitCommandKey -Name '(default)' -Value $SplitCommandValue -PropertyType String -Force | Out-Null

$CombineExtensions = @('.pdf')
foreach ($Extension in $CombineExtensions) {
    $CombineShellRoot = "HKCU:\Software\Classes\SystemFileAssociations\$Extension\shell\RMRRCombineToPDF"
    $CombineCommandKey = Join-Path $CombineShellRoot 'command'

    New-Item -Path $CombineShellRoot -Force | Out-Null
    New-ItemProperty -Path $CombineShellRoot -Name '(default)' -Value 'Combine to PDF...' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CombineShellRoot -Name 'Icon' -Value $PythonwExe -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CombineShellRoot -Name 'MultiSelectModel' -Value 'Player' -PropertyType String -Force | Out-Null

    New-Item -Path $CombineCommandKey -Force | Out-Null
    $CombineCommandValue = '"{0}" "{1}" combine --from-explorer-selection "%1" %*' -f $PythonwExe, $LauncherScript
    New-ItemProperty -Path $CombineCommandKey -Name '(default)' -Value $CombineCommandValue -PropertyType String -Force | Out-Null
}

$StaleImageCombineExtensions = @('.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp')
foreach ($Extension in $StaleImageCombineExtensions) {
    $StaleCombineShellRoot = "HKCU:\Software\Classes\SystemFileAssociations\$Extension\shell\RMRRCombineToPDF"
    if (Test-Path $StaleCombineShellRoot) {
        Remove-Item -Path $StaleCombineShellRoot -Recurse -Force
    }
}

$ImageExtensions = @('.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp')
foreach ($Extension in $ImageExtensions) {
    $ConvertShellRoot = "HKCU:\Software\Classes\SystemFileAssociations\$Extension\shell\RMRRConvertImageToPDF"
    $ConvertCommandKey = Join-Path $ConvertShellRoot 'command'

    New-Item -Path $ConvertShellRoot -Force | Out-Null
    New-ItemProperty -Path $ConvertShellRoot -Name '(default)' -Value 'Convert to PDF...' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $ConvertShellRoot -Name 'Icon' -Value $PythonwExe -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $ConvertShellRoot -Name 'MultiSelectModel' -Value 'Player' -PropertyType String -Force | Out-Null

    New-Item -Path $ConvertCommandKey -Force | Out-Null
    $ConvertCommandValue = '"{0}" "{1}" convert-image --from-explorer-selection "%1" %*' -f $PythonwExe, $LauncherScript
    New-ItemProperty -Path $ConvertCommandKey -Name '(default)' -Value $ConvertCommandValue -PropertyType String -Force | Out-Null
}

$StartupFolder = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
$ShortcutPath = Join-Path $StartupFolder 'PDF Splitter Tray.lnk'

Stop-TrayProcesses

if (Test-Path $ShortcutPath) {
    Remove-Item $ShortcutPath -Force
}

$Wsh = New-Object -ComObject WScript.Shell
$Shortcut = $Wsh.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $PythonwExe
$Shortcut.Arguments = '"{0}"' -f $TrayScript
$Shortcut.WorkingDirectory = $ProjectRoot
$Shortcut.IconLocation = $PythonwExe
$Shortcut.Save()

Start-Process -FilePath $PythonwExe -ArgumentList ('"{0}"' -f $TrayScript) -WorkingDirectory $ProjectRoot -WindowStyle Hidden

for ($Attempt = 0; $Attempt -lt 10; $Attempt++) {
    Start-Sleep -Seconds 1
    $TrayProcesses = @(Get-TrayProcesses)
    if ($TrayProcesses.Count -ge 1) {
        break
    }
}

if ($TrayProcesses.Count -ne 1) {
    if ($TrayProcesses.Count -gt 1) {
        Stop-TrayProcesses
    }
    throw "Tray launch verification failed. Expected 1 running tray process, found $($TrayProcesses.Count)."
}

Write-Host 'Install complete.'
Write-Host 'Context menu entries registered for current user.'
Write-Host 'Tray startup shortcut recreated and app launched.'
