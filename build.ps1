$venvActivatePath = ".\venv\Scripts\activate.ps1"
& $venvActivatePath

Write-Host "Removing old builds..."
foreach ($path in @(".\build", ".\dist")) {
    if (Test-Path $path) {
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue -Path $path
    }
}

Start-Sleep -Seconds 1
Write-Host "Building Caffeine Installer.exe..."
Start-Sleep -Seconds 1

try {
    pyinstaller caffeine_installer.spec

    Start-Sleep -Seconds 1
    Write-Host -ForegroundColor Green "Build process completed successfully."
} catch {
    Write-Host -ForegroundColor Red "ERROR: PyInstaller failed to build Caffeine Installer."
    Write-Host "`nException: $_"
}