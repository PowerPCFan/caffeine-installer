@echo off

CALL %0\..\venv\Scripts\activate.bat
start powershell.exe -NoProfile -NoExit -command "pyinstaller caffeine_installer.spec"