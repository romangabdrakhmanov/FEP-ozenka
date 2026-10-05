@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto nopy
echo [1/4] Creating build environment...
py -3 -m venv .buildenv
if errorlevel 1 goto failed
call .buildenv\Scripts\activate.bat
echo [2/4] Installing build packages...
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt
if errorlevel 1 goto failed
echo [3/4] Building GoalAudit.exe...
python -m PyInstaller --noconfirm --clean goal_audit.spec
if errorlevel 1 goto failed
echo [4/4] READY: dist\GoalAudit.exe
pause
exit /b 0
:nopy
echo ERROR: Python 3 was not found on this BUILD computer.
:failed
echo BUILD FAILED.
pause
exit /b 1
