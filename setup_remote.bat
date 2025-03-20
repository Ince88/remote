@echo off
title Remote Control Setup

:: Check for admin privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Please run this script as Administrator
    echo Right-click the script and select 'Run as Administrator'
    pause
    exit /b 1
)

:: Set colors and clear screen
color 0B
cls

:: Create temp directory
set "tempDir=%TEMP%\remote_setup"
if not exist "%tempDir%" mkdir "%tempDir%"

echo Starting Remote Control Setup...
echo This script will:
echo 1. Download and install Python
echo 2. Install required packages
echo 3. Download and setup the Remote Control server
echo.

:: Download Python
echo Downloading Python...
powershell -Command "& {Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe' -OutFile '%tempDir%\python_installer.exe'}"

:: Install Python
echo Installing Python...
start /wait "" "%tempDir%\python_installer.exe" /quiet InstallAllUsers=1 PrependPath=1

:: Wait for installation
timeout /t 10 /nobreak > nul

:: Check both possible Python installation paths
if exist "%USERPROFILE%\AppData\Local\Programs\Python\Python311\python.exe" (
    set "PYTHON_PATH=%USERPROFILE%\AppData\Local\Programs\Python\Python311"
) else if exist "C:\Program Files\Python311\python.exe" (
    set "PYTHON_PATH=C:\Program Files\Python311"
) else if exist "C:\Python311\python.exe" (
    set "PYTHON_PATH=C:\Python311"
) else (
    echo Python installation not found! Please try running the script again.
    pause
    exit /b 1
)

echo Python found at: %PYTHON_PATH%
set "PATH=%PYTHON_PATH%;%PYTHON_PATH%\Scripts;%PATH%"

:: Install required packages
echo Installing required packages...
"%PYTHON_PATH%\python.exe" -m pip install --upgrade pip
"%PYTHON_PATH%\python.exe" -m pip install keyboard
"%PYTHON_PATH%\python.exe" -m pip install pyautogui
"%PYTHON_PATH%\python.exe" -m pip install pywin32

:: Create program directory
if not exist "%USERPROFILE%\Remote Control Server" mkdir "%USERPROFILE%\Remote Control Server"

:: Download remote_server.py and icon
echo Downloading Remote Control server...
powershell -Command "& {Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Ince88/remote/main/remote_server.py' -OutFile '%USERPROFILE%\Remote Control Server\remote_server.py'}"
powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Ince88/remote/main/icon.ico' -OutFile '%USERPROFILE%\Remote Control Server\icon.ico'}"

:: Create shortcut with custom icon
echo Creating desktop shortcut...
(
    echo Set WshShell = CreateObject^("WScript.Shell"^)
    echo strDesktop = WshShell.SpecialFolders^("Desktop"^)
    echo Set Shortcut = WshShell.CreateShortcut^(strDesktop ^& "\Remote Control Server.lnk"^)
    echo Shortcut.TargetPath = "%PYTHON_PATH%\python.exe"
    echo Shortcut.Arguments = """%USERPROFILE%\Remote Control Server\remote_server.py"""
    echo Shortcut.WorkingDirectory = "%USERPROFILE%\Remote Control Server"
    echo Shortcut.Description = "Remote Control Server"
    echo Shortcut.IconLocation = """%USERPROFILE%\Remote Control Server\icon.ico,0"""
    echo Shortcut.Save
) > "%tempDir%\createShortcut.vbs"

cscript //nologo "%tempDir%\createShortcut.vbs"

:: Cleanup
rmdir /S /Q "%tempDir%"

echo.
echo Setup completed successfully!
echo A shortcut has been created on your desktop: 'Remote Control Server'
echo.
echo When you first start the server:
echo 1. Double-click the 'Remote Control Server' shortcut
echo 2. Note the IP address shown in the window
echo 3. Use that IP address in your phone app
echo.
pause 