@echo off

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed. Installing Python...
    :: Download Python installer (adjust URL to the latest version as needed)
    curl -o python_installer.exe https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe
    start /wait python_installer.exe /quiet InstallAllUsers=1 PrependPath=1
    echo Python installation complete.
) else (
    echo Python is already installed.
)

:: Ensure pip is up to date
python -m ensurepip --upgrade
python -m pip install --upgrade pip

:: Install required libraries
pip install keyboard pycaw comtypes pywin32

:: Navigate to the script directory
cd C:\Users\ceepies\Documents

:: Run the Python script without a console window
start "" pythonw audiocontrol.py

:: Clean up the installer if it was downloaded
if exist python_installer.exe del python_installer.exe

echo Script executed successfully.
