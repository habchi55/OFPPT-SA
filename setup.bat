@echo off

:: Step 1: Download and install Python (skip if Python is already installed)
echo Downloading Python installer...
curl -o python-installer.exe https://www.python.org/ftp/python/3.10.2/python-3.10.2-amd64.exe

echo Installing Python...
start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1

:: Step 2: Verify Python installation
python --version
if %errorlevel% neq 0 (
    echo Python installation failed. Exiting.
    exit /b %errorlevel%
)

:: Step 3: Create a virtual environment and activate it
echo Creating virtual environment...
python -m venv venv
call venv\Scripts\activate

:: Step 4: Install required libraries
echo Installing required Python libraries...
pip install pyinstaller
pip install -r requirements.txt

:: Step 5: Initialize the database
echo Initializing the database...
python -c "from your_script import create_database; create_database()"

:: Step 6: Build the executable using PyInstaller
echo Building the executable...
pyinstaller --onefile --noconsole --name "OFPPT-SA" main.py

:: Step 7: Create a shortcut to run the executable
echo Creating shortcut...

:: Customize the name of the shortcut here
set SHORTCUT_NAME=OFPPT-SA

:: Define paths
set SHORTCUT_PATH="%USERPROFILE%\Desktop\%SHORTCUT_NAME%.lnk"
set TARGET_PATH="%CD%\dist\OFPPT-SA.exe"

:: Create the shortcut using PowerShell
powershell -command "$s=(New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath='%TARGET_PATH%'; $s.Save()"

:: Step 8: Run the application
echo Running the application...
"%TARGET_PATH%"

echo Setup completed successfully.
pause
