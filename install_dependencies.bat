@echo off
echo =================================================
echo  Installing Python dependencies for AutoCut Tool
echo =================================================
echo.

REM Check for python and pip
echo Checking for Python and pip...
python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: python -m pip is not available.
    echo Please ensure Python is installed and that its 'Scripts' directory is in your system's PATH.
    pause
    exit /b 1
)
echo Found pip.
echo.

echo Installing required Python packages...
echo.

python -m pip install customtkinter
if %errorlevel% neq 0 (
    echo ERROR: Failed to install customtkinter.
    pause
    exit /b 1
)

echo.
echo =================================================
echo  Python packages installed successfully!
echo =================================================
echo.
echo.
echo =================================================
echo  IMPORTANT: MANUAL STEP REQUIRED
echo =================================================
echo This script requires FFmpeg (for ffprobe) to function correctly.
echo.
echo Please download and install FFmpeg from: https://ffmpeg.org/download.html
echo.
echo After installation, you MUST add the 'bin' directory of FFmpeg
echo to your system's PATH environment variable.
echo.
echo =================================================
echo.
pause