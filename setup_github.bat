@echo off
echo Setting up GitHub repository connection...
echo.

REM Check if git is installed
where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Git is not installed!
    echo.
    echo Please install Git first:
    echo 1. Go to: https://git-scm.com/download/win
    echo 2. Download and install Git for Windows
    echo 3. Run this script again
    echo.
    pause
    exit /b 1
)

echo Git found! Initializing repository...
echo.

REM Initialize git repository
git init
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to initialize git repository
    pause
    exit /b 1
)

echo Adding files...
git add .
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to add files
    pause
    exit /b 1
)

echo Creating initial commit...
git commit -m "Initial commit: Embroidery Computer project"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to create commit
    pause
    exit /b 1
)

echo Connecting to GitHub repository...
git remote add origin https://github.com/emoney2/EmbroideryComputer.git
if %ERRORLEVEL% NEQ 0 (
    echo Warning: Remote may already exist, trying to set URL...
    git remote set-url origin https://github.com/emoney2/EmbroideryComputer.git
)

echo Setting branch to main...
git branch -M main

echo.
echo ========================================
echo Repository setup complete!
echo.
echo Next step: Push to GitHub
echo You will need to authenticate with GitHub
echo.
echo Run this command to push:
echo   git push -u origin main
echo.
echo Or run: push_to_github.bat
echo ========================================
echo.
pause
