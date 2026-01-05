@echo off
echo Pushing to GitHub...
echo.

REM Check if git is installed
where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Git is not installed!
    echo Please install Git first from: https://git-scm.com/download/win
    pause
    exit /b 1
)

echo Pushing code to GitHub...
echo You may be prompted for your GitHub username and password/token
echo.
git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo SUCCESS! Code pushed to GitHub!
    echo Repository: https://github.com/emoney2/EmbroideryComputer
    echo ========================================
) else (
    echo.
    echo ========================================
    echo Push failed. Common issues:
    echo.
    echo 1. Authentication required:
    echo    - You may need a Personal Access Token instead of password
    echo    - Get one at: https://github.com/settings/tokens
    echo.
    echo 2. Repository might not exist or you don't have access
    echo.
    echo 3. Try running: git push -u origin main
    echo    and enter your credentials when prompted
    echo ========================================
)

echo.
pause
