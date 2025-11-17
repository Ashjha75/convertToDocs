@echo off
REM Installation script for Markdown to PDF Converter (Windows)
REM This script installs all required dependencies

echo ========================================
echo Markdown to PDF Converter - Installation
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://www.python.org/
    pause
    exit /b 1
)

echo [INFO] Python found:
python --version
echo.

REM Upgrade pip
echo [INFO] Upgrading pip...
python -m pip install --upgrade pip
echo.

REM Install dependencies
echo [INFO] Installing dependencies...
pip install -r requirements.txt
echo.

REM Check installation
echo [INFO] Verifying installation...
python -c "import markdown; import weasyprint; import pygments; print('[SUCCESS] All packages installed successfully!')"

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo Installation completed successfully!
    echo ========================================
    echo.
    echo You can now use the converter:
    echo   cd src
    echo   python main.py input.md output.pdf
    echo.
) else (
    echo.
    echo [ERROR] Installation verification failed
    echo Please check the error messages above
    echo.
)

pause
