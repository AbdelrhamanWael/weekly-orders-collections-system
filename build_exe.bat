@echo off
chcp 65001 >nul
REM الانتقال لمجلد البرنامج (حيث يوجد هذا الملف) حتى يعمل البناء من أي مكان
cd /d "%~dp0"

echo ========================================
echo   بناء ملف .exe — نظام الطلبات والتحصيل
echo   Build .exe — Orders ^& Collections System
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python غير مثبت. يرجى تثبيت Python أولاً.
    pause
    exit /b 1
)

REM Install PyInstaller if needed
echo [1/3] التحقق من PyInstaller...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo تثبيت PyInstaller...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo [ERROR] فشل تثبيت PyInstaller.
        pause
        exit /b 1
    )
)
echo [OK] PyInstaller جاهز
echo.

REM Build
echo [2/3] بناء الملف التنفيذي (قد يستغرق دقائق)...
pyinstaller --noconfirm --clean build_exe.spec
if %errorlevel% neq 0 (
    echo [ERROR] فشل البناء.
    pause
    exit /b 1
)
echo.

echo [3/3] انتهى البناء.
echo.
echo الملف الناتج:
echo   dist\OrdersCollectionsSystem.exe
echo.
echo انسخ الملف "OrdersCollectionsSystem.exe" للعميل.
echo المجلدات samples و reports و قاعدة البيانات ستُنشأ بجانب الـ .exe عند التشغيل.
echo ========================================
pause
