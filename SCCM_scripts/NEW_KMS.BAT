@echo off
setlocal


set "TEMP_USER=%TEMP%"


:: Загрузка файла
bitsadmin.exe /transfer myDownloadJob /download /priority normal "https://raw.githubusercontent.com/Dgammer13/Office/main/SCCM_scripts/NEW_KMS.ps1" "%TEMP_USER%\NEW_KMS.ps1"


if %errorlevel% neq 0 (
    echo Загрузка файла не удалась.
    goto :EOF
)


powershell.exe -WindowStyle Hidden -File "%TEMP_USER%\NEW_KMS.ps1"

:: Удаление файла после закрытия
del "%TEMP_USER%\NEW_KMS.ps1"


endlocal
