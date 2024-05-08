@echo off

set "ps1Url=https://raw.githubusercontent.com/Dgammer13/Office/main/SCCM_scripts/NEW_KMS.ps1"
set "tempDir=%TEMP%"

powershell -Command "(New-Object Net.WebClient).DownloadFile('%ps1Url%', '%tempDir%\NEW_KMS.ps1')"
if %errorlevel% neq 0 (
    echo Загрузка файла не удалась.
    goto :EOF
)
powershell.exe -WindowStyle Hidden -File "%tempDir%\NEW_KMS.ps1"
del "%tempDir%\NEW_KMS.ps1"

