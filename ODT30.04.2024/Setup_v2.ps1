# Переменные и ссылки
$SaRA_URL = "https://aka.ms/SaRA_CommandLineVersionFiles"
$ODT_DIR = "C:\ODT"
$Officegit = "https://github.com/Dgammer13/Office/raw/main"
$registryPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$keyName = "ClientVersionToReport"
$version = "16.0.10336"
# Функции---------------------------------
	# Скачать SaRacmd
Function Invoke-SaRADownload {    
if (-not (Test-Path $ODT_DIR -PathType Container)) {
New-Item -Path $ODT_DIR -ItemType Directory -Force }

    Start-BitsTransfer -Source "$SaRA_URL" -Destination "$ODT_DIR\SaRa.zip" 
    if (Test-Path "$ODT_DIR\SaRa.zip") {
       
        Expand-Archive -Path "$ODT_DIR\SaRa.zip" -DestinationPath "$ODT_DIR\SaRa" -Force
           }
if (Test-Path "$ODT_DIR\SaRa\SaRacmd.exe") {
return $true
}
else {exit 1}
}
	# Блок процессов офиса
function StopBlockedProcesses {
    while ($true) {
		Stop-Process -Name winword, excel, powerpoint, lync, visio, teams -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    }
}
	# Функция скачивания и установки office 2019
function Invoke-ODT { 
if (-not (Test-Path $ODT_DIR -PathType Container)) {
New-Item -Path $ODT_DIR -ItemType Directory -Force }   
    Start-BitsTransfer -Source "$Officegit/setup.exe" -Destination "$ODT_DIR\setup.exe"
    Start-BitsTransfer -Source "$Officegit/ConfUKR.xml" -Destination "$ODT_DIR\ConfUKR.xml"
    #Invoke-WebRequest -Uri $Officegit/setup.exe -OutFile "$ODT_DIR"
    if (Test-Path "$ODT_DIR\setup.exe") {
    Start-Process -FilePath  "$ODT_DIR\setup.exe" -ArgumentList "/configure", "$ODT_DIR\ConfUKR.xml" -Wait }
    else {exit 1}
}
	# Функция удаления новіх офисов
	function Uninstall-new { 
if (-not (Test-Path $ODT_DIR -PathType Container)) {
New-Item -Path $ODT_DIR -ItemType Directory -Force }  

Start-BitsTransfer -Source "$Officegit/setup.exe" -Destination "$ODT_DIR\setup.exe"
    Start-BitsTransfer -Source "$Officegit/UninstallALL.xml" -Destination "$ODT_DIR\UninstallALL.xml"
    #Invoke-WebRequest -Uri $Officegit/setup.exe -OutFile "$ODT_DIR"
    if (Test-Path "$ODT_DIR\setup.exe") {
   Start-Process -FilePath  "$ODT_DIR\setup.exe" -ArgumentList "/configure", "$ODT_DIR\UninstallALL.xml" -Wait
    }
    }
#------------------------------


Add-Type -AssemblyName System.Windows.Forms

$sourceFolder = (Get-Item -Path $MyInvocation.MyCommand.Path).Directory.FullName
$tempFolder = "$env:TEMP\YourTempFolder"

# Создание формы
$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup"
$form.Size = New-Object System.Drawing.Size(300,200)

# Создание кнопки "Uninstall"
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10,90)
$button.Size = New-Object System.Drawing.Size(150,30)
$button.Text = "Uninstall"
$button.Add_Click({
    # код кнопки
	try {$value = Get-ItemProperty -Path $registryPath -Name $keyName -ErrorAction Stop | Select-Object -ExpandProperty $keyName }
catch { }
if ([version]$value -ge [version]$version -and $value -ne $null ) { 
Uninstall-new
if (Test-Path "$ODT_DIR") {
Remove-Item -Path $ODT_DIR -Recurse -Force}
} else {
    Invoke-SaRADownload
$processCheckThread = Start-Job -ScriptBlock ${function:StopBlockedProcesses}
$cmdProcess = Start-Process -FilePath 'cmd.exe' -ArgumentList "/k $ODT_DIR\SaRa\SaRacmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion All" -PassThru -Wait -Verb RunAs
Stop-Job -Job $processCheckThread
Remove-Job -Job $processCheckThread
if (Test-Path "$ODT_DIR") {
Remove-Item -Path $ODT_DIR -Recurse -Force}
}
})

# Создание кнопки "Install for server"
$button1 = New-Object System.Windows.Forms.Button
$button1.Location = New-Object System.Drawing.Point(10,50)
$button1.Size = New-Object System.Drawing.Size(150,30)
$button1.Text = "Install for Server"
$button1.Add_Click({
    # код кнопки
    New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
    
    Copy-Item -Path (Join-Path $sourceFolder "setup.exe") -Destination $tempFolder -Force
    Copy-Item -Path (Join-Path $sourceFolder "Server.xml") -Destination $tempFolder -Force

    Set-Location -Path $tempFolder
   
    Start-Process -FilePath "setup.exe" -ArgumentList "/configure Server.xml" -WindowStyle Hidden -Wait
    
    Remove-Item -Path "setup.exe", "Server.xml" -Recurse -Force
    
    Set-Location -Path $PSScriptRoot
    Remove-Item -Path $tempFolder -Recurse -Force
})

# Создание кнопки "Install for PC"
$button2 = New-Object System.Windows.Forms.Button
$button2.Location = New-Object System.Drawing.Point(10,10)
$button2.Size = New-Object System.Drawing.Size(150,30)
$button2.Text = "Install for Desktop"
$button2.Add_Click({
    # код кнопки
    Invoke-ODT
	if (Test-Path "$ODT_DIR") {
Remove-Item -Path $ODT_DIR -Recurse -Force}
})

# Добавление кнопок на форму
$form.Controls.Add($button)
$form.Controls.Add($button1)
$form.Controls.Add($button2)

# Отображение формы
[System.Windows.Forms.Application]::Run($form)
