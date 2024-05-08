 $PowerShellExe = Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'
# Путь к несуществующему скрипту (убить форму ПС)
$ScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "1.ps1"
# Запуск скрипта с параметром -WindowStyle Hidden
Start-Process -FilePath $PowerShellExe -ArgumentList "-WindowStyle Hidden", "-File", $ScriptPath -NoNewWindow


 Add-Type -Assemblyname System.Windows.Forms
 #форма
 $form = New-Object System.Windows.Forms.Form
 $form.Text = 'TGZ-KMS'
#форма метка 1
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Выбери активацию'
$label.Location = New-Object System.Drawing.Point(20, 10)
$label.AutoSize = $true
$form.Controls.Add($label)
#кнопка активации винды
$button = New-Object Windows.Forms.Button
$button.Text = "Windows"
$button.Location = New-Object System.Drawing.Point 20,40
$form.Controls.Add($button)
#событие нажания кнопки 1
 $button.Add_Click({
 try {
 c:\windows\system32\slmgr.vbs /skms 10.160.1.36
c:\windows\system32\slmgr.vbs  /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
c:\windows\system32\slmgr.vbs  /ato
  $label.Text='Команда отправлена'
  } catch {
        $label.Text = 'Ошибка активации: ' + $_.Exception.Message
    }
   })
   #кнопка 2
  $button2 = New-Object Windows.Forms.Button
$button2.Text = "Office"
 $button2.Location = New-Object System.Drawing.Point 120,40
 $form.Controls.Add($button2)
 #событие кнопки 2 
 $button2.Add_Click({

 # переменные путей Office
$OfficePath32 = "C:\Program Files (x86)\Microsoft Office\Office16"
$OfficePath64 = "C:\Program Files\Microsoft Office\Office16"

# Проверяем, установлена ли 64-битная версия Office
if (Test-Path $OfficePath64) {
    $OfficePath = $OfficePath64
}
# Если 64-битная версия Office не найдена, проверяем 32-битную версию
elseif (Test-Path $OfficePath32) {
    $OfficePath = $OfficePath32
}
else {
   $label.Text='Офис не найден'
}

# Переходим в каталог Office
Set-Location -Path $OfficePath

# Выполняем активацию
$licenseFiles = Get-ChildItem -Path "..\root\Licenses16" -Filter "ProPlus2019VL*.xrm-ms"
foreach ($licenseFile in $licenseFiles) {
    & "cscript" "ospp.vbs" "/inslic:..\root\Licenses16\$($licenseFile.Name)"
}
Start-Process -FilePath "cscript" -ArgumentList "ospp.vbs /setprt:1688"
Start-Process -FilePath "cscript" -ArgumentList "ospp.vbs /unpkey:6MWKP >nul"
Start-Process -FilePath "cscript" -ArgumentList "ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
Start-Process -FilePath "cscript" -ArgumentList "ospp.vbs /sethst:10.160.1.36" ### упрощение cscript ospp.vbs /sethst:10.160.1.36 убрав выше переменную КМС
Start-Process -FilePath "cscript" -ArgumentList "ospp.vbs /act" ### упрощение выполнить cscript ospp.vbs /act
$label.Text='Office активирован'
  })
  #форма метка 2
$label2 = New-Object System.Windows.Forms.Label
$label2.Text = 'Альтернативный сервер'
$label2.Location = New-Object System.Drawing.Point(20, 80)
$label2.AutoSize = $true
$form.Controls.Add($label2)
#кнопка 3 активации винды 2
$button3 = New-Object Windows.Forms.Button
$button3.Text = "Windows"
$button3.Location = New-Object System.Drawing.Point 20,110
$form.Controls.Add($button3)
#кнопка 4 активации офиса 2
$button4 = New-Object Windows.Forms.Button
$button4.Text = "Office"
$button4.Location = New-Object System.Drawing.Point 120,110
$form.Controls.Add($button4)
#событие нажания кнопки 3
 $button3.Add_Click({
 try {
 c:\windows\system32\slmgr.vbs /skms e8.us.to
c:\windows\system32\slmgr.vbs  /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
c:\windows\system32\slmgr.vbs  /ato
  $label2.Text='Команда отправлена'
  } catch {
        $label2.Text = 'Ошибка активации: ' + $_.Exception.Message
    }
   })
#событие кнопки 4 
 $button4.Add_Click({

 # переменные путей Office
$OfficePath32 = "C:\Program Files (x86)\Microsoft Office\Office16"
$OfficePath64 = "C:\Program Files\Microsoft Office\Office16"

# Проверяем, установлена ли 64-битная версия Office
if (Test-Path $OfficePath64) {
    $OfficePath = $OfficePath64
}
# Если 64-битная версия Office не найдена, проверяем 32-битную версию
elseif (Test-Path $OfficePath32) {
    $OfficePath = $OfficePath32
}
else {
   $label2.Text='Офис не найден'
}

# Переходим в каталог Office
Set-Location -Path $OfficePath

# Выполняем активацию
Start-Process -FilePath "cscript" -ArgumentList "ospp.vbs /sethst:e8.us.to" ### упрощение cscript ospp.vbs /sethst:10.160.1.36 убрав выше переменную КМС
Start-Process -FilePath "cscript" -ArgumentList "ospp.vbs /act" ### упрощение выполнить cscript ospp.vbs /act
$label2.Text='Office активирован'
  })
$form.Add_FormClosing({
    Write-Host 'exiting'
})
$form.ShowDialog()