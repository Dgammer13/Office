#$regPath = "HKLM:\SOFTWARE\ESET\ESET Security\CurrentVersion\Config\era\plugins\01000400\profile\profile\a1\settings\UPDATE_CFG"
#$regName = "UpdateUrl"

#$value = [byte[]]@(
#    0x68, 0x74, 0x74, 0x70, 0x3a, 0x2f, 0x2f, 0x31,
#    0x30, 0x2e, 0x31, 0x36, 0x30, 0x2e, 0x31, 0x2e,
#    0x34, 0x3A, 0x32, 0x32, 0x32, 0x31
#)

#Set-ItemProperty -Path $regPath -Name $regName -Value $value -Type Binary


$regPath1 = "HKLM:\SOFTWARE\ESET\ESET Security\CurrentVersion\Config\era\plugins\01000400\profile\profile\a1\settings\UPDATE_CFG"
$regPath2 = "HKLM:\SOFTWARE\ESET\ESET Security\CurrentVersion\Config\Plugins\01000400\profile\profile\a1\settings\UPDATE_CFG"
$bytes = [System.Text.Encoding]::ASCII.GetBytes("http://10.160.1.43:2221")

if (Test-Path -Path $regPath1) {
    Set-ItemProperty -Path $regPath1 -Name "UpdateUrl" -Value $bytes -Type Binary
}

if (Test-Path -Path $regPath2) {
    Set-ItemProperty -Path $regPath2 -Name "UpdateUrl" -Value $bytes -Type Binary
Remove-ItemProperty -Path $regPath2 -Name "UpdateFromMirrorPassword" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $regPath2 -Name "UpdateFromMirrorUsername" -ErrorAction SilentlyContinue
$folderPath = "C:\Program Files\ESET\ESETMR"
    New-Item -ItemType Directory -Path $folderPath
}
