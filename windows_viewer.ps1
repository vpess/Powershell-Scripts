#Make Windows Photo Viewer your default image viewer
#Torna o Windows Photo Viewer o seu visualizador padrão de imagens

$Host.UI.RawUI.WindowTitle = "Windows Photo Viewer"
$ErrorActionPreference = "SilentlyContinue"

If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) { #credits to https://github.com/Sycnex/Windows10Debloater/blob/master/Windows10Debloater.ps1
    Write-Host "Você não executou o script com privilégios de administrador. O script será executado novamente solicitando os privilégios necessários."
    Start-Sleep 1
    Write-Host "          3"
    Start-Sleep 1
    Write-Host "          2"
    Start-Sleep 1
    Write-Host "          1"
    Start-Sleep 1
    Start-Process powershell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
    Exit
}

Set-ItemProperty -path "HKLM:\SOFTWARE\Classes\.bmp" -Name "(default)" -Value "PhotoViewer.FileAssoc.Tiff"
Set-ItemProperty -path "HKLM:\SOFTWARE\Classes\.jpg" -Name "(default)" -Value "PhotoViewer.FileAssoc.Tiff"
Set-ItemProperty -path "HKLM:\SOFTWARE\Classes\.jpeg" -Name "(default)" -Value "PhotoViewer.FileAssoc.Tiff"
Set-ItemProperty -path "HKLM:\SOFTWARE\Classes\.ico" -Name "(default)" -Value "PhotoViewer.FileAssoc.Tiff"
Set-ItemProperty -path "HKLM:\SOFTWARE\Classes\.gif" -Name "(default)" -Value "PhotoViewer.FileAssoc.Tiff"
Set-ItemProperty -path "HKLM:\SOFTWARE\Classes\.png" -Name "(default)" -Value "PhotoViewer.FileAssoc.Tiff"
Set-ItemProperty -path "HKLM:\SOFTWARE\Classes\.jfif" -Name "(default)" -Value "PhotoViewer.FileAssoc.Tiff"

Write-Host "GG"
Start-Sleep -s 3
