#------ em construção ------

$Host.UI.RawUI.WindowTitle = "Hyperlink Fix"
$ErrorActionPreference = "SilentlyContinue"

Set-ItemProperty -path "HKCU:\SOFTWARE\Classes\.htm"
Set-ItemProperty -path "HKCU:\SOFTWARE\Classes\.html"
Set-ItemProperty -path "HKCU:\SOFTWARE\Classes\.shtml"
Set-ItemProperty -path "HKCU:\SOFTWARE\Classes\.xht"
Set-ItemProperty -path "HKCU:\SOFTWARE\Classes\.xhtml"

Write-Host "GG"
Start-Sleep -s 1

<#

REG ADD HKEY_CURRENT_USER\Software\Classes\.htm /ve /d htmlfile /F
REG ADD HKEY_CURRENT_USER\Software\Classes\.html /ve /d htmlfile /F
REG ADD HKEY_CURRENT_USER\Software\Classes\.shtml /ve /d htmlfile /F
REG ADD HKEY_CURRENT_USER\Software\Classes\.xht /ve /d htmlfile /F
REG ADD HKEY_CURRENT_USER\Software\Classes\.xhtml /ve /d htmlfile /F

/ve -> specifies that only entries that have no value will be deleted.
/d -> disables execution of AutoRun commands
/F -> force

https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/cmd
https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/reg-delete

#>
