#------ em construção ------

$Host.UI.RawUI.WindowTitle = "Microsoft Outlook Fix" 
$ErrorActionPreference = "SilentlyContinue"

Set-ItemProperty -path "HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Internet" -Name "UseOnlineContent" -Value "2"
Set-ItemProperty -path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Internet" -Name "UseOnlineContent" -Value "2"

Set-ItemProperty -path "HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\SignIn" -Name "SignInOptions" -Value "0"
Set-ItemProperty -path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\SignIn" -Name "SignInOptions" -Value "0"







<# 
    6. Navegue até o caminho: "HKEY_CURRENT_USER\Software\Microsoft\Office\xx.0\Common\Internet"; 
	    Obs.: O espaço reservado XX é 15 para o Office 2013 e 16 para o Office 2016.
    7. Localize e clique duas vezes no seguinte valor: "UseOnlineContent"; 
    7.1. Caso não exista, crie o valor “UseOnlineContent” com o tipo “Valor de Cadeia de Caracteres” no caminho informado acima; 
    7.2. Na caixa "Dados do valor", digite "2" e clique em "OK"; 
    8. Localize e clique na sub chave: "HKEY_CURRENT_USER\Software\Microsoft\Office\xx.0\Common\SignIn"; 
    	Obs.: O espaço reservado XX é 15 para o Office 2013 e 16 para o Office 2016.
    8.1. Localize e clique duas vezes no valor "SignInOptions"; 
    8.2. Na caixa "Dados do valor", digite "0" e clique em "OK";
#>