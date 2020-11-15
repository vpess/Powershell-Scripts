$Host.UI.RawUI.WindowTitle = "Microsoft Teams Fix"
$ErrorActionPreference = "SilentlyContinue"

<#function log {
    $directory = "C:\Windows\Logs\Teams-fix.log" 
    $date = Get-Date -format "dd/MM/yyyy, HH:mm:ss:"
    $msg = "$date $text"

    if ($directory) { $msg | Add-Content $directory }

    else { $msg | Out-File $directory }
}#>

function cleanCache {
    Write-Host "`nFinalizando processos do Teams..."
    Stop-Process -ProcessName teams -Force -ErrorAction SilentlyContinue
    Write-Host "`nEfetuando remoção de arquivos de cache..."
    Remove-Item -Recurse -Force "$ENV:Userprofile\appdata\roaming\Microsoft\Teams\Cache\*" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "$ENV:Userprofile\appdata\roaming\Microsoft\Teams\Application Cache\Cache\*" -ErrorAction SilentlyContinue #>>> pode ser inexistente#>
    Write-Host "`nArquivos de cache do Teams removidos."
    return selec
}

function cleanRoaming {
    Write-Host "`nO Outlook precisa ser finalizado para que essaa correção seja feita. Caso ele esteja aberto, salve seu trabalho e prossiga." -ForegroundColor Red -BackgroundColor Black
    Write-Host -NoNewLine "`nPressione qualquer tecla para continuar...`n";
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    Write-Host "`nFinalizando processos do Teams e Outlook..."
    Stop-Process -ProcessName teams -Force -ErrorAction SilentlyContinue
    Stop-Process -ProcessName outlook -Force -ErrorAction SilentlyContinue
    Write-Host "`nEfetuando remoção de arquivos da %appdata%"
    cleanCache
    Remove-Item -Recurse -Force "$ENV:Userprofile\appdata\roaming\Microsoft\Teams\*" -ErrorAction SilentlyContinue
    Write-Host "`nArquivos do Teams em %appdata% removidos."   
    return selec 
}

function reinstall {
    Write-Host "`nO Outlook precisa ser finalizado para que essa correção seja feita. Caso ele esteja aberto, salve seu trabalho e prossiga." -ForegroundColor Red -BackgroundColor Black
    Write-Host -NoNewLine "`nPressione qualquer tecla para continuar...`n";
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    Write-Host "`nFinalizando processos do Teams e Outlook..."
    Stop-Process -ProcessName teams -Force -ErrorAction SilentlyContinue
    Stop-Process -ProcessName outlook -Force -ErrorAction SilentlyContinue
    cleanCache
    cleanRoaming
    Write-Host "`nEfetuando remoção de demais arquivos do Teams..."
    Remove-Item -Recurse -Force "$Env:userprofile\appdata\local\Microsoft\Teams\*"
    Remove-Item -Recurse -Force "$Env:userprofile\appdata\roaming\Microsoft\Teams\*"
    Remove-Item -Recurse -Force "$Env:userprofile\OneDrive - Petrobras\Desktop\Microsoft Teams.lnk"
    Remove-Item -Recurse -Force "$Env:userprofile\desktop\Microsoft Teams.lnk"
    Remove-Item -Recurse -Force "C:\ProgramData\Microsoft\Microsoft\*"
    Write-Host "`nTeams removido."; Start-Sleep -s 2
    Write-Host "`nEfetuando exclusão de entradas de registro..."
    Remove-ItemProperty -Force "HKCU:\Software\IM Providers\Teams\" -Name *
    Remove-ItemProperty -Force "HKCU:\Software\Microsoft\Office\Teams\" -Name *
    Remove-ItemProperty -Force "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Teams\" -Name *
    Remove-ItemProperty -Force "HKLM:\Software\IM Providers\Teams\" -Name *
    Remove-ItemProperty -Force "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run\" -Name "com.squirrel.Teams.Teams"
    Remove-ItemProperty -Force "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\VREGISTRY_B683C874-A67C-41B4-8750-72BE2153F84C\MACHINE\Software\Wow6432Node\IM Providers\Teams\" -Name * <#>>> pode ser inexistente#>
    Remove-ItemProperty -Force "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\IM Providers\Teams\" -Name * <#>>> pode ser inexistente#>
    Write-Host "`nEntradas de registro do Teams removidas."
    #>>> Download e instalação do Teams
    $url = "https://go.microsoft.com/fwlink/p/?LinkID=869426&clcid=0x416&culture=pt-br&country=BR&lm=deeplink&lmsrc=groupChatMarketingPageWeb&cmpid=directDownloadWin64"
    $output = "$Env:userprofile\Downloads\Teams_x64.exe"
    Write-Host "`nRealizando download do Teams. Aguarde..."
    Invoke-WebRequest -Uri $url -OutFile $output
    Write-Host "`nIniciando a instalação. Aguarde..."
    Invoke-Expression "$ENV:Userprofile\Downloads\Teams_x64.exe"
    Start-Sleep -s 5
    return selec
}

function vdi { #https://autoatendimentotic.petrobras.com.br/visualizar/5100055/16513
    Write-Host "`nO Outlook precisa ser finalizado para que a correção seja feita. Caso ele esteja aberto, salve seu trabalho e prossiga." -ForegroundColor Red -BackgroundColor Black
    Write-Host -NoNewLine "`nPressione qualquer tecla para continuar...`n";
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    Stop-Process -ProcessName teams -Force -ErrorAction SilentlyContinue
    Stop-Process -ProcessName outlook -Force -ErrorAction SilentlyContinue
    Rename-Item "$Env:userprofile\appdata\local\Microsoft\Teams" "$Env:userprofile\appdata\local\Microsoft\Teams_old"
    Remove-Item -Recurse -Force "$ENV:Userprofile\appdata\roaming\Microsoft\Teams\*" -ErrorAction SilentlyContinue
    Set-Location "C:\Program Files (x86)\Teams Installer\" ; .\Teams.exe --checkinstall
}

function selec{

    param (
    [string]$Titulo = 'Menu'
    )
    
    Write-Host "`n============================ Menu de Reparo ============================`n"
    
    Write-Host "	[1] para remover arquivos de cache do Teams"
    Write-Host "	[2] para remover arquivos da Roaming (%appdata%)"
    Write-Host "	[3] para remover todos os resquícios do Microsoft Teams e reinstalá-lo"
    Write-Host "	[4] para executar o reparo do Teams para VDI's"
    Write-Host "	[q] para fechar o script"
    
    Write-Host "`n============================================================================"
    
     $selection = Read-Host "`nSelecione uma das opções acima"
     switch ($selection)
     {
       '1' {cleanCache} 
       
       '2' {cleanRoaming} 

       '3' {reinstall}

       '4' {vdi}

       'q' {
           Write-Output "`nSaindo..."
           Start-Sleep -s 1
           return}
  
       default {
            if ($selection -ige 4 -or $selection -ne 'q'){
                 Write-Host "`n>>> Selecione apenas opções que estejam no menu!`n" -ForegroundColor Red -BackgroundColor Black
                 Start-Sleep -s 2
                 return selec }
                }   
       }
  }

  selec

#Feito por: BJBD
