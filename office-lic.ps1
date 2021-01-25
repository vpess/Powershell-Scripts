$Host.UI.RawUI.WindowTitle = "Ativador do Office"
$ErrorActionPreference = "SilentlyContinue"
<#$logfile = "c:\windows\logs\Office365_license.log"#>

function killall {
    Write-Host "Finalizando os processos do Office...`n"
    Stop-Process -force -ProcessName WINWORD
    Stop-Process -force -ProcessName MSACCESS
    Stop-Process -force -ProcessName ONENOTEM
    Stop-Process -force -ProcessName ONENOTE
    Stop-Process -force -ProcessName OUTLOOK
    Stop-Process -force -ProcessName EXCEL
    Stop-Process -force -ProcessName POWERPNT
    Stop-Process -force -ProcessName VISIO
    Stop-Process -force -ProcessName WINPROJ
    Stop-Process -force -ProcessName OfficeClickToRun
    log "Processos relacionados ao Pacote Office finalizados."
}


function arquitetura {
$arq = $(Write-host "`n-> Informe a arquitetura do Office (32/64): " -ForegroundColor DarkGreen -BackgroundColor Black -NoNewline; Read-Host)
    if ($arq -eq 32){
    Set-Location "${Env:ProgramFiles(x86)}\Microsoft Office\"
}

    elseif ($arq -eq 64){
    Set-Location "$Env:ProgramFiles\Microsoft Office\"
}
    else {Write-Host "`nSelecione uma opção válida.`n" -ForegroundColor Red
    return arquitetura}

    }

function versao {
$ver = $(Write-host "`n-> Informe a versão do Office (10/13/16): " -ForegroundColor DarkGreen -BackgroundColor Black -NoNewline; Read-Host)
    if ($ver -eq 10){
    Write-host "`nProcessando...`n" -ForegroundColor Blue -BackgroundColor Black ; Start-Sleep -s 1
    Set-Location Office14
    Cscript.exe ospp.vbs /act
    Write-Host "`nComando executado." -ForegroundColor Blue -BackgroundColor Black ; Start-Sleep -s 1
    Write-Output "`nSaindo..."
    Start-Sleep -s 3
    return

    }
    
    elseif ($ver -eq 13){
    Write-host "`nProcessando...`n" -ForegroundColor Blue -BackgroundColor Black ; Start-Sleep -s 1
    Set-Location Office15
    Cscript.exe ospp.vbs /act
    Write-Host "`nComando executado." -ForegroundColor Blue -BackgroundColor Black ; Start-Sleep -s 1
    Write-Output "`nSaindo..."
    Start-Sleep -s 3
    return
    }

    elseif ($ver -eq 16){
    Write-host "`nProcessando...`n" -ForegroundColor Blue -BackgroundColor Black ; Start-Sleep -s 1
    Set-Location Office16
    Cscript.exe ospp.vbs /act
    Write-Host "`nComando executado." -ForegroundColor Blue -BackgroundColor Black ; Start-Sleep -s 1
    Write-Output "`nSaindo..."
    Start-Sleep -s 3
    return
    }

    else {Write-Host "Selecione uma opção válida." -ForegroundColor Red
    return versao}

    }

function activation {
arquitetura
versao
}

killall
activation