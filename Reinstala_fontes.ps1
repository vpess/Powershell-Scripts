#Script não criado por mim, inserido aqui no git apenas para finalidade de estudo.

#Script para reinstalacao de fontes do tipo "Arquivo de fonte TrueType*"

#Variaveis para o LOG
$global:DataLog = (Get-Date).tostring("dd-MM-yyyy")
$global:HoraLog = (Get-Date).tostring("HH:mm:ss")

#Funções
function Gravar-Log ([String]$LOG_txt) {
    <#Exemplo de execução:
    Gravar-Log -LOG_txt "Atualizacao realizada"
    #>

    $LOG = (("C:\Windows\Logs\") + ($global:SoftwareName)+ ("_") + ($global:SoftwareVersion) + (".log"))
    $global:HoraLog = (Get-Date).tostring("HH:mm:ss")# Atualização da hora para uso no LOG
    Add-Content -Path $LOG -Value "$global:DataLog - $global:HoraLog : $LOG_txt" -Force -Encoding Unicode
}
	
#DEFINA AQUI O NOME DO SOFTWARE
$global:SoftwareName = "Reinstala_Fontes"

#DEFINA AQUI A VERSÃO DO SOFTWARE
$global:SoftwareVersion = "1"

#---------------------------- Main Program------------------------------

#Arquivo de Log criado
Gravar-Log -LOG_txt "--------------------------------------------------------------------"
Gravar-Log -LOG_txt "Início do script"


Write-Host "Equipamento: $env:COMPUTERNAME"
Gravar-Log "Equipamento: $env:COMPUTERNAME"
#$fontFolder = "C:\Windows.old\WINDOWS\Fonts"
#$fontFolder = "C:\Users\mj8u\Desktop\Fontes_Windows10"
$fontFolder = (Get-Location).Path
$cont = 0
#$openType = "(Open Type)"
$openType = "(TrueType)"
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
$objShell = New-Object -ComObject Shell.Application

$objFolder = $objShell.namespace($fontFolder)
#Write-Host $objFolder
#Gravar-Log $objFolder
    
$destination = "c:\Windows\Fonts" -join ""
foreach ($file in $objFolder.items()){
    $fileType = $($objFolder.getDetailsOf($file, 2))
    #Write-Host $fileType
    if($fileType -like "Arquivo de fonte TrueType*"){
        $fontName = $($objFolder.getDetailsOf($File, 21))
        #Write-Host "Titulo da fonte: $fontName"
        #Gravar-Log "Titulo da fonte: $fontName"
        $regKeyName = $fontName,$openType -join " "
        $regKeyValue = [System.IO.Path]::GetFileName($file.Path)
        #Verificando e removendo a fonte corrompida
        if(Test-Path "C:\Windows\Fonts\$regKeyValue"){
            Write-Host "Identificando o arquivo C:\Windows\Fonts\$regKeyValue"
            Gravar-Log "Identificando o arquivo C:\Windows\Fonts\$regKeyValue"
            try{
                Remove-Item C:\Windows\Fonts\$regKeyValue -Force -ErrorAction Stop
                Write-Host "Fonte antiga removida"
                Gravar-log "Fonte antiga removida"
                Write-Host "Instalando fonte : $fontName , Arquivo: $regKeyValue"
                Gravar-log "Instalando fonte : $fontName , Arquivo: $regKeyValue"
                               
                Copy-Item $file.Path  $destination -Force
                $1 = New-ItemProperty -Path $regPath -Name $regKeyname -Value $regKeyValue -Force
                $cont++
            }
            catch{
                Write-Host "Falha na remocao da fonte : $fontName , Arquivo: $regKeyValue"
                Gravar-log "Falha na remocao da fonte : $fontName , Arquivo: $regKeyValue"
            }
        } 
        
    }
}
Write-Host "Fontes instaladas: $cont"
Gravar-Log "Fontes instaladas: $cont"
Gravar-Log "Fim do script. Sera necessario efetuar um logoff para aplicar as fontes."
[System.Environment]::Exit(0)
    
