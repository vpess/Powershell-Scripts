Write-Host "###Aguarde a coleta de dados###`n"
$i = Get-ADUser -filter * -server transp.biz -properties Name, Displayname, UserPrincipalName, Title, CanonicalName, Department, DistinguishedName | Select-Object Name, Displayname, UserPrincipalName, Title, CanonicalName, Department, DistinguishedName | Where-Object {$_.UserPrincipalName -like '*.INDRA@transpetro.com.br'}
$c = $i | Where-Object {$_.DistinguishedName -like '*OU=SEDE-RJ*'}
Write-Host "###Organizando dados###`n"
$i | Sort-Object -Property Displayname | Export-Csv -Delimiter ';' -Path "$ENV:USERPROFILE\Desktop\indra_transp.csv" -NoTypeInformation | Format-Table
$c | Sort-Object -Property Displayname | Export-Csv -Delimiter ';' -Path "$ENV:USERPROFILE\Desktop\indra_campos_transp.csv" -NoTypeInformation | Format-Table
Write-Host "###Arquivos salvos na Área de Trabalho###"
Start-Sleep -s 3
