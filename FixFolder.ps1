#------ em construção ------

#https://autoatendimentotic.petrobras.com.br/visualizar/5267364/11093


Set-ItemProperty -path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
Rename-Item -Path ".\FileExts\" -NewName "FileExts.old"

<#

1. Acesse o menu "Iniciar";
2.Pesquise por "regedit" e execute-o;
3. Localize a seguinte chave (pasta) "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\";
4. Renomeie a chave "FileExts" para "FileExts.old";
5. Feche o editor de registro do Windows (regedit);
6. Reinicie o Windows.

#>