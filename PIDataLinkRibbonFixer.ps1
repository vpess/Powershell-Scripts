#Script presente em todas as instalações do PI Datalink ( C:\Program Files\PIPC\Excel e também em C:\Program Files (x86)\PIPC\Excel). Copyright no fim do script

Remove-Variable * -ErrorAction SilentlyContinue

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 3) {
    "You are running a non-supported version of PowerShell."
    "The script only supports PowerShell 3.0 or later as Microsoft has deprecated the earlier versions."
    "Executing the script may still fix the missing ribbon issue, but will not generate any logs."
    "Do you still want to continue? (Y/N)"
    $response = read-host
    if ($response -ne "Y") {exit}
    }

$script_start_time = Get-Date
"The script is executed at $script_start_time" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
"`n" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append

$root = "HKLM:","HKCU:"

$path1 = "\SOFTWARE\Microsoft\Office\Excel"
$path2 = "\SOFTWARE\Wow6432Node\Microsoft\Office\Excel"
$path3 = "\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Excel"
$path4 = "\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Excel"
$path5 = "\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Microsoft\Office\Excel"
$path6 = "\SOFTWARE\Microsoft\Office\15.0\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\Excel"
$branch = Get-Variable -Name path* -ValueOnly

$keyword = "PI DataLink"

$OS_bit = (Get-WmiObject Win32_OperatingSystem).OSArchitecture.substring(0,2)

# Check for disabled Excel add-ins.
"Checking if any Excel add-ins are hard-disabled." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
$disabled_addins = @()
Foreach ($i in $root)
{
    Foreach ($j in $branch)
    {
		try {
			$search_results = Get-ChildItem $i$j -Recurse -ErrorAction Stop | Where-Object {($_.Name -clike "*DisabledItems*") -and ($_.ValueCount -gt 0)}
			$disabled_addins += $search_results.Name
		}
		catch {
			$_.Exception.Message | Out-File $PSScriptRoot\RibbonFixer.log -Append
		}
    }
}
# "Disabled Excel add-ins are found at $disabled_addins" | Out-File $PSScriptRoot\RibbonFixer.log -Append
if($disabled_addins -ne $null) {
    "There are hard-disabled Excel add-ins, please follow the Microsoft article provided below to re-enable them." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
    "https://msdn.microsoft.com/en-us/library/ms268871.aspx" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append}
else {
    "No hard-disabled add-ins are found." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append}
"`n" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append

# Find existing registry keys of PI DataLink Excel add-in.
"Searching for existing registry keys of PI DataLink Excel add-in." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
$existing_dir = @()
Foreach ($i in $root)
{
    Foreach ($j in $branch)
    {
		try {
			$search_results = Get-ChildItem $i$j -Recurse -ErrorAction Stop | Where-Object {$_.Name -match $keyword -and $_.Name -notmatch "Notifications" -and $_.Name -notmatch "Legacy"}
			$existing_dir += $search_results.Name
		}
		catch {
			$_.Exception.Message | Out-File $PSScriptRoot\RibbonFixer.log -Append
		}
    }
}
$existing_dir = $existing_dir | ? {$_} 
if($existing_dir.Count -gt 0) {
    $existing_dir_str = ($existing_dir -join ",")
    "Existing PI DataLink addin registry keys are found at $existing_dir_str" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
    "`n" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
    
    # Take a backup of all existing registry keys of PI DataLink Excel add-in for recovery or further investigation.
    $backupFile = "$PSScriptRoot\backup.reg"
    Foreach ($existing_key in $existing_dir)
    {
        "Taking a back up of $existing_key." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        try {
			$index = $existing_dir.IndexOf($existing_key)
			reg export $existing_key "$PSScriptRoot\temp$index.reg" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        }
        catch {
			$_.Exception.Message | Out-File $PSScriptRoot\RibbonFixer.log -Append
        }
    }
    try {
		"Windows Registry Editor Version 5.00" | Set-Content $backupFile
		Get-Content "$PSScriptRoot\temp*.reg" | ? {
			$_ -ne 'Windows Registry Editor Version 5.00'
		} | Add-Content $backupFile
    }
    catch {
		$_.Exception.Message | Out-File $PSScriptRoot\RibbonFixer.log -Append
    }
    try {
		Get-ChildItem $PSScriptRoot -Filter temp*.reg | Remove-Item
    }
    catch {
		$_.Exception.Message | Out-File $PSScriptRoot\RibbonFixer.log -Append
    }
}
else {
    "No existing PI DataLink addin registry keys are found. Please make sure PI DataLink is installed or contact escalation for help if it already is." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
    "`n" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append}

# Check the current LoadBehavior of each existing PI DataLink Excel addin registry key.
# If it's 3, do nothing.
# For any other values, reset it to 3.
# If the entry doesn't exist, delete the entire PI DataLink key and re-create it as well as the LoadBehavior and Manifest entries with the correct values.
Foreach ($dir in $existing_dir)
{
    $parent_dir = $dir.TrimEnd("\PI DataLink")
    $currentLoadBehavior = (Get-ItemProperty -Path Registry::$dir -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior
    if ($currentLoadBehavior -eq 3) {
        "The LoadBehavior of $dir is currently set to $currentLoadBehavior." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        "No change is to be made." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
    }
    ElseIf ($currentLoadBehavior -ne $null) {
        "The LoadBehavior of $dir is currently set to $currentLoadBehavior." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        "Resetting LoadBehavior to 3." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        try
        {
			Set-ItemProperty -Path Registry::$dir -Name LoadBehavior -Value 3 -ErrorAction Stop
			$dir + "|LoadBehavior is set to 3." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        }
        catch {
			$_.Exception.Message | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        }
    }
    Else {
        "No LoadBehavior entry is found in $dir." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        "Deleting the entire key and re-creating it as well as the LoadBehavior and Manifest entries with the correct values." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
        try
		{
			Remove-Item -Path Registry::$dir -ErrorAction Stop
			"$dir is removed." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
			New-Item -Path Registry::$parent_dir -Name $keyword -ErrorAction Stop | Out-File $PSScriptRoot\RibbonFixer.log -Append
			"$dir is re-created." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
			New-ItemProperty -Path Registry::$dir -Name LoadBehavior -Value 3 -Force | Out-File $PSScriptRoot\RibbonFixer.log -Append
			"$dir|LoadBehavior is created and set to 3." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
			if ($OS_bit -eq 32 -or $dir.Contains("6432") -eq 1) {
				New-ItemProperty -Path Registry::$dir -Name Manifest -Value "file:\\$env:PIHOME\Excel\OSIsoft.PIDataLink.UI.vsto|vstolocal" -Force | Out-File $PSScriptRoot\RibbonFixer.log -Append
			}
			Else {
				New-ItemProperty -Path Registry::$dir -Name Manifest -Value "file:\\$env:PIHOME64\Excel\OSIsoft.PIDataLink.UI.vsto|vstolocal" -Force | Out-File $PSScriptRoot\RibbonFixer.log -Append
			}
			"$dir|Manifest is created and reset." | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
		}
		catch {
			$_.Exception.Message | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
		}
    }
    "`n" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
}

$script_end_time = Get-Date
"The execution is finished at $script_end_time" | Tee-Object -File $PSScriptRoot\RibbonFixer.log -Append
"`n" | Out-File $PSScriptRoot\RibbonFixer.log -Append

# SIG # Begin signature block
# MIIbygYJKoZIhvcNAQcCoIIbuzCCG7cCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDu7GCr0eCrjHj0
# LM0vmin3gzI0sQsM4jHdyPsRPa8QlKCCCnowggUwMIIEGKADAgECAhAECRgbX9W7
# ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBa
# Fw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/l
# qJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fT
# eyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqH
# CN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+
# bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLo
# LFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIB
# yTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAK
# BggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwA
# AgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAK
# BghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0j
# BBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7s
# DVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGS
# dQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6
# r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo
# +MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qz
# sIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHq
# aGxEMrJmoecYpJpkUe8wggVCMIIEKqADAgECAhAF/ovfX92ZLF/1euv5PeP1MA0G
# CSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0
# IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMTcwNjE1MDAwMDAw
# WhcNMTkwNjIwMTIwMDAwWjB/MQswCQYDVQQGEwJVUzELMAkGA1UECBMCQ0ExFDAS
# BgNVBAcTC1NhbiBMZWFuZHJvMRUwEwYDVQQKEwxPU0lzb2Z0LCBMTEMxFTATBgNV
# BAMTDE9TSXNvZnQsIExMQzEfMB0GCSqGSIb3DQEJARYQY29kZUBvc2lzb2Z0LmNv
# bTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAL65qeF9YvzlQJqumgPX
# U3k0QnCXX9nLAyOeWFBXnyh9m6ckQbcqoSSOXWMHvBsBP4zZM0swMlMzv4dBwRgw
# rKMFQNnuHA/iyT00h7PfcxI3RNNTiOg0rk/Efs1drtZRNjbm+VCAwougvfjcbSqs
# wbHH5OUeC7y80qLZcF/rJJ0TdMEbGg773efQ6fakv4+RrUsxdPoin3mMHiRji2ee
# CMtAk9xt27yvwwCn5M09TlJF65YgA4q5Nve94kNQonjn8U5Vs+ryIja4309KxaAQ
# aZO5EiBi8Opo23MOjdCekhSE5As7tiHTXMx/VCYzjx9/flnPz71KMzCYk+aCbfCx
# BEkCAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5LfZldQ5Y
# MB0GA1UdDgQWBBR0qWp9jhqvhmnUje7R3szZhKzmeTAOBgNVHQ8BAf8EBAMCB4Aw
# EwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAzoDGGL2h0
# dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMEwG
# A1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3
# LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4MHYwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEFBQcwAoZC
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJ
# RENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQAD
# ggEBAL6YG3rJGv5A2eSk6L08mnwYujvcka+kef37DUKTXjZqGCltOrhzwy0y7ybn
# cbx73eg3zDY3TLVYEH868qwRUXa41N2HbvK2qw2AaFQR0WIbYyqCXTGaPuSg5L0Z
# OidUiFkIlGwKKA2cK/KySPLewQsKDDiUy9EN/uBFB+GIJ0CBJkYQMlbOVELi5AXY
# PNFzp7kX7WUQvk9t1iH277QxgG0sxJBvwvJiwwzNGdb8dHTjNSca3OvagJW94OxS
# 3K0ivNM3lVgd6rMAJdp2YqC4ly6IGwni4MUimBkzxemIrrtNFGIeyDp85ITgGBzN
# wcebhsBA9DEiKKpjAKyfPy8BGLwxghCmMIIQogIBATCBhjByMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWdu
# aW5nIENBAhAF/ovfX92ZLF/1euv5PeP1MA0GCWCGSAFlAwQCAQUAoIGwMBkGCSqG
# SIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3
# AgEVMC8GCSqGSIb3DQEJBDEiBCByU7eb1S9C4Ph2XvlV4YCo802k0PpWnqRYxWtv
# SZ8NUjBEBgorBgEEAYI3AgEMMTYwNKAygDAAUABJACAARABhAHQAYQBMAGkAbgBr
# ACAAUgBpAGIAYgBvAG4AIABGAGkAeABlAHIwDQYJKoZIhvcNAQEBBQAEggEAaRdA
# b/GE2uNkR32BvdL7cVg6KeM92Kti04kYY8htOwzfL7zuofpvCV65mji+zE8Px3M5
# FUFRW2z3OhOPrXKlDFUr4X/mcl6W3+BRyzARw2aYYS8BI9WFSeuUKGCMKAyiKLyg
# sDsAIUaVZYpN2z4jJIGAX55L+SNBQH11q1w+NjoKcLO2IHqGXk59sRkXx8cayP9q
# n6+f1SeuQTqAF65NJK+qqUH3DE3vHcdb6g4chzdzWYGvEZ/V9g+oF7iUaTQP0WMz
# f7HIDWg0mmT8y0cpxkY58TSTuhRfmz/2lstFJLAC06E3hPyfdFd2Wug5x4IVZD2S
# 0KeeAnaHYaYfwXqhPaGCDj0wgg45BgorBgEEAYI3AwMBMYIOKTCCDiUGCSqGSIb3
# DQEHAqCCDhYwgg4SAgEDMQ0wCwYJYIZIAWUDBAIBMIIBDwYLKoZIhvcNAQkQAQSg
# gf8EgfwwgfkCAQEGC2CGSAGG+EUBBxcDMDEwDQYJYIZIAWUDBAIBBQAEIMDzrrBT
# p+CQKit77HqZOehiGic5hm2BjjlQYB4z/HOMAhUAxZvY6e9Z52z2y0HCisszGVSs
# mAQYDzIwMTcxMTE3MDExMDI3WjADAgEeoIGGpIGDMIGAMQswCQYDVQQGEwJVUzEd
# MBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xHzAdBgNVBAsTFlN5bWFudGVj
# IFRydXN0IE5ldHdvcmsxMTAvBgNVBAMTKFN5bWFudGVjIFNIQTI1NiBUaW1lU3Rh
# bXBpbmcgU2lnbmVyIC0gRzKgggqLMIIFODCCBCCgAwIBAgIQewWx1EloUUT3yYnS
# nBmdEjANBgkqhkiG9w0BAQsFADCBvTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZl
# cmlTaWduLCBJbmMuMR8wHQYDVQQLExZWZXJpU2lnbiBUcnVzdCBOZXR3b3JrMTow
# OAYDVQQLEzEoYykgMjAwOCBWZXJpU2lnbiwgSW5jLiAtIEZvciBhdXRob3JpemVk
# IHVzZSBvbmx5MTgwNgYDVQQDEy9WZXJpU2lnbiBVbml2ZXJzYWwgUm9vdCBDZXJ0
# aWZpY2F0aW9uIEF1dGhvcml0eTAeFw0xNjAxMTIwMDAwMDBaFw0zMTAxMTEyMzU5
# NTlaMHcxCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlv
# bjEfMB0GA1UECxMWU3ltYW50ZWMgVHJ1c3QgTmV0d29yazEoMCYGA1UEAxMfU3lt
# YW50ZWMgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALtZnVlVT52Mcl0agaLrVfOwAa08cawyjwVrhponADKXak3J
# ZBRLKbvC2Sm5Luxjs+HPPwtWkPhiG37rpgfi3n9ebUA41JEG50F8eRzLy60bv9iV
# kfPw7mz4rZY5Ln/BJ7h4OcWEpe3tr4eOzo3HberSmLU6Hx45ncP0mqj0hOHE0Xxx
# xgYptD/kgw0mw3sIPk35CrczSf/KO9T1sptL4YiZGvXA6TMU1t/HgNuR7v68kldy
# d/TNqMz+CfWTN76ViGrF3PSxS9TO6AmRX7WEeTWKeKwZMo8jwTJBG1kOqT6xzPnW
# K++32OTVHW0ROpL2k8mc40juu1MO1DaXhnjFoTcCAwEAAaOCAXcwggFzMA4GA1Ud
# DwEB/wQEAwIBBjASBgNVHRMBAf8ECDAGAQH/AgEAMGYGA1UdIARfMF0wWwYLYIZI
# AYb4RQEHFwMwTDAjBggrBgEFBQcCARYXaHR0cHM6Ly9kLnN5bWNiLmNvbS9jcHMw
# JQYIKwYBBQUHAgIwGRoXaHR0cHM6Ly9kLnN5bWNiLmNvbS9ycGEwLgYIKwYBBQUH
# AQEEIjAgMB4GCCsGAQUFBzABhhJodHRwOi8vcy5zeW1jZC5jb20wNgYDVR0fBC8w
# LTAroCmgJ4YlaHR0cDovL3Muc3ltY2IuY29tL3VuaXZlcnNhbC1yb290LmNybDAT
# BgNVHSUEDDAKBggrBgEFBQcDCDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGlt
# ZVN0YW1wLTIwNDgtMzAdBgNVHQ4EFgQUr2PWyqNOhXLgp7xB8ymiOH+AdWIwHwYD
# VR0jBBgwFoAUtnf6aUhHn1MS1cLqBzJ2B9GXBxkwDQYJKoZIhvcNAQELBQADggEB
# AHXqsC3VNBlcMkX+DuHUT6Z4wW/X6t3cT/OhyIGI96ePFeZAKa3mXfSi2VZkhHEw
# Kt0eYRdmIFYGmBmNXXHy+Je8Cf0ckUfJ4uiNA/vMkC/WCmxOM+zWtJPITJBjSDlA
# IcTd1m6JmDy1mJfoqQa3CcmPU1dBkC/hHk1O3MoQeGxCbvC2xfhhXFL1TvZrjfdK
# er7zzf0D19n2A6gP41P3CnXsxnUuqmaFBJm3+AZX4cYO9uiv2uybGB+queM6AL/O
# ipTLAduexzi7D1Kr0eOUA2AKTaD+J20UMvw/l0Dhv5mJ2+Q5FL3a5NPD6itas5VY
# VQR9x5rsIwONhSrS/66pYYEwggVLMIIEM6ADAgECAhBUWPKq10HWRLyEqXugllLm
# MA0GCSqGSIb3DQEBCwUAMHcxCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRl
# YyBDb3Jwb3JhdGlvbjEfMB0GA1UECxMWU3ltYW50ZWMgVHJ1c3QgTmV0d29yazEo
# MCYGA1UEAxMfU3ltYW50ZWMgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTAeFw0xNzAx
# MDIwMDAwMDBaFw0yODA0MDEyMzU5NTlaMIGAMQswCQYDVQQGEwJVUzEdMBsGA1UE
# ChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xHzAdBgNVBAsTFlN5bWFudGVjIFRydXN0
# IE5ldHdvcmsxMTAvBgNVBAMTKFN5bWFudGVjIFNIQTI1NiBUaW1lU3RhbXBpbmcg
# U2lnbmVyIC0gRzIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCZ8/zY
# BAkDhvnXXKaTwEJ86nxjz10A4o7zwJDfjyn1GOqUt5Ll17Cgc4Ho6QqbSnwB/52P
# pDmnDupF9CIMOnDtOUWL5MUbXPBFaEYkBWN2mxz8nmwqsVblin9Sca7yNdVGIwYc
# z0gtHbTNuNl2I44c/z6/uwZcaQemZQ74Xq59Lu1NrjXvydcAQv0olQ6fXXJCCbzD
# 2kTS7cxHhOT8yi2sWL6u967ZRA0It8J31hpDcNFuA95SksQQCHHZuiJV8h+87Zud
# O+JeHUyD/5cPewvnVYNO0g3rvtfsrm5HuZ/fpdZRvARV7f8ncEzJ7SpLE+GxuUwP
# yQHuVWVfaQJ4Zss/AgMBAAGjggHHMIIBwzAMBgNVHRMBAf8EAjAAMGYGA1UdIARf
# MF0wWwYLYIZIAYb4RQEHFwMwTDAjBggrBgEFBQcCARYXaHR0cHM6Ly9kLnN5bWNi
# LmNvbS9jcHMwJQYIKwYBBQUHAgIwGRoXaHR0cHM6Ly9kLnN5bWNiLmNvbS9ycGEw
# QAYDVR0fBDkwNzA1oDOgMYYvaHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20v
# c2hhMjU2LXRzcy1jYS5jcmwwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0P
# AQH/BAQDAgeAMHcGCCsGAQUFBwEBBGswaTAqBggrBgEFBQcwAYYeaHR0cDovL3Rz
# LW9jc3Aud3Muc3ltYW50ZWMuY29tMDsGCCsGAQUFBzAChi9odHRwOi8vdHMtYWlh
# LndzLnN5bWFudGVjLmNvbS9zaGEyNTYtdHNzLWNhLmNlcjAoBgNVHREEITAfpB0w
# GzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtNTAdBgNVHQ4EFgQUCbXB/pZylylD
# msngArqu+P0vuvYwHwYDVR0jBBgwFoAUr2PWyqNOhXLgp7xB8ymiOH+AdWIwDQYJ
# KoZIhvcNAQELBQADggEBABezCojpXFpeIGs7ChWybMWpijKH07H0HFOuhb4/m//X
# vLeUhbTHUn6U6L3tYbLUp5nkw8mTwTU9C+hoCl1WmL2xIjvRRHrXv/BtUTKK1SPf
# OAE39uJTK3orEY+3TWx6MwMbfGsJlBe75NtY1CETZefs0SXKLHWanH/8ybsqaKvE
# fbTPo8lsp9nEAJyJCneR9E2i+zE7hm725h9QA4abv8tCq+Z2m3JaEQGKxu+lb5Xn
# 3a665iJl8BhZGxHJzYC32JdHH0II+KxxH7BGU7PUstWjq1B1SBIXgq3P4EFPMn7N
# lRy/kYoIPaSnZwKW3yRMpdBBwIJgo4oXMkvTvM+ktIwxggJaMIICVgIBATCBizB3
# MQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xHzAd
# BgNVBAsTFlN5bWFudGVjIFRydXN0IE5ldHdvcmsxKDAmBgNVBAMTH1N5bWFudGVj
# IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0ECEFRY8qrXQdZEvISpe6CWUuYwCwYJYIZI
# AWUDBAIBoIGkMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0B
# CQUxDxcNMTcxMTE3MDExMDI3WjAvBgkqhkiG9w0BCQQxIgQg5/yHv8RcCTXRo5OI
# YdkxKpU8EZuXs8c0QvW9cOwYQ+IwNwYLKoZIhvcNAQkQAi8xKDAmMCQwIgQgz3rB
# etBH7NX9w2giAxsS1O8Hi28rTF5rpB+P8s9LrWcwCwYJKoZIhvcNAQEBBIIBABen
# mKrWg/tNhI1Kaa2mZYs9KZ/sX2K1mrUnkc6yIZ7sILRJQ7FBDZccf1BbvGZdL5Lf
# ObynC0ynjWq5P7ye6nrzfsd/CCqDcmPKbFzZib6UccmF0e3iKTHjl6LikhUV6v4i
# wppyEeRxIPD4LeU7Wg+zt14vnyrI++M27/LuMGvXqmtxNdyedDeIgY1vhTC1xS0n
# rP4A/SbbpEVLmD8c1zC8hChZ8Wz9lENmkg2o2J2kJVVlglVJDw/wOA5Z6X+ogzXp
# +b6zTBOtD6dpsenImHy617xv3b08BUTkZ3+Mr54Qy6i0TvQU3pJSgLXYZlwvtgPi
# NovLMmJ6J3SKE0arEvs=
# SIG # End signature block


# Copyright (c) 2017 OSIsoft, LLC All rights reserved.

# THIS SOFTWARE CONTAINS CONFIDENTIAL INFORMATION AND TRADE SECRETS OF
# OSIsoft, LLC  USE, DISCLOSURE, OR REPRODUCTION IS PROHIBITED WITHOUT
# THE PRIOR EXPRESS WRITTEN PERMISSION OF OSIsoft, LLC

# RESTRICTED RIGHTS LEGEND
# Use, duplication, or disclosure by the Government is subject to restrictions
# as set forth in subparagraph (c)(1)(ii) of the Rights in Technical Data and
# Computer Software clause at DFARS 252.227.7013

# OSIsoft, LLC
# 1600 Alvarado St., San Leandro, CA 94577

# FileVersion("1.1")