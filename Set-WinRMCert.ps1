<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.179
	 Created on:   	8/13/2020 10:15 AM
	 Created by:   	Richard Stoddart
	 Organization: 	Geico
	 Filename:     	Set-WinRMCert.ps1
	===========================================================================
	.DESCRIPTION
		Enables Remote PS (WinRM) SSL HTTPS service.
		Sets certificate to default WinRM  port. Port 5986
#>

#Requires -RunAsAdministrator
#enable WinRM HTTPS service
& winrm quickconfig -transport:https -q
Start-Sleep -Seconds 2

#get Certificate
$CertPath = "Cert:\LocalMachine\My"
$cert = ((Get-ChildItem $CertPath | Sort-Object NotAfter )[-1] )
if (!$cert) { Write-error "Certificate not found in $CertPath "; return }

# Create text for CMD file
'CSCRIPT ' +
$env:SystemRoot +
'\System32\winrm.vbs ' + 
'set winrm/config/Listener?Address=*+Transport=HTTPS @{Hostname="'+
$Cert.Subject.TrimStart('CN=') + 
'";CertificateThumbprint="' + 
$cert.Thumbprint +
'"}' |
Out-File -FilePath .\TempSetWinRM.cmd -Encoding oem

# Add below line Will turn off the HTTP WinRM port if not controled by GPO
# + "`n" +'winrm set winrm/config/Listener?Address=*+Transport=HTTP @{Enabled="false"}'

# Execute TempSetWinRM.CMD file
$Out = & .\TempSetWinRM.cmd

#vailidate config worked, catch error
$CertInstalled =
	(($out | ? { $_.trim() -like "CertificateThumbprint*" }).split("=")[1]).trim()

If ($CertInstalled -eq $cert.Thumbprint)
	{ Write-Output "Sucess: Certificate $($cert.Thumbprint) $($Cert.Subject)" }
Else { Write-Error $out[0];  return}
Write-Output "`n"

#Output WinRM settings
Write-Output "Winrm Setting results `n --------------------------"

& Winrm enumerate winrm/config/listener

Remove-Item -Path '.\TempSetWinRM.cmd'

#GI WSMan:\localhost\Service\CertificateThumbprint | Set-Item -Value ""



