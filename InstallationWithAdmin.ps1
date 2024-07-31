$DefaultInstallFolder = "C:\Aspose.Cells.WebAddin"
$DefaultInstallPackagePath = "$(Get-Location).Path\Aspose.Cells.WebAddIn.zip"
$HostName = HOSTNAME


$certName = "CN=localhost"
$curPath= $(Get-Location).Path
$certPath = "${curPath}\self-signed.pfx"
Write-Host $certPath
$certPassword = ConvertTo-SecureString -String "MyPassword123!" -Force -AsPlainText
# Create a new self-signed certificate
$cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "Cert:\LocalMachine\My"
# Export the certificate to a PFX file
$null = Export-PfxCertificate -Cert "Cert:\LocalMachine\My\$($cert.Thumbprint)" -FilePath $certPath -Password $certPassword

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath, "MyPassword123!")
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()
Write-Output "Build self-signed.pfx."


if ( -not $(Test-Path -Path $DefaultInstallPackagePath) ) 
{
    $DefaultInstallPackagePath = ".\Aspose.Cells.WebAddIn.zip"
    if(-not $(Test-Path -Path $DefaultInstallPackagePath))
    {
        Write-Host "Aspose.Cells Excel Web Add-In Package not found, please download it from https://github.com/aspose-cells/Aspose.Cells-for-Excel-Web-AddIn/releases/download/v1.0.0/Aspose.Cells.for.Excel.Web.AddIn.zip"
        Exit
    }
}

if( -not $(Test-Path $DefaultInstallFolder) )
{
    New-Item -Path $DefaultInstallFolder -ItemType Directory
}

Expand-Archive -Path $DefaultInstallPackagePath -DestinationPath $DefaultInstallFolder -Force
Copy-Item -Path .\self-signed.pfx -Destination $DefaultInstallFolder\server\self-signed.pfx
New-SmbShare -Name  "AsposeCellsWebAddIn" -Path "$DefaultInstallFolder" -FullAccess "Everyone"
New-Service -Name "AsposeCellsExcelAddInServices" -BinaryPathName "$DefaultInstallFolder\server\Aspose.Cells.WebAddin.Server.exe" -StartupType Automatic
Start-Service -Name "AsposeCellsExcelAddInServices" 
Start-Sleep 2
$guid = New-Guid
Write-Host "[$guid]" , "[$HostName]"
Write-Host "[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$guid}]"
New-Item -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$guid}" -Force
Write-Host """Id""=""{$guid}"""
Set-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$guid}" -Name "Id" -Value "{$guid}" -Type String
Write-Host """Url""=""file://$HostName/AsposeCellsWebAddIn/add-ins/"""
Set-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$guid}" -Name "Url" -Value "file://$HostName/AsposeCellsWebAddIn/add-ins" -Type String
Write-Host """Flags""=dword:00000001"
Set-ItemProperty -Path "Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$guid}" -Name "Flags" -Value "1" -Type DWord

Write-Host "Start to check installation result:"
Start-Sleep 4
. .\CheckEnv.ps1
