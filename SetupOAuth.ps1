Write-host "checking if OAuth Servers are present"

$AuthServers = Get-AuthServer

if($AuthServers -lt 1){
    Write-host "OAuth Servers not found, setting them up now..."

    New-AuthServer -Name "WindowsAzureACS" -AuthMetadataUrl "https://accounts.accesscontrol.windows.net/vwaknr.onmicrosoft.com/metadata/json/1"
    New-AuthServer -Name "evoSTS" -Type AzureAD -AuthMetadataUrl "https://login.windows.net/vwaknr.onmicrosoft.com/federationmetadata/2007-06/federationmetadata.xml"
}

else {
    Write-Host "Auth-Servers are present..."
}

write-host "checking if Partner Application for ExchangeOnline is present..."

$partnerApplication = Get-PartnerApplication

if($partnerApplication.applicationIdentitfier -eq "00000002-0000-0ff1-ce00-000000000000"){
    Write-host "Parnter Application for Exchange Online is present... Moving On."
}
else {
    Write-Host "Partner Application was not found, setting it up now."

    Get-PartnerApplication |  Where-Object {$_.ApplicationIdentifier -eq "00000002-0000-0ff1-ce00-000000000000" -and $_.Realm -eq ""} | Set-PartnerApplication -Enabled $true
}

$defaultPath = "$env:SystemDrive\OAuthConfig\OauthCert.cer"

write-host "Exporting Certificate for OAuth configuration. Certificate will be saved to: $($defaultPath)"

$thumbprint = (Get-AuthConfig).CurrentCertificateThumbprint
if((Test-Path $env:SYSTEMDRIVE\OAuthConfig) -eq $false)
{
   New-Item -Path $env:SYSTEMDRIVE\OAuthConfig -Type Directory
}
Set-Location -Path $env:SYSTEMDRIVE\OAuthConfig
$oAuthCert = (dir Cert:\LocalMachine\My) | Where-Object {$_.Thumbprint -match $thumbprint}
$certType = [System.Security.Cryptography.X509Certificates.X509ContentType]::Cert
$certBytes = $oAuthCert.Export($certType)
$CertFile = "$env:SYSTEMDRIVE\OAuthConfig\OAuthCert.cer"
[System.IO.File]::WriteAllBytes($CertFile, $certBytes)

Write-host "Please Connect to M365 as global admin to upload the certificate"

start-sleep -Seconds 3

try {
    write-host "checking if GraphModule is present"

    import-module -name Microsoft.Graph.Applications -ErrorAction Stop
}
catch {
    Write-host "Graph Module is not present, installing it..."

    try {
        Install-module -Name Microsoft.Graph.Applications -ErrorAction Stop

        write-host "Graph Module successfully installed, uploading Certificate"
    }
    catch {
        Write-host "Graph Module could not be installed. Script will exit with error $($error[0].exception.message)" -BackgroundColor red -ForegroundColor White

        Start-Sleep -Seconds 5

        Exit
    }
    
}

try {

    write-host "Connecting to Graph Module"
    Connect-MgGraph -Scopes Application.ReadWrite.All -ErrorAction Stop
}
catch {
    Write-host "Connection to Graph moduled failed with error $($error[0].Exception.Message), script will exit" -BackgroundColor red -ForegroundColor White

    Start-Sleep -Seconds 5

    exit
}


$CertFile = "$env:SYSTEMDRIVE\OAuthConfig\OAuthCert.cer"
$objFSO = New-Object -ComObject Scripting.FileSystemObject
$CertFile = $objFSO.GetAbsolutePathName($CertFile)
$cer = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertFile)
$binCert = $cer.GetRawCertData()
$credValue = [System.Convert]::ToBase64String($binCert)
$ServiceName = "00000002-0000-0ff1-ce00-000000000000"
Write-Host "[+] Trying to query the service principals for service: $ServiceName" -ForegroundColor Cyan
$p = Get-MgServicePrincipal -Filter "AppId eq '$ServiceName'"
Write-Host "[+] Trying to query the keyCredentials for service: $ServiceName" -ForegroundColor Cyan
$servicePrincipalKeyInformation = Get-MgServicePrincipal -Filter "AppId eq '$ServiceName'" -Select "keyCredentials"

$keyCredentialsLength = $servicePrincipalKeyInformation.KeyCredentials.Length
if ($keyCredentialsLength -gt 0) {
   Write-Host "[+] $keyCredentialsLength existing key(s) found - we keep them if they have not expired" -ForegroundColor Cyan

$newCertAlreadyExists = $false
   $servicePrincipalObj = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal
   $keyCredentialsArray = @()

foreach ($cred in $servicePrincipalKeyInformation.KeyCredentials) {
      $thumbprint = [System.Convert]::ToBase64String($cred.CustomKeyIdentifier)

Write-Host "[+] Processing existing key: $($cred.DisplayName) thumbprint: $thumbprint" -ForegroundColor Cyan

if ($newCertAlreadyExists -ne $true) {
         $newCertAlreadyExists = ($cer.Thumbprint).Equals($thumbprint, [System.StringComparison]::OrdinalIgnoreCase)
      }

if ($cred.EndDateTime -lt (Get-Date)) {
         Write-Host "[+] This key has expired on $($cred.EndDateTime) and will not be retained" -ForegroundColor Yellow
         continue
      }

$keyCredential = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential
      $keyCredential.Type = "AsymmetricX509Cert"
      $keyCredential.Usage = "Verify"
      $keyCredential.Key = $cred.Key

$keyCredentialsArray += $keyCredential
   }

if ($newCertAlreadyExists -eq $false) {
      Write-Host "[+] New key: $($cer.Subject) thumbprint: $($cer.Thumbprint) will be added" -ForegroundColor Cyan
      $keyCredential = New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential
      $keyCredential.Type = "AsymmetricX509Cert"
      $keyCredential.Usage = "Verify"
      $keyCredential.Key = [System.Text.Encoding]::ASCII.GetBytes($credValue)

$keyCredentialsArray += $keyCredential

$servicePrincipalObj.KeyCredentials = $keyCredentialsArray
      Update-MgServicePrincipal -ServicePrincipalId $p.Id -BodyParameter $servicePrincipalObj
   } else {
      Write-Host "[+] New key: $($cer.Subject) thumbprint: $($cer.Thumbprint) already exists and will not be uploaded again" -ForegroundColor Yellow
   }
} else {
   $params = @{
      type = "AsymmetricX509Cert"
      usage = "Verify"
      key = [System.Text.Encoding]::ASCII.GetBytes($credValue)
   }

Write-Host "[+] This is the first key which will be added to this service principal" -ForegroundColor Cyan
   Update-MgServicePrincipal -ServicePrincipalId $p.Id -KeyCredentials $params
}

write-host "Registering services..."

$MailServer = Read-host "Please enter hybrid server FQDN - example: https://mail.contoso.com:"

$AutodiscoverFQDN = Read-host "Please enter Autodiscover-FQDN - example: https://autodiscover.contoso.com"

$ServiceName = "00000002-0000-0ff1-ce00-000000000000";
$x = Get-MgServicePrincipal -Filter "AppId eq '$ServiceName'"
$x.ServicePrincipalNames += $MailServer
$x.ServicePrincipalNames += $AutodiscoverFQDN

Update-MgservicePrincipal -ServicePrincipalId $x.id -ServicePrincipalNames $x.ServicePrincipalNames

Write-Host "Here is a list of all services that were added"

Get-MgServicePrincipal -Filter "AppId eq '$ServiceName'" | Select-Object -ExpandProperty ServicePrincipalNames | Sort-Object
 
Write-host "Setting up local IntraOrganizationConnector..."

$ServiceDomain = (Get-AcceptedDomain | Where-Object {$_.DomainName -like "*.mail.onmicrosoft.com"}).DomainName.Address
New-IntraOrganizationConnector -Name ExchangeHybridOnPremisesToOnline -DiscoveryEndpoint https://outlook.office365.com/autodiscover/autodiscover.svc -TargetAddressDomains $ServiceDomain

Write-host "Setting up cloud IntraOrganizationConnector..."

try {
    Write-host "checking if module exists..."

    import-module ExchangeOnlineManagement -ErrorAction Stop
}
catch {
    install-module ExchangeOnlineManagement
}

Write-host "Trying to connect to Exchange Online..."

try {
    Connect-ExchangeOnline -Prefix Cloud -ErrorAction Stop
}
catch {
     Write-host "Could not connect to ExchangeOnline with error $($error[0].Exception.message)"

     Start-Sleep -Seconds 5

     Exit
}

$Domains = Read-Host "please enter your mail domainName: - example: Contoso.com"

New-CloudIntraOrganizationConnector -Name ExchangeHybridOnlineToOnPremises -DiscoveryEndpoint $AutodiscoverFQDN -TargetAddressDomains $domains
