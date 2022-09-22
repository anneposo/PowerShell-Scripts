# https://adamtheautomator.com/powershell-graph-api/ 
# https://docs.microsoft.com/en-us/azure/active-directory/develop/howto-create-self-signed-certificate 

# Create a self signed cert to use for an Azure app registration 
# to run PowerShell scripts with app-only authentication.

# Your tenant name (can something more descriptive as well)
$TenantName        = "domain.onmicrosoft.com"

# Where to export the certificate without the private key
$CerOutputPath     = "C:\Users\user\Desktop\PowerShellGraphCert.cer"
$PrivateKeyPath    = "C:\Users\user\Desktop\PowerShellGraphCert.pfx"

# What cert store you want it to be in
$StoreLocation     = "Cert:\CurrentUser\My"

# Expiration date of the new certificate
$ExpirationDate    = (Get-Date).AddYears(2)


# Splat for readability
$CreateCertificateSplat = @{
    FriendlyName      = "AzureApp"
    DnsName           = $TenantName
    CertStoreLocation = $StoreLocation
    NotAfter          = $ExpirationDate
    KeyExportPolicy   = "Exportable"
    KeySpec           = "Signature"
    Provider          = "Microsoft Enhanced RSA and AES Cryptographic Provider"
    HashAlgorithm     = "SHA256"
}

# Create certificate
$Certificate = New-SelfSignedCertificate @CreateCertificateSplat

# Get certificate path
$CertificatePath = Join-Path -Path $StoreLocation -ChildPath $Certificate.Thumbprint

# Export certificate
Export-Certificate -Cert $CertificatePath -FilePath $CerOutputPath

# Create a password for certificate's private key
$mypwd = ConvertTo-SecureString -String "{myPassword}" -Force -AsPlainText  ## Replace {myPassword}

# Use the password to export your private key file
Export-PfxCertificate -Cert $Certificate -FilePath $PrivateKeyPath -Password $mypwd 