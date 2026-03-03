# Graph-EWS-TestApplflow.ps1
# Generted using CoPilot  
#
# This is a sample script to demonstrate using EWS Managed API with app-only OAuth (client secret) 
# and impersonation to access another user's mailbox. It includes example calls to list folders 
# and read inbox items.
#
# References:           
# - EWS OAuth + app-only + impersonation + X-AnchorMailbox guidance (Microsoft Learn)       
# - ImpersonatedUserId usage (Microsoft docs)   

<#
EWS Managed API + App-only OAuth (client secret) + Impersonation
No parameters: edit CONFIG section.

References:
- EWS OAuth + app-only + impersonation + X-AnchorMailbox guidance (Microsoft Learn)
- ImpersonatedUserId usage (Microsoft docs)
#>

# =========================
# CONFIG (EDIT THESE)
# =========================
$TenantId = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
$ClientId = "11111111-2222-3333-4444-555555555555"
$ClientSecretPlainText = "PASTE_CLIENT_SECRET_VALUE_HERE"   # Consider storing securely in production

# Mailbox you want to access via impersonation (SMTP address)
$ImpersonateSmtpAddress = "user@contoso.com"


 
# EWS endpoint (EXO)
$EwsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"

# Path to EWS Managed API DLL
$EwsManagedApiDll = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

# =========================


function Get-EwsAppOnlyAccessToken {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][string]$ClientSecretPlainText
    )

    # OAuth2 v2 token endpoint
    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    # App-only EWS uses the resource ".default" scope for outlook.office365.com (per Microsoft)
    $scope = "https://outlook.office365.com/.default"  # [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)

    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecretPlainText
        grant_type    = "client_credentials"
        scope         = $scope
    }

    $resp = Invoke-RestMethod -Method Post -Uri $tokenEndpoint `
        -ContentType "application/x-www-form-urlencoded" -Body $body

    if (-not $resp.access_token) {
        throw "Token response did not contain access_token."
    }

    return $resp.access_token
}

# 1) Load EWS Managed API
if (-not (Test-Path $EwsManagedApiDll)) {
    throw "EWS Managed API DLL not found at: $EwsManagedApiDll`nInstall EWS Managed API or update `$EwsManagedApiDll."
}
Add-Type -Path $EwsManagedApiDll

# 2) Acquire app-only access token using client secret
$accessToken = Get-EwsAppOnlyAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecretPlainText $ClientSecretPlainText

# 3) Create and configure ExchangeService
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(
    [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
)

$service.Url = [Uri]$EwsUrl

# Attach OAuth token to EWS client (OAuthCredentials)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($accessToken)  # [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)

# 4) App-only requires explicit impersonation of target mailbox (SMTP/UPN/SID)
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
    [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
    $ImpersonateSmtpAddress
)  # [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)[2](https://github.com/MicrosoftDocs/office-developer-exchange-docs/blob/main/docs/exchange-web-services/how-to-identify-the-account-to-impersonate.md)

# 5) When using impersonation, include X-AnchorMailbox = impersonated SMTP
if ($service.HttpHeaders.ContainsKey("X-AnchorMailbox")) {
    $service.HttpHeaders["X-AnchorMailbox"] = $ImpersonateSmtpAddress
} else {
    $service.HttpHeaders.Add("X-AnchorMailbox", $ImpersonateSmtpAddress)
}  # [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)

# =========================
# DEMO CALLS
# =========================

# A) List top 10 folders under mailbox root
try {
    $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10)
    $folders = $service.FindFolders(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,
        $view
    )

    "Top folders for impersonated mailbox: $ImpersonateSmtpAddress"
    $folders.Folders | Select-Object DisplayName, TotalCount, ChildFolderCount | Format-Table -AutoSize
}
catch {
    throw "EWS FindFolders failed: $($_.Exception.Message)"
}

# B) Read 5 newest inbox items
try {
    $inboxId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,
        $ImpersonateSmtpAddress
    )

    $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(5)
    $itemView.OrderBy.Add(
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
        [Microsoft.Exchange.WebServices.Data.SortDirection]::Descending
    )

    $items = $service.FindItems($inboxId, $itemView)

    # Load common properties (subject/from/received etc.)
    $service.LoadPropertiesForItems($items, [Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties)

    $items.Items | ForEach-Object {
        [pscustomobject]@{
            Subject   = $_.Subject
            From      = $_.From.Address
            Received  = $_.DateTimeReceived
            ItemClass = $_.ItemClass
        }
    } | Format-Table -AutoSize
}
catch {
    throw "EWS FindItems failed: $($_.Exception.Message)"
}
