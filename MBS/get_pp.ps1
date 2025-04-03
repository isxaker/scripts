# Variables
$ClientId = "<>"          # Application (Client) ID from Azure AD
$TenantId = "<>"          # Consuming Tenant ID from Azure AD
$ClientSecret = "<>"  # Client Secret from Azure AD
$GraphEndpoint = "https://graph.microsoft.com/v1.0" # Microsoft Graph API endpoint

$body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}

Write-Output "Getting all protection policies..."
# Get an access token for Microsoft Graph API
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $body
$accessToken = $tokenResponse.access_token

# Example: Get protection policies (replace with the correct Graph API endpoint for MBS protection policies)
$PoliciesEndpoint = "$GraphEndpoint/solutions/backupRestore/protectionPolicies"
$response = Invoke-RestMethod -Uri $PoliciesEndpoint -Headers @{ Authorization = "Bearer $($accessToken)" }

# Step 3: Output Policies
$currentTimestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
$outputFileName = "protection_policies__$currentTimestamp.json"
$outputPath = Join-Path -Path "." -ChildPath $outputFileName
$response | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputPath  -Encoding utf8
Write-Output "Data written to $outputPath."
Write-Output "Total number of items: $($response.Value.Count)"