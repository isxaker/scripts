# Variables
$ClientId     = ""          # Application (Client) ID from Azure AD
$TenantId     = ""          # Consuming Tenant ID from Azure AD
$ClientSecret = ""          # Client Secret from Azure AD
$GraphApiVersion = "v1.0"   # Or "beta" if needed
$RestoreSessionId = ""      # Restore Session Id

# Microsoft Graph API endpoint
$GraphEndpoint = "https://graph.microsoft.com/$GraphApiVersion"

# Request body for token
$body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $ClientId
    Client_Secret = $ClientSecret
}

# Get Access Token
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body
$accessToken = $tokenResponse.access_token

# 1. Call Exchange Restore Artifacts endpoint
$MailboxArtifactsEndpoint = "$GraphEndpoint/solutions/backupRestore/exchangeRestoreSessions/$RestoreSessionId/mailboxRestoreArtifacts"
$mailboxArtifactsResponse = Invoke-RestMethod -Uri $MailboxArtifactsEndpoint -Headers @{ Authorization = "Bearer $($accessToken)" }

# Save mailbox artifacts response
$currentTimestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
$mailboxArtifactsFileName = "mailboxRestoreArtifacts__$currentTimestamp.json"
$mailboxArtifactsFilePath = Join-Path -Path "." -ChildPath $mailboxArtifactsFileName
$mailboxArtifactsResponse | ConvertTo-Json -Depth 10 | Out-File -FilePath $mailboxArtifactsFilePath -Encoding utf8
Write-Output "Mailbox Restore Artifacts written to $mailboxArtifactsFilePath."

# 2. Call Restore Session endpoint
$RestoreSessionEndpoint = "$GraphEndpoint/solutions/backupRestore/restoreSessions/$RestoreSessionId"
$restoreSessionResponse = Invoke-RestMethod -Uri $RestoreSessionEndpoint -Headers @{ Authorization = "Bearer $($accessToken)" }

# Save restore session response
$restoreSessionFileName = "restoreSession__$currentTimestamp.json"
$restoreSessionFilePath = Join-Path -Path "." -ChildPath $restoreSessionFileName
$restoreSessionResponse | ConvertTo-Json -Depth 10 | Out-File -FilePath $restoreSessionFilePath -Encoding utf8
Write-Output "Restore Session info written to $restoreSessionFilePath."

# Optionally, display counts or summary
Write-Output "Mailbox Restore Artifacts count: $($mailboxArtifactsResponse.Value.Count)"
Write-Output "Restore Session properties: $($restoreSessionResponse | Get-Member | Where-Object {$_.MemberType -eq 'NoteProperty'} | Select-Object -ExpandProperty Name)"