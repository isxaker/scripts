# ==== CONFIGURATION ====
#
# Displays out total number of items in the specified mailbox
#
# 1 Define Azure AD and Microsoft Graph details
$tenantId     = ""
# 2 Define Azure AD and Microsoft Graph details
$clientId     = ""
# 3 Define Azure AD and Microsoft Graph details
$clientSecret = ""
# 4 Define the target user mailbox to count all items
$userId       = ""

$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# Authenticate and get access token
$body = @{
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}
$response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
$accessToken = $response.access_token

# --- Count All Items in a Mailbox ---
$headers = @{
    "Authorization"    = "Bearer $accessToken"
    "ConsistencyLevel" = "eventual"
}

# Get all mail folders
$foldersUrl = "https://graph.microsoft.com/v1.0/users/$userId/mailFolders"
$folders = @()
do {
    $resp = Invoke-RestMethod -Headers $headers -Uri $foldersUrl -Method Get
    $folders += $resp.value
    $foldersUrl = $resp.'@odata.nextLink'
} while ($foldersUrl)

# Count messages in each folder
$totalCount = 0
foreach ($folder in $folders) {
    $folderId = $folder.id
    $messagesUrl = "https://graph.microsoft.com/v1.0/users/$userId/mailFolders/$folderId/messages?`$count=true&`$top=1"
    $resp = Invoke-RestMethod -Headers $headers -Uri $messagesUrl -Method Get
    $count = $resp.'@odata.count'
    if ($count) { $totalCount += $count }
}


Write-Host "Total items in mailbox ${userId}: $totalCount"