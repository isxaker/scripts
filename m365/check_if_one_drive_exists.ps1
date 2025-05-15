# 1 Define Azure AD and Microsoft Graph details
$tenantId = ""
# 2 Define Azure AD and Microsoft Graph details
$clientId = ""
# 3 Define Azure AD and Microsoft Graph details
$clientSecret = ""
# 4 Define Azure AD and Microsoft Graph details
$resource = "https://graph.microsoft.com/"
# 5 Define Azure AD and Microsoft Graph details
$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# 6 Authenticate and get access token
$body = @{
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}

# 7 Authenticate and get access token
$response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
# 8 Authenticate and get access token
$accessToken = $response.access_token

# 9 Get all users in the tenant
$usersEndpoint = "https://graph.microsoft.com/v1.0/users"
# 10 Get all users in the tenant
$usersResponse = Invoke-RestMethod -Uri $usersEndpoint -Headers @{Authorization = "Bearer $accessToken"}

# 11 Initialize user index
$userIndex = 0

# 12 Iterate over each user and check OneDrive availability
foreach ($user in $usersResponse.value) {
    # 13 Increment user index
    $userIndex++
    # 14 Get user details
    $userId = $user.id
    # 15 Get user details
    $userPrincipalName = $user.userPrincipalName

    # 16 Attempt to get the user's OneDrive
    try {
        # 17 Attempt to get the user's OneDrive
        $driveEndpoint = "https://graph.microsoft.com/v1.0/users/$userId/drive"
        # 18 Attempt to get the user's OneDrive
        $driveResponse = Invoke-RestMethod -Uri $driveEndpoint -Headers @{Authorization = "Bearer $accessToken"}
        # 19 Output result
        Write-Host "$userIndex User $userPrincipalName OneDrive [true]"
    } catch {
        # 20 Output result
        Write-Host "$userIndex User $userPrincipalName OneDrive [false]"
    }
}