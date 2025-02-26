param (
    [string]$tenantId,             # Parameter for tenant ID
    [string]$clientId,             # Parameter for client ID
    [string]$url,                  # Parameter for the target URL
    [string]$outputPath = ""       # Optional parameter for the output file path; defaults to timestamped file in current directory
)

# Prompt for client secret securely
$secureClientSecret = Read-Host "Enter Client Secret" -AsSecureString

# Convert the secure string to plain text for use in the request
# Note: Be cautious with plain text handling of sensitive data
$clientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                   [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureClientSecret))

# Generate the current timestamp for the output file name if no path is specified
if (-not $outputPath) {
    $currentTimestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
    $outputFileName = "result_$currentTimestamp.json"
    $outputPath = Join-Path -Path "." -ChildPath $outputFileName
}

# Prepare the body for the token request
$body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}

# Get an access token for Microsoft Graph API
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $body
$accessToken = $tokenResponse.access_token

# Initialize a list to hold all results and a counter for pages
$allResults = @()
$pageCount = 0

# Loop to handle pagination
do {
    Write-Output "Processing page number: $pageCount"

    # Perform the API request for the current page
    $response = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $($accessToken)" }
    $allResults += $response.value
    $url = $response.'@odata.nextLink'  # Continue if there's a nextLink

    # Increment the page counter
    $pageCount++
} while ($url -ne $null)

# Convert results to JSON and write to the specified file path
$allResults | ConvertTo-Json -Depth 4 | Out-File -FilePath $outputPath -Encoding utf8

# Output the total number of processed pages and items
Write-Output "Data written to $outputPath."
Write-Output "Total number of items: $($allResults.Count)"
Write-Output "Total number of processed pages: $pageCount"