Enum Workloads {
    Exchange
    OneDrive
}

$tenantId = ""
$clientId = ""
$clientSecret = ""

# Define the workload type
$workloadType = [Workloads]::Exchange

# Define the initial API endpoint for protection units
$PolicyId = ""

# Select the URL based on workload type
switch ($workloadType) {
    Exchange {
        $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/exchangeProtectionPolicies/$PolicyId/mailboxProtectionUnits"
    }
    OneDrive {
        $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/oneDriveForBusinessProtectionPolicies/$PolicyId/driveProtectionUnit"
    }
}

# Verify URL has been set
if (-not $url) {
    Write-Error "URL is not set. Check workload type or switch logic."
    exit
}

Write-Output "Selected Workload Type: $workloadType"
Write-Output "Using URL: $url"

# Prepare the body for the token request
$tokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}

# Get an access token for Microsoft Graph API
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $tokenBody
$accessToken = $tokenResponse.access_token

# Initialize a list to hold all PUs and keep track of the number of processed units
$allResults = @()
$pageCountPU = 0 # Page counter for protection units
$processedPUs = 0 # Total number of processed PUs

# Loop to handle pagination for protection units
do {
    # Perform the API request
    $response = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $($accessToken)" }
    $allResults += $response.value
    $url = $response.'@odata.nextLink'  # Continue if there's a nextLink
    
    # Update processed PUs count
    $processedPUs += $response.value.Count

    # Output the current status after processing each page
    Write-Output "Fetching protection units, page number: $pageCountPU. Processed PUs: $processedPUs"

    # Increment page counter
    $pageCountPU++
} while ($url -ne $null)

$pageCountUser = 0 # Index for tracking each protection unit

# Now process each protection unit to get user information
foreach ($unit in $allResults) {
    Write-Output "Fetching user details for protection unit, index: $pageCountUser"

    $userId = $unit.directoryObjectId
    if ($userId) {
        # Get user information
        $userResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$userId`?`$select=id,deletedDateTime,createdDateTime,accountEnabled,userPrincipalName,mail,displayName" -Headers @{ Authorization = "Bearer $($accessToken)" }
        $unit | Add-Member -MemberType NoteProperty -Name user -Value $userResponse
    }
    
    # Add processed_index property
    $unit | Add-Member -MemberType NoteProperty -Name processed_index -Value $pageCountUser
    
    $pageCountUser++
}

$currentTimestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
$outputFileName = "result_$currentTimestamp.json"
$outputPath = Join-Path -Path "." -ChildPath $outputFileName

# Convert results to JSON and write to a file
$allResults | ConvertTo-Json -Depth 4 | Out-File -FilePath $outputPath -Encoding utf8

# Output the total number of processed pages and items
Write-Output "Data written to $outputPath."
Write-Output "Total number of items: $($allResults.Count)"