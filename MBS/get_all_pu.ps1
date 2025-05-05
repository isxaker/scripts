param(
    [string]$WorkloadType,              # Specifies the type of workload (e.g., SharePoint, OneDrive, Exchange)
    [bool]$OnlyCurrentPolicy = $true,   # Indicates whether to enumerate only using the current policy
    [bool]$Detailed = $true             # Determines whether to fetch detailed information about an item under protection
)

$tenantId = ""
#$tenantId = "zpnrb.onmicrosoft.com"
$clientId = ""
$clientSecret = ""

# Check if WorkloadType is provided
if (-not $WorkloadType) {
    Write-Host "WorkloadType is required. Specify SharePoint, OneDrive, or Exchange." -ForegroundColor Red
    exit
}

Write-Output "Detailed processing mode: $Detailed"
Write-Output "Checking OnlyCurrentPolicy: $OnlyCurrentPolicy"

# Prepare the body for the token request
$tokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}

# Get an access token for Microsoft Graph API
try {
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $tokenBody
    if (-not $tokenResponse.access_token) {
        Write-Host "Failed to authenticate. Access token is missing." -ForegroundColor Red
        exit
    }
    $accessToken = $tokenResponse.access_token
    Write-Host "Authentication successful." -ForegroundColor Green
} catch {
    Write-Host "Authentication failed: $_" -ForegroundColor Red
    exit
}

# Function to determine Policy ID by Workload Type from all protection policies
function GetPolicyIdByWorkloadType($accessToken, $WorkloadType) {
    $url = "https://graph.microsoft.com/beta/solutions/backupRestore/ProtectionPolicies"
    try {
        $response = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $($accessToken)" } -Method Get -ErrorAction Stop
        if ($response -and $response.value) {
            # Determine @odata.type based on WorkloadType
            $typeMapping = @{
                SharePoint = "#microsoft.graph.sharePointProtectionPolicy"
                OneDrive = "#microsoft.graph.oneDriveForBusinessProtectionPolicy"
                Exchange = "#microsoft.graph.exchangeProtectionPolicy"
            }
            $odataType = $typeMapping[$WorkloadType]

            # Find the policy with the specified WorkloadType
            $policy = $response.value | Where-Object { $_.'@odata.type' -eq $odataType }
            if ($policy) {
                Write-Host "Detected Policy ID for ${WorkloadType}: $(${policy.id})" -ForegroundColor Green
                return $policy.id
            } else {
                Write-Host "No policy found for WorkloadType $WorkloadType." -ForegroundColor Red
                exit
            }
        } else {
            Write-Host "No valid response received from protection policies endpoint." -ForegroundColor Red
            exit
        }
    } catch {
        Write-Host "Error accessing protection policies endpoint: $_" -ForegroundColor Red
        exit
    }
}

if($OnlyCurrentPolicy){

    # Determine the Policy ID based on Workload Type
    $PolicyId = GetPolicyIdByWorkloadType -accessToken $accessToken -WorkloadType $WorkloadType

    # Print the Policy ID to debug
    Write-Host "Policy ID for workload type ${WorkloadType}: $PolicyId" -ForegroundColor Yellow

    # Select the URL based on detected workload type
    switch ($WorkloadType) {
        "OneDrive" {
            $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/oneDriveForBusinessProtectionPolicies/$PolicyId/driveProtectionUnits"
        }
        "SharePoint" {
            $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/sharePointProtectionPolicies/$PolicyId/siteProtectionUnits"
        }
        "Exchange" {
            $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/exchangeProtectionPolicies/$PolicyId/mailboxProtectionUnits"
        }
        default {
            Write-Host "Unsupported workload type detected: $WorkloadType" -ForegroundColor Red
            exit
        }
    }
}
else{
    # Select the URL based on detected workload type
    switch ($WorkloadType) {
        "OneDrive" {
            $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.driveProtectionUnit"
        }
        "SharePoint" {
            $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.siteProtectionUnit"
        }
        "Exchange" {
            $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.mailboxProtectionUnit"
        }
        default {
            Write-Host "Unsupported workload type detected: $WorkloadType" -ForegroundColor Red
            exit
        }
    }

}

Write-Host "Using URL: $url" -ForegroundColor Green

# Function to get site information for SharePoint
function Get-SiteDetails($unit, $accessToken) {
    $siteId = $unit.siteId
    if ($siteId) {
        # Get site information
        try {
            $siteUrl = "https://graph.microsoft.com/v1.0/sites/${siteId}"
            $siteResponse = Invoke-RestMethod -Uri $siteUrl -Headers @{ Authorization = "Bearer $($accessToken)" }

            # Check if siteResponse contains expected fields
            if ($siteResponse -and $siteResponse.id) {
                $unit | Add-Member -MemberType NoteProperty -Name site -Value @{
                    id = $siteResponse.id
                    displayName = $siteResponse.displayName
                    webUrl = $siteResponse.webUrl
                    createdDateTime = $siteResponse.createdDateTime
                    lastModifiedDateTime = $siteResponse.lastModifiedDateTime
                }
            } else {
                $errorMsg = @{
                    message = "No valid site data returned"
                    siteId = $siteId
                }
                #Write-Host ($errorMsg | ConvertTo-Json -Depth 4) -ForegroundColor Red
                $unit | Add-Member -MemberType NoteProperty -Name siteError -Value $errorMsg
            }

        } catch {
            $errorMsg = @{
                message = "Error fetching site details"
                siteId = $siteId
                errorDetail = $_
            }
            #Write-Host ($errorMsg | ConvertTo-Json -Depth 4) -ForegroundColor Red
            $unit | Add-Member -MemberType NoteProperty -Name siteError -Value $errorMsg
        }
    } else {
        $errorMsg = @{
            message = "Invalid siteId"
            siteId = $siteId
        }
        #Write-Host ($errorMsg | ConvertTo-Json -Depth 4) -ForegroundColor Red
        $unit | Add-Member -MemberType NoteProperty -Name siteError -Value $errorMsg
    }
}

# Function to get user information for Exchange/OneDrive
function Get-UserDetails($unit, $accessToken) {
    $userId = $unit.directoryObjectId
    if ($userId) {
        # Get user information
        try {
            $userResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$userId`?`$select=id,deletedDateTime,createdDateTime,accountEnabled,userPrincipalName,mail,displayName" -Headers @{ Authorization = "Bearer $($accessToken)" }
            $unit | Add-Member -MemberType NoteProperty -Name user -Value $userResponse
        } catch {
            $errorMsg = @{
                message = "Error fetching user details"
                userId = $userId
                errorDetail = $_.Response
            }
            #Write-Host ($errorMsg | ConvertTo-Json -Depth 4) -ForegroundColor Red
            $unit | Add-Member -MemberType NoteProperty -Name userError -Value $errorMsg
        }
    } else {
        $errorMsg = @{
            message = "Invalid userId"
            userId = $userId
        }
        #Write-Host ($errorMsg | ConvertTo-Json -Depth 4) -ForegroundColor Red
        $unit | Add-Member -MemberType NoteProperty -Name userError -Value $errorMsg
    }
}

# Initialize a list to hold all PUs and keep track of the number of processed units
$allResults = @()
$pageCountPU = 0 # Page counter for protection units
$processedPUs = 0 # Total number of processed PUs

# Loop to handle pagination for protection units
do {
    # Perform the API request
    try {
        $response = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $($accessToken)" }
        $allResults += $response.value
        $url = $response.'@odata.nextLink'  # Continue if there's a nextLink
        
        # Update processed PUs count
        $processedPUs += $response.value.Count

        # Output the current status after processing each page
        Write-Host "Fetching protection units, page number: $pageCountPU. Processed PUs: $processedPUs" -ForegroundColor Green

        # Increment page counter
        $pageCountPU++
    } catch {
        Write-Host "Failed to fetch protection units: $_" -ForegroundColor Red
        exit
    }
} while ($url -ne $null)

$pageCountUser = 0 # Index for tracking each protection unit

# Process each protection unit based on workload type
foreach ($unit in $allResults) {
    Write-Host "Fetching details for protection unit, index: $pageCountUser" -ForegroundColor Green

    if($Detailed){
        if ($WorkloadType -eq "SharePoint") {
            Get-SiteDetails -unit $unit -accessToken $accessToken
        } else {
            Get-UserDetails -unit $unit -accessToken $accessToken
        }
    }
    
    # Add processed_index property
    $unit | Add-Member -MemberType NoteProperty -Name processed_index -Value $pageCountUser
    
    $pageCountUser++
}

# Only create a file if processing was successful
if ($allResults.Count -gt 0) {
    $currentTimestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
    $outputFileName = "result_${WorkloadType}_${PolicyId}_$currentTimestamp.json"
    $outputPath = Join-Path -Path "." -ChildPath $outputFileName

    # Convert results to JSON and write to a file
    $allResults | ConvertTo-Json -Depth 4 | Out-File -FilePath $outputPath -Encoding utf8

    # Output the total number of processed pages and items
    Write-Host "Data written to $outputPath." -ForegroundColor Green
    Write-Host "Total number of items: $($allResults.Count)" -ForegroundColor Green
} else {
    Write-Host "No protection units were processed. No file created." -ForegroundColor Yellow
}