param(
    [string]$WorkloadType = "",             # Specifies the type of workload (e.g., SharePoint, OneDrive, Exchange). If empty, all workloads will be processed
    [bool]$UseProtectionPolicies = $true,   # Indicates whether to enumerate PUs through protection policies
    [bool]$Detailed = $true,                # Determines whether to fetch detailed information about an item under protection
    [bool]$IncludeHeaders = $false,         # Determines whether to include request/response headers in output
    [bool]$CreateEmptyPolicyFiles = $false  # Determines whether to create JSON files for policies that have no protection units
)

$tenantId = ""
#$tenantId = "zpnrb.onmicrosoft.com"
$clientId = ""
$clientSecret = ""

# Generate a single timestamp for all output files in this run
$currentTimestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"

# Define supported workload types
$SupportedWorkloads = @("SharePoint", "OneDrive", "Exchange")

# Validate WorkloadType if provided
if ($WorkloadType -and $WorkloadType -notin $SupportedWorkloads) {
    Write-Host "Invalid WorkloadType. Supported values: $($SupportedWorkloads -join ', ')" -ForegroundColor Red
    exit
}

Write-Output "Detailed processing mode: $Detailed"
Write-Output "Using protection policies: $UseProtectionPolicies"

# Determine workloads to process
$workloadsToProcess = if ($WorkloadType) { @($WorkloadType) } else { $SupportedWorkloads }
Write-Output "Processing workload(s): $($workloadsToProcess -join ', ')"

# Prepare the body for the token request
$tokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}

# Get an access token for Microsoft Graph API
try {
    Write-Output "Sending token request..."
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $tokenBody -ErrorAction Stop
    
    if ($tokenResponse.access_token) {
        Write-Output "Successfully got access token"
        $accessToken = $tokenResponse.access_token
        $url = "https://graph.microsoft.com/beta/solutions/backupRestore/ProtectionPolicies"
        Write-Output "Trying to get policies from: $url"
        $response = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $($accessToken)" } -Method Get -ErrorAction Stop
        Write-Output "Found policies:"
        foreach ($policy in $response.value) {
            Write-Output "- $($policy.'@odata.type') (ID: $($policy.id))"
        }
    } else {
        Write-Output "No access token in response"
    }
} catch {
    Write-Output "Error occurred: $($_.Exception.Message)"
    Write-Output "Full error: $_"
    exit 1
}

# Function to filter headers to only whitelisted values
function Filter-Headers($headers) {
    $filteredHeaders = @{}
    if ($headers -and $headers.Count -gt 0) {
        # Whitelist of headers we want to keep
        $whitelistedHeaders = @("request-id", "client-request-id")
        foreach ($key in $whitelistedHeaders) {
            if ($headers.ContainsKey($key)) {
                $filteredHeaders[$key] = $headers[$key]
            }
        }
    }
    return $filteredHeaders
}

# Function to create standardized headers structure
function Create-HeadersStructure($url, $includeHeaders, $responseHeaders = $null, $statusCode = $null) {
    if ($includeHeaders) {
        return @{
            request = @{
                url = $url
            }
            response = @{
                statusCode = $statusCode
                headers = Filter-Headers $responseHeaders
            }
        }
    } else {
        return @{
            request = @{}
            response = @{}
        }
    }
}

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

# Function to get all Protection Policy IDs for a given Workload Type
function GetPolicyIdsByWorkloadType($accessToken, $WorkloadType, $IncludeHeaders) {
    $url = "https://graph.microsoft.com/beta/solutions/backupRestore/ProtectionPolicies"
    $requestHeaders = @{ Authorization = "Bearer $($accessToken)" }
    
    try {
        $invokeParams = @{ Uri = $url; Headers = $requestHeaders; Method = "Get"; ErrorAction = "Stop" }
        if ($IncludeHeaders) {
            $invokeParams["ResponseHeadersVariable"] = "respHeaders"
            $invokeParams["StatusCodeVariable"] = "respStatus"
        }
        
        $response = Invoke-RestMethod @invokeParams
        
        $policyHeaders = Create-HeadersStructure -url $url -includeHeaders $IncludeHeaders -responseHeaders $respHeaders -statusCode $respStatus
        
        if ($response -and $response.value) {
            # Determine @odata.type based on WorkloadType
            $typeMapping = @{
                SharePoint = "#microsoft.graph.sharePointProtectionPolicy"
                OneDrive = "#microsoft.graph.oneDriveForBusinessProtectionPolicy"
                Exchange = "#microsoft.graph.exchangeProtectionPolicy"
            }
            $odataType = $typeMapping[$WorkloadType]

            # Show all policies for debugging
            Write-Host "All available policies:" -ForegroundColor Yellow
            $response.value | ForEach-Object {
                Write-Host "Policy type: $($_.'@odata.type'), ID: $($_.id)" -ForegroundColor Yellow
            }
            
            # Find all policies for the specified WorkloadType
            $policies = $response.value | Where-Object { $_.'@odata.type' -eq $odataType }
            if ($policies) {
                $policyIds = @($policies | Select-Object -ExpandProperty id)  # Force array with @()
                Write-Host "Detected Policy IDs for ${WorkloadType}: $($policyIds -join ', ')" -ForegroundColor Green
                return @{
                    policyIds = $policyIds
                    policies = $policies
                    headers = $policyHeaders
                }
            } else {
                Write-Host "No policy found for WorkloadType $WorkloadType." -ForegroundColor Yellow
                return $null
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

# Initialize array to store all results across all workloads and policies
$allResults = @()

# Process workloads based on whether we're using protection policies or not
if ($UseProtectionPolicies) {
    # Process each workload with policies
    $policyPagesTable = @{}
    foreach ($currentWorkload in $workloadsToProcess) {
        Write-Host "`nProcessing workload: $currentWorkload" -ForegroundColor Cyan
        # Get all Policy IDs for the current workload
        $policyResult = GetPolicyIdsByWorkloadType -accessToken $accessToken -WorkloadType $currentWorkload -IncludeHeaders $IncludeHeaders
        # Skip this workload if no policies found and continue with the next workload
        if ($null -eq $policyResult) {
            Write-Host "Skipping $currentWorkload workload due to no policies found" -ForegroundColor Yellow
            continue
        }
        $PolicyIds = $policyResult.policyIds
        $policies = $policyResult.policies
        $policyRetrievalHeaders = $policyResult.headers
        
        # Create a lookup table for policy data by ID
        $policyLookup = @{}
        foreach ($policy in $policies) {
            $policyLookup[$policy.id] = $policy
        }
        
        foreach ($PolicyId in $PolicyIds) {
            $pages = @()
            # Print the current Policy ID
            Write-Host "Processing Policy ID for workload type $currentWorkload`: $PolicyId" -ForegroundColor Yellow
            # Select the URL based on detected workload type
            switch ($currentWorkload) {
                "OneDrive" {
                    $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/oneDriveForBusinessProtectionPolicies/$PolicyId/driveProtectionUnits"
                }
                "SharePoint" {
                    $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/sharePointProtectionPolicies/$PolicyId/siteProtectionUnits"
                }
                "Exchange" {
                    $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/exchangeProtectionPolicies/$PolicyId/mailboxProtectionUnits"
                }
            }
            Write-Host "Using URL: $url" -ForegroundColor Green
            $pageCountPU = 0
            $processedPUs = 0
            $nextUrl = $url
            do {
                $requestHeaders = @{ Authorization = "Bearer $($accessToken)" }
                $pageInfo = @{ page = $pageCountPU }
                $pageValues = @()
                try {
                    $invokeParams = @{ Uri = $nextUrl; Headers = $requestHeaders }
                    if ($IncludeHeaders) {
                        $invokeParams["ResponseHeadersVariable"] = "respHeaders"
                        $invokeParams["StatusCodeVariable"] = "respStatus"
                    }
                    $response = Invoke-RestMethod @invokeParams
                    foreach ($item in $response.value) {
                        $item | Add-Member -MemberType NoteProperty -Name "policyId" -Value $PolicyId -Force
                    }
                    $pageValues = $response.value
                    $nextUrl = $response.'@odata.nextLink'
                    $processedPUs += $response.value.Count
                    Write-Host "Fetching protection units for policy $PolicyId, page number: $pageCountPU. Processed PUs: $processedPUs" -ForegroundColor Green
                    $pageCountPU++
                    $pageHeaders = Create-HeadersStructure -url $invokeParams.Uri -includeHeaders $IncludeHeaders -responseHeaders $respHeaders -statusCode $respStatus
                } catch {
                    Write-Host "Failed to fetch protection units: $_" -ForegroundColor Red
                    exit
                }
                $pageInfo = @{ 
                    page = $pageCountPU
                    headers = $pageHeaders
                }
                $pages += @{ info = $pageInfo; values = $pageValues }
                
                # Add to allResults for detailed processing if needed
                $allResults += $pageValues
            } while ($nextUrl -ne $null)
            
            # Process detailed information for this policy's protection units
            if ($Detailed -and $allResults.Count -gt 0) {
                Write-Host "Processing detailed information for $($allResults.Count) protection units in policy $PolicyId" -ForegroundColor Cyan
                $pageCountUser = 0
                foreach ($unit in $allResults) {
                    Write-Host "Fetching details for $currentWorkload protection unit, index: $pageCountUser" -ForegroundColor Green
                    if ($unit.PSObject.Properties["siteId"] -and $unit.siteId) {
                        Get-SiteDetails -unit $unit -accessToken $accessToken
                    } else {
                        Get-UserDetails -unit $unit -accessToken $accessToken
                    }
                    $pageCountUser++
                }
                # Clear allResults for next policy
                $allResults = @()
            }
            
            $policyPagesTable["${currentWorkload}_$PolicyId"] = @{
                pages = $pages
                policyHeaders = $policyRetrievalHeaders
                policyData = $policyLookup[$PolicyId]
            }
        }
    } # End of foreach workload
} # <-- Add this closing brace for the if ($UseProtectionPolicies) block

# Handle non-policy based queries
if (-not $UseProtectionPolicies) {
    Write-Host "`nProcessing without protection policies" -ForegroundColor Cyan
    $workloadResultsTable = @{}
    foreach ($currentWorkload in $workloadsToProcess) {
        Write-Host "Processing workload: $currentWorkload" -ForegroundColor Cyan
        # Select the URL based on detected workload type
        switch ($currentWorkload) {
            "OneDrive" {
                $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.driveProtectionUnit"
            }
            "SharePoint" {
                $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.siteProtectionUnit"
            }
            "Exchange" {
                $url = "https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.mailboxProtectionUnit"
            }
        }
        Write-Host "Using URL: $url" -ForegroundColor Green
        $pageCountPU = 0 # Page counter for protection units
        $processedPUs = 0 # Total number of processed PUs
        $pages = @()
        $nextUrl = $url
        do {
            $requestHeaders = @{ Authorization = "Bearer $($accessToken)" }
            $pageValues = @()
            $respHeaders = $null
            $respStatus = $null
            try {
                $invokeParams = @{ Uri = $nextUrl; Headers = $requestHeaders }
                if ($IncludeHeaders) {
                    $invokeParams["ResponseHeadersVariable"] = "respHeaders"
                    $invokeParams["StatusCodeVariable"] = "respStatus"
                }
                $response = Invoke-RestMethod @invokeParams
                $pageValues = $response.value
                $nextUrl = $response.'@odata.nextLink'
                $processedPUs += $response.value.Count
                Write-Host "Fetching protection units for $currentWorkload, page number: $pageCountPU. Processed PUs: $processedPUs" -ForegroundColor Green
                $pageHeaders = Create-HeadersStructure -url $invokeParams.Uri -includeHeaders $IncludeHeaders -responseHeaders $respHeaders -statusCode $respStatus
            } catch {
                Write-Host "Failed to fetch protection units: $_" -ForegroundColor Red
                exit
            }
            $pageInfo = @{ 
                page = $pageCountPU
                headers = $pageHeaders
            }
            $pageCountPU++
            
            # Add to allResults for detailed processing if needed
            $allResults += $pageValues
            
            $pages += @{ info = $pageInfo; values = $pageValues }
        } while ($nextUrl -ne $null)
        
        # Process detailed information for this workload's protection units
        if ($Detailed -and $allResults.Count -gt 0) {
            Write-Host "Processing detailed information for $($allResults.Count) protection units in workload $currentWorkload" -ForegroundColor Cyan
            $pageCountUser = 0
            foreach ($unit in $allResults) {
                Write-Host "Fetching details for $currentWorkload protection unit, index: $pageCountUser" -ForegroundColor Green
                if ($unit.PSObject.Properties["siteId"] -and $unit.siteId) {
                    Get-SiteDetails -unit $unit -accessToken $accessToken
                } else {
                    Get-UserDetails -unit $unit -accessToken $accessToken
                }
                $pageCountUser++
            }
            # Clear allResults for next workload
            $allResults = @()
        }
        
        $workloadResultsTable[$currentWorkload] = $pages
    }
}

# Create files for processed results
if ($UseProtectionPolicies) {
    $anyWritten = $false
    foreach ($policyKey in $policyPagesTable.Keys) {
        $policyData = $policyPagesTable[$policyKey]
        $pages = $policyData.pages
        $policyHeaders = $policyData.policyHeaders
        $fullPolicyData = $policyData.policyData
        
        # Check if there are any values in any page
        $hasValues = $false
        foreach ($page in $pages) {
            if ($page.values.Count -gt 0) { $hasValues = $true; break }
        }
        
        # Create file if it has values OR if CreateEmptyPolicyFiles is true
        if ($hasValues -or $CreateEmptyPolicyFiles) {
            # No need to add script_internal index anymore
            
            $workloadType, $policyId = $policyKey -split '_', 2
            $outputObject = @{
                info = @{ 
                    policy = $fullPolicyData
                    timestamp = $currentTimestamp
                    policyHeaders = $policyHeaders
                }
                pages = $pages
            }
            $outputFileName = "result_${workloadType}_${policyId}_$currentTimestamp.json"
            $outputPath = Join-Path -Path "." -ChildPath $outputFileName
            $outputObject | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputPath -Encoding utf8
            
            $itemCount = ($pages | ForEach-Object { $_.values.Count } | Measure-Object -Sum).Sum
            if ($itemCount -eq 0) {
                Write-Host "Data for $workloadType policy $policyId written to $outputPath (Empty policy - no protection units)" -ForegroundColor Yellow
            } else {
                Write-Host "Data for $workloadType policy $policyId written to $outputPath (Pages: $($pages.Count), Items: $itemCount)" -ForegroundColor Green
            }
            $anyWritten = $true
        }
    }
    if (-not $anyWritten) {
        Write-Host "No protection units were processed. No file created." -ForegroundColor Yellow
    }
} else {
    # For non-policy queries, output one file per workload, with page structure and headers
    if ($workloadResultsTable.Count -gt 0) {
        foreach ($workloadName in $workloadResultsTable.Keys) {
            $pages = $workloadResultsTable[$workloadName]
            # No need to add script_internal index anymore
            $outputObject = @{
                info = @{ workload = $workloadName; timestamp = $currentTimestamp }
                pages = $pages
            }
            $outputFileName = "result_${workloadName}_$currentTimestamp.json"
            $outputPath = Join-Path -Path "." -ChildPath $outputFileName
            $outputObject | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputPath -Encoding utf8
            
            $totalItems = ($pages | ForEach-Object { $_.values.Count } | Measure-Object -Sum).Sum
            Write-Host "Data for workload $workloadName written to $outputPath (Pages: $($pages.Count), Items: $totalItems)" -ForegroundColor Green
        }
    } else {
        Write-Host "No protection units were processed. No file created." -ForegroundColor Yellow
    }
}