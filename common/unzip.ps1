param (
    [string]$InputDirectory = (Get-Location),             # Default to the current directory if not specified
    [string]$DestinationDirectory = $null                 # Optional, if not specified, defaults to InputDirectory
)

# Validate input directory
if (-Not (Test-Path -Path $InputDirectory)) {
    Write-Error "The input directory '$InputDirectory' does not exist."
    exit
}

# Set DestinationDirectory to InputDirectory if not specified
if (-not $DestinationDirectory) {
    $DestinationDirectory = $InputDirectory
}

# Validate that the DestinationDirectory can be accessed or created
if (-Not (Test-Path -Path $DestinationDirectory)) {
    try {
        New-Item -Path $DestinationDirectory -ItemType Directory -ErrorAction Stop
        Write-Output "Created destination directory at '$DestinationDirectory'."
    }
    catch {
        Write-Error "Cannot create destination directory '$DestinationDirectory'."
        exit
    }
}

# Find all .zip files in the input directory
$zipFiles = Get-ChildItem -Path $InputDirectory -Filter *.zip

# Loop through each zip file and extract it
foreach ($zipFile in $zipFiles) {
    # Unzip the file's base name
    $extractFolder = $zipFile.BaseName

    # Create the full extraction path
    $extractPath = Join-Path -Path $DestinationDirectory -ChildPath $extractFolder

    # Create the extraction directory if it does not exist
    if (-Not (Test-Path -Path $extractPath)) {
        New-Item -Path $extractPath -ItemType Directory
    }

    # Unzip the file
    Write-Output "Extracting $($zipFile.Name) to $($extractPath)"
    Expand-Archive -Path $zipFile.FullName -DestinationPath $extractPath -Force
}

Write-Output "All zip files have been extracted."