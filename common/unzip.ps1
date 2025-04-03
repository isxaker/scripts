param (
    [string]$InputDirectory,    # Root folder containing .zip files
    [string]$OutputDirectory    # Folder for extracted files
)

# Validate input directory
if (-Not (Test-Path -Path $InputDirectory)) {
    Write-Error "The input directory '$InputDirectory' does not exist."
    exit
}

# Validate or create output directory
if (-Not (Test-Path -Path $OutputDirectory)) {
    try {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
        Write-Output "Created output directory '$OutputDirectory'."
    }
    catch {
        Write-Error "Cannot create output directory '$OutputDirectory'."
        exit
    }
}

# Global variable to track total number of processed items
$script:TotalProcessed = 0

function Extract-ZipFilesRecursively {
    param (
        [string]$SourceFolder,    # Folder containing zip files
        [string]$DestinationFolder # Folder to extract files into
    )

    # Find all zip files in the current folder (excluding the output directory)
    $zipFiles = Get-ChildItem -Path $SourceFolder -Recurse -File -Filter *.zip

    foreach ($zipFile in $zipFiles) {
        # Create a folder for the extracted contents
        $extractFolder = Join-Path -Path $DestinationFolder -ChildPath $zipFile.BaseName
        if (-Not (Test-Path -Path $extractFolder)) {
            New-Item -Path $extractFolder -ItemType Directory -Force | Out-Null
        }

        # Extract the zip file
        Write-Output "Extracting $($zipFile.FullName) to $($extractFolder)"
        try {
            Expand-Archive -Path $zipFile.FullName -DestinationPath $extractFolder -Force
            $script:TotalProcessed++ # Increment the counter for successfully processed item
        }
        catch {
            Write-Error "Failed to extract $($zipFile.FullName): $($_.Exception.Message)"
        }

        # Check for nested zip files in the extracted folder
        Extract-ZipFilesRecursively -SourceFolder $extractFolder -DestinationFolder $extractFolder
    }

    # Clean up: Remove .zip files from the output directory
    Get-ChildItem -Path $DestinationFolder -Recurse -File -Filter *.zip | Remove-Item -Force
}

# Start the extraction process
Extract-ZipFilesRecursively -SourceFolder $InputDirectory -DestinationFolder $OutputDirectory

Write-Output "All zip files have been extracted."
Write-Output "Total number of zip files processed: $script:TotalProcessed"