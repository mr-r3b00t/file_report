# PowerShell script to enumerate files and folders recursively using parallel processing and record details in a CSV file

# Requires PowerShell 7 or later for ForEach-Object -Parallel

# Get current timestamp for the CSV filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputCsv = "FileReport_$timestamp.csv"

# Prompt for the root directory to scan (default to current directory)
$rootPath = Read-Host "Enter the root directory to scan (press Enter for current directory)"
if ([string]::IsNullOrWhiteSpace($rootPath)) {
    $rootPath = Get-Location
}

# Resolve the path to ensure it's valid
$rootPath = Resolve-Path -Path $rootPath -ErrorAction Stop

# Define CSV headers
$headers = "Type,FileName,FolderName,FullPath,CreatedTimestamp,LastModifiedTimestamp,FileSizeBytes,MD5Hash,HasMarkOfTheWeb,DownloadURL"
$headers | Out-File -FilePath $outputCsv -Encoding UTF8

# Create a thread-safe collection for results
$results = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()

# Get the number of logical processors for throttle limit (default to core count)
$throttleLimit = [Environment]::ProcessorCount

# Recursively enumerate files and folders
$items = Get-ChildItem -Path $rootPath -Recurse -Force -ErrorAction SilentlyContinue
$items | ForEach-Object -Parallel {
    # Define helper functions within the parallel block to ensure availability
    function Get-MD5Hash {
        param ($filePath)
        try {
            $hash = Get-FileHash -Path $filePath -Algorithm MD5 -ErrorAction Stop
            return $hash.Hash
        }
        catch {
            return "N/A"
        }
    }

    function Get-MarkOfTheWeb {
        param ($filePath)
        try {
            $zoneInfo = Get-Content -Path "$filePath:Zone.Identifier" -ErrorAction Stop
            $hasMotW = $true
            # Look for ReferrerUrl or HostUrl in Zone.Identifier
            $downloadUrl = $zoneInfo | Where-Object { $_ -match "ReferrerUrl|HostUrl" }
            if ($downloadUrl) {
                $url = ($downloadUrl -split "=")[1].Trim()
            }
            else {
                $url = "N/A"
            }
            return $hasMotW, $url
        }
        catch {
            return $false, "N/A"
        }
    }

    try {
        $item = $_
        $type = if ($item.PSIsContainer) { "Folder" } else { "File" }
        $fileName = $item.Name
        $folderName = if ($item.PSIsContainer) { $item.Name } else { Split-Path $item.DirectoryName -Leaf }
        $fullPath = $item.FullName
        $createdTimestamp = $item.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
        $lastModifiedTimestamp = $item.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        $fileSizeBytes = if ($item.PSIsContainer) { "N/A" } else { $item.Length }

        # Initialize variables
        $md5Hash = "N/A"
        $hasMotW = $false
        $downloadUrl = "N/A"

        # If it's a file, compute MD5 and check for Mark of the Web
        if (-not $item.PSIsContainer) {
            $md5Hash = Get-MD5Hash -filePath $item.FullName
            $motwResult = Get-MarkOfTheWeb -filePath $item.FullName
            $hasMotW = $motwResult[0]
            $downloadUrl = $motwResult[1]
        }

        # Create result object
        $row = [PSCustomObject]@{
            Type = $type
            FileName = $fileName
            FolderName = $folderName
            FullPath = $fullPath
            CreatedTimestamp = $createdTimestamp
            LastModifiedTimestamp = $lastModifiedTimestamp
            FileSizeBytes = $fileSizeBytes
            MD5Hash = $md5Hash
            HasMarkOfTheWeb = $hasMotW
            DownloadURL = $downloadUrl
        }

        # Add to thread-safe collection
        ($using:results).Add($row)
    }
    catch {
        Write-Warning "Error processing $($item.FullName): $_"
    }
} -ThrottleLimit $throttleLimit

# Write all results to CSV
$results | Export-Csv -Path $outputCsv -Append -NoTypeInformation -Encoding UTF8

Write-Host "Enumeration complete. Results saved to $outputCsv"
