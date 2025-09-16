param (
    [Parameter(Mandatory=$true)]
    [string[]]$FilePaths,

    [Parameter(Mandatory=$false)]
    [switch]$AsJson = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$DebugOutput = $false
)

# Set output encoding to UTF-8 for consistency
$OutputEncoding = [System.Text.UTF8Encoding]::new($false) # $false for no BOM

$results = @{
    deleted = @()
    failed = @()
}

# Debug output
if ($DebugOutput) {
    Write-Host "DEBUG: RemoveScript Start" -ForegroundColor Cyan
    Write-Host "DEBUG: FilePaths Count: $($FilePaths.Count)" -ForegroundColor Cyan
    if ($FilePaths.Count -gt 0) {
        Write-Host "DEBUG: First path: $($FilePaths[0])" -ForegroundColor Cyan
    }
}

# Input validation
if ($FilePaths.Count -eq 0) {
    Write-Error "No file paths provided for deletion."
    if ($AsJson) {
        Write-Output ($results | ConvertTo-Json -Depth 3)
    } else {
        return $results
    }
    exit 1
}

foreach ($filePath in $FilePaths) {
    try {
        if ($DebugOutput) {
            Write-Host "DEBUG: Processing $filePath" -ForegroundColor Cyan
        }
        
        # Test if file exists before trying to get metadata
        if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
            $results.failed += @{ 
                Path = $filePath
                Error = "File not found or is not a regular file." 
            }
            continue
        }
        
        # Now we can safely get file info
        $fileObj = Get-Item -LiteralPath $filePath -ErrorAction Stop
        
        # Validate path is not a system path or critical area
        $invalidPaths = @(
            $env:windir, 
            "$env:windir\System32", 
            "$env:windir\System",
            "$env:ProgramFiles",
            "$env:ProgramFiles(x86)",
            "$env:ProgramData"
        )
        
        # Check if file is in a forbidden/system directory
        $isProtectedLocation = $false
        foreach ($protectedPath in $invalidPaths) {
            if ($fileObj.FullName.StartsWith($protectedPath)) {
                $isProtectedLocation = $true
                break
            }
        }
        
        if ($isProtectedLocation) {
            $results.failed += @{ 
                Path = $filePath
                Error = "Cannot delete files in system directories for security reasons." 
            }
            continue
        }
        
        # Make sure file matches one of Visio temporary file patterns
        $isVisioTempFile = $false
        $visioPatterns = @('~$*.vssx', '~$*.vsdx', '~$*.vstx', '~$*.vsdm', '~$*.vsd')
        
        foreach ($pattern in $visioPatterns) {
            if ($fileObj.Name -like $pattern) {
                $isVisioTempFile = $true
                break
            }
        }
        
        if (-not $isVisioTempFile) {
            $results.failed += @{ 
                Path = $filePath
                Error = "File does not match Visio temporary file pattern for safety." 
            }
            continue
        }
        
        # Now we can safely remove the file
        Remove-Item -LiteralPath $filePath -Force -ErrorAction Stop
        $results.deleted += $filePath
        
        if ($DebugOutput) {
            Write-Host "DEBUG: Successfully deleted $filePath" -ForegroundColor Green
        }
    }
    catch {
        $results.failed += @{ 
            Path = $filePath
            Error = $_.Exception.Message 
        }
        
        if ($DebugOutput) {
            Write-Host "DEBUG: Error with $filePath : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

if ($AsJson) {
    # Ensure clean JSON output
    Write-Output ($results | ConvertTo-Json -Depth 3)
} else {
    return $results
} 