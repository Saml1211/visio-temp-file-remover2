# Test script for visio-temp-file-remover API

# Function to test API endpoints
function Test-API {
    Write-Host "Starting server..." -ForegroundColor Cyan
    $serverJob = Start-Job -ScriptBlock {
        # Use current directory instead of hardcoded path
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
        cd $scriptDir
        node app.js
    }
    
    # Wait for server to start
    Start-Sleep -Seconds 3
    
    try {
        # Test scan API
        Write-Host "`nTesting scan API..." -ForegroundColor Cyan
        $scanResponse = Invoke-RestMethod -Uri 'http://localhost:3000/api/scan' -Method POST -ContentType 'application/json' -Body '{"directory":"Z:\\ENGINEERING TEMPLATES\\VISIO SHAPES 2025"}'
        Write-Host "Scan API Response:" -ForegroundColor Green
        $scanResponse | ConvertTo-Json -Depth 5
        
        # If files were found, test the delete API with the first file
        if ($scanResponse.files -and $scanResponse.files.Count -gt 0) {
            $sampleFile = $scanResponse.files[0].FullName
            Write-Host "`nFound $($scanResponse.files.Count) file(s). First file: $sampleFile" -ForegroundColor Yellow
            
            $deleteTestPrompt = Read-Host "Do you want to test file deletion for this file? (y/n)"
            if ($deleteTestPrompt -eq 'y') {
                Write-Host "`nTesting delete API..." -ForegroundColor Cyan
                $deleteBody = @{
                    files = @($sampleFile)
                } | ConvertTo-Json
                
                $deleteResponse = Invoke-RestMethod -Uri 'http://localhost:3000/api/delete' -Method POST -ContentType 'application/json' -Body $deleteBody
                Write-Host "Delete API Response:" -ForegroundColor Green
                $deleteResponse | ConvertTo-Json
            }
        } else {
            Write-Host "No files found to test deletion." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error testing API: $_" -ForegroundColor Red
    }
    finally {
        # Stop the server
        Write-Host "`nStopping server..." -ForegroundColor Cyan
        Stop-Job -Job $serverJob
        Remove-Job -Job $serverJob -Force
    }
}

# Run the test
Test-API

Write-Host "`nTests completed." -ForegroundColor Cyan
