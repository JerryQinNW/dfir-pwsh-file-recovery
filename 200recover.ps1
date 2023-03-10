<#
    PowerShell for Digital Forensics

    Case 200: File Recovery
    Recover deleted files from a suspect's computer

    Author:  Denise Case
    Date:    2023-03-09

#>

# Clear the terminal window
Clear-Host

# Write header to the console
Write-Output ""
Write-Output "------------------------------------------------------"
Write-Output "Starting CASE 200 RECOVER DELETED EVIDENCE script"
Write-Output "------------------------------------------------------"
Write-Output "In a PowerShell terminal, run:"
Write-Output ".\200recover.ps1"
Write-Output ""
Write-Output "------------------------------------------------------"

$evidencePath = ".\200_evidence"
$evidencePathNoSlash = "200_evidence"
$recoveredFilesPath = "$evidencePath-recovered"
$reportedFilesPath = "$evidencePath-reported"
$reportFolder = "$evidencePathNoSlash-reported"
$currentDir = Get-Location

if (-not (Test-Path $recoveredFilesPath)) {
    New-Item -ItemType Directory -Path $recoveredFilesPath | Out-Null
}
if (-not (Test-Path $reportedFilesPath)) {
    New-Item -ItemType Directory -Path $reportedFilesPath | Out-Null
}

# Set the path to the recovered files word temporary folder
$wordFolderPath = Join-Path "C:\Users" $env:USERNAME "AppData\Roaming\Microsoft\Word"

# Set the path to the recovered files excel temporary folder
$excelFolderPath = Join-Path $env:USERPROFILE "AppData\Roaming\Microsoft\Excel"

# Get all the deleted Word files in the specified time range
$wordDeletedFiles = Get-ChildItem -Path $wordFolderPath -Filter "*.asd" -Recurse 
| Where-Object { $_.LastWriteTime -ge (Get-Date).AddDays(-7) }

# Get all the deleted Excel files in the specified time range and append them to the Word deleted files array
$excelDeletedFiles = Get-ChildItem -Path $excelFolderPath -Filter "*.xl*" -Recurse 
| Where-Object { $_.LastWriteTime -ge (Get-Date).AddDays(-7) }

Write-Output("Found $($wordDeletedFiles.Count) deleted files in the last 7 days (Word)")
Write-Output("Found $($excelDeletedFiles.Count) deleted files in the last 7 days (Excel)")
Write-Output "------------------------------------------------------"
Write-Output ""

try {
    $word = New-Object -ComObject Word.Application
    $excel = New-Object -ComObject Excel.Application
    $recoverySummaryContent += "========================`r`n"
    $recoverySummaryContent += "Recovered Word Documents`r`n"
    $recoverySummaryContent += "========================`r`n"
    $recoverySummaryContent += "`r`n"

    foreach ($file in $wordDeletedFiles) {
        $wordDoc = $word.Documents.Open($file.FullName, $false, $true)
        if ($null -ne $wordDoc) { 
            $text = $wordDoc.Content.Text
            $wordDoc.Close() 
        }
        $recoverySummaryContent += "RECOVERED: $($file.Name)`r`n`r`n"
        $recoverySummaryContent += "PATH: `r`n"
        $recoverySummaryContent += "$file`r`n`r`n"
        $recoverySummaryContent += "CONTENT:`r`n"
        $recoverySummaryContent += "$text`r`n`r`n"
        $destinationPath = Join-Path $recoveredFilesPath $file.Name
        if (!(Test-Path $destinationPath)) {
            Copy-Item $file.FullName $destinationPath -Force
        }
    }

    $recoverySummaryContent += "=========================`r`n"
    $recoverySummaryContent += "Recovered Excel Documents`r`n"
    $recoverySummaryContent += "=========================`r`n"
    $recoverySummaryContent += "`r`n"

    foreach ($file in $wordDeletedFiles) {
        $workbook = $excel.Workbooks.Open($newName, $false, $true)
        if ($null -ne $workbook) {
            $worksheet = $workbook.Worksheets.Item(1)
            $text = $worksheet.Cells.Item(1, 1).Value.ToString()
            $workbook.Close()
        }
        $recoverySummaryContent += "RECOVERED: $($file.Name)`r`n`r`n"
        $recoverySummaryContent += "PATH: `r`n"
        $recoverySummaryContent += "$file`r`n`r`n"
        $recoverySummaryContent += "CONTENT:`r`n"
        $recoverySummaryContent += "$text`r`n`r`n"
        $destinationPath = Join-Path $recoveredFilesPath $file.Name
        if (!(Test-Path $destinationPath)) {
            Copy-Item $file.FullName $destinationPath -Force
        }
    }
}
catch {
    Write-Error "Error occurred while processing file $($file.FullName): $_"
}
finally {
    # Clean up the COM objects
    if ($word -and -not [System.Runtime.InteropServices.Marshal]::IsComObject([System.__ComObject]$word)) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if ($excel -and -not [System.Runtime.InteropServices.Marshal]::IsComObject([System.__ComObject]$excel)) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$summaryFileName = "RecoverySummary$(Get-Date -Format 'yyyyMMdd').txt"
$summaryFilePath = Join-Path $currentDir $reportFolder $summaryFileName
Set-Content -Path $summaryFilePath -Value $recoverySummaryContent
Write-Output("Created summary file: $summaryFilePath")

$destinationPath = Join-Path $recoveredFilesPath $file.Name
Copy-Item $file.FullName $destinationPath -Force
