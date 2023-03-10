<#
    PowerShell for Digital Forensics

    Case 200: File Recovery
    Create example evidence files

    Author:  Denise Case
    Date:    2023-03-09

#>


# Clear the terminal window
Clear-Host

# Write header to the console
Write-Output ""
Write-Output "------------------------------------------------------"
Write-Output "Starting CASE 200 CREATE EVIDENCE script"
Write-Output "------------------------------------------------------"
Write-Output "In a PowerShell terminal, run:"
Write-Output ".\200create.ps1"
Write-Output ""
Write-Output "------------------------------------------------------"
Write-Output "Creating evidence... please be patient."
Write-Output "------------------------------------------------------"

# Create the CASE200\documents folder representing
# evidence siezed from the suspect's computer
$countOfFilesPerFolder = 3
$evidencePath = ".\200_evidence"
$evidencePathNoSlash = "200_evidence"

$folderNames = @("inbox", "sent", "drafts", "software", "jobs", "school", "hobby", "vault", "work", "home")
$subjects = @('Riley', 'Casey', 'she', 'he', 'Pat')
$nouns = @('car', 'house', 'book', 'computer', 'phone')
$verbs = @('run', 'jump', 'sing', 'read', 'write')
$adjectives = @('happy', 'sad', 'angry', 'tired', 'excited')
$adverbs = @('barely', 'slowly', 'loudly', 'quietly', 'happily')
$prepositions = "with", "on", "for", "to", "from", "by", "about", "of"
$folderMessage = @( "Hey, did you see this news article?", "Thanks for meeting with me today.", "Running late, can we reschedule?", "Got the goods, we leave tonight!",
  "This is a rough draft of my proposal.", "I need your help with a project.", "Sorry for the confusion, let me clarify.", "Goods are six Ducati Superleggera V4.",
  "I have a new idea I want to run by you.", "I think we need to adjust our approach." )

$folderIndex = 0
$currentDir = Get-Location

try {

  $word = New-Object -ComObject Word.Application
  $excel = New-Object -ComObject Excel.Application

  foreach ($folder in $folderNames) {
    $folderPath = Join-Path $evidencePath $folder
    if (-not (Test-Path -Path $folderPath -PathType Container)) {
      New-Item -ItemType Directory -Path $folderPath  | Out-Null
    }
    $folderIndex += 1
    
    for ($iFile = 1; $iFile -le $countOfFilesPerFolder; $iFile++) {
      $noun = $nouns | Get-Random
      $verb = $verbs | Get-Random
      $adjective = $adjectives | Get-Random
      $adverb = $adverbs | Get-Random
      $subject = $noun.Substring(0, 1).ToUpper() + $noun.Substring(1)
      $subject = $subjects | Get-Random
      $subject = $subject.Substring(0, 1).ToUpper() + $subject.Substring(1)
      $predicate = $adverb + " " + $verb + "s"
      $object = $adjective + " " + $noun
      $preposition = $prepositions | Get-Random
      switch (Get-Random -Minimum 0 -Maximum 5) {
        0 { $message = "$subject $predicate $preposition $object." }
        1 { $message = "$subject $predicate $object." }
        2 { $message = "$subject $verb $preposition $object." }
        3 { $message = "$subject $verb $object." }
        4 { $message = "$preposition $object $verb $subject." }
      }
      $message = $message.Substring(0, 1).ToUpper() + $message.Substring(1)

      if ($iFile -eq 2) {
        $fileName = "Document$($iFile*$folderIndex).docx"
        $filePath = Join-Path $currentDir $evidencePathNoSlash $folder $fileName
        if (Test-Path -Path $filePath -PathType Leaf) { Remove-Item $filePath }
        $document = $word.Documents.Add()
        #$document.AutoSaveOn = $true
        if ($folderIndex -eq 4) {
          $m = $folderMessage[$folderIndex - 1]
          $document.Content.Text = $m
        }
        else { $document.Content.Text = $message }
        $document.SaveAs2($filePath)
        $document.Close()
      }
      elseif ($iFile -eq 3) {
        $fileName = "Spreadsheet$($iFile*$folderIndex).xlsx"
        $filePath = Join-Path $currentDir $evidencePathNoSlash $folder $fileName
        if (Test-Path -Path $filePath -PathType Leaf) { Remove-Item $filePath }
        $workbook = $excel.Workbooks.Add()
        #$workbook.AutoSaveOn = $true
        $worksheet = $workbook.Worksheets.Item(1)
        if ($folderIndex -eq 8) {
          $m = $folderMessage[$folderIndex - 1]
          $worksheet.Cells.Item(1, 1) = $m
        }
        else { $worksheet.Cells.Item(1, 1) = $message }
        $workbook.SaveAs($filePath)
        $workbook.Close()
      }
      else {
        $fileName = "file$($iFile*$folderIndex).txt"
        $filePath = Join-Path $evidencePath $folder $fileName
        if (Test-Path -Path $filePath -PathType Leaf) { Remove-Item $filePath }
        New-Item -ItemType File -Path $filePath | Out-Null
        Set-Content -Path $filePath -Value $message
      } 
    }
    Write-Output "Creating evidence... please be patient."
    Write-Output "------------------------------------------------------"
  }
  $targetFolder1 = Join-Path $evidencePath "software"
  $targetFolder2 = Join-Path $evidencePath "vault"

  Get-ChildItem -Path $targetFolder1 -Directory | ForEach-Object {
    # Delete immediate children and the folder
    Get-ChildItem $_.FullName | Remove-Item -Force
    Remove-Item $_.FullName  -Force
    Write-Output "Folder '$($_.FullName)' deleted."
  }

  Get-ChildItem -Path $targetFolder2 -Directory | ForEach-Object {
    # Delete immediate children and the folder
    Get-ChildItem $_.FullName | Remove-Item -Force
    Remove-Item $_.FullName  -Force
    Write-Output "Folder '$($_.FullName)' deleted."
  }

  Remove-Item -Path $targetFolder1 -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -Path $targetFolder2 -Recurse -Force -ErrorAction SilentlyContinue
  Write-Output "------------------------------------------------------"
  Write-Output("Evidence created in $evidencePath")
  Write-Output "------------------------------------------------------"
}
catch {
  Write-Error "Error with COM apps"
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
