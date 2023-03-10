<#
    PowerShell for Digital Forensics

    Case 200: File Recovery
    Remove example files

    Author:  Denise Case
    Date:    2023-03-09

#>

function DeleteEvidenceFolder($evPath) {
  if (Test-Path $evPath -PathType Container) {
    $folderNames = @("inbox", "sent", "drafts", "software", "jobs", "school", "hobby", "vault", "work", "home")
      # Loop through the folders in this evPath
      Get-ChildItem -Path $evPath -Directory | ForEach-Object {
          $name = $_.Name
          $isOurFolder = $name -in $folderNames
          if ($isOurFolder) {
              # Delete immediate children and the folder
              Get-ChildItem $_.FullName | Remove-Item -Force
              Remove-Item $_.FullName  -Force
              Write-Output "Folder '$($_.FullName)' deleted."
          }
      } 
      Remove-Item $evPath -Force
  }
}

Clear-Host

Write-Output ""
Write-Output "------------------------------------------------------"
Write-Output "Starting CASE 200 CREATE EVIDENCE script"
Write-Output "------------------------------------------------------"
Write-Output "In a PowerShell terminal, run:"
Write-Output ".\200create.ps1"
Write-Output ""
Write-Output "------------------------------------------------------"

$evidencePath = ".\200_evidence"
$folderNames = @("inbox", "sent", "drafts", "software", "jobs", "school", "hobby", "vault", "work", "home")
$recoveredFilesPath = "$evidencePath-recovered"
$reportedFilesPath = "$evidencePath-reported"

DeleteEvidenceFolder  $evidencePath 
DeleteEvidenceFolder  $recoveredFilesPath
DeleteEvidenceFolder  $reportedFilesPath 
Write-Output "Done."
Write-Output "------------------------------------------------------"
