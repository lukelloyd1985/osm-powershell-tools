# ==========================================================
# Online Scout Manager (OSM) Tools
# ==========================================================

# Settings
$shell = New-Object -ComObject Shell.Application
$downloads = $shell.Namespace('shell:Downloads')
$downloadsPath = $downloads.Self.Path

# Import OSM API
Import-Module .\Osm.PowerShell.Api.ps1

# Functions
function New-OsmParentRota {
  param (
    [int]$sectionId,
    [switch]$print
  )
  if ($sectionTerm.sectionId -notcontains $sectionId) {
    Write-Error "❌ Not a valid sectionId" -ErrorAction Stop
  }
  $sectionName = ($sectionTerm | Where-Object { $_.sectionId -eq $sectionId }).sectionName
  $sectionNameFile = $sectionName.Replace(" ", "_").ToLower()
  Write-Output $sectionName | Out-File $downloadsPath\rota_$sectionNameFile.txt
  if ($print) {
    Get-Content $downloadsPath\rota_$sectionNameFile.txt | Out-Printer
  }
}

# Main
$sectionTerm = @()
$terms = Invoke-OsmApi -url $termsUrl
$userRoles = Invoke-OsmApi -url $userRolesUrl
$userRoles | ForEach-Object {
  $sectionId = $_.sectionid
  $sectionName = $_.sectionname
  $thisTerm = $terms.$sectionId | Where-Object { (Get-Date $_.enddate) -gt (Get-Date) }
  $sectionTerm += [PSCustomObject]@{
    sectionId   = $sectionId
    sectionName = $sectionName
    termId      = $thisTerm.termid
    termName    = $thisTerm.name
  }
}

Write-Output $sectionTerm | Format-Table -AutoSize