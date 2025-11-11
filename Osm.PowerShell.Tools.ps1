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
  if ($sections.sectionId -notcontains $sectionId) {
    Write-Error "❌ Not a valid sectionId" -ErrorAction Stop
  }
  $section = $sections | Where-Object { $_.sectionId -eq $sectionId }
  $termId = $section.termId
  $sectionName = $section.sectionName
  $sectionNameFile = $sectionName.Replace(" ", "_").ToLower()
  $membersListUrl = $membersListUrl + "&sectionid=$sectionId&termid=$termId"
  $membersList = (Invoke-OsmApi -url $membersListUrl).items
  # $membersList | Out-File $downloadsPath\rota_$sectionNameFile.json
  if ($print) {
    # commenting out until file
    # Get-Content $downloadsPath\rota_$sectionNameFile.html | Out-Printer
  }
}

# Main
$sections = @()
$terms = Invoke-OsmApi -url $termsUrl
$userRoles = Invoke-OsmApi -url $userRolesUrl
$userRoles | ForEach-Object {
  $sectionId = $_.sectionid
  $sectionName = $_.sectionname
  $thisTerm = $terms.$sectionId | Where-Object { (Get-Date $_.enddate) -gt (Get-Date) }
  $sections += [PSCustomObject]@{
    sectionId   = $sectionId
    sectionName = $sectionName
    termId      = $thisTerm.termid
    termName    = $thisTerm.name
  }
}

Write-Output $sections | Format-Table -AutoSize