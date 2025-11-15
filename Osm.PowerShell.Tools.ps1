# ==========================================================
# Online Scout Manager (OSM) Tools
# ==========================================================

# Settings
$shell = New-Object -ComObject Shell.Application
$downloads = $shell.Namespace('shell:Downloads')
$downloadsPath = $downloads.Self.Path
$htmlStyle = @'
<style>
table {
  width: 600px;
  border-collapse: collapse;
  border-width: 2px;
  border-style: solid;
  border-color: black;
  color: black;
  font-size: 24px;
  text-align: center;
}

th {
  background-color: #0000ff;
  color: white;
}
</style>
'@

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
  $termName = $section.termName
  $sectionName = $section.sectionName
  $sectionNameFile = $sectionName.Replace(" ", "_").ToLower()
  
  # Members
  $membersListUrl = $membersListUrl + "&sectionid=$sectionId&termid=$termId"
  $membersList = (Invoke-OsmApi -url $membersListUrl).items
  $excludeMembers = Get-Content $downloadsPath\exclude_$sectionNameFile.txt -ErrorAction SilentlyContinue
  $filteredMembers = $membersList | Sort-Object lastname -Unique | Where-Object { $excludeMembers -notcontains $_.lastname -and $_.patrolid -gt 0 }
  $initials = foreach ($member in $filteredMembers) {
    $fname = $member.firstname
    $lname = $member.lastname
    $finit = ($fname[0].ToString().ToUpper() + $fname[1].ToString().ToLower())
    if ($lname -match "-") {
      $linit = ($lname -split "-" | ForEach-Object { $_[0].ToString().ToUpper() }) -join "-"
    } else {
      $linit = $lname[0].ToString().ToUpper()
    }
    "$finit$linit"
  }

  # Programme
  $programmeSummaryUrl = $programmeSummaryUrl + "&sectionid=$sectionId&termid=$termId"
  $programmeSummary = (Invoke-OsmApi -url $programmeSummaryUrl).items
  $futureMeetings = $programmeSummary | Where-Object { [datetime]$_.meetingdate -gt (Get-Date) }

  # Randomly assign 2 members initials per meeting (with no re-use)
  $shuffledInitials = Get-Random -InputObject $initials -Count $initials.Count
  $assignments = @()
  $index = 0
  foreach ($meeting in $futureMeetings) {
    $dateUK = (Get-Date $meeting.meetingdate -Format "dd-MM-yyyy")
    if ($shuffledInitials.Count -ge 2) {
      $assigned = $shuffledInitials[0..1]
      $shuffledInitials = $shuffledInitials[2..($shuffledInitials.Count-1)]
      $assignedText = ($assigned -join " & ")
    } else {
      $assignedText = "None"
    }
    $assignments += [PSCustomObject]@{
      Date     = $dateUK
      Title    = $meeting.title
      Assigned = $assignedText
    }
    $index++
  }

  # Output rota
  $assignments | Format-Table -AutoSize
  $htmlParams = @{
    Head = $htmlStyle
    Title = "$sectionName Parent Rota"
    PreContent = "<h1>$sectionName parent rota for $termName</h1>"
  }
  $assignments | ConvertTo-Html @htmlParams | Out-File $downloadsPath\parent_rota_$sectionNameFile.html

  if ($print) {
    Get-Content $downloadsPath\parent_rota_$sectionNameFile.html | Out-Printer
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