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
  $leadersSurnames = ($membersList | Where-Object { $_.patrolid -lt 0 }).lastname
  $filteredMembers = $membersList | Sort-Object lastname -Unique | Where-Object { $excludeMembers -notcontains $_.lastname -and $leadersSurnames -notcontains $_.lastname -and $_.patrolid -gt 0 }
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

  # Randomly assign 2 members initials per meeting (with no re-use, replenishing when empty)
  $shuffledInitials = Get-Random -InputObject $initials -Count $initials.Count
  $assignments = @()
  $index = 0
  foreach ($meeting in $futureMeetings) {
    $dateUK = (Get-Date $meeting.meetingdate -Format "dd-MM-yyyy")

    # Replenish shuffledInitials if we don't have enough for assignment
    if ($shuffledInitials.Count -lt 2) {
      $shuffledInitials = Get-Random -InputObject $initials -Count $initials.Count
    }

    $assigned = $shuffledInitials[0..1]
    $shuffledInitials = $shuffledInitials[2..($shuffledInitials.Count-1)]
    $assignedText = ($assigned -join " & ")

    $assignments += [PSCustomObject]@{
      Date     = $dateUK
      Title    = $meeting.title
      Assigned = $assignedText
    }
    $index++
  }

  # Output rota
  Write-Output $assignments | Format-Table -AutoSize
  $htmlParams = @{
    Head = $htmlStyle
    Title = "$sectionName Parent Rota"
    PreContent = "<h1>$sectionName parent rota for $termName</h1>"
  }
  $assignments | ConvertTo-Html @htmlParams | Out-File $downloadsPath\parent_rota_$sectionNameFile.html

  if ($print) {
    Out-Printer -Name "$downloadsPath\parent_rota_$sectionNameFile.html"
  }
}
function Get-OsmPaperRegister {
  param (
    [int]$sectionId,
    [ValidateSet("firstname", "lastname", "dob", "patrolid")]
    [string]$order = "patrolid",
    [switch]$print
  )
  
  if ($sections.sectionId -notcontains $sectionId) {
    Write-Error "❌ Not a valid sectionId" -ErrorAction Stop
  }
  
  $section = $sections | Where-Object { $_.sectionId -eq $sectionId }
  $termId = $section.termId
  $sectionName = $section.sectionName
  $sectionNameFile = $sectionName.Replace(" ", "_").ToLower()
  $printRegisterUrl = $printRegisterUrl + "&sectionid=$sectionId&termid=$termId"

  # Set site preferences
  $body = @{
    preference = "sort"
    value      = $order
  }
  $sortOrder = Invoke-OsmApi -url $accountPreferences -Method "POST" -Body $body

  # Download register
  Invoke-OsmApi -url $printRegisterUrl -method "DOWNLOAD" -file "$downloadsPath\paper_register_$sectionNameFile.pdf"
  Write-Output "✅ Register downloaded to $downloadsPath\paper_register_$sectionNameFile.pdf"

  if ($print) {
    Out-Printer -Name "$downloadsPath\paper_register_$sectionNameFile.pdf"
  }
}
function New-OsmMeetings {
  param (
    [int]$sectionId,
    [ValidateSet("mon", "tue", "wed", "thu", "fri")]
    [string]$day
  )

  if ($sections.sectionId -notcontains $sectionId) {
    Write-Error "❌ Not a valid sectionId" -ErrorAction Stop
  }

  if (!$day) {
    Write-Error "❌ Provide day of week for meetings" -ErrorAction Stop
  }

  $section = $sections | Where-Object { $_.sectionId -eq $sectionId }
  $termId = $section.termId
  $termName = $section.termName
  $thisTerm = $terms.$sectionId | Where-Object { $_.termid -eq $termId }
  $termStartDate = [datetime]$thisTerm.startdate
  $termEndDate = [datetime]$thisTerm.enddate
  $termStartDay = $termStartDate.DayOfWeek.ToString().ToLower().Substring(0, 3)

  # Get the first occurrence of $day from $termStartDate for $firstMeetingDate
  $days = @("mon", "tue", "wed", "thu", "fri", "sat", "sun")
  $dayIndex = $days.IndexOf($day)
  $termIndex = $days.IndexOf($termStartDay)
  $daysToAdd = ($dayIndex - $termIndex + 7) % 7
  if ($daysToAdd -eq 0 -and $termStartDay -ne $day) {
    $daysToAdd = 7
  }
  $firstMeetingDate = $termStartDate.AddDays($daysToAdd)

  Write-Output "Term Start Date: $($termStartDate.ToString('dd-MM-yyyy'))"
  Write-Output "Term End Date: $($termEndDate.ToString('dd-MM-yyyy'))"
  Write-Output "Term Start Day: $termStartDay"
  Write-Output "Selected Day: $day"
  Write-Output "First Meeting Date: $($firstMeetingDate.ToString('dd-MM-yyyy'))"

  # Create meetings
  $body = @{
    sectionid = $sectionId
    title     = "Planning..."
    start     = $firstMeetingDate.ToString('yyyy-MM-dd')
    end       = $termEndDate.ToString('yyyy-MM-dd')
    repeat    = 7
  }
  $meetings = Invoke-OsmApi -url $programmeAddMeetingUrl -Method "POST" -Body $body
  Write-Output "✅ Meetings created for $termName"
}
function Copy-OsmMeetings {
  param (
    [int]$fromSectionId,
    [int]$toSectionId,
    [ValidateSet("mon", "tue", "wed", "thu", "fri")]
    [string]$day
  )

  if ($sections.sectionId -notcontains $fromSectionId) {
    Write-Error "❌ Not a valid fromSectionId" -ErrorAction Stop
  }

  if ($sections.sectionId -notcontains $toSectionId) {
    Write-Error "❌ Not a valid toSectionId" -ErrorAction Stop
  }

  if (!$day) {
    Write-Error "❌ Provide day of week for meetings" -ErrorAction Stop
  }

  $fromSection = $sections | Where-Object { $_.sectionId -eq $fromSectionId }
  $fromSectionName = $fromSection.sectionName
  $fromTermId = $fromSection.termId
  $fromTermName = $fromSection.termName
  $toSection = $sections | Where-Object { $_.sectionId -eq $toSectionId }
  $toSectionName = $toSection.sectionName
  $toTermId = $toSection.termId
  $toThisTerm = $terms.$toSectionId | Where-Object { $_.termid -eq $toTermId }
  $toTermStartDate = [datetime]$toThisTerm.startdate
  $toTermStartDay = $toTermStartDate.DayOfWeek.ToString().ToLower().Substring(0, 3)
  $programmeShareUrl = $programmeShareUrl + "&sectionid=$fromSectionId&termid=$fromTermId&target=$toSectionId"
  $programmeShareAcceptUrl = $programmeShareAcceptUrl + "&sectionid=$toSectionId"

  # Get the first occurrence of $day from $toTermStartDate for $toFirstMeetingDate
  $days = @("mon", "tue", "wed", "thu", "fri", "sat", "sun")
  $dayIndex = $days.IndexOf($day)
  $toTermIndex = $days.IndexOf($toTermStartDay)
  $daysToAdd = ($dayIndex - $toTermIndex + 7) % 7
  if ($daysToAdd -eq 0 -and $toTermStartDay -ne $day) {
    $daysToAdd = 7
  }
  $toFirstMeetingDate = $toTermStartDate.AddDays($daysToAdd)

  Write-Output "Source Section Name: $fromSectionName"
  Write-Output "Target Section Name: $toSectionName"
  Write-Output "Selected Day: $day"
  Write-Output "Target Section First Meeting Date: $($toFirstMeetingDate.ToString('dd-MM-yyyy'))"

  # Copy meetings
  $share = Invoke-OsmApi -url $programmeShareUrl -Method "GET"
  $body = @{
    startdate = $toFirstMeetingDate.ToString('yyyy-MM-dd')
    starttime = $null
    endtime   = $null
  }
  $shareAccept = Invoke-OsmApi -url $programmeShareAcceptUrl -Method "POST" -Body $body
  Write-Output "✅ Meetings copied for $fromTermName"
}

# Main
$sections = @()
$terms = Invoke-OsmApi -url $termsUrl
$userRoles = Invoke-OsmApi -url $userRolesUrl
$userRoles | ForEach-Object {
  $sectionId = $_.sectionid
  $sectionName = $_.sectionname
  $thisTerm = $terms.$sectionId | Where-Object { (Get-Date $_.enddate) -gt (Get-Date) -and $_.past -eq "False" }
  $sections += [PSCustomObject]@{
    sectionId   = $sectionId
    sectionName = $sectionName
    termId      = $thisTerm.termid
    termName    = $thisTerm.name
  }
}

Write-Output $sections | Format-Table -AutoSize