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

# Helper Functions
function Get-FirstMeetingDate {
  <#
  .SYNOPSIS
  Calculates the first meeting date for a given day of week from a term start date.

  .PARAMETER termStartDate
  The start date of the term.

  .PARAMETER day
  The three-letter day abbreviation (mon, tue, wed, thu, fri, sat, sun).

  .OUTPUTS
  DateTime - The first meeting date.
  #>
  param(
    [datetime]$termStartDate,
    [string]$day
  )

  $days = @("mon", "tue", "wed", "thu", "fri", "sat", "sun")
  $dayIndex = $days.IndexOf($day)
  $termIndex = $days.IndexOf($termStartDate.DayOfWeek.ToString().ToLower().Substring(0, 3))
  $daysToAdd = ($dayIndex - $termIndex + 7) % 7
  if ($daysToAdd -eq 0 -and $termStartDate.DayOfWeek.ToString().ToLower().Substring(0, 3) -ne $day) {
    $daysToAdd = 7
  }
  return $termStartDate.AddDays($daysToAdd)
}

function Assert-ValidSection {
  <#
  .SYNOPSIS
  Validates that a section ID exists in the available sections.

  .PARAMETER sectionId
  The section ID to validate.

  .PARAMETER paramName
  Optional parameter name to include in error message (default: "sectionId").

  .OUTPUTS
  None - Throws an error if validation fails.
  #>
  param(
    [int]$sectionId,
    [string]$paramName = "sectionId"
  )

  if ($sections.sectionId -notcontains $sectionId) {
    $validIds = ($sections | ForEach-Object { "$($_.sectionId) ($($_.sectionName))" }) -join ", "
    Write-Error "❌ Not a valid $paramName. Valid options: $validIds" -ErrorAction Stop
  }
}

# Functions
function New-OsmParentRota {
  <#
  .SYNOPSIS
  Creates a parent rota for section meetings.

  .DESCRIPTION
  Generates a randomized parent rota assigning two parents per future meeting.
  The rota excludes leaders and members in the exclusion file. Output is saved
  as an HTML file in the Downloads folder and optionally sent to a printer.

  .PARAMETER sectionId
  The OSM section ID to generate the rota for.

  .PARAMETER print
  Optional switch to send the output directly to the default printer.

  .EXAMPLE
  New-OsmParentRota -sectionId 12345

  .EXAMPLE
  New-OsmParentRota -sectionId 12345 -print
  #>
  param (
    [int]$sectionId,
    [switch]$print
  )

  Assert-ValidSection -sectionId $sectionId
  
  $section = $sections | Where-Object { $_.sectionId -eq $sectionId }
  $termId = $section.termId
  $termName = $section.termName
  $sectionName = $section.sectionName
  $sectionNameFile = $sectionName.Replace(" ", "_").ToLower()

  # Members
  Write-Host "🔄 Fetching member list from OSM..."
  $membersListUrl = $membersListUrl + "&sectionid=$sectionId&termid=$termId"
  $membersList = (Invoke-OsmApi -url $membersListUrl).items
  Write-Host "✅ Retrieved $($membersList.Count) members"
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
  Write-Host "🔄 Fetching programme summary from OSM..."
  $programmeSummaryUrl = $programmeSummaryUrl + "&sectionid=$sectionId&termid=$termId"
  $programmeSummary = (Invoke-OsmApi -url $programmeSummaryUrl).items
  $futureMeetings = $programmeSummary | Where-Object { [datetime]$_.meetingdate -gt (Get-Date) }
  Write-Host "✅ Found $($futureMeetings.Count) future meetings"

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
  $htmlParams = @{
    Head = $htmlStyle
    Title = "$sectionName Parent Rota"
    PreContent = "<h1>$sectionName parent rota for $termName</h1>"
  }
  $outputFile = "$downloadsPath\parent_rota_$sectionNameFile.html"
  $assignments | ConvertTo-Html @htmlParams | Out-File $outputFile

  if ($print) {
    Out-Printer -Name $outputFile
  }

  Write-Host "✅ Parent rota saved to $outputFile"
  return $assignments
}
function Get-OsmPaperRegister {
  <#
  .SYNOPSIS
  Downloads a paper register PDF from OSM.

  .DESCRIPTION
  Downloads a paper register for future meetings from Online Scout Manager.
  The register is saved as a PDF in the Downloads folder and can optionally
  be sent to the default printer. The sort order can be customized.

  .PARAMETER sectionId
  The OSM section ID to download the register for.

  .PARAMETER order
  The sort order for the register. Valid values: firstname, lastname, dob, patrolid.
  Default is patrolid.

  .PARAMETER print
  Optional switch to send the PDF directly to the default printer.

  .EXAMPLE
  Get-OsmPaperRegister -sectionId 12345

  .EXAMPLE
  Get-OsmPaperRegister -sectionId 12345 -order lastname -print
  #>
  param (
    [int]$sectionId,
    [ValidateSet("firstname", "lastname", "dob", "patrolid")]
    [string]$order = "patrolid",
    [switch]$print
  )

  Assert-ValidSection -sectionId $sectionId
  
  $section = $sections | Where-Object { $_.sectionId -eq $sectionId }
  $termId = $section.termId
  $sectionName = $section.sectionName
  $sectionNameFile = $sectionName.Replace(" ", "_").ToLower()
  $printRegisterUrl = $printRegisterUrl + "&sectionid=$sectionId&termid=$termId"

  # Set site preferences
  Write-Host "🔄 Setting register sort preferences..."
  $body = @{
    preference = "sort"
    value      = $order
  }
  $sortOrder = Invoke-OsmApi -url $accountPreferences -Method "POST" -Body $body

  # Download register
  Write-Host "🔄 Downloading register PDF..."
  $outputFile = "$downloadsPath\paper_register_$sectionNameFile.pdf"
  Invoke-OsmApi -url $printRegisterUrl -method "DOWNLOAD" -file $outputFile

  if ($print) {
    Out-Printer -Name $outputFile
  }

  Write-Host "✅ Register downloaded to $outputFile"
  return [PSCustomObject]@{
    FilePath = $outputFile
    SectionName = $sectionName
    SortOrder = $order
  }
}
function New-OsmMeetings {
  <#
  .SYNOPSIS
  Creates weekly meetings for a section's current term.

  .DESCRIPTION
  Creates a series of weekly meetings in OSM starting from the first occurrence
  of the specified day of week within the term, and repeating weekly until the
  term end date. All meetings are initially created with the title "Planning...".

  .PARAMETER sectionId
  The OSM section ID to create meetings for.

  .PARAMETER day
  The three-letter day abbreviation for meetings. Valid values: mon, tue, wed, thu, fri.

  .EXAMPLE
  New-OsmMeetings -sectionId 12345 -day tue

  .EXAMPLE
  New-OsmMeetings -sectionId 12345 -day fri
  #>
  param (
    [int]$sectionId,
    [ValidateSet("mon", "tue", "wed", "thu", "fri")]
    [string]$day
  )

  Assert-ValidSection -sectionId $sectionId

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
  $firstMeetingDate = Get-FirstMeetingDate -termStartDate $termStartDate -day $day

  Write-Output "Term Start Date: $($termStartDate.ToString('dd-MM-yyyy'))"
  Write-Output "Term End Date: $($termEndDate.ToString('dd-MM-yyyy'))"
  Write-Output "Term Start Day: $termStartDay"
  Write-Output "Selected Day: $day"
  Write-Output "First Meeting Date: $($firstMeetingDate.ToString('dd-MM-yyyy'))"

  # Create meetings
  Write-Host "🔄 Creating meetings in OSM..."
  $body = @{
    sectionid = $sectionId
    title     = "Planning..."
    start     = $firstMeetingDate.ToString('yyyy-MM-dd')
    end       = $termEndDate.ToString('yyyy-MM-dd')
    repeat    = 7
  }
  $meetings = Invoke-OsmApi -url $programmeAddMeetingUrl -Method "POST" -Body $body
  Write-Host "✅ Meetings created for $termName"

  return [PSCustomObject]@{
    SectionName = $section.sectionName
    TermName = $termName
    Day = $day
    FirstMeetingDate = $firstMeetingDate
    TermEndDate = $termEndDate
    MeetingsCreated = $meetings
  }
}
function Copy-OsmMeetings {
  <#
  .SYNOPSIS
  Copies meetings from one section to another section.

  .DESCRIPTION
  Shares the programme from a source section and accepts it into a target section.
  The meetings are scheduled starting from the first occurrence of the specified
  day of week in the target section's term.

  .PARAMETER fromSectionId
  The OSM section ID to copy meetings from (source section).

  .PARAMETER toSectionId
  The OSM section ID to copy meetings to (target section).

  .PARAMETER day
  The three-letter day abbreviation for meetings in the target section.
  Valid values: mon, tue, wed, thu, fri.

  .EXAMPLE
  Copy-OsmMeetings -fromSectionId 12345 -toSectionId 67890 -day wed

  .EXAMPLE
  Copy-OsmMeetings -fromSectionId 12345 -toSectionId 67890 -day thu
  #>
  param (
    [int]$fromSectionId,
    [int]$toSectionId,
    [ValidateSet("mon", "tue", "wed", "thu", "fri")]
    [string]$day
  )

  Assert-ValidSection -sectionId $fromSectionId -paramName "fromSectionId"
  Assert-ValidSection -sectionId $toSectionId -paramName "toSectionId"

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
  $toFirstMeetingDate = Get-FirstMeetingDate -termStartDate $toTermStartDate -day $day

  Write-Output "Source Section Name: $fromSectionName"
  Write-Output "Target Section Name: $toSectionName"
  Write-Output "Selected Day: $day"
  Write-Output "Target Section First Meeting Date: $($toFirstMeetingDate.ToString('dd-MM-yyyy'))"

  # Copy meetings
  Write-Host "🔄 Sharing programme from source section..."
  $share = Invoke-OsmApi -url $programmeShareUrl -Method "GET"
  Write-Host "🔄 Accepting shared programme in target section..."
  $body = @{
    startdate = $toFirstMeetingDate.ToString('yyyy-MM-dd')
    starttime = $null
    endtime   = $null
  }
  $shareAccept = Invoke-OsmApi -url $programmeShareAcceptUrl -Method "POST" -Body $body
  Write-Host "✅ Meetings copied for $fromTermName"

  return [PSCustomObject]@{
    FromSection = $fromSectionName
    ToSection = $toSectionName
    TermName = $fromTermName
    Day = $day
    FirstMeetingDate = $toFirstMeetingDate
    ShareResult = $shareAccept
  }
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