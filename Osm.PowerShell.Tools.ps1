# ==========================================================
# Online Scout Manager (OSM) Tools
# ==========================================================

# Settings
$shell = New-Object -ComObject Shell.Application
$downloads = $shell.Namespace('shell:Downloads')
$downloadsPath = $downloads.Self.Path

# Load HTML style from external CSS file
$cssPath = "$PSScriptRoot\default.css"
if (Test-Path $cssPath) {
  $cssContent = Get-Content $cssPath -Raw
  $htmlStyle = "<style>`n$cssContent`n</style>"
} else {
  Write-Warning "⚠️ CSS file not found at $cssPath. Using minimal default style."
  $htmlStyle = "<style>table { border-collapse: collapse; } th { background-color: #0000ff; color: white; }</style>"
}

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

function Add-QueryParams {
  <#
  .SYNOPSIS
  Appends query parameters to a URL.

  .PARAMETER url
  The base URL to append parameters to.

  .PARAMETER params
  A hashtable of parameter names and values.

  .OUTPUTS
  String - The URL with appended parameters.
  #>
  param(
    [string]$url,
    [hashtable]$params
  )

  $queryString = ($params.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
  return "$url&$queryString"
}

function Get-MemberInitials {
  <#
  .SYNOPSIS
  Generates member initials from first and last name.

  .PARAMETER firstName
  The member's first name.

  .PARAMETER lastName
  The member's last name.

  .OUTPUTS
  String - The member's initials in the format "FiLi" (e.g., "JoS" for John Smith).
  #>
  param(
    [string]$firstName,
    [string]$lastName
  )

  $finit = ($firstName[0].ToString().ToUpper() + $firstName[1].ToString().ToLower())
  if ($lastName -match "-") {
    $linit = ($lastName -split "-" | ForEach-Object { $_[0].ToString().ToUpper() }) -join "-"
  } else {
    $linit = $lastName[0].ToString().ToUpper()
  }
  return "$finit$linit"
}

function Get-OsmMembers {
  <#
  .SYNOPSIS
  Retrieves member list from OSM for a given section and term.

  .PARAMETER sectionId
  The OSM section ID.

  .PARAMETER termId
  The OSM term ID.

  .OUTPUTS
  Array - The list of members from OSM.
  #>
  param(
    [int]$sectionId,
    [int]$termId
  )

  Write-Host "🔄 Fetching member list from OSM..."
  $membersListUrlWithParams = Add-QueryParams -url $membersListUrl -params @{
    sectionid = $sectionId
    termid = $termId
  }
  $membersList = (Invoke-OsmApi -url $membersListUrlWithParams).items
  Write-Host "✅ Retrieved $($membersList.Count) members"
  return $membersList
}

function ConvertTo-PdfForPrinting {
  <#
  .SYNOPSIS
  Converts an HTML file to PDF using Microsoft Edge or Google Chrome in headless mode.

  .PARAMETER htmlPath
  The path to the HTML file to convert.

  .PARAMETER pdfPath
  The path where the PDF should be saved.

  .OUTPUTS
  String - The path to the generated PDF file.
  #>
  param(
    [string]$htmlPath,
    [string]$pdfPath
  )

  Write-Host "🔄 Converting HTML to PDF using Microsoft Edge..."

  # Try Microsoft Edge first
  $edgePath = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
  if (-not (Test-Path $edgePath)) {
    $edgePath = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
  }

  # Fall back to Google Chrome
  $chromePath = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
  if (-not (Test-Path $chromePath)) {
    $chromePath = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
  }

  $browserPath = $null
  $browserName = $null

  if (Test-Path $edgePath) {
    $browserPath = $edgePath
    $browserName = "Microsoft Edge"
  } elseif (Test-Path $chromePath) {
    $browserPath = $chromePath
    $browserName = "Google Chrome"
  } else {
    throw "Neither Microsoft Edge nor Google Chrome found. Please install one of these browsers."
  }

  # Convert HTML to PDF using headless browser
  # Suppress stderr by redirecting it to null
  $process = Start-Process -FilePath $browserPath `
    -ArgumentList "--headless", "--disable-gpu", "--print-to-pdf=`"$pdfPath`"", "`"$htmlPath`"" `
    -NoNewWindow `
    -Wait `
    -PassThru `
    -RedirectStandardOutput "$env:TEMP\edge-stdout.txt" `
    -RedirectStandardError "$env:TEMP\edge-stderr.txt"

  # Clean up stderr temp file
  if (Test-Path "$env:TEMP\edge-stderr.txt") {
    Remove-Item "$env:TEMP\edge-stderr.txt" -Force -ErrorAction SilentlyContinue
  }

  # Clean up stdout temp file
  if (Test-Path "$env:TEMP\edge-stdout.txt") {
    Remove-Item "$env:TEMP\edge-stdout.txt" -Force -ErrorAction SilentlyContinue
  }

  if ($process.ExitCode -ne 0) {
    throw "Failed to convert HTML to PDF using $browserName (exit code: $($process.ExitCode))"
  }

  if (-not (Test-Path $pdfPath)) {
    throw "PDF file was not created at $pdfPath"
  }

  Write-Host "✅ PDF created successfully"
  return $pdfPath
}

function Invoke-PrintReport {
  <#
  .SYNOPSIS
  Prints one or more HTML reports by converting to PDF and sending to the default printer.

  .PARAMETER htmlPaths
  Array of HTML file paths to print.

  .PARAMETER pdfPaths
  Array of PDF file paths to create (must match htmlPaths length).

  .OUTPUTS
  None - Sends print jobs to the default printer.
  #>
  param(
    [string[]]$htmlPaths,
    [string[]]$pdfPaths
  )

  if ($htmlPaths.Count -ne $pdfPaths.Count) {
    throw "htmlPaths and pdfPaths must have the same length"
  }

  try {
    # Convert all HTML files to PDF
    for ($i = 0; $i -lt $htmlPaths.Count; $i++) {
      ConvertTo-PdfForPrinting -htmlPath $htmlPaths[$i] -pdfPath $pdfPaths[$i]
    }

    # Print all PDFs
    foreach ($pdfPath in $pdfPaths) {
      $printProcess = Start-Process -FilePath $pdfPath -Verb Print -PassThru

      # Wait for print dialog to appear, then close the PDF viewer
      Start-Sleep -Seconds 5
      if (-not $printProcess.HasExited) {
        if (-not $printProcess.CloseMainWindow()) {
          Stop-Process -Id $printProcess.Id -Force
        }
      }
    }

    Write-Host "✅ Print job sent successfully"
  } catch {
    Write-Warning "⚠️ Failed to print: $_"
  }
}

# Functions
function New-OsmParentRota {
  <#
  .SYNOPSIS
  Creates a parent rota for section meetings.

  .DESCRIPTION
  Generates a randomized parent rota assigning two parents per future meeting.
  The rota excludes leaders and members in the exclusion file. The assigned parents
  are automatically updated in OSM programme meetings (shown as full names on agendas).
  Output is also saved as an HTML file in the Downloads folder and optionally sent to a printer.

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
  $membersList = Get-OsmMembers -sectionId $sectionId -termId $termId

  # Check for exclusion file and provide guidance if missing
  $excludeFile = "$downloadsPath\exclude_$sectionNameFile.txt"
  if (Test-Path $excludeFile) {
    $excludeMembers = Get-Content $excludeFile
    Write-Host "ℹ️  Loaded exclusion file with $($excludeMembers.Count) excluded surnames"
  } else {
    Write-Host "ℹ️  No exclusion file found. Create $excludeFile to exclude members from rota."
    $excludeMembers = @()
  }
  $leadersSurnames = ($membersList | Where-Object { $_.patrolid -lt 0 }).lastname
  $filteredMembers = $membersList | Sort-Object lastname -Unique | Where-Object { $excludeMembers -notcontains $_.lastname -and $leadersSurnames -notcontains $_.lastname -and $_.patrolid -gt 0 }

  # Create member objects with both initials and full names for rota assignment
  $membersForRota = foreach ($member in $filteredMembers) {
    [PSCustomObject]@{
      Initials = Get-MemberInitials -firstName $member.firstname -lastName $member.lastname
      FullName = $member.full_name
    }
  }

  # Programme
  Write-Host "🔄 Fetching programme summary from OSM..."
  $programmeSummaryUrlWithParams = Add-QueryParams -url $programmeSummaryUrl -params @{
    sectionid = $sectionId
    termid = $termId
  }
  $programmeSummary = (Invoke-OsmApi -url $programmeSummaryUrlWithParams).items
  $futureMeetings = $programmeSummary | Where-Object { [datetime]$_.meetingdate -gt (Get-Date) }
  Write-Host "✅ Found $($futureMeetings.Count) future meetings"

  # Randomly assign 2 members per meeting (with no re-use, replenishing when empty)
  $shuffledMembers = Get-Random -InputObject $membersForRota -Count $membersForRota.Count
  $assignments = @()
  $updateCount = 0
  $updateErrors = 0

  Write-Host "🔄 Assigning parents to $($futureMeetings.Count) future meetings..."

  foreach ($meeting in $futureMeetings) {
    $dateUK = (Get-Date $meeting.meetingdate -Format "dd-MM-yyyy")

    # Replenish shuffledMembers if we don't have enough for assignment
    if ($shuffledMembers.Count -lt 2) {
      $shuffledMembers = Get-Random -InputObject $membersForRota -Count $membersForRota.Count
    }

    # Safely extract assignments
    if ($shuffledMembers.Count -ge 2) {
      $assigned = $shuffledMembers[0..1]
      if ($shuffledMembers.Count -gt 2) {
        $shuffledMembers = $shuffledMembers[2..($shuffledMembers.Count-1)]
      } else {
        $shuffledMembers = @()
      }
    } else {
      # Handle edge case with fewer than 2 members
      $assigned = $shuffledMembers
      $shuffledMembers = @()
    }

    $assignedInitials = ($assigned | ForEach-Object { $_.Initials }) -join " & "
    $assignedFullNames = ($assigned | ForEach-Object { $_.FullName }) -join " & "

    # Update the meeting in OSM with the assigned parents
    try {
      $body = @{
        sectionid = $sectionId
        eveningid = $meeting.eveningid
        adults = $assignedFullNames
      }
      $updateResult = Invoke-OsmApi -url $programmeUpdateUrl -Method "POST" -Body $body
      $updateCount++
    } catch {
      Write-Warning "⚠️ Failed to update meeting on $dateUK`: $_"
      $updateErrors++
    }

    $assignments += [PSCustomObject]@{
      Date     = $dateUK
      Title    = $meeting.title
      Assigned = $assignedInitials
    }
  }

  if ($updateCount -gt 0) {
    Write-Host "✅ Updated $updateCount meeting(s) in OSM with assigned parents"
  }
  if ($updateErrors -gt 0) {
    Write-Warning "⚠️ $updateErrors meeting(s) could not be updated in OSM"
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
    $pdfPath = "$downloadsPath\parent_rota_$sectionNameFile.pdf"
    Invoke-PrintReport -htmlPaths @($outputFile) -pdfPaths @($pdfPath)
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
  $printRegisterUrlWithParams = Add-QueryParams -url $printRegisterUrl -params @{
    sectionid = $sectionId
    termid = $termId
  }

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
  Invoke-OsmApi -url $printRegisterUrlWithParams -method "DOWNLOAD" -file $outputFile

  if ($print) {
    # Print PDF directly without conversion (already a PDF)
    try {
      $printProcess = Start-Process -FilePath $outputFile -Verb Print -PassThru

      # Wait for print dialog to appear, then close the PDF viewer
      Start-Sleep -Seconds 5
      if (-not $printProcess.HasExited) {
        if (-not $printProcess.CloseMainWindow()) {
          Stop-Process -Id $printProcess.Id -Force
        }
      }

      Write-Host "✅ Print job sent successfully"
    } catch {
      Write-Warning "⚠️ Failed to print: $_"
    }
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
  $programmeShareUrlWithParams = Add-QueryParams -url $programmeShareUrl -params @{
    sectionid = $fromSectionId
    termid = $fromTermId
    target = $toSectionId
  }
  $programmeShareAcceptUrlWithParams = Add-QueryParams -url $programmeShareAcceptUrl -params @{
    sectionid = $toSectionId
  }

  # Get the first occurrence of $day from $toTermStartDate for $toFirstMeetingDate
  $toFirstMeetingDate = Get-FirstMeetingDate -termStartDate $toTermStartDate -day $day

  Write-Output "Source Section Name: $fromSectionName"
  Write-Output "Target Section Name: $toSectionName"
  Write-Output "Selected Day: $day"
  Write-Output "Target Section First Meeting Date: $($toFirstMeetingDate.ToString('dd-MM-yyyy'))"

  # Copy meetings
  Write-Host "🔄 Sharing programme from source section..."
  $share = Invoke-OsmApi -url $programmeShareUrlWithParams -Method "GET"
  Write-Host "🔄 Accepting shared programme in target section..."
  $body = @{
    startdate = $toFirstMeetingDate.ToString('yyyy-MM-dd')
    starttime = $null
    endtime   = $null
  }
  $shareAccept = Invoke-OsmApi -url $programmeShareAcceptUrlWithParams -Method "POST" -Body $body
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
function Get-OsmPhotoConsent {
  <#
  .SYNOPSIS
  Retrieves photograph consent of members from OSM.

  .DESCRIPTION
  Generates a report of members with no photograph consent.
  Output is saved as an HTML file in the Downloads folder
  and optionally sent to a printer.

  .PARAMETER sectionId
  The OSM section ID to generate the report for.

  .PARAMETER print
  Optional switch to send the output directly to the default printer.

  .EXAMPLE
  Get-OsmPhotoConsent -sectionId 12345

  .EXAMPLE
  Get-OsmPhotoConsent -sectionId 12345 -print
  #>
  param (
    [int]$sectionId,
    [switch]$print
  )

  Assert-ValidSection -sectionId $sectionId
  
  $section = $sections | Where-Object { $_.sectionId -eq $sectionId }
  $termId = $section.termId
  $sectionName = $section.sectionName
  $sectionNameFile = $sectionName.Replace(" ", "_").ToLower()

  # Members
  $membersList = Get-OsmMembers -sectionId $sectionId -termId $termId

  # Members photograph consent
  Write-Host "🔄 Fetching members photograph consent from OSM..."
  $photoConsentFullname = @()
  $photoConsentInitials = @()
  foreach ($member in $membersList) {
    $memberFullname = $member.full_name
    $memberInitials = Get-MemberInitials -firstName $member.firstname -lastName $member.lastname
    $memberId = $member.scoutid
    $membersDataUrlWithParams = Add-QueryParams -url $membersDataUrl -params @{
      section_id = $sectionId
      associated_id = $memberId
      associated_type = "member"
      context = "members"
    }
    $memberData = (Invoke-OsmApi -url $membersDataUrlWithParams).data
    $memberPhotoConsent = (($memberData | where { $_.identifier -eq "consents" }).columns | where { $_.label -eq "Photographs" } | Select varname, value)
    $memberPhotoInt = ($memberPhotoConsent | where { $_.varname -eq "photographs_internal" }).value
    $memberPhotoExt = ($memberPhotoConsent | where { $_.varname -eq "photographs_tsa" }).value

    if ($memberPhotoInt -eq "No" -or $memberPhotoExt -eq "No") {
      $photoConsentFullname += [PSCustomObject]@{
        Name      = $memberFullname
        PhotosInt = $memberPhotoInt
        PhotosExt = $memberPhotoExt
      }
      $photoConsentInitials += [PSCustomObject]@{
        Name      = $memberInitials
        PhotosInt = $memberPhotoInt
        PhotosExt = $memberPhotoExt
      }
    }
  }
  Write-Host "✅ Retrieved $($photoConsentFullname.Count) members with no photograph consent"

  # Output report
  $htmlParams = @{
    Head = $htmlStyle
    Title = "$sectionName Photograph Consent"
    PreContent = "<h1>$sectionName Photograph Consent</h1>"
  }
  $outputFileFullname = "$downloadsPath\photograph_consent_fullname_$sectionNameFile.html"
  $outputFileInitials = "$downloadsPath\photograph_consent_initials_$sectionNameFile.html"
  $photoConsentFullname | ConvertTo-Html @htmlParams | Out-File $outputFileFullname
  $photoConsentInitials | ConvertTo-Html @htmlParams | Out-File $outputFileInitials

  if ($print) {
    $pdfPathFullname = "$downloadsPath\photograph_consent_fullname_$sectionNameFile.pdf"
    $pdfPathInitials = "$downloadsPath\photograph_consent_initials_$sectionNameFile.pdf"
    Invoke-PrintReport -htmlPaths @($outputFileFullname, $outputFileInitials) -pdfPaths @($pdfPathFullname, $pdfPathInitials)
  }

  Write-Host "✅ Photograph consent report saved to $outputFileFullname"
  Write-Host "✅ Photograph consent report saved to $outputFileInitials"
  return $photoConsentFullname
}
function Get-OsmDietary {
  <#
  .SYNOPSIS
  Retrieves allergies & dietary requirements of members from OSM.

  .DESCRIPTION
  Generates a report of members with allergies & dietary requirements.
  Output is saved as an HTML file in the Downloads folder
  and optionally sent to a printer.

  .PARAMETER sectionId
  The OSM section ID to generate the report for.

  .PARAMETER print
  Optional switch to send the output directly to the default printer.

  .EXAMPLE
  Get-OsmDietary -sectionId 12345

  .EXAMPLE
  Get-OsmDietary -sectionId 12345 -print
  #>
  param (
    [int]$sectionId,
    [switch]$print
  )

  Assert-ValidSection -sectionId $sectionId
  
  $section = $sections | Where-Object { $_.sectionId -eq $sectionId }
  $termId = $section.termId
  $sectionName = $section.sectionName
  $sectionNameFile = $sectionName.Replace(" ", "_").ToLower()

  # Members
  $membersList = Get-OsmMembers -sectionId $sectionId -termId $termId

  # Members allergies & dietary requirements
  Write-Host "🔄 Fetching members allergies & dietary requirements from OSM..."
  $dietaryFullname = @()
  $dietaryInitials = @()
  foreach ($member in $membersList) {
    $memberFullname = $member.full_name
    $memberInitials = Get-MemberInitials -firstName $member.firstname -lastName $member.lastname
    $memberId = $member.scoutid
    $membersDataUrlWithParams = Add-QueryParams -url $membersDataUrl -params @{
      section_id = $sectionId
      associated_id = $memberId
      associated_type = "member"
      context = "members"
    }
    $memberData = (Invoke-OsmApi -url $membersDataUrlWithParams).data
    $memberAllergies = (($memberData | where { $_.identifier -eq "standard_fields" }).columns | where { $_.varname -eq "allergies" }).value
    $memberDietary = (($memberData | where { $_.identifier -eq "standard_fields" }).columns | where { $_.varname -eq "dietary" }).value

    if (($memberAllergies -ne "N/A" -and $memberAllergies -ne "None" -and $memberAllergies -ne "" -and $memberAllergies -ne "NKDA") -or ($memberDietary -ne "N/A" -and $memberDietary -ne "None" -and $memberDietary -ne "")) {
      $dietaryFullname += [PSCustomObject]@{
        Name      = $memberFullname
        Allergies = $memberAllergies
        Dietary   = $memberDietary
      }
      $dietaryInitials += [PSCustomObject]@{
        Name      = $memberInitials
        Allergies = $memberAllergies
        Dietary   = $memberDietary
      }
    }
  }
  Write-Host "✅ Retrieved $($dietaryFullname.Count) members with allergies & dietary requirements"

  # Output report
  $htmlParams = @{
    Head = $htmlStyle
    Title = "$sectionName Allergies & Dietary Requirements"
    PreContent = "<h1>$sectionName Allergies & Dietary Requirements</h1>"
  }
  $outputFileFullname = "$downloadsPath\allergies_dietary_fullname_$sectionNameFile.html"
  $outputFileInitials = "$downloadsPath\allergies_dietary_initials_$sectionNameFile.html"
  $dietaryFullname | ConvertTo-Html @htmlParams | Out-File $outputFileFullname
  $dietaryInitials | ConvertTo-Html @htmlParams | Out-File $outputFileInitials

  if ($print) {
    $pdfPathFullname = "$downloadsPath\allergies_dietary_fullname_$sectionNameFile.pdf"
    $pdfPathInitials = "$downloadsPath\allergies_dietary_initials_$sectionNameFile.pdf"
    Invoke-PrintReport -htmlPaths @($outputFileFullname, $outputFileInitials) -pdfPaths @($pdfPathFullname, $pdfPathInitials)
  }

  Write-Host "✅ Allergies & dietary requirements report saved to $outputFileFullname"
  Write-Host "✅ Allergies & dietary requirements report saved to $outputFileInitials"
  return $dietaryFullname
}

# Main
try {
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
}
catch {
  Write-Warning "⚠️ Failed to initialize OSM Tools: $($_.Exception.Message)"
  Write-Warning "This may be due to missing/invalid credentials or network issues."
  Write-Warning "Please check your credentials with Export-OsmCredentials or verify network connectivity."
}