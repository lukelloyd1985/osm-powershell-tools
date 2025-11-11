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

# Main
Write-Host $sectionTerm