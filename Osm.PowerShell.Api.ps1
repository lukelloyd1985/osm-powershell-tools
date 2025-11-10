# ==========================================================
# Online Scout Manager (OSM) OAuth 2.0 API
# ==========================================================

# Settings
$osmAppPath = "$env:LOCALAPPDATA\osm-powershell-tools"
New-Item -Path $osmAppPath -ItemType Directory -Force | Out-Null
$credentialsFile = "$osmAppPath\osm_credentials.json"
$tokenFile = "$osmAppPath\osm_token.json"
$tokenUrl = "https://www.onlinescoutmanager.co.uk/oauth/token"
$userRolesUrl = "https://www.onlinescoutmanager.co.uk/api.php?action=getUserRoles"
$termsUrl = "https://www.onlinescoutmanager.co.uk/api.php?action=getTerms"

# Functions
function Request-Credentials {
  $clientId = Read-Host "Enter your OSM OAuth Client ID"
  $clientSecret = Read-Host "Enter your OSM OAuth Client Secret"
  Export-OsmCredentials -clientId $clientId -clientSecret $clientSecret
  return @{ clientId = $clientId; clientSecret = $clientSecret }
}
function Export-OsmCredentials {
  param($clientId, $clientSecret)
  $data = @{
    clientId     = $clientId
    clientSecret = $clientSecret
  }
  $data | ConvertTo-Json | Set-Content -Path $credentialsFile -Encoding UTF8
  Write-Host "✅ Credentials saved to $credentialsFile"
}
function Import-OsmCredentials {
  if (Test-Path $credentialsFile) {
    return Get-Content -Raw -Path $credentialsFile | ConvertFrom-Json
  }
  else {
    Write-Warning "⚠️ Failed to load credentials file. Requesting credentials..."
    return Request-Credentials
  }
}
function Export-OsmToken {
  param($response)
  $data = @{
    access_token  = $response.access_token
    refresh_token = $response.refresh_token
    expires_at    = (Get-Date).AddSeconds($response.expires_in)
    token_type    = $response.token_type
  }
  $data | ConvertTo-Json | Set-Content -Path $tokenFile -Encoding UTF8
  Write-Host "✅ Token saved to $tokenFile"
}
function Import-OsmToken {
  param($clientId, $clientSecret)
  if (Test-Path $tokenFile) {
    return Get-Content -Raw -Path $tokenFile | ConvertFrom-Json
  }
  else {
    Write-Warning "⚠️ Failed to load token file. Requesting token..."
    return New-OsmToken -clientId $clientId -clientSecret $clientSecret
  }
}
function Update-OsmToken {
  param($refreshToken, $clientId, $clientSecret)
  Write-Host "🔄 Refreshing token..."
  $body = @{
    grant_type    = "refresh_token"
    refresh_token = $refreshToken
    client_id     = $clientId
    client_secret = $clientSecret
  }
  try {
    $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body
    Export-OsmToken $response
    return $response.access_token
  }
  catch {
    Write-Error "❌ Token refresh failed: $($_.Exception.Message)" -ErrorAction Stop
  }
}
function New-OsmToken {
  param($clientId, $clientSecret)
    
  $body = @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = "section:member:read section:programme:read section:event:read"
  }

  try {
    $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body
    Export-OsmToken $response
    return @{ access_token = $response.access_token; expires_at = (Get-Date).AddSeconds($response.expires_in) }
  }
  catch {
    Write-Error "❌ Token creation failed: $($_.Exception.Message)" -ErrorAction Stop
  }
}
function Get-OsmToken {
  param($clientId, $clientSecret)
  $osmToken = Import-OsmToken -clientId $clientId -clientSecret $clientSecret
  if ((Get-Date) -lt (Get-Date $osmToken.expires_at)) {
    Write-Host "✅ Using existing valid token."
    return $osmToken.access_token
  }
  else {
    return Update-OsmToken -refreshToken $osmToken.refresh_token -clientId $clientId -clientSecret $clientSecret
  }
}
function Invoke-OsmApi {
  param (
    [string]$url,
    [string]$method = "GET",
    [hashtable]$body = $null
  )

  $OsmCredentials = Import-OsmCredentials
  $OsmToken = Get-OsmToken -clientId $OsmCredentials.clientId -clientSecret $OsmCredentials.clientSecret
  $headers = @{ Authorization = "Bearer $OsmToken" }

  try {
    if ($method -eq "POST") {
      return Invoke-RestMethod -Uri $url -Headers $headers -Body $body -Method Post
    }
    else {
      return Invoke-RestMethod -Uri $url -Headers $headers -Method Get
    }
  }
  catch {
    Write-Error "❌ API call failed: $($_.Exception.Message)" -ErrorAction Stop
  }
}

# Main
Invoke-OsmApi -url $userRolesUrl