<# 
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. We grant You a nonexclusive, royalty-free right to use and modify the 
 Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that
 You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the
 Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which
 the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers
 from and against any claims or lawsuits, including attorneys' fees, that arise or result from the
 use or distribution of the Sample Code.
#>

function Get-Token {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)][string]$Tenant,
    [Parameter(Mandatory = $true)][System.Guid]$ClientID,
    [Parameter(Mandatory = $true)][string]$ClientSecret,
    [Parameter(Mandatory = $true)][string]$Resource
  )
  
  begin {
  }

  process {
    $ReqBody = @{
      Grant_Type    = "client_credentials"
      Scope         = "$Resource.default"
      client_Id     = $ClientID
      Client_Secret = $ClientSecret
    } 
    
    $RequestToken = @{
      ContentType = 'application/x-www-form-urlencoded'
      Method      = 'POST'
      Body        = $ReqBody
      Uri         = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token"
    }

    try {
      $response = Invoke-RestMethod @RequestToken
    }
    catch {
      Write-Error $_.Exception
    }
  }
  end {
    $response
  }
}

function Get-AuthenticationHeaders {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)][string]$AccessToken
  )

  begin {
  }
  process {
    @{
      'Content-Type'  = 'application/json'
      'Authorization' = "Bearer $($AccessToken)"
    }    
  }
  end {
  }
}

function Get-Users {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)][string]$AccessToken
  )

  begin {
    $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member'"
    $users = @()
  }
  process {
    try {
      $response = Invoke-RestMethod -Uri $uri -Headers $headers -Body $Payload -Method Get
      $users = $response.value
    }
    catch {
      Write-Error "Error sending batch request. Error: $($_.Exception)"
    }
  }
  end {
    $users
  }
}

function Get-Events {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)][string]$UserPrincipalName,
    [Parameter(Mandatory = $true)][string]$AccessToken
  )

  begin {
    $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
    $uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/events?`$top=3"
  }
  process {
    try {
      $response = Invoke-RestMethod -Uri $uri -Headers $headers -Body $Payload -Method Get
      return @($response.value)
    }
    catch {
      if ($_.Exception.Message.Contains("(403) Forbidden")) {
        Write-Host "  [ACCESS DENIED TO EVENTS]" -ForegroundColor Yellow
      }
      else {
        Write-Error "Request Error: $($_.Exception)"        
      }
      throw $_
    }
  }
}

# script requries Users.Read.All, Calendars.ReadWrite
$tenant = "contoso.onmicrosoft.com"
$clientId = "3617fd1b-xxxx-xxxx-xxxx-057668e52db0"                # aka "Application ID" in Azure Portal > Azure Active Directory > App Registrations
$clientSecret = "abc123secret"              # aka "Keys" in Azure Portal > Azure Active Directory > App Registrations

$token = Get-Token -Tenant $tenant -ClientID $clientId -ClientSecret $clientSecret -Resource "https://graph.microsoft.com/"

if ($token.access_token) {
  
  $users = Get-Users -AccessToken $token.access_token

  foreach ($user in $users | ? { $_.mail -ne $null }) {
    Write-Host "[USER] $($user.userPrincipalName)"
    try {
      $events = Get-Events -AccessToken $token.access_token -UserPrincipalName $user.userPrincipalName
      
      if ($null -ne $events -and $events.Count -gt 0) {
        foreach ($event in $events) {
          Write-Host "  [EVENT] $($event.subject)" -ForegroundColor Green
        }
      }
      else {
        Write-Host "  [NO EVENTS FOUND]" -ForegroundColor Blue
      }
    }
    catch {}
    Write-Host
  }
}
