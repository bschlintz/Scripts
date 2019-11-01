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

function Split-ArrayIntoChunks {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)][object[]]$Array,
    [Parameter(Mandatory = $true)][int]$ChunkSize
  )

  begin {
    $counter = [pscustomobject] @{ Value = 0 }
  }
  process {
    $Array | Group-Object -Property { [math]::Floor($counter.Value++ / $ChunkSize) }
  }
  end {
  }
}

function CreateBatch-NewO365Group {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)]$Groups
  )

  begin {
  }
  process {
    @{
      "requests" = @(foreach ($group in $Groups) { 
          @{
            "id"      = $group.mailNickname
            "url"     = "/groups"
            "method"  = "POST"
            "headers" = @{
              "Content-Type" = "application/json"
            }
            "body"    = $group
          }
        })
    } | ConvertTo-Json -Depth 5 -Compress
  }
  end {
  }
}


function Process-GraphBatch {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)][string]$Payload,
    [Parameter(Mandatory = $true)][string]$AccessToken
  )

  begin {
    $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
    $uri = "https://graph.microsoft.com/v1.0/`$batch"
  }
  process {
    try {
      $response = Invoke-RestMethod -Uri $uri -Headers $headers -Body $Payload -Method Post

      foreach ($batchResponse in $response.responses) {
        if ($batchResponse.status -ge 400) {
          Write-Error "Error in batch request: $($batchResponse.id). Error: $(ConvertTo-Json $batchResponse.body.error -Depth 5)"
        }
      }             
    }
    catch {
      Write-Error "Error sending batch request. Error: $($_.Exception)"
    }
  }
  end {
    $response
  }
}


# script requries Group.ReadWrite.All and Directory.ReadWrite.All
$tenant = "contoso.onmicrosoft.com"
$clientId = "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxx"                # aka "Application ID" in Azure Portal > Azure Active Directory > App Registrations
$clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"     # aka "Keys" in Azure Portal > Azure Active Directory > App Registrations

# Graph $batch limit is 20
# https://developer.microsoft.com/en-us/graph/docs/concepts/known_issues#json-batching
$batchSize = 20

$groupStartIndex = 1
$groupEndIndex = 500
$groupPrefix = "zDemoGroup"
$groupDefaultOwner = "admin@$tenant"
$groupDefaultMember = "alexw@$tenant"

if ( -not $token -or $token.ExpiresOn.DateTime -lt (Get-Date) ) {
  $token = Get-Token -Tenant $tenant -ClientID $clientId -ClientSecret $clientSecret -Resource "https://graph.microsoft.com/"
}

if ( $token.access_token ) {
  $groups = @()
  $groupStartIndex..$groupEndIndex | ForEach-Object {
    $groupName = "$($groupPrefix)$_"
    $groups += @{
      "mailNickname"       = "$groupName" 
      "displayName"        = "$groupName"
      "description"        = "$groupName - Created by Graph API"
      "mailEnabled"        = $true
      "securityEnabled"    = $true
      "groupTypes"         = @("Unified")
      "owners@odata.bind"  = @(
        "https://graph.microsoft.com/v1.0/users/$groupDefaultOwner"
      )
      "members@odata.bind" = @(
        "https://graph.microsoft.com/v1.0/users/$groupDefaultMember"
      )
    }
  }

  Write-Host "Splitting $($groups.Count) groups into $([Math]::Ceiling($groups.Count / $batchSize)) batches..."
  $batches = Split-ArrayIntoChunks -Array $groups -ChunkSize $batchSize

  # Iterate each batch of users to create Graph $batch JSON payload
  for ($idx = 0; $idx -lt $batches.Length; $idx++) {
    $batch = $batches[$idx].Group

    Write-Host "Processing Batch: $($idx + 1) of $($batches.Length) [Batch Size: $($batch.Count)]"

    $batchPayload = CreateBatch-NewO365Group -Groups $batch

    Process-GraphBatch -Payload $batchPayload -AccessToken $token.access_token | Out-Null
  }
}
