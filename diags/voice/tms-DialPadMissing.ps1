# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# DEPENDENCIES: Connect-MsolService, Connect-MicrosoftTeams
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USAGE: 
# Connect-MsolService
# Connect-MicrosoftTeams
# cd PATH_TO_SCRIPT
# .\tms-DialPadMissing.ps1 user@domain.com
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Back it up with https://aka.ms/TeamsDialPadMissingDiag
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN
)

Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'User:'$UPN
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# FUNCTIONS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
function Get-OfficeUserLicense {
  param (
    [string]$UserPrincipalName
  )


  $SKUs = (Get-MsolUser -UserPrincipalName $UserPrincipalName).Licenses.AccountSkuId
  $ServicePlans = (Get-MsolUser -UserPrincipalName $UserPrincipalName).Licenses.ServiceStatus.ServicePlan.ServiceName

  $licenses = @{
    isLicensed = (Get-MsolUser -UserPrincipalName $UserPrincipalName).isLicensed
    SKU = @()
    ServicePlans = ''
  }

  if ($SKUs.length -gt 1) {
    foreach ($SKU in $SKUs) {
      $licenses.SKU += $SKU.split(":")[1]
    }
  } else {
    try {
      $licenses.SKU += $SKUs.split(":")[1]
    } catch {} # exception if user is unlicensed
  }

  $licenses.ServicePlans = $ServicePlans

  return $licenses
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TENANT
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$tenant = (Get-CsTenant)

Write-Host 'Checking if the tenant exists:'
if ($tenant) {
  Write-Host -ForegroundColor Green 'The tenant exists.'
} else {
  return Write-Host -ForegroundColor Red 'The tenant does not exist.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USER
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$user = Get-MsolUser -UserPrincipalName $UPN

Write-Host 'Checking if the user exists:'
if ($user) {
  Write-Host -ForegroundColor Green 'The user exists.'
} else {
  return Write-Host -ForegroundColor Red 'The user does not exist.'
}

$userLicense = Get-OfficeUserLicense $UPN

Write-Host 'Checking if the user is licensed:'
if ($userLicense.isLicensed) {
  Write-Host -ForegroundColor Green 'The user is licensed.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed.'
}

Write-Host 'Checking user licenses:'

# TEAMS (TEAMS1)
if ($userLicense.ServicePlans -contains 'TEAMS1') {
  Write-Host -ForegroundColor Green 'The user is licensed for Teams.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Teams.'
}

# SfBO (MCOSTANDARD)
if ($userLicense.ServicePlans -contains 'MCOSTANDARD') {
  Write-Host -ForegroundColor Green 'The user is licensed for Skype for Business Online.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Skype for Business Online.'
}

# MCOEV
if ($userLicense.ServicePlans -contains 'MCOEV') {
  Write-Host -ForegroundColor Green 'The user is licensed for Phone System.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Phone System.'
}

$tmsUser = (Get-CsOnlineUser $UPN)

Write-Host 'Checking if the user is SIP enabled:'
if ($tmsUser.IsSipEnabled) {
  Write-Host -ForegroundColor Green 'The user is SIP enabled.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not SIP enabled.'
}

Write-Host 'Checking if the user is homed online:'
if ($tmsUser.HostingProvider -eq 'sipfed.online.lync.com') {
  Write-Host -ForegroundColor Green 'The user is homed online.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not homed online.'
}

Write-Host 'Checking if the user is set up for Teams:'
if ($tmsUser.TeamsUpgradeEffectiveMode -eq 'TeamsOnly') {
  Write-Host -ForegroundColor Green 'The user is set up for TeamsOnly.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not set up for TeamsOnly.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# VOICE
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
Write-Host 'Checking if the user is Enterprise Voice enabled:'
if ($tmsUser.EnterpriseVoiceEnabled -eq $true) {
  Write-Host -ForegroundColor Green 'The user is Enterprise Voice enabled.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not Enterprise Voice enabled.'
}

$callingPolicyName = $tmsUser.TeamsCallingPolicy

if ($callingPolicyName -eq $null) {
  $callingPolicyName = 'Global'
}

$callingPolicy = (Get-CsTeamsCallingPolicy -Identity $callingPolicyName)

Write-Host 'Checking if the user is assigned a Teams calling policy that allows private calls:'
if ($callingPolicy.AllowPrivateCalling) {
  Write-Host -ForegroundColor Green 'The user is assigned a Teams calling policy that allows private calls.'
} else {
  return Write-Host -ForegroundColor Red 'The user is assigned a Teams calling policy that does not allow private calls.'
}

$voiceUser = (Get-CsOnlineVoiceUser -Identity $UPN)
$PSTNType = ($voiceUser.PSTNConnectivity.Value | Out-String).trim()
$isCP = $false
$isDR = $false

Write-Host 'Checking if the user is a Calling Plan or Direct Routing user:'
if ($PSTNType -eq 'Online') {
  $isCP = $true
  Write-Host 'The user is a Calling Plan user.'
} elseif ($PSTNType -eq 'OnPremises') { 
  $isDR = $true
  Write-Host 'The user is a Direct Routing user.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# CALLING PLAN
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
if ($isCP) {
  $lineUri = $tmsUser.LineURI

  Write-Host 'Checking if the user has a phone number assigned:'
  if ($lineUri) {
    Write-Host -ForegroundColor Green 'The user has a phone number assigned.'

    $voicePolicy = $tmsUser.VoicePolicy

    Write-Host "Checking if the user's voice policy is set to BusinessVoice:"
    if ($voicePolicy -eq 'BusinessVoice') {
      Write-Host -ForegroundColor Green "The user's voice policy is set to BusinessVoice."
    } else {
      return Write-Host -ForegroundColor Red "The user's voice policy is not set to BusinessVoice."
    }
  } else {
    return Write-Host -ForegroundColor Red 'The user does not have a phone number assigned.'
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# DIRECT ROUTING
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
if ($isDR) {
  $lineUri = $tmsUser.LineURI
  
  Write-Host 'Checking if the user has a phone number assigned:'
  if ($lineUri) {
    Write-Host -ForegroundColor Green 'The user has a phone number assigned.'

    $voicePolicy = $tmsUser.VoicePolicy

    Write-Host "Checking if the user's voice policy is set to HybridVoice:"
    if ($voicePolicy -eq 'HybridVoice') {
      Write-Host -ForegroundColor Green "The user's voice policy is set to HybridVoice."
    } else {
      return Write-Host -ForegroundColor Red "The user's voice policy is not set to HybridVoice."
    }

    $voiceRoutingPolicy = $tmsUser.OnlineVoiceRoutingPolicy

    if ($voiceRoutingPolicy -eq $null) {
      $voiceRoutingPolicy = 'Global'
    }

    $pstnUsageName = ((Get-CsOnlineVoiceRoutingPolicy -Identity $voiceRoutingPolicy).OnlinePstnUsages | Out-String).trim()
    $voiceRoute = (Get-CsOnlineVoiceRoute | ? {$_.OnlinePstnUsages -contains $pstnUsageName})

    Write-Host "Checking if the user's voice routing policy has a gateway route:"
    if ($voiceRoute.OnlinePstnGatewayList) {
      Write-Host -ForegroundColor Green "The user's voice routing policy has a gateway route."
    } else {
      return Write-Host -ForegroundColor Red "The user's voice routing policy does not have a gateway route."
    }
  } else {
    return Write-Host -ForegroundColor Red 'The user does not have a phone number assigned.'
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TRANSCRIPT
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\DialPadMissing.txt"
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'User:'$UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) (TMS):"
  Get-CsOnlineUser $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) (MSO):"
  Get-MsolUser -UserPrincipalName $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User licenses (MSO):"
  Write-Host "Subscriptions:"$userLicense.SKU
  Write-Host "Service plans:"$userLicense.ServicePlans
  if ($isDR) {
    Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
    Write-Host 'Voice routing policies (TMS):'
    Get-CsOnlineVoiceRoutingPolicy
    Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
    Write-Host 'PSTN usages (TMS):'
    Get-CsOnlinePstnUsage
    Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
    Write-Host 'Voice routes (TMS):'
    Get-CsOnlineVoiceRoute
  } 
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 