# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# DEPENDENCIES: Connect-MsolService, Connect-AzureAD, Connect-MicrosoftTeams
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USAGE: 
# Connect-MsolService
# Connect-AzureAD
# Connect-MicrosoftTeams
# cd PATH_TO_SCRIPT
# .\tms-CannotCallOboResourceAccount.ps1 user@domain.com CallQueueName ResourceAccountUPN TeamName
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN,
  $CQName,
  $OboRA,
  $TeamName
)

Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'User:'$UPN
Write-Host 'Call queue:'$CQName
Write-Host 'OBO resource account:'$OboRA
Write-Host 'Team name:'$TeamName

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
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'TENANT'
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
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
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'USER'
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
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
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'USER VOICE'
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
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
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'CALLING PLAN'
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
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
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'DIRECT ROUTING'
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
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
# CALL QUEUE
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'CALL QUEUE'
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
$CQ = Get-CsCallQueue -NameFilter $CQName

Write-Host 'Checking if the call queue exists:'
if ($CQ) {
  Write-Host -ForegroundColor Green 'The call queue exists.'
} else {
  return Write-Host -ForegroundColor Red "No call queue with the provided name $($CQName) exists."
}

$RAs = @()

Write-Host 'Checking the resource accounts linked this call queue:'
if ($CQ.ApplicationInstances) {
  if (!$CQ.ApplicationInstances[1]) {
    Write-Host -ForegroundColor Green 'The call queue is linked to a single resource account:'
    $RAs += (Get-CsOnlineUser $CQ.ApplicationInstances[0])
    Write-Host $RAs[0].UserPrincipalName
  } else {
    Write-Host -ForegroundColor Green 'The call queue is linked to multiple resource accounts:'
    foreach ($appInstance in $CQ.ApplicationInstances) {
      $RAs += (Get-CsOnlineUser $appInstance)
    }

    foreach ($RAcc in $RAs) {
      Write-Host $RAcc.UserPrincipalName
    }
  }
} else {
  return Write-Host -ForegroundColor Red 'The call queue is not linked to any resource account.'
}

$channelId = $CQ.ChannelId

Write-Host 'Checking if the call queue is linked to a Teams channel:'
if ($channelId) {
  Write-Host -ForegroundColor Green 'The call queue is linked to a Teams channel.'
} else {
  return Write-Host -ForegroundColor Red 'The call queue is not linked to any Teams channel.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TEAM & CHANNEL
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'TEAM & CHANNEL'
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'

$team = (Get-Team -DisplayName $TeamName)

Write-Host 'Checking if the provided team exists:'
if ($team) {
  Write-Host -ForegroundColor Green 'The provided team exists.'
} else {
  return Write-Host -ForegroundColor Red 'The provided team does not exist.'
}

$teamId = $team.GroupId

$teamChannelIds = (Get-TeamChannel -GroupId $teamId).Id

Write-Host 'Checking if the linked channel exists in the provided team:'
if ($teamChannelIds -contains $channelId) {
  Write-Host -ForegroundColor Green 'The linked channel exists in the provided team.'
} else {
  return Write-Host -ForegroundColor Red 'The linked channel does not exist in the provided team.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# M365 GROUP MEMBERSHIP
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'M365 GROUP MEMBERSHIP'
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# If no proper membership (I.e. added as Owner in AC):
# diagnosticsCode: {"callControllerCode":403,"callControllerSubCode":10105,"phrase":"On-behalf-of authorization failed.","resultCategories":["UnexpectedClientError"]}
$groupMembers = (Get-AzureADGroupMember -ObjectId $teamId).UserPrincipalName
# write-host $groupMembers
Write-Host 'Checking if the user has proper M365 group member permissions:'
if ($groupMembers -contains $UPN) {
  Write-Host -ForegroundColor Green 'The user is properly added as a member to the M365 group.'
} else {
  $toAddAsMember = Read-Host 'The user is not properly added as a member to the M365 group. Would you like to add the user as member? [Y/N]'
  if ($toAddAsMember -eq 'Y') {
    $userObjId = (Get-AzureAdUser -ObjectID $UPN).ObjectId
    Add-AzureAdGroupMember -ObjectId $teamId -RefObjectId $userObjId
    Write-Host -ForegroundColor Green 'The user was successfully added as a member to the M365 group.'
  } else {
    return Write-Host -ForegroundColor Red 'The user is not properly added as a member to the M365 group.'
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# ON-BEHALF-OF RESOURCE ACCOUNT
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'OBO RESOURCE ACCOUNT'
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
$raLicense = Get-OfficeUserLicense $OboRA

Write-Host 'Checking if the resource account is licensed:'
if ($raLicense.isLicensed) {
  Write-Host -ForegroundColor Green 'The resource account is licensed.'
} else {
  return Write-Host -ForegroundColor Red 'The resource account is not licensed.'
}

Write-Host 'Checking resource account licenses:'

# MCOEV_VIRTUALUSER
if ($raLicense.ServicePlans -contains 'MCOEV_VIRTUALUSER') {
  Write-Host -ForegroundColor Green 'The resource account is licensed for Phone System - Virtual user.'
} else {
  return Write-Host -ForegroundColor Red 'The resource account is not licensed for Phone System - Virtual User.'
}

# NUMBER TYPE
$raLineUri = (Get-CsOnlineUser $OboRA).LineUri

if ($raLineUri[0] -ne '+') {
  $raLineUri = "+$($raLineUri)"
}

Write-Host 'Checking if the resource account has a phone number assigned:'
if ($raLineUri) {
  Write-Host -ForegroundColor Green 'The resource account has a phone number assigned.'

  Write-Host 'Checking if the phone number is Online or On-premises:'
  $raLineUriTrimmed = $raLineUri.split('+')[1]
  $isOnline = (Get-CsOnlineTelephoneNumber -TelephoneNumber $raLineUriTrimmed)
  if ($isOnline) {
    Write-Host 'The phone number is an Online (Microsoft) number.'
  } else {
    Write-Host 'The phone number is an On-Premises (Direct Routing) number.'

    $raVRP = (Get-CsOnlineUser $OboRA).OnlineVoiceRoutingPolicy

    if ($raVRP -eq $null) {
      $raVRP = 'Global'
    }

    $pstnUsageName = ((Get-CsOnlineVoiceRoutingPolicy -Identity $raVRP).OnlinePstnUsages | Out-String).trim()
    $voiceRoute = (Get-CsOnlineVoiceRoute | ? {$_.OnlinePstnUsages -contains $pstnUsageName})

    Write-Host "Checking if the resource account's voice routing policy has a gateway route:"
    if ($voiceRoute.OnlinePstnGatewayList) {
      Write-Host -ForegroundColor Green "The resource account's voice routing policy has a gateway route."
    } else {
      return Write-Host -ForegroundColor Red "The resource account's voice routing policy does not have a gateway route."
    }
  }
} else {
  return Write-Host -ForegroundColor Red 'The resource account does not have a phone number assigned.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TRANSCRIPT
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\CannotCallOboResourceAccount.txt"
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'User:'$UPN
  Write-Host 'Call queue:'$CQName
  Write-Host 'OBO resource account:'$OboRA
  Write-Host 'Team name:'$TeamName
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
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Call queue (TMS):"
  Get-CsCallQueue -NameFilter $CQName
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "OBO Resource account (TMS):"
  Get-CsOnlineUser $OboRA
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Team (TMS):"
  $teamTR = Get-Team -DisplayName $TeamName
  Get-Team -DisplayName $TeamName
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Team channels (TMS):"
  Get-TeamChannel -GroupId $teamTR.GroupId
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 