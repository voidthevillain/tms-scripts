# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# DEPENDENCIES: Connect-MsolService, Connect-ExchangeOnline, Connect-MicrosoftTeams
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USAGE: 
# Connect-MsolService
# Connect-ExchangeOnline
# Connect-MicrosoftTeams
# .\tms-MissingOWATeamsAddIn user@domain.com
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Back it up with https://aka.ms/TeamsAdd-inDiag
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
function Get-TenantAssignedPlans {
  param (
    $APS
  )

  $tAPS = @()

  foreach ($ap in $APS) {
    $tAPS += ($ap.Capability | Out-String).trim()
  }

  return $tAPS
}

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

# TENANT TEAMS PLAN
$tenantAssignedPlans = (Get-CsTenant).AssignedPlan

Write-Host 'Checking if the tenant is licensed for Teams:'
$parsedTenantAssignedPlans = Get-TenantAssignedPlans $tenantAssignedPlans
if ($parsedTenantAssignedPlans -contains 'Teams') {
  Write-Host -ForegroundColor Green 'The tenant is licensed for Teams.'
} else {
  return Write-Host -ForegroundColor Red 'The tenant is not licensed for Teams.'
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

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# EXCHANGE HOMING
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
$recipientType = (Get-Mailbox -Identity $UPN).RecipientTypeDetails
Write-Host 'Checking the Exchange homing of the user:'

if ($recipientType -eq 'UserMailbox') {
  Write-Host -ForegroundColor Green 'The mailbox is hosted in Exchange Online.'
} elseif ($recipientType -eq 'MailUser') {
  return Write-Host -ForegroundColor Red 'The mailbox is hosted in Exchange On-Premises.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USER LICENSES
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
$userLicense = Get-OfficeUserLicense $UPN

Write-Host 'Checking if the user is licensed:'
if ($userLicense.isLicensed) {
  Write-Host -ForegroundColor Green 'The user is licensed.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed.'
}

Write-Host 'Checking user licenses:'

# TEAMS1
if ($userLicense.ServicePlans -contains 'TEAMS1') {
  Write-Host -ForegroundColor Green 'The user is licensed for Teams.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Teams.'
}

# MCOSTANDARD
if ($userLicense.ServicePlans -contains 'MCOSTANDARD' -OR $userLicense.ServicePlans -contains 'MCO_TEAMS_IW' ) {
  Write-Host -ForegroundColor Green 'The user is licensed for Skype for Business Online.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Skype for Business Online.'
}

# EXO (EXCHANGE_S_STANDARD, EXCHANGE_S_ENTERPRISE) 
if ($userLicense.ServicePlans -contains 'EXCHANGE_S_STANDARD' -OR $userLicense.ServicePlans -contains 'EXCHANGE_S_ENTERPRISE') {
  Write-Host -ForegroundColor Green 'The user is licensed for Exchange Online.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Exchange Online.'
}

# SIP ENABLED
$tmsUser = (Get-CsOnlineUser $UPN)

Write-Host 'Checking if the user is SIP enabled:'
if ($tmsUser.IsSipEnabled) {
  Write-Host -ForegroundColor Green 'The user is SIP enabled.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not SIP enabled.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# COEXISTENCE MODE
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
$coexistenceMode = $tmsUser.TeamsUpgradeEffectiveMode

Write-Host 'Checking if the user is assigned an upgrade mode which allows hosing meetings in Teams:'
if (($coexistenceMode -eq 'TeamsOnly') -OR ($coexistenceMode -eq 'Islands') -OR ($coexistenceMode -eq 'SfBWithTeamsCollabAndMeetings') -OR ($coexistenceMode -eq 'SfBWithTeamsCollabAndMeetingsWithNotify')) {
  Write-Host -ForegroundColor Green 'The user is assigned an upgrade mode which allows hosting meetings in Teams.'
} else {
  return Write-Host -ForegroundColor Red 'The user is assigned an upgrade mode which does not allow hosting meetings in Teams.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# MEETING POLICY
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
$meetingPolicyName = $tmsUser.TeamsMeetingPolicy.Name

if ($meetingPolicyName -eq $null) {
  $meetingPolicyName = 'Global'
}

Write-Host "Checking if the user's meeting policy allows the Outlook Add-In:"
$meetingPolicy = (Get-CsTeamsMeetingPolicy -Identity $meetingPolicyName)
if ($meetingPolicy.AllowOutlookAddIn) {
  Write-Host -ForegroundColor Green "The user's meeting policy allows the Outlook Add-In."
} else {
  $toAllowOutlookAddIn = Read-Host "The user's meeting policy does not allow the Outlook Add-In. Would you like to allow it? [Y/N]"
  if ($toAllowOutlookAddIn -eq 'Y') {
    Set-CsTeamsMeetingPolicy -Identity $meetingPolicyName -AllowOutlookAddIn $true
    Write-Host -ForegroundColor Green "The user's meeting policy now allows the Outlook Add-In."
  } else {
    return Write-Host -ForegroundColor Red "The user's meeting policy does not allow the Outlook Add-In."
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# CALENDAR CONFIGURATION
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
$calendarConfig = (Get-MailboxCalendarConfiguration -Identity $UPN)

Write-Host "Checking if the user's default meeting provider is set to TeamsForBusiness:"
if ($calendarConfig.DefaultOnlineMeetingProvider -eq 'TeamsForBusiness') {
  Write-Host -ForegroundColor Green "The user's default meeting provider is set to TeamsForBusiness."
} else {
  $toSetTfB = Read-Host "The user's default meeting provider is not set to TeamsForBusiness. Would you like to set it? [Y/N]"
  if ($toSetTfB -eq 'Y') {
    Set-MailboxCalendarConfiguration -Identity $UPN -DefaultOnlineMeetingProvider 'TeamsForBusiness'
  } else {
    return Write-Host -ForegroundColor Red "The user's default meeting provider is not set to TeamsForBusiness."
  }
}

Write-Host "Checking if the user's meeting providers list is populated:"
if ($calendarConfig.AllowedOnlineMeetingProviders) {
  Write-Host -ForegroundColor Green "The user's meeting providers list is populated."
  Write-Host "Checking if the user's meeting providers list includes TeamsForBusiness:"
  if ($calendarConfig.AllowedOnlineMeetingProviders) {
    Write-Host -ForegroundColor Green "The user's meeting providers list contains TeamsForBusiness."
  } else {  
    return Write-Host -ForegroundColor Red "The user's meeting providers list does not contain TeamsForBusiness. Please re-assign the Teams license to the user and run the script again in 24 hours."
  }
} else {
  return Write-Host -ForegroundColor Red "The user's meeting providers list is not populated. Please schedule and join a few meetings in Teams and run the script again in 24 hours."
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TRANSCRIPT
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\MissingOWATeamsAddIn.txt"
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
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Meeting policy (TMS):"
  Get-CsTeamsMeetingPolicy -Identity $meetingPolicyName
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Calendar configuration (EXO):"
  Get-MailboxCalendarConfiguration -Identity $UPN
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 