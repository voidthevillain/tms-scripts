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
# .\tms-AnonCannotJoinMeetings.ps1 user@domain.com
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
$user = (Get-MsolUser -UserPrincipalName $UPN)

Write-Host 'Checking if the user exists:'
if ($user) {
  Write-Host -ForegroundColor Green 'The user exists.'
} else {
  return Write-Host -ForegroundColor Red 'The user does not exist.'
}

# USER LICENSES
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

# SIP ENABLED
$tmsUser = (Get-CsOnlineUser $UPN)

Write-Host 'Checking if the user is SIP enabled:'
if ($tmsUser.IsSipEnabled) {
  Write-Host -ForegroundColor Green 'The user is SIP enabled.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not SIP enabled.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# MEETING SETTINGS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$meetingSettings = (Get-CsTeamsMeetingConfiguration)

Write-Host 'Checking if the tenant meeting configuration allows anonymous users to join meetings:'
if (!$meetingSettings.DisableAnonymousJoin) {
  Write-Host -ForegroundColor Green 'The tenant meeting configuration allows anonymous users to join meetings.'
} else {
  $toEnableAnonJoinTenant = Read-Host 'The tenant meeting configuration does not allow anonymous users to join meetings. Would you like to allow anonymous users? [Y/N]'
  if ($toEnableAnonJoinTenant -eq 'Y') {
    Set-CsTeamsMeetingConfiguration -DisableAnonymousJoin $false
    Write-Host -ForegroundColor Green 'The tenant meeting configuration now allows anonymous users to join meetings.'
  } else {
    return Write-Host -ForegroundColor Red 'The tenant meeting configuration does not allow anonymous users to join meetings.'
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# MEETING POLICY
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$policyName = $tmsUser.TeamsMeetingPolicy

if ($policyName -eq $null) {
  $policyName = 'Global'
}

Write-Host "Checking if the user's assigned meeting policy allows anonymous users to join meetings:"
$meetingPolicy = (Get-CsTeamsMeetingPolicy -Identity $policyName)
if ($meetingPolicy.AllowAnonymousUsersToJoinMeeting) {
  Write-Host -ForegroundColor Green "The user's meeting policy allows anonymous users to join meetings."
} else {
  $toEnableAnonJoinPolicy = Read-Host "The user's meeting policy does not allow anonymous users to join meetings. Would you like to allow anonymous users? [Y/N]"
  if ($toEnableAnonJoinPolicy -eq 'Y') {
    Set-CsTeamsMeetingPolicy -Identity $policyName -AllowAnonymousUsersToJoinMeeting $true
    Write-Host -ForegroundColor Green "The user's meeting policy now allows anonymous users to join meetings."
  } else {
    return Write-Host -ForegroundColor Red "The user's meeting policy does not allow anonymous users to join meetings."
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
  Start-Transcript -Path "$($desktopPath)\AnonCannotJoinMeetings.txt"
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
  Write-Host 'Meeting settings (TMS):'
  Get-CsTeamsMeetingConfiguration
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Meeting policy (TMS):"
  Get-CsTeamsMeetingPolicy -Identity $policyName
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 