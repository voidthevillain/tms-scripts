# requires Connect-MicrosoftTeams, Connect-MsolService

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
# Tenant existence
$tenant = (Get-CsTenant)

Write-Host 'Checking if the tenant exists:'
if ($tenant) {
  Write-Host -ForegroundColor Green 'The tenant exists.'
} else {
  return Write-Host -ForegroundColor Red 'The tenant does not exist.'
}

# Tenant Teams plan
$tenantAssignedPlans = (Get-CsTenant).AssignedPlan

Write-Host 'Checking if the tenant is licensed for Teams:'
$parsedTenantAssignedPlans = Get-TenantAssignedPlans $tenantAssignedPlans
if ($parsedTenantAssignedPlans -contains 'Teams') {
  Write-Host -ForegroundColor Green 'The tenant is licensed for Teams.'
} else {
  return Write-Host -ForegroundColor Red 'The tenant is not licensed for Teams.'
}

# User 
$user = (Get-MsolUser -UserPrincipalName $UPN)

Write-Host 'Checking if the user exists:'
if ($user) {
  Write-Host -ForegroundColor Green 'The user exists.'
} else {
  return Write-Host -ForegroundColor Red 'The user does not exist.'
}

# User licenses
$userLicense = Get-OfficeUserLicense $UPN

Write-Host 'Checking if the user is licensed:'
if ($userLicense.isLicensed) {
  Write-Host -ForegroundColor Green 'The user is licensed.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed.'
}

Write-Host 'Checking user licenses:'

# CHECK TEAMS (TEAMS1)
if ($userLicense.ServicePlans -contains 'TEAMS1') {
  Write-Host -ForegroundColor Green 'The user is licensed for Teams.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Teams.'
}

# CHECK SfBO (MCOSTANDARD, MCO_TEAMS_IW)
if ($userLicense.ServicePlans -contains 'MCOSTANDARD' -OR $userLicense.ServicePlans -contains 'MCO_TEAMS_IW' ) {
  Write-Host -ForegroundColor Green 'The user is licensed for Skype for Business Online.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Skype for Business Online.'
}

# check if Sip enabled
$tmsUser = (Get-CsOnlineUser $UPN)

Write-Host 'Checking if the user is SIP enabled:'
if ($tmsUser.Enabled) {
  Write-Host -ForegroundColor Green 'The user is SIP enabled.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not SIP enabled.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Meeting policy settings
$policyName = $tmsUser.TeamsMeetingPolicy

if ($policyName -eq $null) {
  $policyName = 'Global'
}

$meetingPolicy = (Get-CsTeamsMeetingPolicy -Identity $policyName)

Write-Host 'Checking if the user is assigned a Teams meeting policy that allows to chat in meetings:'
if ($meetingPolicy.MeetingChatEnabledType -eq 'Enabled') {
  Write-Host -ForegroundColor Green 'The user is assigned a Teams meeting policy that allows to chat in meetings.'
} else {
  $toEnableMeetingChat = Read-Host 'The user is assigned a Teams meeting policy that does not allow to chat in meetings. Would you like to enable meeting chat on the policy? [Y/N]'
  if ($toEnableMeetingChat -eq 'Y') {
    Set-CsTeamsMeetingPolicy -Identity $policyName -MeetingChatEnabledType 'Enabled'
    Write-Host -ForegroundColor Green 'The user is now assigned a Teams meeting policy that allows to chat in meetings.'
  } else {
    return Write-Host -ForegroundColor Red 'The user is assigned a Teams meeting policy that does not allow to chat in meetings.'
  }
}

$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\CannotChatInMeetings.txt"
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
  Get-CsTeamsMeetingPolicy -Identity $policyName
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 