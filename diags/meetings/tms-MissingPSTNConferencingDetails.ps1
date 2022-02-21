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
# .\tms-MissingAudioConferencingNumber.ps1 user@domain.com
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
# EXCHANGE HOMING - ????
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
# $recipientType = (Get-Mailbox -Identity $UPN).RecipientTypeDetails
# Write-Host 'Checking the Exchange homing of the user:'

# if ($recipientType -eq 'UserMailbox') {
#   Write-Host -ForegroundColor Green 'The mailbox is hosted in Exchange Online.'
# } elseif ($recipientType -eq 'MailUser') {
#   return Write-Host -ForegroundColor Yellow 'The mailbox is hosted in Exchange On-Premises. See https://docs.microsoft.com/en-us/microsoftteams/troubleshoot/exchange-integration/teams-exchange-interaction-issue'
# }

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
# if ($userLicense.ServicePlans -contains 'EXCHANGE_S_STANDARD' -OR $userLicense.ServicePlans -contains 'EXCHANGE_S_ENTERPRISE') {
#   Write-Host -ForegroundColor Green 'The user is licensed for Exchange Online.'
# } else {
#   return Write-Host -ForegroundColor Red 'The user is not licensed for Exchange Online.'
# }

# SIP ENABLED
$tmsUser = (Get-CsOnlineUser $UPN)

Write-Host 'Checking if the user is SIP enabled:'
if ($tmsUser.IsSipEnabled) {
  Write-Host -ForegroundColor Green 'The user is SIP enabled.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not SIP enabled.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUDIO CONFERENCING
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
# Write-Host 'Checking if the user is licensed for PSTN conferencing:'
# if ($userLicense.ServicePlans -contains 'MCOMEETADV' -OR $userLicense.ServicePlans -contains 'MCOMEETADD') {
#   Write-Host -ForegroundColor Green 'The user is licensed for Audio Conferencing.'
# } else {
#   return Write-Host -ForegroundColor Red 'The user is not licensed for Audio Conferencing.'
# }

Write-Host 'Checking if the user is enabled for PSTN conferencing:'
$conferencingUser = (Get-CsOnlineDialInConferencingUser -Identity $UPN)
if ($conferencingUser) {
  Write-Host -ForegroundColor Green 'The user is enabled for PSTN conferencing.'
} else {
  $toEnable = Read-Host 'The user is not enabled for PSTN conferencing. Would you like to enable the user for PSTN conferencing? [Y/N]'
  if ($toEnable -eq 'Y') {
    try {
      Enable-CsOnlineDialInConferencingUser -Identity $UPN
    } catch { return Write-Host -ForegroundColor Red 'An error occurred:'$_}
    Write-Host -ForegroundColor Green 'The user is now enabled for PSTN conferencing.'
  } else {
      return Write-Host -ForegroundColor Red 'The user is not enabled for PSTN conferencing.'    
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
  Start-Transcript -Path "$($desktopPath)\MissingPSTNConferencingDetails.txt"
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
  Write-Host "Conferencing user (TMS):"
  Get-CsOnlineDialInConferencingUser -Identity $UPN
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 