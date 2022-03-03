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
# .\tms-AutoAttendantNotReceivingCalls.ps1 autoattendant@domain.com
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Back it up with https://aka.ms/TeamsAADiag
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TODO: LineUri check (? Nested ?)

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN
)

Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'Auto attendant:'$UPN

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
# RESOURCE ACCOUNT
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'RESOURCE ACCOUNT'
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
$user = Get-MsolUser -UserPrincipalName $UPN

Write-Host 'Checking if the resource account exists:'
if ($user) {
  Write-Host -ForegroundColor Green 'The resource account exists.'
} else {
  return Write-Host -ForegroundColor Red 'The resource account does not exist.'
}

$userLicense = Get-OfficeUserLicense $UPN

Write-Host 'Checking if the resource account is licensed:'
if ($userLicense.isLicensed) {
  Write-Host -ForegroundColor Green 'The resource account is licensed.'
} else {
  return Write-Host -ForegroundColor Red 'The resource account is not licensed.'
}

Write-Host 'Checking resource account licenses:'

# TEAMS (TEAMS1)
if ($userLicense.ServicePlans -contains 'MCOEV_VIRTUALUSER') {
  Write-Host -ForegroundColor Green 'The resource account is licensed for Phone Standard - Virtual User.'
} else {
  return Write-Host -ForegroundColor Red 'The resource account is not licensed for Phone Standard - Virtual User.'
}


Write-Host 'Checking if the resource eaccount account is enabled:'

if ($user.BlockCredential) {
  Write-Host -ForegroundColor Green 'The resource account is not enabled.'
} else {
  return Write-Host -ForegroundColor Red 'The resource account is enabled.'
}

Write-Host 'Checking if the Department property is valid:'

if ($user.Department -eq 'Microsoft Communication Application Instance') {
  Write-Host -ForegroundColor Green 'The Department property is valid.'
} else {
  return Write-Host -ForegroundColor Red 'The Department property is not valid.'
}

Write-Host 'Checking if there is a SIP address set:'

$tmsUser = (Get-CsOnlineUser $UPN)

if ($tmsUser.SipAddress -eq $null) {
  Write-Host -ForegroundColor Green 'There is not a SIP address set.'
} else {
  return Write-Host -ForegroundColor Red 'There is a SIP address set.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TRANSCRIPT
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\AutoAttendantNotReceivingCalls.txt"
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Auto attendant:'$UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Resource account $($UPN) (TMS):"
  Get-CsOnlineUser $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Resource account $($UPN) (MSO):"
  Get-MsolUser -UserPrincipalName $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Resource account licenses (MSO):"
  Write-Host "Subscriptions:"$userLicense.SKU
  Write-Host "Service plans:"$userLicense.ServicePlans
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 