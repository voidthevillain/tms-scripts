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
# .\tms-CannotRecordMeetings.ps1 user@domain.com
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Back it up with https://aka.ms/MeetingRecordingDiag
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
$user = (Get-MsolUser -UserPrincipalName $UPN)

Write-Host 'Checking if the user exists:'
if ($user) {
  Write-Host -ForegroundColor Green 'The user exists.'
} else {
  return Write-Host -ForegroundColor Red 'The user does not exist.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# POLICY
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$policyName = (Get-CsOnlineUser $UPN).TeamsMeetingPolicy.Name

if (!$policyName)
{
  $policyName = 'Global'
}

$meetingPolicy = (Get-CsTeamsMeetingPolicy -Identity $policyName)

$allowCloudRecording = $meetingPolicy.AllowCloudRecording
Write-Host 'Checking if the user is assigned a Teams meeting polichy that allows to record meetings:'


if ($allowCloudRecording) {
  Write-Host -ForegroundColor Green 'The user is assigned a Teams meeting policy that allows to record meetings.'
} else {
  $toEnableCloudRecording = Read-Host 'The user is assigned a Teams meeting policy that does not allow to record meetings. Would you like to allow recording meetings? [Y/N]'
  if ($toEnableCloudRecording -eq 'Y') {
    Set-CsTeamsMeetingPolicy -Identity $policyName -AllowCloudRecording $true
    Write-Host -ForegroundColor Green 'The user is now assigned a Teams meeting policy that allows to record meetings.'
  } else {
    return Write-Host -ForegroundColor Red 'The user is assigned a Teams meeting policy that does not allow to record meetings.'
  }
}

$recordingStorageMode = $meetingPolicy.RecordingStorageMode
Write-Host 'Checking the recording storage location:'

if ($recordingStorageMode -eq 'OneDriveForBusiness') {
  Write-Host 'The meeting recordings for the user are stored in OneDrive/SharePoint.'
} elseif ($recordingStorageMode -eq 'Stream') {
  Write-Host 'The meeting recordings for the user are stored in Stream.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TRANSCRIPT
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\CannotRecordMeetings.txt"
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'User:'$UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) (TMS):"
  Get-CsOnlineUser $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) (MSO):"
  Get-MsolUser -UserPrincipalName $UPN
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