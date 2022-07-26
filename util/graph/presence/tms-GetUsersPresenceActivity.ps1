# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USAGE: 
# cd PATH_TO_SCRIPT
# Connect-MicrosoftTeams
# Select-MgProfile -Name "beta"
# Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All", "Presence.Read", "Presence.Read.All"
# Import-Module Microsoft.Graph.CloudCommunications
# .\tms-GetUsersPresenceAvailability.ps1
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

# get a list of users that have Teams
$users = get-csonlineuser | ? {$_.assignedplan.capability -contains 'Teams'} | select userprincipalname, identity
$presences = @()
$size = $users.count

if ($size % 1500 -eq 0) {
    Write-Host 'Users count above throttle'
    $count = 0

    foreach ($user in $users) {
        $count += 1
        $userPresence = Get-MgUserPresence -UserId $user.Identity
        $newUser = New-Object -TypeName PSObject
        $newUser | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $user.UserPrincipalName
        $newUser | Add-Member -MemberType NoteProperty -Name Id -Value $user.Identity
        $newUser | Add-Member -MemberType NoteProperty -Name Availability -Value $userPresence.Availability
        $newUser | Add-Member -MemberType NoteProperty -Name Activity -Value $userPresence.Activity

        if ($count % 1500 -eq 0) {
            Write-Host 'Reached throttle, sleeping for 32 seconds'
            Start-Sleep -Seconds 32
        }

        $presences += $newUser
    }
} else {
    Write-Host 'Users count below throttle'
    foreach ($user in $users) {
        $userPresence = Get-MgUserPresence -UserId $user.Identity
        $newUser = New-Object -TypeName PSObject
        $newUser | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $user.UserPrincipalName
        $newUser | Add-Member -MemberType NoteProperty -Name Id -Value $user.Identity
        $newUser | Add-Member -MemberType NoteProperty -Name Availability -Value $userPresence.Availability
        $newUser | Add-Member -MemberType NoteProperty -Name Activity -Value $userPresence.Activity

        $presences += $newUser
    }
}

$presences

# export a CSV
# $presences | export-csv '.\presences.csv'