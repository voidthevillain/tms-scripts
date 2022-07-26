# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

# get a list of users that have Teams
$users = get-csonlineuser | ? {$_.assignedplan.capability -contains 'Teams'} | select userprincipalname, identity
$usersOffline = @()
$size = $users.count

# bypass 1500 request x 30s throttle
if ($size % 1500 -eq 0) {
    Write-Host 'Users count above throttle'
    $count = 0

    foreach ($user in $users) {
        $count += 1
        $userPresence = Get-MgUserPresence -UserId $user.Identity

        if ($count % 1500 -eq 0) {
            Write-Host 'Reached throttle, sleeping for 32 seconds'
            Start-Sleep -Seconds 32
        }

        if ($userPresence.Availability -eq 'Offline' -AND $userPresence.Activity -eq 'OffWork') {
            $usersOffline += $user.userprincipalname
        }
    }
} else {
    Write-Host 'Users count below throttle'
    foreach ($user in $users) {
        $userPresence = Get-MgUserPresence -UserId $user.Identity

        if ($userPresence.Availability -eq 'Offline' -AND $userPresence.Activity -eq 'OffWork') {
            $usersOffline += $user.userprincipalname
        }
    }
}

Write-Host "The following users have their presence manually set to offline:"
Write-Host "$($usersOffline)"