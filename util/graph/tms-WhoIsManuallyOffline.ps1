# get a list of users that have Teams
$users = get-csonlineuser | ? {$_.assignedplan.capability -contains 'Teams'} | select userprincipalname, identity
$usersOffline = @()
$size = $users.count

# PS C:\Users\a-mihaifilip> Select-MgProfile -Name "beta"
# PS C:\Users\a-mihaifilip> Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All", "Presence.Read", "Presence.Read.All"
# Welcome To Microsoft Graph!
# PS C:\Users\a-mihaifilip> Import-Module Microsoft.Graph.CloudCommunications
# PS C:\Users\a-mihaifilip> $userId = "578ba82c-d809-4094-992c-37144af98b51"
# PS C:\Users\a-mihaifilip> Get-MgUserPresence -UserId $userId

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