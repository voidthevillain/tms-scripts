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

$presences | export-csv '.\presences.csv'