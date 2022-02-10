# Functions
Function Get-ProfileScript {
  $path = "$($HOME)\Documents\WindowsPowerShell"
  $exists = $false

  gci -path $path | foreach {
    if ($_.Name -eq 'tms-CustomProfiles.ps1') {
      $exists = $true
    } 
  }

  return $exists
}

Function New-ProfileScript {
  $path = "$($HOME)\Documents\WindowsPowerShell"
  $name = "tms-CustomProfiles.ps1"

  $url = "https://raw.githubusercontent.com/voidthevillain/tms-resources/main/util/custom-profiles/tms-CustomProfiles.ps1"

  try {
    (curl $url).Content | Out-File -FilePath "$($path)\$($name)"
    return $true
  } catch {
    return $false
  }
}

Function Get-CustomProfiles {
  $profiles = gci -path "$($env:localappdata)\Microsoft\Teams\CustomProfiles"

  return $profiles
}

$profileScript = Get-ProfileScript

if ($profileScript) {
  Write-Host -ForegroundColor Green "Profile script exists in $($HOME)\Documents\WindowsPowerShell."
} else {
  Write-Host 'Profile script does not exist. Downloading script from https://github.com/voidthevillain/tms-scripts/blob/main/util/profiles/tms-CustomProfiles.ps1'
  $isCreated = New-ProfileScript
  if ($isCreated) {
    Write-Host -ForegroundColor Green "Successfully downloaded and installed script in $($HOME)\Documents\WindowsPowerShell."
  } else {
    return Write-Host -ForegroundColor Red "Could not download or install script from https://github.com/voidthevillain/tms-scripts/blob/main/util/profiles/tms-CustomProfiles.ps1"
  }
}


<# 
.NAME
    PManager
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$TMSProfileManager               = New-Object system.Windows.Forms.Form
$TMSProfileManager.ClientSize    = New-Object System.Drawing.Point(679,355)
$TMSProfileManager.text          = "TEAMS - Profile Manager"
$TMSProfileManager.TopMost       = $false
$TMSProfileManager.icon          = "$($HOME)\Documents\WindowsPowerShell\teams.ico"
$TMSProfileManager.BackColor     = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$TMSProfileManager.FormBorderStyle    = 'Fixed3D'
$TMSProfileManager.MaximizeBox        = $false
$TMSProfileManager.MinimizeBox        = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Profiles:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(10,10)
$Label1.Font                     = New-Object System.Drawing.Font('Segoe UI',10)

$listProfiles                    = New-Object system.Windows.Forms.ListBox
$listProfiles.text               = "listBox"
$listProfiles.width              = 659
$listProfiles.height             = 200
$listProfiles.location           = New-Object System.Drawing.Point(10,35)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Selected: "
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(10,255)
$Label2.Font                     = New-Object System.Drawing.Font('Segoe UI',10)

$lblSelectedProfile              = New-Object system.Windows.Forms.Label
$lblSelectedProfile.text         = "(no profile selected)"
$lblSelectedProfile.AutoSize     = $true
$lblSelectedProfile.width        = 25
$lblSelectedProfile.height       = 10
$lblSelectedProfile.location     = New-Object System.Drawing.Point(80,255)
$lblSelectedProfile.Font         = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$btnStart                        = New-Object system.Windows.Forms.Button
$btnStart.text                   = "Start"
$btnStart.width                  = 100
$btnStart.height                 = 30
$btnStart.location               = New-Object System.Drawing.Point(346,250)
$btnStart.Font                   = New-Object System.Drawing.Font('Segoe UI',10)

$btnCache                        = New-Object system.Windows.Forms.Button
$btnCache.text                   = "Cache"
$btnCache.width                  = 100
$btnCache.height                 = 30
$btnCache.location               = New-Object System.Drawing.Point(459,250)
$btnCache.Font                   = New-Object System.Drawing.Font('Segoe UI',10)

$btnLocation                     = New-Object system.Windows.Forms.Button
$btnLocation.text                = "Location"
$btnLocation.width               = 100
$btnLocation.height              = 30
$btnLocation.location            = New-Object System.Drawing.Point(569,250)
$btnLocation.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "New profile:"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(10,325)
$Label3.Font                     = New-Object System.Drawing.Font('Segoe UI',10)

$txtNewProfile                   = New-Object system.Windows.Forms.TextBox
$txtNewProfile.multiline         = $false
$txtNewProfile.width             = 151
$txtNewProfile.height            = 20
$txtNewProfile.location          = New-Object System.Drawing.Point(94,320)
$txtNewProfile.Font              = New-Object System.Drawing.Font('Segoe UI',10)

$btnCreate                       = New-Object system.Windows.Forms.Button
$btnCreate.text                  = "Create"
$btnCreate.width                 = 100
$btnCreate.height                = 30
$btnCreate.location              = New-Object System.Drawing.Point(569,316)
$btnCreate.Font                  = New-Object System.Drawing.Font('Segoe UI',10)

$txtProfileName                  = New-Object system.Windows.Forms.TextBox
$txtProfileName.multiline        = $false
$txtProfileName.width            = 100
$txtProfileName.height           = 20
$txtProfileName.visible          = $false
$txtProfileName.location         = New-Object System.Drawing.Point(316,320)
$txtProfileName.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$customProfiles = Get-CustomProfiles

if ($customProfiles) {
  foreach ($cp in $customProfiles) {
    $listProfiles.Items.Add($cp.Name)
  }
} else {
  $listProfiles.Items.Add("No custom profiles found.")
}

$listProfiles.Add_SelectedIndexChanged({
  $txtProfileName.text = $listProfiles.selectedItems
  $lblSelectedProfile.text = $txtProfileName.text
})

$btnStart.Add_Click({
  $scriptPath = "$($HOME)\Documents\WindowsPowerShell\tms-CustomProfiles.ps1"

  if ($txtProfileName.Text -eq "") {
    Write-Host 'No profile selected to start.'
  } else {
    Write-Host 'Starting profile.'
    PowerShell.exe -File $scriptPath $txtProfileName.Text
  } 
})

$btnCache.Add_Click({
  $path = "$($env:localappdata)\Microsoft\Teams\CustomProfiles\$($txtProfileName.Text)\AppData\Roaming\Microsoft\Teams"

  if ($txtProfileName.Text -eq "") {
    Write-Host 'No profile selected to clear cache.'
  } else {
    gci -path $path | foreach { Remove-Item $_.FullName -Recurse -Force }
    Write-Host 'Cache cleared successfully.'
  } 
})

$btnLocation.Add_Click({
  if ($txtProfileName.Text -eq "") {
    Write-Host 'No profile selected to open location.'
  } else {
    Invoke-Item "$($env:localappdata)\Microsoft\Teams\CustomProfiles\$($txtProfileName.Text)"
    Write-Host 'Location opened successfully.'
  }
})

$btnCreate.Add_Click({
  $scriptPath = "$($HOME)\Documents\WindowsPowerShell\tms-CustomProfiles.ps1"
  $profileName = $txtNewProfile.Text

  if ($profileName -eq "") {
    Write-Host 'No profile name to create profile.'
  } else {
    write-host 'New profile created successfully'
    $listProfiles.Items.Add($profileName)
    PowerShell.exe -File $scriptPath $profileName
  }
})

$TMSProfileManager.controls.AddRange(@($Label1,$listProfiles,$Label2,$lblSelectedProfile,$btnStart,$btnCache,$btnLocation,$Label3,$txtNewProfile,$btnCreate,$txtProfileName))

#region Logic 

#endregion

[void]$TMSProfileManager.ShowDialog()