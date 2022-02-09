Add-Type -AssemblyName PresentationFramework

# Function Get-MainWindow {
#   $path = "$($HOME)\Documents\WindowsPowerShell"
#   $exists = $false

#   gci -path $path | foreach {
#     if ($_.Name -eq 'MainWindow.xaml') {
#       $exists = $true
#     } 
#   }

#   return $exists
# }

# Function New-MainWindow {
#   $path = "$($HOME)\Documents\WindowsPowerShell"
#   $name = "MainWindow.xaml"


# }

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

Function Get-CustomProfiles {
  $profiles = gci -path "$($env:localappdata)\Microsoft\Teams\CustomProfiles"

  return $profiles
}

# $xamlFile = "C:\Users\voidt\source\repos\Teams - Profile Manager\Teams - Profile Manager\MainWindow.xaml"
$xamlFile = "C:\Users\voidt\Desktop\MainWindow.xaml"

#create window
$inputXML = Get-Content $xamlFile -Raw
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[XML]$XAML = $inputXML

#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    Write-Warning $_.Exception
    throw
}

# Create variables based on form control names.
# Variable will be named as 'var_<control name>'

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)"
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}
Get-Variable var_*

$customProfiles = Get-CustomProfiles

if ($customProfiles) {
  foreach ($cp in $customProfiles) {
    $var_listProfiles.Items.Add($cp.Name)
  }
} else {
  $var_listProfiles.Items.Add("No custom profiles found.")
}

$var_listProfiles.Add_SelectionChanged({
  $var_txtSelectedProfile.Text = $var_listProfiles.selectedItems
  $var_lblSelectedProfile.Content = $var_txtSelectedProfile.Text
})

$var_btnStart.Add_Click({
  $scriptPath = "$($HOME)\Documents\WindowsPowerShell\tms-CustomProfiles.ps1"

  PowerShell.exe -File $scriptPath $var_txtSelectedProfile.Text
})

$var_btnClearCache.Add_Click({
  $path = "$($env:localappdata)\Microsoft\Teams\CustomProfiles\$($var_txtSelectedProfile.Text)\AppData\Roaming\Microsoft\Teams"

  gci -path $path | foreach { Remove-Item $_.FullName -Recurse -Force }

  write-host 'Cache cleared successfully.'
})


$var_btnOpenLocation.Add_Click({
  Invoke-Item "$($env:localappdata)\Microsoft\Teams\CustomProfiles\$($var_txtSelectedProfile.Text)"
})

$var_btnCreate.Add_Click({
  $scriptPath = "$($HOME)\Documents\WindowsPowerShell\tms-CustomProfiles.ps1"
  $profileName = $var_txtNewProfileName.Text

  $var_listProfiles.Items.Add($profileName)

  PowerShell.exe -File $scriptPath $profileName
})


# BAD idea cause permission errors
# $var_btnDelete.Add_Click({
#   $pathToProfile = "$($env:localappdata)\Microsoft\Teams\CustomProfiles\$($var_txtSelectedProfile.Text)"

#   try { 
#     Remove-Item $pathToProfile -Recurse -Force
#   } catch {
#     Write-Host -ForegroundColor Red 'Missing permissions to fully delete this profile.'
#   }
#   Write-Host 'Deleted profile'

#   $var_listProfiles.Items.RemoveAt($var_listProfiles.SelectedIndex)
# })



# $var_btnLaunch.Add_Click({
#   # $scriptPath = "$($HOME)\Documents\WindowsPowerShell\tms-CustomProfiles.ps1"
  
#   # PowerShell.exe -File $scriptPath $selectedProfile
#   write-host $selectedProfile
# })

$Null = $window.ShowDialog()