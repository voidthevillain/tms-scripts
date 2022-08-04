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
# .\tms-ProvisionCallingPlanUser.ps1 -UserPrincipalName user@tms-ninja.com -DisplayName userOne -FirstName user -LastName one -PhoneNumber +40744XXXXXX
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# PARAMS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]
    [String]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [String]$DisplayName = $UserPrincipalName.split('@')[0],

    [Parameter(Mandatory = $false)]
    [String]$FirstName = $UserPrincipalName.split('@')[0][0],

    [Parameter(Mandatory = $false)]
    [String]$LastName = $UserPrincipalName.split('@')[0][$UserPrincipalName.split('@')[0].length - 1],

    [Parameter(Mandatory = $true)]
    [String]$PhoneNumber,

    [Parameter(Mandatory = $false)]
    [String]$UsageLocation = 'US'
)

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# GLOBAL CONSTANTS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
$ErrorActionPreference = "SilentlyContinue"
$PBXBaseSKUs = @('O365_BUSINESS_ESSENTIALS', 'O365_BUSINESS_PREMIUM', 'SPB', 'SPE_E3', 'SPE_E5', 'STANDARDPACK', 'ENTERPRISEPACK', 'ENTERPRISEPREMIUM')
# $PSTNSKUs = @('MCOPSTN1', 'MCOPSTN2', 'MCOPSTN5', 'MCOPSTN6', 'MCOPSTN8', 'MCOPSTN9')
# $PBXSKU = 'MCOEV'

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# FUNCTIONS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
function Get-DomainFromUPN {
    param (
        [string]$UPN
    )

    return $UPN.split('@')[1]
}

function Get-VerifiedDomain {
    param (
        [string]$Domain
    )

    $MSODomain = Get-MsolDomain -DomainName $Domain

    if ($MSODomain.Name -eq $Domain) {
        return $true
    }
    else {
        return $false
    }
}

function Get-UserExistsInMSO {
    param (
        [string] $UPN
    )

    $MSOUser = Get-MsolUser -UserPrincipalName $UPN

    if ($MSOUser.UserPrincipalName -eq $UPN) {
        return $true
    }
    else {
        return $false
    }
}

function Get-TenantSKUs {
    return Get-MsolAccountSku
}

function Get-TenantPBXSupportedSKUs {
    param (
        $TenantSKUs
    )

    $tPBXSupportedSKUs = @()

    foreach ($SKU in $TenantSKUs) {
        if ($PBXBaseSKUs -contains $SKU.AccountSkuId.split(':')[1] -AND $SKU.ConsumedUnits -lt $SKU.ActiveUnits) {
            $tPBXSupportedSKUs += $SKU
        }
    }

    return $tPBXSupportedSKUs
}

function Get-TenantPBXSKUs {
    param (
        $TenantSKUs
    )

    $TenantPBXSKUs = @()

    foreach ($SKU in $TenantSKUs) {
        if ($SKU.AccountSkuId.split(':')[1] -eq 'MCOEV') {
            $TenantPBXSKUs += $SKU
        }
    }

    return $TenantPBXSKUs
}

function Get-TenantPSTNSKUs {
    param (
        $TenantSKUs
    )

    $TenantPSTNSKUs = @()

    # 'MCOPSTN1', 'MCOPSTN2', 'MCOPSTN5', 'MCOPSTN6', 'MCOPSTN8', 'MCOPSTN9'
    foreach ($SKU in $TenantSKUs) {
        if ($SKU.AccountSkuId.split(':')[1] -eq 'MCOPSTN1' -OR $SKU.AccountSkuId.split(':')[1] -eq 'MCOPSTN2' -OR $SKU.AccountSkuId.split(':')[1] -eq 'MCOPSTN5' -OR $SKU.AccountSkuId.split(':')[1] -eq 'MCOPSTN6' -OR $SKU.AccountSkuId.split(':')[1] -eq 'MCOPSTN8' -OR $SKU.AccountSkuId.split(':')[1] -eq 'MCOPSTN9' -OR $SKU.AccountSkuId.split(':')[1] -eq 'Microsoft_Teams_Calling_Plan_pay_as_you_go_(country_zone_1)') {
            $TenantPSTNSKUs += $SKU
        }
    }

    return $TenantPSTNSKUs
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
# LOGIC
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.
Write-Host "...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -."
Write-Host "---INPUT VALIDATIONS"
Write-Host "...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -."

Write-Host "Checking if the domain is verified:"
$domain = Get-DomainFromUPN -UPN $UserPrincipalName
$isValidDomain = Get-VerifiedDomain -Domain $domain

if ($isValidDomain) {
    Write-Host -ForegroundColor Green "The domain is verified in the directory."
}
else {
    return Write-Host -ForegroundColor Red "The domain is not verified in the directory."
}

Write-Host "Checking if the user already exists:"
$userExists = Get-UserExistsInMSO -UPN $UserPrincipalName

if (!$userExists) {
    Write-Host -ForegroundColor Green "The user does not exist in the directory."
    Write-Host "...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -."
    Write-Host "---USER CREATION"
    Write-Host "...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -."
    Write-Host "Creating user $($UserPrincipalName) in $($domain) directory:"

    $user = New-MsolUser -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName -FirstName $FirstName -LastName $LastName -UsageLocation $UsageLocation

    Write-Host -ForegroundColor Green "Successfully created user $($user.DisplayName):"
    Write-Host "...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -."
    Write-Host "UserPrincipalName: $($user.UserPrincipalName)"
    Write-Host "DisplayName: $($user.DisplayName)"
    Write-Host "FirstName: $($user.FirstName)"
    Write-Host "LastName: $($user.LastName)"
    Write-Host "UsageLocation: $($user.UsageLocation)"

    Write-Host "...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -."
    Write-Host "---LICENSING"
    Write-Host "...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -."
    Write-Host "Checking if there are supported subscriptions for Teams Phone in the directory:"
    $tenantSKUs = Get-TenantSKUs
    $tenantPBXSupportedSKUs = Get-TenantPBXSupportedSKUs -TenantSKUs $tenantSKUs

    if (!($tenantPBXSupportedSKUs.length -eq 0)) {
        Write-Host -ForegroundColor Green "Found $($tenantPBXSupportedSKUs.length) supported subscriptions for Teams Phone available."

        #check if among these subs there is M365 E5 or O365 E5
        Write-Host "Checking if among these subscriptions there is any that contains Teams Phone:"
        $SKUsWithMCOEV = @()
        

        foreach ($SKU in $tenantPBXSupportedSKUs) {
            if ($SKU.AccountSkuId.split(':')[1] -eq 'SPE_E5' -OR $SKU.AccountSkuId.split(':')[1] -eq 'ENTERPRISEPREMIUM') {
                $SKUsWithMCOEV += $SKU.AccountSkuId
            }
        }

        if (!($SKUsWithMCOEV.length -eq 0)) {
            Write-Host -ForegroundColor Green "Found $($SKUsWithMCOEV.length) subscriptions that contain Teams Phone."
            Write-Host "Assigning subscription $($SKUsWithMCOEV[0].split(':')[1]) to $($UserPrincipalName)."

            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $SKUsWithMCOEV[0]

            # CP NOW
            Write-Host "Checking if there is any Calling Plan license available:"
            $tenantPSTNSKUs = Get-TenantPSTNSKUs -TenantSKUs $tenantSKUs

            if (!($tenantPSTNSKUs.length -eq 0)) {
                Write-Host -ForegroundColor Green "Found $($tenantPSTNSKUs.length) available Calling Plan licenses."
                Write-Host "Assigning $($tenantPSTNSKUs[0].AccountSkuId.split(':')[1]) to $($UserPrincipalName)"

                Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $tenantPSTNSKUs[0].AccountSkuId
            }
            else {
                return Write-Host -ForegroundColor Red "There are no Calling Plan licenses available."
            }

            # Phone number
            # NEED SLEEB BEFORE LICENSING AND PHONE NUMBER ASSIGNMENT ~ 1 minute
            $phoneN = Get-CsPhoneNumberAssignment -TelephoneNumber $PhoneNumber

            Write-Host "Checking if the phone number is present in the inventory:"
            if ($phoneN) {
                Write-Host -ForegroundColor Green "The phone number is present in the inventory."

                Write-Host "Checking if the phone number is of correct type:"
                if ($phoneN.NumberType -eq "CallingPlan") {
                    Write-Host -ForegroundColor Green "The phone number is of correct type."

                    Write-Host "Checking if the number is assigned:"
                    if ($phoneN.PstnAssignmentStatus -eq 'Unassigned') {
                        Write-Host -ForegroundColor Green "The phone number is not assigned."

                        Write-Host "Assigning phone number $($PhoneNumber) to $($UserPrincipalName)."
                        Set-CsPhoneNumberAssignment -Identity $UserPrincipalName -PhoneNumber $PhoneNumber -PhoneNumberType "CallingPlan"

                        Write-Host -ForegroundColor Green "Successfully enabled user $($UserPrincipalName) for Calling Plan!"
                    }
                    else {
                        return Write-Host -ForegroundColor Green "The phone number is already assigned."
                    }
                }
                else {
                    return Write-Host -ForegroundColor Red "The phone number is of incorrect type."
                }
            }
            else {
                return Write-Host -ForegroundColor Red "The phone number $($PhoneNumber) is not present in the inventory."
            }
        }
        else {
            Write-Host "Found no subscriptions that contain Teams Phone."
            Write-Host "Assigning subscription $($tenantPBXSupportedSKUs[0].AccountSkuId) to $($UserPrincipalName)."

            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $tenantPBXSupportedSKUs[0].AccountSkuId

            # MCOEV NOW
            Write-Host "Checking if there is any Teams Phone license available:"
            $tenantPBXSKUs = Get-TenantPBXSKUs -TenantSKUs $tenantSKUs

            if (!($tenantPBXSKUs.length -eq 0)) {
                Write-Host -ForegroundColor Green "Found $($tenantPBXSKUs.length) Teams Phone licenses available."
                Write-Host "Assigning $($tenantPBXSKUs[0].AccountSkuId.split(':')[1]) to $($UserPrincipalName)"
            }
            else {
                return Write-Host -ForegroundColor Red "No Teams Phone licenses available."
            }

            # CP NOW
            Write-Host "Checking if there is any Calling Plan license available:"
            $tenantPSTNSKUs = Get-TenantPSTNSKUs -TenantSKUs $tenantSKUs

            if (!($tenantPSTNSKUs.length -eq 0)) {
                Write-Host -ForegroundColor Green "Found $($tenantPSTNSKUs.length) available Calling Plan licenses."
                Write-Host "Assigning $($tenantPSTNSKUs[0].AccountSkuId.split(':')[1]) to $($UserPrincipalName)"
                
                Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $tenantPSTNSKUs[0].AccountSkuId
            }
            else {
                return Write-Host -ForegroundColor Red "There are no Calling Plan licenses available."
            }

            # phone number
            # NEED SLEEB BEFORE LICENSING AND PHONE NUMBER ASSIGNMENT ~ 1 minute
            $phoneN = Get-CsPhoneNumberAssignment -TelephoneNumber $PhoneNumber

            Write-Host "Checking if the phone number is present in the inventory:"
            if ($phoneN) {
                Write-Host -ForegroundColor Green "The phone number is present in the inventory."
            
                Write-Host "Checking if the phone number is of correct type:"
                if ($phoneN.NumberType -eq "CallingPlan") {
                    Write-Host -ForegroundColor Green "The phone number is of correct type."
            
                    Write-Host "Checking if the number is assigned:"
                    if ($phoneN.PstnAssignmentStatus -eq 'Unassigned') {
                        Write-Host -ForegroundColor Green "The phone number is not assigned."
            
                        Write-Host "Assigning phone number $($PhoneNumber) to $($UserPrincipalName)."
                        Set-CsPhoneNumberAssignment -Identity $UserPrincipalName -PhoneNumber $PhoneNumber -PhoneNumberType "CallingPlan"
            
                        Write-Host -ForegroundColor Green "Successfully enabled user $($UserPrincipalName) for Calling Plan!"
                    }
                    else {
                        return Write-Host -ForegroundColor Green "The phone number is already assigned."
                    }
                }
                else {
                    return Write-Host -ForegroundColor Red "The phone number is of incorrect type."
                }
            }
            else {
                return Write-Host -ForegroundColor Red "The phone number $($PhoneNumber) is not present in the inventory."
            }
        }

    }
    else {
        return Write-Host -ForegroundColor Red "There are no supported subscriptions for Teams Phone available in the directory."
    }

    # logic to prefer skus with MCOEV inside

}
else {
    return Write-Host -ForegroundColor Red "The user already exists in the directory."
}