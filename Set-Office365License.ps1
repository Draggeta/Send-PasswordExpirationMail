Import-Module C:\Scripts\SupportScripts\CredentialManager\CredentialManager.psd1
Import-Module C:\Scripts\SupportScripts\Office365CustomTools
Import-Module ActiveDirectory

[xml]$configFile = Get-Content -Path C:\Scripts\ConfigFile.xml
$o365Credentials = Get-StoredCredential -Target $configFile.Settings.Credentials.Office365Credentials

Login-Office365 -Credential $o365Credentials -MsOnline

#region Specify Licenses

    $Licenses = @{ 
        
        'tenantname:STANDARDWOFFPACK_FACULTY' = @{
        
            'sg_o365_disabled-education_license' = @{

                DisabledPlans = 'PROJECTWORKMANAGEMENT','SWAY','INTUNE_O365','YAMMER_EDU','SHAREPOINTWAC_EDU','MCOSTANDARD','SHAREPOINTSTANDARD_EDU'
                UsageLocation = 'NL'

            }
            'sg_o365_service-education_license' = @{

                DisabledPlans = 'PROJECTWORKMANAGEMENT','SWAY','INTUNE_O365','YAMMER_EDU','SHAREPOINTWAC_EDU','MCOSTANDARD','SHAREPOINTSTANDARD_EDU'
                UsageLocation = 'NL'

            }
            'sg_o365_staff-education_license' = @{

                DisabledPlans = ''
                UsageLocation = 'NL'

            }
        
        }
        'tenantname:ENTERPRISEPACK_FACULTY' = @{
        
            'sg_o365_staff-education-e3_license' = @{

                DisabledPlans = 'RMS_S_ENTERPRISE'
                UsageLocation = 'NL'

            }
        
        }
        'tenantname:STANDARDWOFFPACK_STUDENT' = @{
        
            'sg_o365_student-education_license' = @{

                DisabledPlans = 'PROJECTWORKMANAGEMENT','INTUNE_O365','YAMMER_EDU','SHAREPOINTWAC_EDU','SHAREPOINTSTANDARD_EDU'
                UsageLocation = 'NL'

            }

        }
        'tenantname:OFFICESUBSCRIPTION_STUDENT' = @{

            'sg_o365_student-proplus_license' = @{

                DisabledPlans = 'INTUNE_O365'
                UsageLocation = 'NL'

            }

        }
        
    }

#endregion

#region Licenses and Unlicense Users

    foreach ($AccountSkuID in $Licenses.Keys) {
        
        $UsageLocation = 'NL'
        $Groups = $Licenses.$AccountSkuID.Keys
        $CurrentlyLicensedUsers = ((Get-MsolUser -All).Where{ $_.Licenses.AccountSkuID -Contains $AccountSkuID }).UserPrincipalName
        $ReferenceLicensedUsers = @()

        foreach ($Group in $Groups) {

            $DisabledPlans = $Licenses.$AccountSkuID.$Group.DisabledPlans
                        
            $LicensedADUsers = (Get-ADGroupMember -Identity $Group -Recursive | Get-ADUser).UserPrincipalName
            $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuID -DisabledPlans $DisabledPlans
            
            foreach ($ADUser in $LicensedADUsers) {
                
                $MsolUser = Get-MsolUser -UserPrincipalName $ADUser
                $CurrentUsageLocation = $MsolUser.UsageLocation
                $CurrentUserLicenses = $MsolUser.Licenses.AccountSkuId
                
                If ($CurrentUsageLocation -ne $UsageLocation) {
                    Set-MsolUser -UserPrincipalName $MsolUser.UserPrincipalName -UsageLocation $UsageLocation -Verbose
                }
                If ($CurrentUserLicenses -contains $AccountSkuID){
                    Set-MsolUserLicense -UserPrincipalName $MsolUser.UserPrincipalName -LicenseOptions $LicenseOptions -Verbose
                }
                If ($CurrentUserLicenses -notcontains $AccountSkuID) {
                    Set-MsolUserLicense -UserPrincipalName $MsolUser.UserPrincipalName -AddLicenses $AccountSkuID -LicenseOptions $LicenseOptions -Verbose
                }

            }

            $ReferenceLicensedUsers += $LicensedADUsers

        }
        foreach ($CurrentlyLicensedUser in $CurrentlyLicensedUsers) {
            
            If ($CurrentlyLicensedUser -notin $ReferenceLicensedUsers) {
                
                Set-MsolUserLicense -UserPrincipalName $CurrentlyLicensedUser -RemoveLicenses $AccountSkuID

            }
        
        }

    }

#endregion

#region Log and Mail
<#
#region Mail Parameters

[string]$body = @(
    if (!$eventlogerrors) {"The script ran successfully. No errors occured. Any changes made will be listed below.`n"}
    elseif ($eventlogerrors) {"The script completed with errors. Any changes and errors will be listed below.`n"}
    if ($eventlogmovedusers) {"Moved the following accounts into the correct OU:`n $eventlogmovedusers"}
    if ($eventlogaddedusers) {"Created the following accounts:`n $eventlogaddedusers"}
    if ($eventlogrenamedusers) {"Renamed the following accounts:`n $eventlogdeletedusers"}
    if ($eventlogdisabledusers) {"Disabled the following accounts:`n $eventlogdisabledusers"}
    if ($eventlogdeletedusers) {"Deleted the following accounts:`n $eventlogdeletedusers"}
    if ($eventlogerrors) {"The following errors occured:`n $eventlogerrors"}
)
[string]$subject = @(
    if (!$eventlogerrors) {"[Success] Manage-Student: Script ran succesfully"}
    elseif ($eventlogerrors) {"[Warning] Manage-Student: Script completed with errors"}
)
$emailparams += @{
    Subject = $subject
    Body = $body
}

#endregion

Send-MailMessage @emailparams
#>
#endregion