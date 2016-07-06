#Import the necessary modules for this script.
Import-Module C:\Scripts\SupportScripts\CredentialManager\CredentialManager.psd1
Import-Module O365-Tools
Import-Module ActiveDirectory
#Load the configuration file, credentials and the file containing the license data for use in the script.
[xml]$configFile = Get-Content -Path C:\Scripts\ConfigFile.xml
$o365Credentials = Get-StoredCredential -Target $configFile.Settings.Credentials.Office365Credentials
$Licenses = $configFile.Setings.ScriptSettings.SetOffice365License.LicenseData
#login to Office 365. May need to be changed.
Try {
    Connect-O365Session -Credential $o365Credentials -AzureAD
} Catch {
    Write-Warning "Could not connect to Azure AD/MS Online. Script will abort."
    Break
}

#region Licenses and Unlicense Users

    ForEach ($AccountSkuID in $Licenses.Keys) {
        #Set a few base variables for each sku set in the license file.
        $UsageLocation = $Licenses.$AccountSkuID.Values.UsageLocation
        $Groups = $Licenses.$AccountSkuID.Keys
        #Make an array of all users currently licensed with this SKU. This is handy for when you have multiple options per SKU.
        $CurrentlyLicensedUsers = ((Get-MsolUser -All).Where{ $_.Licenses.AccountSkuID -Contains $AccountSkuID }).UserPrincipalName
        #Create an empty array to store all users that should be licensed.
        $ReferenceLicensedUsers = @()

        ForEach ($Group in $Groups) {
            #Retrieve the disabled plans. Then create license objects for use with the Set-MsolUserLicense cmdlet later on in this script.
            $DisabledPlans = $Licenses.$AccountSkuID.$Group.DisabledPlans
            $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuID -DisabledPlans $DisabledPlans
            #Get all users who should have a license. Add their userprincipalnames to the ReferenceLicensedUsers array.
            $LicensedADUsers = (Get-ADGroupMember -Identity $Group -Recursive | Get-ADUser).UserPrincipalName
            $ReferenceLicensedUsers += $LicensedADUsers
            
            ForEach ($ADUser in $LicensedADUsers) {
                #Set a few base variables for each of the users found in the group.
                $MsolUser = Get-MsolUser -UserPrincipalName $ADUser
                $CurrentUsageLocation = $MsolUser.UsageLocation
                $CurrentUserLicenses = $MsolUser.Licenses.AccountSkuId
                
                #Set the usage location to the correct value if incorrect.
                If ($CurrentUsageLocation -ne $UsageLocation) {
                    Set-MsolUser -UserPrincipalName $MsolUser.UserPrincipalName -UsageLocation $UsageLocation -Verbose
                }
                #Set the license options to the correct value if this license is already assigned.
                If ($CurrentUserLicenses -contains $AccountSkuID){
                    Set-MsolUserLicense -UserPrincipalName $MsolUser.UserPrincipalName -LicenseOptions $LicenseOptions -Verbose
                }
                #Set the license and license options if the SKU hasn't been assigned yet.
                If ($CurrentUserLicenses -notcontains $AccountSkuID) {
                    Set-MsolUserLicense -UserPrincipalName $MsolUser.UserPrincipalName -AddLicenses $AccountSkuID -LicenseOptions $LicenseOptions -Verbose
                }
            }
        }
        #Test if a user who is currently licensed is in the list of users who should be licensed. If not, remove this SKU.
        ForEach ($CurrentlyLicensedUser in $CurrentlyLicensedUsers) {
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