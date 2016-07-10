Function Set-O365License {
    
    [CmdletBinding()]
    Param (
        [Parameter()]
        [String]$ConfigFile,

        [Parameter()]
        [ValidateSet('ActiveDirectory','AzureAD')]
        [String]$LicenseSource = 'AzureAD',

        [Parameter()]
        [PSCredential]$Credential

    )
    BEGIN {
        #Load the configuration file, credentials and the file containing the license data for use in the script.
        $ConfigData = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        #Login to Office 365. May need to be changed.
        Try {
            Connect-MsolService -Credential $Credential
        } Catch {
            Write-Warning "Could not connect to Azure AD/MS Online. Script will abort."
            Break
        }
        $AllMsolUser = Get-MsolUser -All
    }
    PROCESS {
        ForEach ($AccountSkuID in $ConfigData.Licenses.PSObject.Properties.Name) {
            #Set a few base variables for each sku set in the license file.
            [System.Collections.ArrayList]$Groups = @()
            #Create an empty array to store all users that should be licensed.
            [System.Collections.ArrayList]$ReferenceLicensedUsers = @()
            #Create an empty array to store all users that are currently licensed.
            [System.Collections.ArrayList]$DifferenceLicensedUsers = @()
            $Groups = [Array]$ConfigData.Licenses.$AccountSkuID.PSObject.Properties.Name
            #Fill an array with all users currently licensed with this SKU. This is handy for when you have multiple options per SKU.
            $DifferenceLicensedUsers = [Array]($AllMsolUser.Where{ $_.Licenses.AccountSkuID -Contains $AccountSkuID }).UserPrincipalName
            <#
            $NetRights = @{}
            ForEach ($Group in $Groups) {
                $DisabledPlans = $ConfigData.Licenses.$AccountSkuID.$Group.DisabledPlans
                $Users = (Get-ADGroupMember -Identity $Group -Recursive | Get-ADUser).UserPrincipalName
                
                ForEach ($User in $Users) {
                    $NetRights.Add($user,$DisabledPlans)
                }
            }
            #>
            ForEach ($Group in $Groups) {
                #Retrieve the usage location for this group.
                $UsageLocation = $ConfigData.Licenses.$AccountSkuID.$Group.UsageLocation
                #Retrieve the disabled plans. Then create license objects for use with the Set-MsolUserLicense cmdlet later on in this script.
                $DisabledPlans = $ConfigData.Licenses.$AccountSkuID.$Group.DisabledPlans
                #Create the license options objects to assign to users.
                $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuID -DisabledPlans $DisabledPlans
                #Get all users who should have a license. Add their userprincipalnames to the ReferenceLicensedUsers array.
                If ($LicenseSource -eq 'ActiveDirectory') {
                    $LicensedUsers = (Get-ADGroupMember -Identity $Group -Recursive | Get-ADUser).UserPrincipalName
                }
                ElseIf ($LicenseSource -eq 'AzureAD') {
                    $GroupId = (Get-MsolGroup -All).Where{ $_.DisplayName -eq $Group }
                    $LicensedUsers = (Get-MsolGroupMember -GroupObjectId $GroupId.ObjectId -All).EmailAddress
                }
                If ($LicensedUsers) {
                    $ReferenceLicensedUsers.AddRange($LicensedUsers)
                }
                ForEach ($LicensedUser in $LicensedUsers) {
                    #Set a few base variables for each of the users found in the group.
                    $MsolUser = $AllMsolUser.Where{ $_.UserPrincipalName -eq $LicensedUser }
                    $CurrentUsageLocation = $MsolUser.UsageLocation
                    $CurrentUserLicenses = $MsolUser.Licenses.AccountSkuId
                    $CurrentUserOptions = $MsolUser.Licenses.ServiceStatus
                    #Set the usage location to the correct value if incorrect.
                    If ($CurrentUsageLocation -ne $UsageLocation) {
                        Set-MsolUser -UserPrincipalName $MsolUser.UserPrincipalName -UsageLocation $UsageLocation -Verbose
                    }
                    #Set the license options to the correct value if this license is already assigned.
                    If ($CurrentUserLicenses -contains $AccountSkuID) {
                        Set-MsolUserLicense -UserPrincipalName $MsolUser.UserPrincipalName -LicenseOptions $LicenseOptions -Verbose
                    }
                    #Set the license and license options if the SKU hasn't been assigned yet.
                    ElseIf ($CurrentUserLicenses -notcontains $AccountSkuID) {
                        Set-MsolUserLicense -UserPrincipalName $MsolUser.UserPrincipalName -AddLicenses $AccountSkuID -LicenseOptions $LicenseOptions -Verbose
                    }
                }
            }
            #Test if a user who is currently licensed is in the list of users who should be licensed. If not, remove this SKU.
            ForEach ($CurrentlyLicensedUser in $DifferenceLicensedUsers) {
                If ($CurrentlyLicensedUser -notin $ReferenceLicensedUsers) {
                    Set-MsolUserLicense -UserPrincipalName $CurrentlyLicensedUser -RemoveLicenses $AccountSkuID
                }
            }
        }
    }
    END {
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
    }
}
