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
        #Login to Office 365. May need to be changed to use the Azure AD preview cmdlets. Stop execution if logging in
        #fails. May require tests as well to check if the module(s) are installed.
        Try {
            Connect-MsolService -Credential $Credential
        } Catch {
            Write-Warning "Could not connect to Azure AD. Script will abort."
            Break
        }
    }
    PROCESS {
        ForEach ($AccountSkuID in $ConfigData.Licenses.PSObject.Properties.Name) {
            #Create an empty array to store all groups that grant this license.
            [System.Collections.ArrayList]$Groups = @()
            #Create an empty array to store all users that are currently licensed.
            [System.Collections.ArrayList]$CurrentlyLicensedUsers = @()

            #Query for all users. This allows the comparison between users that should have licenses and the actually
            #assigned licenses. Also reduces the amount of Get-MsolUser commands run.
            $AllMsolUser = Get-MsolUser -All
            #Retrieve the usage location for this license.
            $UsageLocation = $ConfigData.Licenses.$AccountSkuID.UsageLocation
            #Supersedence was implemented because sometimes different licenses have options which don't work well
            #together such as two licenses that grant the user an Exchange account. In those cases, the simple solution
            #for now is to remove the superseded license.
            #Retrieve the licenses that are superseded by this license.
            $SupersededLicenses = $ConfigData.Licenses.$AccountSkuID.Supersedes
            #Retrieve the licenses that supersede this license.
            $SupersededByLicenses = $ConfigData.Licenses.$AccountSkuID.SupersededBy

            #Fill the $Groups array/variable with all groups that assign this license.
            $Groups.AddRange($ConfigData.Licenses.$AccountSkuID.Groups.PSObject.Properties.Name)
            #Fill the $CurrentlyLicensedUsers array with all users currently licensed with this SKU. This is used for 
            #when you have multiple license options per SKU that can be assigned to one user.
            $CurrentlyLicensedUsers.AddRange(($AllMsolUser.Where{ $_.Licenses.AccountSkuID -Contains $AccountSkuID }).UserPrincipalName)
            
            #Create an empty hashtable to store users and their net license options for this license SKU.
            $LicensedUsers = @{}

            #Find all users and their net license options by comparing the licenses between groups and discarding all
            #options that don't appear in all license options.
            ForEach ($Group in $Groups) {
                #Find all members of the currently iterated group. Different commands are run depending on if AzureAD
                #or Active Directory is used as primary source.
                If ($LicenseSource -eq 'ActiveDirectory') {
                    $Members = (Get-ADGroupMember -Identity $Group -Recursive | Get-ADUser).UserPrincipalName
                    #$Members = Get-ADObject -LDAPFilter "(memberOf=$($(Get-ADGroup $Group).DistinguishedName))" -Properties UserPrincipalName
                }
                ElseIf ($LicenseSource -eq 'AzureAD') {
                    $GroupId = (Get-MsolGroup -All).Where{ $_.DisplayName -eq $Group }
                    $Members = (Get-MsolGroupMember -GroupObjectId $GroupId.ObjectId -All).EmailAddress
                }
                #Get all users who should have a license. Add their UserPrincipalNames to the $LicensedUsers hashtable.
                ForEach ($Member in $Members) {
                    #If the user isn't in the hashtable yet, add them and the license options denied by this group.
                    If (-not $LicensedUsers.ContainsKey($Member)) {
                        $LicensedUsers.Add($Member, $ConfigData.Licenses.$AccountSkuID.Groups.$Group.DisabledPlans)
                    }
                    #If the user is already in the hashtable, compare the options and keep only the ones that
                    #are equal between license options.
                    ElseIf ($LicensedUsers.ContainsKey($Member)) {
                        $CompareArray = Compare-Object -ReferenceObject $LicensedUsers.Get_Item($Member) -DifferenceObject $ConfigData.Licenses.$AccountSkuID.Groups.$Group.DisabledPlans -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                        $LicensedUsers.Set_Item($Member, $CompareArray.InputObject)
                    }
                }
            }

            #Test if a user who is currently licensed is in the list of users who should be licensed. If not, remove 
            #this SKU.
            ForEach ($User in $CurrentlyLicensedUsers) {
                If ($LicensedUsers.ContainsKey($User) -eq $False) {
                    Set-MsolUserLicense -UserPrincipalName $CurrentlyLicensedUser -RemoveLicenses $AccountSkuID
                }
            }

            #Assign and revoke (based on supersedence) the license and net license options. 
            ForEach ($LicensedUser in $LicensedUsers.GetEnumerator()) {
                #Set a few base variables for each of the users found in the group.
                $CurrentUser = $AllMsolUser.Where{ $_.UserPrincipalName -eq $LicensedUser }
                $CurrentUsageLocation = $CurrentUser.UsageLocation
                $CurrentUserLicenses = $CurrentUser.Licenses.AccountSkuId
                $CurrentUserOptions = $CurrentUser.Licenses.ServiceStatus
                #Set the usage location to the correct value if incorrect.
                If ($CurrentUsageLocation -ne $UsageLocation) {
                    Set-MsolUser -UserPrincipalName $CurrentUser.UserPrincipalName -UsageLocation $UsageLocation -Verbose
                }
                #Compare the currently assigned and "superseded by" licenses. If there is no match, assign the license.
                $SkipLicense = Compare-Object -ReferenceObject $SupersededByLicenses -DifferenceObject $CurrentUserLicenses -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                If ((-not $SupersededByLicenses) -or (-not $SkipLicense)) {
                    #Splat the default paramters used in Set-MsolUserLicense.
                    $SetMsolUserLicenseParams = @{
                        UserPrincipalName = $CurrentUser.UserPrincipalName
                        LicenseOptions = (New-MsolLicenseOptions -AccountSkuId $AccountSkuID -DisabledPlans $LicensedUser.Value -Verbose)
                    }
                    #If the user has a superseded license configured, add a remove parameter to the splat to remove
                    #this license.
                    $RemoveSupersededLicense = Compare-Object -ReferenceObject $SupersededLicenses -DifferenceObject $CurrentUserLicenses -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                    If ($RemoveSupersededLicense) {
                        $SetMsolUserLicenseParams.RemoveLicenses = $RemoveSupersededLicense.InputObject 
                    }
                    #If the user hasn't been granted the license yet, add an add parameter to the splat to add the
                    #license to this user.
                    If ($CurrentUserLicenses -notcontains $AccountSkuID) {
                        $SetMsolUserLicenseParams.AddLicenses = $AccountSkuID
                    }
                    #Run the command with the required parameters and, if available, the optional ones.
                    Set-MsolUserLicense @SetMsolUserLicenseParams
                }
                #If there is a match, don't assign the license as it is superseded.
                ElseIf ($SkipLicense) {
                    Write-Output "Skipped license $AccountSkuID as it was superseded for this user by license(s) $($SkipLicense.InputObject)"
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
