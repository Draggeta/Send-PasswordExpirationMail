Function Set-O365License {
    
    [CmdletBinding()]
    Param (

        [Parameter(Mandatory = $True)]
        [ValidateScript({ Test-Path -Path $_ })]
        [Alias('ConfigFile')]
        [String]$ConfigurationFilePath,

        [Parameter()]
        [ValidateSet('ActiveDirectory','AzureAD')]
        [String]$LicenseSource = 'AzureAD',

        [Parameter(Mandatory = $True)]
        [PSCredential]$Credential

    )
    BEGIN {

        #region Import required modules
            #Import the modules required to securely pull the passwords from the credential manager, manage ActiveDirectory
            #and Azure AD.
            Import-Module CredentialManager
            Import-Module MSOnline
            If ($LicenseSource -eq 'ActiveDirectory') {
                Import-Module ActiveDirectory
            }
        #endregion

        #region Create empty general logging arrays
            #Create empty arrays for the logs. These logs collect the specific SKU logs so they can be sent at the end.
            [System.Collections.ArrayList]$LogErrorVariable = @()
            [System.Collections.ArrayList]$LogLicensesAssigned = @()
            [System.Collections.ArrayList]$LogLicensesChanged = @()
            [System.Collections.ArrayList]$LogLicensesRemoved = @()
            [System.Collections.ArrayList]$LogSupersededAssigned = @()
            [System.Collections.ArrayList]$LogSupersededRemoved = @()
        #endregion

        #region Load configuration from file
            #Load the configuration file, and convert from JSON. If it fails, stop the script execution.
            Try {
                $ConfigData = Get-Content $ConfigurationFilePath -Raw | ConvertFrom-Json
            }
            Catch [ArgumentException] {
                #$LogErrorVariable = "Error Message: $($_.Exception.Message)`nFailed Item: $($_.Exception.ItemName)"
                Break
            }
            #Set the email parameters.
            $EmailParams = @{}
            $EmailParams.From = $ConfigData.Settings.EmailServerSettings.From
            $EmailParams.SmtpServer =  $ConfigData.Settings.EmailServerSettings.SmtpServer
            $EmailParams.Port = $ConfigData.Settings.EmailServerSettings.Port
            $EmailParams.UseSsl = $ConfigData.Settings.EmailServerSettings.UseSsl
            $EmailParams.Credential = Get-StoredCredential -Target $ConfigData.Settings.Credentials.MailCredentials
        #endregion

        #region Log in to Azure AD
            #Login to Office 365. May need to be changed to use the Azure AD preview cmdlets. Stop execution if logging in
            #fails. May require tests as well to check if the module(s) are installed.
            Try {
                Connect-MsolService -Credential $Credential
            } 
            Catch {
                $LogErrorVariable = "Error Message: $($_.Exception.Message)`nFailed Item: $($_.Exception.ItemName)"
                $EmailParams.Subject = '[Error] Set-O365License: Failed to log in to Azure AD'
                $EmailParams.Body = "Logging in to Azure AD failed. See detailed error(s) below.`n $LogErrorVariable"
                $EmailParams.To = $ConfigData.Settings.EmailServerSettings.To
                Send-MailMessage @EmailParams
                Break
            }
        #endregion

    }
    PROCESS {
        
        #For each license in the configuration file, process it.
        ForEach ($AccountSkuID in $ConfigData.Licenses.PSObject.Properties.Name) {
            #region Create empty per SKU log arrays
                #Create empty arrays for the SKU logs. These logs are added to the earlier created log arrays.
                [System.Collections.ArrayList]$LogSkuLicensesAssigned = @()
                [System.Collections.ArrayList]$LogSkuLicensesChanged = @()
                [System.Collections.ArrayList]$LogSkuLicensesRemoved = @()
                [System.Collections.ArrayList]$LogSkuSupersededAssigned = @()
                [System.Collections.ArrayList]$LogSkuSupersededRemoved = @()
                #Create an empty array to store all groups that grant this license.
                [System.Collections.ArrayList]$Groups = @()
                #Create an empty array to store all users that are currently licensed.
                [System.Collections.ArrayList]$CurrentlyLicensedUsers = @()
            #region

            #region Examine available licenses
                #Query for all users. This allows the comparison between users that should have licenses and the 
                #actually assigned licenses. Also reduces the amount of Get-MsolUser commands run.
                $AllMsolUser = Get-MsolUser -All
                #Retrieve the usage location for this license.
                $UsageLocation = $ConfigData.Licenses.$AccountSkuID.UsageLocation
                #Supersedence was implemented because sometimes different licenses have options which don't work well
                #together, such as two licenses that grant the user an Exchange account. In those cases, the simple
                #solution for now is to remove the superseded license.
                #Retrieve the licenses that are superseded by this license.
                $SupersededLicenses = $ConfigData.Licenses.$AccountSkuID.Supersedes
                #Retrieve the licenses that supersede this license.
                $SupersededByLicenses = $ConfigData.Licenses.$AccountSkuID.SupersededBy

                #Fill the $Groups array/variable with all groups that assign this license.
                $Groups.AddRange([Array]$ConfigData.Licenses.$AccountSkuID.Groups.PSObject.Properties.Name)
                #Fill the $CurrentlyLicensedUsers array with all users currently licensed with this SKU. This is used
                #for when you have multiple license options per SKU that can be assigned to one user.
                $CurrentlyLicensedUsers.AddRange([Array]($AllMsolUser.Where{ $_.Licenses.AccountSkuID -Contains $AccountSkuID }).UserPrincipalName)
            #endregion
            
            #region Examine current/reference licenses
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
                    #This currently doesn't do nested groups.
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
                        #If the user is already in the hashtable, compare the options and keep only the ones that are
                        #equal between license options.
                        ElseIf ($LicensedUsers.ContainsKey($Member)) {
                            $CompareArray = Compare-Object -ReferenceObject $LicensedUsers.Get_Item($Member) -DifferenceObject $ConfigData.Licenses.$AccountSkuID.Groups.$Group.DisabledPlans -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                            $LicensedUsers.Set_Item($Member, $CompareArray.InputObject)
                        }
                    }
                }
            #endregion

            #region Add/remove license
                #Test if a user who is currently licensed is in the list of users who should be licensed. If not, remove 
                #this SKU.
                ForEach ($CurrentlyLicensedUser in $CurrentlyLicensedUsers) {
                    If ($LicensedUsers.ContainsKey($CurrentlyLicensedUser) -eq $False) {
                        Set-MsolUserLicense -UserPrincipalName $CurrentlyLicensedUser -RemoveLicenses $AccountSkuID -ErrorVariable LogErrorVariable
                        $LogSkuLicensesRemoved.Add("$CurrentlyLicensedUser`n")
                    }
                }

                #Assign and revoke (based on supersedence) the license and net license options. 
                ForEach ($LicensedUser in $LicensedUsers.GetEnumerator()) {
                    #Set a few base variables for each of the users found in the group.
                    $CurrentUser = $AllMsolUser.Where{ $_.UserPrincipalName -eq $LicensedUser.Key }
                    $CurrentUsageLocation = $CurrentUser.UsageLocation
                    $CurrentUserLicenses = $CurrentUser.Licenses.AccountSkuId
                    $CurrentUserOptions = $CurrentUser.Licenses.Where{ $_.AccountSkuId -eq $AccountSkuID }.ServiceStatus.Where{ $_.ProvisioningStatus -eq 'Disabled' }.ServicePlan.ServiceName
                    #$CurrentUserOptions = $CurrentUser.Licenses.Where{ $_.AccountSkuId -eq $AccountSkuID }.ServiceStatus.Where{ $_.ProvisioningStatus -eq 'Disabled' -or $_.ProvisioningStatus -eq 'PendingActivation' }.ServicePlan.ServiceName
                    #Set the usage location to the correct value if incorrect.
                    If ($CurrentUsageLocation -ne $UsageLocation) {
                        Set-MsolUser -UserPrincipalName $CurrentUser.UserPrincipalName -UsageLocation $UsageLocation -ErrorVariable LogErrorVariable
                    }
                    #Compare the currently assigned and "superseded by" licenses. If there is no match, assign the license.
                    #Check if the user has the license already assigned.
                    $AssignLicenses = $CurrentUserLicenses -notcontains $AccountSkuID
                    #Check if there are any superseded licenses currently assigned to the user. Return only the ones that
                    #appear in both arrays. These superseded licenses will be removed.
                    $RemoveSupersededLicenses = Compare-Object -ReferenceObject $SupersededLicenses -DifferenceObject $CurrentUserLicenses -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                    #Check if any licenses that supersede this license are assigned to the user. If that is the case, 
                    #skip adding this license.
                    $SkippedLicenses = Compare-Object -ReferenceObject $SupersededByLicenses -DifferenceObject $CurrentUserLicenses -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                    #If the license options are both in teh config file and currently assigned to the user, check if there
                    #are differences or not. If there are differences, the license options will be set.
                    If ($LicensedUser.Value -and $CurrentUserOptions) {
                        $ChangeLicensesOptions = Compare-Object -ReferenceObject $LicensedUser.Value -DifferenceObject $CurrentUserOptions
                    }
                    #Else if only one of the variables has options, then options need to be changed. If that is the case,
                    #just set the variable to true.
                    ElseIf ($LicensedUser.Value -or $CurrentUserOptions) {
                        $ChangeLicensesOptions = $True
                    }
                    #If the license doesn't needs to be skipped, and either the license has not been assigned, is not
                    #superseded or won't be superseded by this license, run the script below.
                    If ((-not $SkippedLicenses) -and ($AssignLicenses -or $RemoveSupersededLicenses -or $ChangeLicensesOptions)) {
                        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuID -DisabledPlans $LicensedUser.Value -ErrorVariable LogErrorVariable
                        #Splat the default paramters used in Set-MsolUserLicense.
                        $SetMsolUserLicenseParams = @{}
                        $SetMsolUserLicenseParams.UserPrincipalName = $CurrentUser.UserPrincipalName
                        #If license options need to be changed, add the parameter/value to the splat variable.
                        If ($ChangeLicensesOptions) {
                            $SetMsolUserLicenseParams.LicenseOptions = $LicenseOptions
                            If ($LicenseOptions.DisabledServicePlans) {
                                $LogSkuLicensesChanged.Add("$($LicensedUser.Key) - $($LicenseOptions.DisabledServicePlans)`n")
                            }
                            ElseIf (-not $LicenseOptions.DisabledServicePlans) {
                                $LogSkuLicensesChanged.Add("$($LicensedUser.Key) - No disabled options`n")
                            }
                        }
                        #If the user has a superseded license configured, add a remove parameter to the splat to remove
                        #this license.
                        If ($RemoveSupersededLicenses) {
                            $SetMsolUserLicenseParams.RemoveLicenses = $RemoveSupersededLicenses.InputObject
                            $LogSkuSupersededRemoved.Add("$($LicensedUser.Key) - $($RemoveSupersededLicenses.InputObject)`n")
                        }
                        #If the user hasn't been granted the license yet, add an add parameter to the splat to add the
                        #license to this user.
                        If ($AssignLicenses) {
                            $SetMsolUserLicenseParams.AddLicenses = $AccountSkuID
                            If ($LicenseOptions.DisabledServicePlans) {
                                $LogSkuLicensesAssigned.Add("$($LicensedUser.Key) - $($LicenseOptions.DisabledServicePlans)`n")
                            }
                            ElseIf (-not $LicenseOptions.DisabledServicePlans) {
                                $LogSkuLicensesAssigned.Add("$($LicensedUser.Key) - No disabled options`n")
                            }
                        }
                        #Run the command with the required parameters and, if available, the optional ones.
                        Set-MsolUserLicense @SetMsolUserLicenseParams -ErrorVariable LogErrorVariable
                    }
                    #If the license needs to be skipped, don't assign the license as it is superseded.
                    ElseIf ($SkippedLicenses) {
                        ForEach ($SkippedLicense in $SkippedLicenses){
                            $LogSkuSupersededAssigned.Add("$($LicensedUser.Key) - $($SkippedLicense.InputObject)`n")
                        }
                    }
                }
            #endregion

            #region Log performed activities
                #Add the individual logs for the Skus to the logs for the whole function.
                If ($LogSkuLicensesAssigned) { $LogLicensesAssigned.Add("Assigned $AccountSkuID to the following users:`n $LogSkuLicensesAssigned`n") }
                If ($LogSkuLicensesChanged) { $LogLicensesChanged.Add("Changed $AccountSkuID options for the following users:`n $LogSkuLicensesChanged`n") }
                If ($LogSkuLicensesRemoved) { $LogLicensesRemoved.Add("Removed $AccountSkuID from the following users:`n $LogSkuLicensesRemoved`n") }
                If ($LogSkuSupersededAssigned) { $LogSupersededAssigned.Add("License $AccountSkuID is assigned but superseded for the following users. Please check:`n $LogSkuSupersededAssigned`n") }
                If ($LogSkuSupersededRemoved) { $LogSupersededRemoved.Add("Due to supersedence by $AccountSkuID, the following users had licenses removed:`n $LogSkuSupersededRemoved`n") }
            #endregion
        }

    }
    END {

        #region Prepare and send email message
            #Compose the body from all collected logs. Only logs with entries will be displayed.
            [string]$Body = @(
                If (-not $LogErrorVariable) {"The script ran successfully. No errors occured. Any changes made will be listed below.`n"}
                ElseIf ($LogErrorVariable) {"`nThe script completed with errors. Any changes and errors will be listed below.`n $LogErrorVariable`n"}
                If ($LogLicensesAssigned) {"`nAssigned the following licenses:`n $LogLicensesAssigned`n"}
                If ($LogLicensesChanged) {"`nChanged the following license options:`n $LogLicensesChanged`n"}
                If ($LogLicensesRemoved) {"`nRemoved the following licenses:`n $LogLicensesRemoved`n"}
                If ($LogSupersededAssigned) {"`nThe following licenses are assigned but superseded:`n $LogSupersededAssigned`n"}
                If ($LogSupersededRemoved) {"`nRemoved the following superseded licenses:`n $LogSupersededRemoved`n"}
            )
            #Compose the subject. Subject description depends on if errors occured or not.
            [string]$Subject = @(
                If (-not $LogErrorVariable) {"[Success] Set-O365License: Script ran succesfully"}
                ElseIf ($LogErrorVariable) {"[Warning] Set-O365License: Script completed with errors"}
            )
            #Set the email parameters and send the message.
            $EmailParams.Subject = $Subject
            $EmailParams.Body = $Body
            $EmailParams.To = $ConfigData.Settings.EmailServerSettings.To
            Send-MailMessage @EmailParams
        #endregion

    }
}
