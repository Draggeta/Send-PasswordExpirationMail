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
            if ($LicenseSource -eq 'ActiveDirectory') {
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
            try {
                $ConfigData = Get-Content $ConfigurationFilePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop 
            }
            catch {
                #$LogErrorVariable = "Error Message: $($_.Exception.Message)`nFailed Item: $($_.Exception.ItemName)"
                break
            }
            #Set the email parameters.
            $EmailParams = @{}
            $EmailParams.From = $ConfigData.Settings.EmailServerSettings.From
            $EmailParams.SmtpServer =  $ConfigData.Settings.EmailServerSettings.SmtpServer
            $EmailParams.Port = $ConfigData.Settings.EmailServerSettings.Port
            $EmailParams.UseSsl = $ConfigData.Settings.EmailServerSettings.UseSsl
            $EmailParams.BodyAsHtml = $ConfigData.Settings.EmailServerSettings.BodyAsHtml
            $EmailParams.Credential = Get-StoredCredential -Target $ConfigData.Settings.Credentials.MailCredentials
        #endregion

        #region Log in to Azure AD
            #Login to Office 365. May need to be changed to use the Azure AD preview cmdlets. Stop execution if logging in
            #fails. May require tests as well to check if the module(s) are installed.
            try {
                Connect-MsolService -Credential $Credential
            } 
            catch {
                $LogErrorVariable = "Error Message: $($_.Exception.Message)`nFailed Item: $($_.Exception.ItemName)"
                $EmailParams.Subject = '[Error] Set-O365License: Failed to log in to Azure AD'
                $EmailParams.Body = "Logging in to Azure AD failed. See detailed error(s) below.`n $LogErrorVariable"
                $EmailParams.To = $ConfigData.Settings.EmailServerSettings.To
                Send-MailMessage @EmailParams
                break
            }
        #endregion

    }
    PROCESS {
        
        #For each license in the configuration file, process it.
        foreach ($AccountSkuID in $ConfigData.Licenses.PSObject.Properties.Name) {
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
                foreach ($Group in $Groups) {
                    #Find all members of the currently iterated group. Different commands are run depending on if AzureAD
                    #or Active Directory is used as primary source.
                    if ($LicenseSource -eq 'ActiveDirectory') {
                        $Members = (Get-ADGroupMember -Identity $Group -Recursive | Get-ADUser).UserPrincipalName
                        #$Members = Get-ADObject -LDAPFilter "(memberOf=$($(Get-ADGroup $Group).DistinguishedName))" -Properties UserPrincipalName
                    }
                    #This currently doesn't do nested groups.
                    elseif ($LicenseSource -eq 'AzureAD') {
                        $GroupId = (Get-MsolGroup -All).Where{ $_.DisplayName -eq $Group }
                        $Members = (Get-MsolGroupMember -GroupObjectId $GroupId.ObjectId -All).EmailAddress
                    }
                    #Get all users who should have a license. Add their UserPrincipalNames to the $LicensedUsers hashtable.
                    foreach ($Member in $Members) {
                        #If the user isn't in the hashtable yet, add them and the license options denied by this group.
                        if (-not $LicensedUsers.ContainsKey($Member)) {
                            $LicensedUsers.Add($Member, $ConfigData.Licenses.$AccountSkuID.Groups.$Group.DisabledPlans)
                        }
                        #If the user is already in the hashtable, compare the options and keep only the ones that are
                        #equal between license options.
                        elseif ($LicensedUsers.ContainsKey($Member)) {
                            $CompareArray = Compare-Object -ReferenceObject $LicensedUsers.Get_Item($Member) -DifferenceObject $ConfigData.Licenses.$AccountSkuID.Groups.$Group.DisabledPlans -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                            $LicensedUsers.Set_Item($Member, $CompareArray.InputObject)
                        }
                    }
                }
            #endregion

            #region Add/remove license
                #Test if a user who is currently licensed is in the list of users who should be licensed. If not, remove 
                #this SKU.
                foreach ($CurrentlyLicensedUser in $CurrentlyLicensedUsers) {
                    if ($LicensedUsers.ContainsKey($CurrentlyLicensedUser) -eq $False) {
                        try {
                            Set-MsolUserLicense -UserPrincipalName $CurrentlyLicensedUser -RemoveLicenses $AccountSkuID -ErrorAction Stop
                            $LogSkuLicensesRemoved.Add("$CurrentlyLicensedUser`n")
                        }
                        catch {
                            $LogErrorVariable.Add("Failed to remove license $AccountSkuID from $CurrentlyLicensedUser")
                        }
                    }
                }

                #Assign and revoke (based on supersedence) the license and net license options. 
                foreach ($LicensedUser in $LicensedUsers.GetEnumerator()) {
                    #Set a few base variables for each of the users found in the group.
                    $CurrentUser = $AllMsolUser.Where{ $_.UserPrincipalName -eq $LicensedUser.Key }
                    $CurrentUsageLocation = $CurrentUser.UsageLocation
                    $CurrentUserLicenses = $CurrentUser.Licenses.AccountSkuId
                    $CurrentUserOptions = $CurrentUser.Licenses.Where{ $_.AccountSkuId -eq $AccountSkuID }.ServiceStatus.Where{ $_.ProvisioningStatus -eq 'Disabled' }.ServicePlan.ServiceName
                    #$CurrentUserOptions = $CurrentUser.Licenses.Where{ $_.AccountSkuId -eq $AccountSkuID }.ServiceStatus.Where{ $_.ProvisioningStatus -eq 'Disabled' -or $_.ProvisioningStatus -eq 'PendingActivation' }.ServicePlan.ServiceName
                    #Set the usage location to the correct value if incorrect.
                    if ($CurrentUsageLocation -ne $UsageLocation) {
                        try {
                            Set-MsolUser -UserPrincipalName $CurrentUser.UserPrincipalName -UsageLocation $UsageLocation -ErrorAction Stop
                        }
                        catch {
                            $LogErrorVariable.Add("Failed to set usage location $UsageLocation for $($CurrentUser.UserPrincipalName)")
                        }
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
                    #If the license options are both in the config file and currently assigned to the user, check if there
                    #are differences or not. If there are differences, the license options will be set.
                    if ($LicensedUser.Value -and $CurrentUserOptions) {
                        $ChangeLicensesOptions = Compare-Object -ReferenceObject $LicensedUser.Value -DifferenceObject $CurrentUserOptions
                    }
                    #Else if only one of the variables has options, then options need to be changed. If that is the case,
                    #just set the variable to true.
                    elseif ($LicensedUser.Value -or $CurrentUserOptions) {
                        $ChangeLicensesOptions = $True
                    }
                    #If the license doesn't needs to be skipped, and either the license has not been assigned, is not
                    #superseded or won't be superseded by this license, run the script below.
                    if ((-not $SkippedLicenses) -and ($AssignLicenses -or $RemoveSupersededLicenses -or $ChangeLicensesOptions)) {
                        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuID -DisabledPlans $LicensedUser.Value -ErrorVariable LogErrorVariable
                        #Splat the default paramters used in Set-MsolUserLicense.
                        $SetMsolUserLicenseParams = @{}
                        $SetMsolUserLicenseParams.UserPrincipalName = $CurrentUser.UserPrincipalName
                        #If license options need to be changed, add the parameter/value to the splat variable.
                        if ($ChangeLicensesOptions) {
                            $SetMsolUserLicenseParams.LicenseOptions = $LicenseOptions
                        }
                        #If the user has a superseded license configured, add a remove parameter to the splat to remove
                        #this license.
                        if ($RemoveSupersededLicenses) {
                            $SetMsolUserLicenseParams.RemoveLicenses = $RemoveSupersededLicenses.InputObject
                        }
                        #If the user hasn't been granted the license yet, add an add parameter to the splat to add the
                        #license to this user.
                        if ($AssignLicenses) {
                            $SetMsolUserLicenseParams.AddLicenses = $AccountSkuID
                        }
                        #Run the command with the required parameters and, if available, the optional ones.
                        try {
                            Set-MsolUserLicense @SetMsolUserLicenseParams -ErrorVariable LogErrorVariable
                            switch ($SetMsolUserLicenseParams) {
                                { $_.RemoveLicenses }                                               { $LogSkuSupersededRemoved.Add("$($RemoveSupersededLicenses.InputObject  -join ', ') - $($LicensedUser.Key)") }
                                { $_.AddLicenses -and $LicenseOptions.DisabledServicePlans }        { $LogSkuLicensesAssigned.Add("$($LicenseOptions.DisabledServicePlans -join ', ') - $($LicensedUser.Key)"); break }
                                { $_.AddLicenses -and -not $LicenseOptions.DisabledServicePlans }   { $LogSkuLicensesAssigned.Add("No disabled options - $($LicensedUser.Key)"); break }
                                { $_.LicenseOptions -and $LicenseOptions.DisabledServicePlans }     { $LogSkuLicensesChanged.Add("$($LicenseOptions.DisabledServicePlans -join ', ') - $($LicensedUser.Key)") }
                                { $_.LicenseOptions -and -not$LicenseOptions.DisabledServicePlans } { $LogSkuLicensesChanged.Add("No disabled options - $($LicensedUser.Key)") }
                            }
                        }
                        catch {
                            switch ($SetMsolUserLicenseParams) {
                                { $_.RemoveLicenses }   { $LogErrorVariable.Add("Failed to remove superseded license $($SetMsolUserLicenseParams.RemoveLicenses) from $($LicensedUser.Key)`n") } 
                                { $_.AddLicenses }      { $LogErrorVariable.Add("Failed to add license $($SetMsolUserLicenseParams.AddLicenses) to $($LicensedUser.Key)`n") }
                                { $_.LicenseOptions }   { $LogErrorVariable.Add("Failed to change license options $($SetMsolUserLicenseParams.LicenseOptions -join ', ') for $($LicensedUser.Key)`n") }
                            }
                        }
                    }
                    #If the license needs to be skipped, don't assign the license as it is superseded.
                    elseif ($SkippedLicenses) {
                        foreach ($SkippedLicense in $SkippedLicenses){
                            $LogSkuSupersededAssigned.Add("$($SkippedLicense.InputObject -join ', ') - $($LicensedUser.Key)")
                        }
                    }
                }
            #endregion

            #region Log performed activities
                #Add the individual logs for the Skus to the logs for the whole function.
                if ($LogSkuLicensesAssigned) { $LogLicensesAssigned.Add("<p>Assigned $AccountSkuID to the following users:</br>$($LogSkuLicensesAssigned -join '</br>')</p>") }
                if ($LogSkuLicensesChanged) { $LogLicensesChanged.Add("<p>Changed $AccountSkuID options for the following users:</br>$($LogSkuLicensesChanged -join '</br>')</p>") }
                if ($LogSkuLicensesRemoved) { $LogLicensesRemoved.Add("<p>Removed $AccountSkuID from the following users:</br>$($LogSkuLicensesRemoved -join '</br>')</p>") }
                if ($LogSkuSupersededAssigned) { $LogSupersededAssigned.Add("<p>License $AccountSkuID is assigned but superseded for the following users. Please check:</br>$($LogSkuSupersededAssigned -join '</br>')</p>") }
                if ($LogSkuSupersededRemoved) { $LogSupersededRemoved.Add("<p>Due to supersedence by $AccountSkuID, the following users had licenses removed:</br>$($LogSkuSupersededRemoved -join '</br>')</p>") }
            #endregion
        }

    }
    END {

        #region Prepare and send email message
            #Compose the body from all collected logs. Only logs with entries will be displayed.
            [String]$Body = @(
                if (-not $LogErrorVariable) {"The script ran successfully. No errors occured. Any changes made will be listed below.</br></br>"}
                elseif ($LogErrorVariable) {"<p>The script completed with errors. Any changes and errors will be listed below.</br>$LogErrorVariable</p>"}
                if ($LogLicensesAssigned) {"<p>Assigned the following licenses:</br>$LogLicensesAssigned</p>"}
                if ($LogLicensesChanged) {"<p>Changed the following license options:</br>$LogLicensesChanged</p>"}
                if ($LogLicensesRemoved) {"<p>Removed the following licenses:</br>$LogLicensesRemoved</p>"}
                if ($LogSupersededAssigned) {"<p>The following licenses are assigned but superseded:</br>$LogSupersededAssigned</p>"}
                if ($LogSupersededRemoved) {"<p>Removed the following superseded licenses:</br>$LogSupersededRemoved</p>"}
            )
            #Compose the subject. Subject description depends on if errors occured or not.
            [String]$Subject = @(
                if (-not $LogErrorVariable) {"[Success] Set-O365License: Script ran succesfully"}
                elseif ($LogErrorVariable) {"[Warning] Set-O365License: Script completed with errors"}
            )
            #Set the email parameters and send the message.
            $EmailParams.Subject = $Subject
            $EmailParams.Body = $Body
            $EmailParams.To = $ConfigData.Settings.EmailServerSettings.To
            Send-MailMessage @EmailParams
        #endregion

    }
}
