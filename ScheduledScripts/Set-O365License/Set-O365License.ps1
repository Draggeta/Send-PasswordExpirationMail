Function Set-O365License {
    
    [CmdletBinding()]
    Param (

        [Parameter(Mandatory = $True)]
        [ValidateScript({ Test-Path -Path $_ })]
        [Alias('ConfigFile')]
        [String]$ConfigurationFilePath,

        [Parameter()]
        [ValidateScript({ Test-Path -Path $_ })]
        [Alias('EmailHtml')]
        [String]$EmailHtmlFilePath,

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
            try {
                Import-Module CredentialManager
                Import-Module MSOnline
                if ($LicenseSource -eq 'ActiveDirectory') {
                    Import-Module ActiveDirectory
                }
            }
            catch {
                break
            } 
        #endregion

        #region Create empty general logging arrays
            #Create empty arrays for the logs. These logs collect the other logs that fall into their categories.
            [System.Collections.ArrayList]$LogErrorVariable = @()
            [System.Collections.ArrayList]$LogLicensesStatus = @()
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
                break
            }
            #Set the email parameters.
            $EmailParams = @{}
            $EmailParams.From = $ConfigData.Settings.EmailServerSettings.From
            $EmailParams.To = $ConfigData.Settings.EmailServerSettings.To
            $EmailParams.SmtpServer =  $ConfigData.Settings.EmailServerSettings.SmtpServer
            $EmailParams.Port = $ConfigData.Settings.EmailServerSettings.Port
            $EmailParams.UseSsl = $ConfigData.Settings.EmailServerSettings.UseSsl
            $EmailParams.BodyAsHtml = $ConfigData.Settings.EmailServerSettings.BodyAsHtml
            $EmailParams.Credential = Get-StoredCredential -Target $ConfigData.Settings.Credentials.MailCredentials
        #endregion

        #region Log in to Azure AD
            #Login to Office 365. May need to be changed to use the Azure AD preview cmdlets. Stop execution if logging in
            #fails.
            try {
                Connect-MsolService -Credential $Credential
            } 
            catch {
                $LogErrorVariable = "$($_.Exception.Message)"
                $EmailParams.Subject = '[Error] Set-O365License: Failed to log in to Azure AD'
                $EmailParams.Body = "<p>Logging in to Azure AD failed. See detailed error(s) below.<br>$LogErrorVariable</p>"
                Send-MailMessage @EmailParams
                break
            }
        #endregion

    }
    PROCESS {
        
        #For each license in the configuration file, process it.
        foreach ($AccountSkuID in $ConfigData.Licenses.PSObject.Properties.Name) {
            #region Create empty per SKU log arrays
                #Create empty arrays for the current SKU logs. These logs are added to the earlier created log arrays.
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
                #actually assigned licenses. Also reduces the amount of Get-MsolUser commands run. The reason this is
                #run every SKU, is that assigned licenses may change during the execution of this script.
                $AllMsolUser = Get-MsolUser -All
                #Retrieve the usage location for this license.
                $UsageLocation = $ConfigData.Licenses.$AccountSkuID.UsageLocation
                #Supersedence was implemented because sometimes different licenses have options that don't work well
                #together. e.g. two licenses that grant a user an Exchange account. In those cases, the simple
                #solution for now is to remove the superseded license. This is done later in the script.
                #Retrieve the licenses that are superseded by this license.
                $SupersededLicenses = $ConfigData.Licenses.$AccountSkuID.Supersedes
                #Retrieve the licenses that supersede this license.
                $SupersededByLicenses = $ConfigData.Licenses.$AccountSkuID.SupersededBy
                #Fill the $Groups array/variable with all groups that assign this license. These are retrieved from the
                #configuration file. Make sure it is correct.
                $Groups.AddRange([Array]$ConfigData.Licenses.$AccountSkuID.Groups.PSObject.Properties.Name)
                #Fill the $CurrentlyLicensedUsers array with all users currently licensed with this SKU.
                $CurrentlyLicensedUsers.AddRange([Array]($AllMsolUser.Where{ $_.Licenses.AccountSkuID -Contains $AccountSkuID }).UserPrincipalName)
            #endregion
            
            #region Examine current/reference licenses
                #Create an empty hashtable to store users and their net license options for this license SKU.
                $LicensedUsers = @{}

                #Find all users and their net license options by comparing the license options denied by groups and
                #discarding all options that don't appear in all license options.
                foreach ($Group in $Groups) {
                    #Find all members of the currently iterated group. Different commands are run depending on if AzureAD
                    #or Active Directory is used as primary source.
                    if ($LicenseSource -eq 'ActiveDirectory') {
                        $Members = (Get-ADGroupMember -Identity $Group -Recursive | Get-ADUser).UserPrincipalName
                    }
                    #This currently doesn't do nested groups. May change it in the future
                    elseif ($LicenseSource -eq 'AzureAD') {
                        $GroupId = (Get-MsolGroup -All).Where{ $_.DisplayName -eq $Group }
                        $Members = (Get-MsolGroupMember -GroupObjectId $GroupId.ObjectId -All).EmailAddress
                    }
                    #Get all users who should have a license. Add their UserPrincipalNames to the $LicensedUsers
                    #hashtable.
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
                #Test if a user who is currently licensed is in the list of users who should be licensed. If not,
                #remove this SKU from the account.
                foreach ($CurrentlyLicensedUser in $CurrentlyLicensedUsers) {
                    if ($LicensedUsers.ContainsKey($CurrentlyLicensedUser) -eq $False) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $CurrentlyLicensedUser -RemoveLicenses $AccountSkuID -ErrorAction Stop
                            #Log the removal of the license by adding it as an HTML list item to the current SKU 
                            #removed licenses array.
                            $LogSkuLicensesRemoved.Add("<li>$CurrentlyLicensedUser</li>")
                        }
                        catch {
                            #Log the failure to remove the licenses by adding it as an HTML list item to the general
                            #error log.
                            $LogErrorVariable.Add("<li>Failed to remove license $($AccountSkuID -replace ".*:") from $CurrentlyLicensedUs</li>")
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
                            #Set-MsolUser -UserPrincipalName $CurrentUser.UserPrincipalName -UsageLocation $UsageLocation -ErrorAction Stop
                        }
                        catch {
                            #Log the failure to set the usage location by adding it as an HTML list item to the general
                            #error log.
                            $LogErrorVariable.Add("Failed to set usage location $UsageLocation for $($CurrentUser.UserPrincipalName)")
                        }
                    }
                    #Compare the currently assigned and "superseded by" licenses. If there is no match, assign the
                    #license.
                    #Check if the user has the license already assigned.
                    $AssignLicenses = $CurrentUserLicenses -notcontains $AccountSkuID
                    #Check if there are any superseded licenses currently assigned to the user. Return only the ones
                    #that appear in both arrays. These superseded licenses will be removed.
                    $RemoveSupersededLicenses = Compare-Object -ReferenceObject $SupersededLicenses -DifferenceObject $CurrentUserLicenses -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                    #Check if any licenses that supersede this license are assigned to the user. If it is the case, 
                    #skip adding the current license SKU.
                    $SkippedLicenses = Compare-Object -ReferenceObject $SupersededByLicenses -DifferenceObject $CurrentUserLicenses -IncludeEqual -ExcludeDifferent -ErrorAction SilentlyContinue
                    #If disabled license options are specified in the config file and options are disabled for the
                    #user, check for differences. If there are differences, the license options will be set.
                    if ($LicensedUser.Value -and $CurrentUserOptions) {
                        $ChangeLicensesOptions = Compare-Object -ReferenceObject $LicensedUser.Value -DifferenceObject $CurrentUserOptions
                    }
                    #Else if only one of the variables has options, then options need to be changed. If that is the case,
                    #just set the variable to true.
                    elseif ($LicensedUser.Value -or $CurrentUserOptions) {
                        $ChangeLicensesOptions = $True
                    }
                    #If the license doesn't needs to be skipped (due to it being superseded) and either:
                    #the license has not been assigned yet
                    #a superseded license needs to be removed 
                    #license options need to be changed
                    #run this block of code to perform any or all of the changes mentioned above
                    if ((-not $SkippedLicenses) -and ($AssignLicenses -or $RemoveSupersededLicenses -or $ChangeLicensesOptions)) {
                        #Create the license options that may be assigned
                        $LicenseOptions = New-MsolLicenseOptions -AccountSkuId $AccountSkuID -DisabledPlans $LicensedUser.Value -ErrorVariable LogErrorVariable
                        #Splat the default paramters used in #Set-MsolUserLicense.
                        $SetMsolUserLicenseParams = @{}
                        $SetMsolUserLicenseParams.UserPrincipalName = $CurrentUser.UserPrincipalName
                        #If license options need to be changed, add the parameter/value to the splat variable.
                        if ($ChangeLicensesOptions) {
                            $SetMsolUserLicenseParams.LicenseOptions = $LicenseOptions
                        }
                        #If the user has a superseded license configured, add a remove parameter to the splat to remove
                        #this license.
                        if ($RemoveSupersededLicenses) {
                            $SetMsolUserLicenseParams.RemoveLicenses = $RemoveSupersededLicenses.InputObject -join ','
                        }
                        #If the user hasn't been granted the license yet, add an add parameter to the splat to add the
                        #license to this user.
                        if ($AssignLicenses) {
                            $SetMsolUserLicenseParams.AddLicenses = $AccountSkuID
                        }
                        #Run the command with the required parameters and, if available, the optional ones.
                        try {
                            #Set-MsolUserLicense @SetMsolUserLicenseParams -ErrorVariable LogErrorVariable
                            #Log the performed changes by adding them as an HTML list item to their respective current
                            #SKU log array. There is a break purposefully on the $LogSkuLicensesAssigned array as
                            #they basically are a more specific subset of the $LogSkuLicensesChanged array. Anyone who
                            #had his license changed, had his options changed.
                            switch ($SetMsolUserLicenseParams) {
                                { $_.RemoveLicenses }                                               { $LogSkuSupersededRemoved.Add("<li>$($RemoveSupersededLicenses.InputObject -replace ".*:") - $($LicensedUser.Key)</li>") }
                                { $_.AddLicenses -and $LicenseOptions.DisabledServicePlans }        { $LogSkuLicensesAssigned.Add("<li>$($LicenseOptions.DisabledServicePlans -join ', ') - $($LicensedUser.Key)</li>"); break }
                                { $_.AddLicenses -and -not $LicenseOptions.DisabledServicePlans }   { $LogSkuLicensesAssigned.Add("<li>No disabled options - $($LicensedUser.Key)</li>"); break }
                                { $_.LicenseOptions -and $LicenseOptions.DisabledServicePlans }     { $LogSkuLicensesChanged.Add("<li>$($LicenseOptions.DisabledServicePlans -join ', ') - $($LicensedUser.Key)</li>") }
                                { $_.LicenseOptions -and -not$LicenseOptions.DisabledServicePlans } { $LogSkuLicensesChanged.Add("<li>No disabled options - $($LicensedUser.Key)</li>") }
                            }
                        }
                        catch {
                            #Log the failures by adding them as an HTML list item to the general error log array.
                            switch ($SetMsolUserLicenseParams) {
                                { $_.RemoveLicenses }   { $LogErrorVariable.Add("<li>Failed to remove superseded license $($SetMsolUserLicenseParams.RemoveLicenses) from $($LicensedUser.Key)</li>") } 
                                { $_.AddLicenses }      { $LogErrorVariable.Add("<li>Failed to add license $($SetMsolUserLicenseParams.AddLicenses) to $($LicensedUser.Key)</li>") }
                                { $_.LicenseOptions }   { $LogErrorVariable.Add("<li>Failed to change license options $($SetMsolUserLicenseParams.LicenseOptions -join ', ') for $($LicensedUser.Key)</li>") }
                            }
                        }
                    }
                    #If the current license SKU needs to be skipped, don't assign the license and log it as an HTML
                    #list item.
                    elseif ($SkippedLicenses) {
                        foreach ($SkippedLicense in $SkippedLicenses){
                            $LogSkuSupersededAssigned.Add("<li>$($SkippedLicense.InputObject -replace ".*:") - $($LicensedUser.Key)</li>")
                        }
                    }
                }
            #endregion

            #region Log performed activities
                #Add the individual current SKU logs to the general logs created at the start of the script as HTML
                #paragraphs.
                if ($LogSkuLicensesAssigned)    { $LogLicensesAssigned.Add("<p><h4>Assigned $($AccountSkuID -replace ".*:") to the following users:</h4><ul>$($LogSkuLicensesAssigned)</ul></p>") }
                if ($LogSkuLicensesChanged)     { $LogLicensesChanged.Add("<p><h4>Changed $($AccountSkuID -replace ".*:") options for the following users:</h4><ul>$($LogSkuLicensesChanged)</ul></p>") }
                if ($LogSkuLicensesRemoved)     { $LogLicensesRemoved.Add("<p><h4>Removed $($AccountSkuID -replace ".*:") from the following users:</h4><ul>$($LogSkuLicensesRemoved)</ul></p>") }
                if ($LogSkuSupersededAssigned)  { $LogSupersededAssigned.Add("<p><h4>License $($AccountSkuID -replace ".*:") is assigned but superseded for the following users. Please check:</h4><ul>$($LogSkuSupersededAssigned)</ul></p>") }
                if ($LogSkuSupersededRemoved)   { $LogSupersededRemoved.Add("<p><h4>Due to supersedence by $($AccountSkuID -replace ".*:"), the following users had licenses removed:</h4><ul>$($LogSkuSupersededRemoved)</ul></p>") }
            #endregion
        }

    }
    END {

        #region Check license availability
            #Get all licenses currently available for this tennant
            $LicenseSkus = Get-MsolAccountSku
            foreach ($LicenseSku in $LicenseSkus) {
                #Create a hashtable with all licenses, total licenses, used licenses and the amount of available ones.
                $LicenseHashTable = @{
                    License = $LicenseSku.AccountSkuId -replace ".*:"
                    Total = $LicenseSku.ActiveUnits
                    Used = $LicenseSku.ConsumedUnits
                    Available = ($LicenseSku.ActiveUnits - $LicenseSku.ConsumedUnits)
                }
                #Add the status of the currently iterated SKU as a table row to the logs.
                $LogLicensesStatus.Add("<tr><td>$($LicenseHashTable.License)</td>")
                $LogLicensesStatus.Add("<td>$($LicenseHashTable.Total)</td>")
                $LogLicensesStatus.Add("<td>$($LicenseHashTable.Used)</td>")
                $LogLicensesStatus.Add("<td>$($LicenseHashTable.Available)</td></tr>")
                #Add an entry to the general error log if the amount of licenses is lower than (or equal to) the
                #specified value.
                if ($LicenseHashTable.Available -le 3) {
                    $LogErrorVariable.Add("<li>$($LicenseHashTable.License) - $($LicenseHashTable.Available) out of $($LicenseHashTable.Total) licenses available.</li>")
                }
            }
        #endregion

        #region Prepare and send email message
            #Check if something is placed in the error array and set the correct message
            if (-not $LogErrorVariable)     { $ScriptStatus =  "<p>The script ran successfully. No errors occured. Any changes made will be listed below.</p>" }
            elseif ($LogErrorVariable)      { $ScriptStatus =  "<p>The script completed with errors. Any changes and warnings will be listed below.</p>"
                                              $ScriptAction += "<p><h2>Warnings</h2><ul>$LogErrorVariable</ul></p>" }
            #As I wanted to keep the html file simple, I specified the License status table design in the script.
            #May not be the best solution, as it doesn't take into consideration the fact that no template may be used
            #as well.
            if ($LogLicensesStatus)         { $ScriptAction += "<p><h2>Licenses status</h2>
                                                                    <table align='center' width='100%' style='color: #606060; font-family: Calibri; border-collapse: collapse;'>
                                                                    <tr>
                                                                    <th align='left'>License</th>
                                                                    <th align='left'>Total</th>
                                                                    <th align='left'>Used</th>
                                                                    <th align='left'>Available</th>
                                                                    </tr>
                                                                    $LogLicensesStatus
                                                                    </table>
                                                                </p>" }
            #Fill the arrays for the other logs.
            if ($LogLicensesAssigned)       { $ScriptAction += "<p><h2>Assigned licenses</h2>$LogLicensesAssigned</p>" }
            if ($LogLicensesChanged)        { $ScriptAction += "<p><h2>Changed license options</h2>$LogLicensesChanged</p>" }
            if ($LogLicensesRemoved)        { $ScriptAction += "<p><h2>Removed licenses</h2>$LogLicensesRemoved</p>" }
            if ($LogSupersededAssigned)     { $ScriptAction += "<p><h2>Assigned but superseded</h2>$LogSupersededAssigned</p>" }
            if ($LogSupersededRemoved)      { $ScriptAction += "<p><h2>Removed superseded licenses</h2>$LogSupersededRemoved</p>" }
            #If an HTML template has been specified, use it. This section may need to be changed to work with your
            #desired template.
            if ($EmailHtmlFilePath) {
                $EmailHtmlFile = Get-Content $EmailHtmlFilePath -Raw
                $EmailHtmlFile = $EmailHtmlFile.Replace('SCRIPTTITLE','Set-O365License')
                $EmailHtmlFile = $EmailHtmlFile.Replace('SCRIPTSTATUS',$ScriptStatus)
                $EmailHtmlFile = $EmailHtmlFile.Replace('SCRIPTACTION',$ScriptAction)
                $EmailHtmlFile = $EmailHtmlFile.Replace('DATETIME',(Get-Date -UFormat "%Y-%m-%d %T %Z"))
                [String]$Body  = $EmailHtmlFile
            }
            #If no template is specified, send the email as is.
            elseif (-not $EmailHtmlFilePath) {
                #Compose the body from all collected logs. Only logs with entries will be displayed.
                [String]$Body = @(
                    $ScriptStatus
                    $ScriptAction
                )
            }
            #Compose the subject. Subject description depends on if errors/warnings occured or not.
            [String]$Subject = @(
                if (-not $LogErrorVariable) { "[Success] Set-O365License: Script ran succesfully" }
                elseif ($LogErrorVariable)  { "[Warning] Set-O365License: Script completed with warnings" }
            )
            #Set the email parameters and send the message.
            $EmailParams.Subject = $Subject
            $EmailParams.Body = $Body
            Send-MailMessage @EmailParams -Encoding UTF8
        #endregion

    }
}
