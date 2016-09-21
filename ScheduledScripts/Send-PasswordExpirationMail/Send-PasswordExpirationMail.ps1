function Write-StatusToEventLog {

    If ($EventLogName -and $EventLogSource) {

        # The values here below can be used to have it write to event logs.
        $WriteEventLog = @{
        
            LogName = $EventLogName
            Source = $EventLogSource
        
        }
        
        If ($EventLogErrors) {
        
            $WriteEventLog += @{

                EventId = $EventWarningID
                EntryType = 'Warning'
                Message = "One or more errors occured while running the Send-PasswordExpirationMail Powershell script. See the errors below:`n`n $EventLogErrors"

            }

        } Else {

            $WriteEventLog += @{

                EventId = $EventInformationID
                EntryType = 'Information'
                Message = "The Send-PasswordExpirationMail Powershell script ran succesfully.`n $EventLogMessage"

            }
            
        }

        Write-EventLog @WriteEventLog

    }

}

function Send-StatusMessage {

    $MailAttributes = @{

        To = $AdminEmail
        From = $From
        Port = $Port
        
    }
    If ($UseSsl) {$MailAttributes += @{UseSsl = $True}}
    
    If ($MailCredential) {$MailAttributes += @{Credential = $MailCredential}}
    
    If ($EventLogErrors) {

        $MailAttributes += @{
            Subject = 'Send-PasswordExpirationMail script execution failed'
            Body = "One or more errors occured while running the Send-PasswordExpirationMail Powershell script. See the errors below:`n`n $EventLogErrors"
        }

    } Else {

        $MailAttributes += @{
            Subject = 'Send-PasswordExpirationMail script execution succeeded'
            Body = "The Send-PasswordExpirationMail Powershell script ran succesfully.`n $EventLogMessage"
        }

    }

    Send-MailMessage @MailAttributes

}

Function Send-PEMailMessage {
    <#
.SYNOPSIS
    Sends out password expiration notices.
.DESCRIPTION
    The Send-PasswordExpirationMail.ps1 script sends out password expiration notices to users. While it is a bit convoluted it is written to work in most environments without changing much, if anything outside of the parameters.
	To use the EventLog variables, you may need to create your own Logname or Source with the New-EventLog cmdlet if not using a default Windows log and source.
	The AD Filter should suffice for most organizations. There is no smtp server specified in the Send-MailMessage cmdlet as we are using the global $PSEmailServer variable.
.PARAMETER ConfigFile
    Specifies the location of the config file.
.PARAMETER RemindOn
    Specifies when to send a reminder, in days. A single value can be specified or an array of values separated by commas.
.PARAMETER SmtpServer
    The name of the server which will send email. Can be just the hostname of the FQDN of the server.
.PARAMETER From
    The from address which will send the email.
.PARAMETER Port
    The port on which to connect to the SMTP server. If not specified, the default port 25 is used.
.PARAMETER SearchBase
    The OU where the revelant users can be found in the Active Directory environment. If not specified, the script will search through the whole AD to find users.
.PARAMETER EventLogName
    The name of the event log where errors and successes will be written to. If specified, the EventLogSource parameter must also be specified. 
    If the event log name doesn't exits it must be created with the New-EventLog cmdlet.
.PARAMETER EventLogSource
    The source of the errors and successes to be written to the event log. If specified, the EventLogName parameter must also be specified. 
    If the event log source doesn't exits it must be created with the New-EventLog cmdlet.
.PARAMETER EventWarningID
    The EventID to be used if one or more parts of the script fails.
.PARAMETER EventInformationID
    The EventID to be used if the script runs successfully.
If those are not specified either (commented/removed), the emails will be sent anonymously.
.PARAMETER Credential
    Specify credentials allowed to send emails from the from address. This is necessary if authentication is needed and the account running the task isn't allowed to send via that address. 
    Can be used with the Get-Credential cmdlet. If not specified it will use the credentials specified in the script. Those can be specified plain text below or pulled from an a file.
.EXAMPLE
    Send-PasswordExpirationMail -ConfigFile 'C:\Scripts\ConfigFile.xml'
    Description
    
    -----------

    This command sends out emails according to the settings specified in the configuration file.
.EXAMPLE
    Send-PasswordExpirationMail -RemindOn 1,3,7 -SmtpServer exchange.domain.com -From relay@domain.com
    Description
    
    -----------
    
    This command sends out emails on one, three and seven days before the password expires, via server exchange.domain.com from the relay@domain.com. This is the minimum required and will send emails via port 25 and anonymously
.EXAMPLE
    Send-PasswordExpirationMail -RemindOn 1,7,14 -SmtpServer mail.contoso.com -From noreply@contoso.com -Port 587 -UseSsl -EventLogName Application -EventLogSource MyScripts -EventWarningId 2001 -EventInformationID 2002
    Description
    
    -----------
    
    This command sends out emails encrypted and via port 587. It will also write errors and successes to the event log with the specified EventIDs.
.EXAMPLE
    $Cred = (Get-Credential)
    PS C:\>Send-PasswordExpirationMail -RemindOn 0,1,7 -SmtpServer smtp.company.com -From no-reply@company.com -Port 587 -SearchBase 'OU=Users,OU=Company,DC=Company,DC=local' -UseSsl -EventLogName Application -EventLogSource MyScripts -Credential $Cred


    Description
    
    -----------
    
    This command sends out emails authenticated and encrypted, while also logging to the event log with the default EventIDs. Credentials can be passed by using the Get-Credential cmdlet.
.INPUTS
	None. You cannot pipe objects to Send-PasswordExpirationMail.ps1
.OUTPUTS
	None. Send-PasswordExpirationMail.ps1 only outputs to the EventLog
.NOTES
    Author:   Tony Fortes Ramos
    Created:  April 24, 2014
    Modified: November 23, 2015
.LINK
	New-EventLog
	Write-EventLog
    Get-Credential
    Get-ADUser
    Send-MailMessage
#>
    
    [CmdletBinding(DefaultParameterSetName='ConfigFile')]
    Param (
        
        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True, Position = 1)]
        [Array]$RemindOn,
        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True, Position = 2)]
        [String]$SmtpServer,
        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True, Position = 3)]
        [String]$From,
        [Parameter (ParameterSetName = 'Parameter')]
        [ValidateRange(1,65535)]
        [Int]$Port = '25',
        [Parameter (ParameterSetName = 'Parameter')]
        [Switch]$UseSsl,
        [Parameter (ParameterSetName = 'Parameter')]
        [String]$SearchBase = (Get-ADDomain).DistinguishedName,
        [Parameter (ParameterSetName = 'Parameter')]
        [String]$EventLogName,
        [Parameter (ParameterSetName = 'Parameter')]
        [String]$EventLogSource,
        [Parameter (ParameterSetName = 'Parameter')]
        [Object]$MailCredential,
        [Parameter (ParameterSetName = 'ConfigFile')]
        [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
        [String]$ConfigFile
    
    )
    BEGIN {
    If ($ConfigFile) {
    [xml]$ConfigFile = (Get-Content -Path $ConfigFile) 

    $PSEmailServer = $ConfigFile.Settings.EmailServerSettings.SmtpServer
    $Port = $ConfigFile.Settings.EmailServerSettings.Port
    $From = $ConfigFile.Settings.EmailServerSettings.From
    $AdminEmail = ($ConfigFile.Settings.EmailServerSettings.AdminEmail).Split(',')
    $MailCredential = $ConfigFile.Settings.Credentials.MailCredentials
    [bool]$UseSsl = [int]$ConfigFile.Settings.EmailServerSettings.UseSsl
    
    $EventLogName = $ConfigFile.Settings.EventLogSettings.EventLogName
    $EventLogSource = $ConfigFile.Settings.EventLogSettings.EventLogSource
    $EventInformationID = $ConfigFile.Settings.EventLogSettings.EventLogInformationID
    $EventWarningID = $ConfigFile.Settings.EventLogSettings.EventLogWarningID
    
    $SearchBase = $ConfigFile.Settings.DomainSettings.UserSearchBase
    
    $RemindOn = [array]($ConfigFile.Settings.ScriptSettings.SendPasswordExpirationMail.RemindOn).Split(',')
}
    
    If ($MailCredential) {
    Import-Module CredentialManager -ErrorVariable +EventLogErrors
    $MailCredential = Get-StoredCredential -Target $MailCredential
}
    
    # Create an empty array for the messages which are to be written to the Event Logs.
    $EventLogMessage = @()
    
    # Import the necessary modules.  
    Import-Module ActiveDirectory -ErrorVariable +EventLogErrors
    
    # Set a few default variables, most of these needn't be changed unless you use fine-grained password policies.
    $GpoMaxAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
    $GpoMinLength = (Get-ADDefaultDomainPasswordPolicy).MinPasswordLength
    $GpoPasswordHistory = (Get-ADDefaultDomainPasswordPolicy).PasswordHistoryCount
    $GpoComplexityEnabled = (Get-ADDefaultDomainPasswordPolicy).ComplexityEnabled
    $Today = Get-Date 
    
    # The filters used to limit the queried users. Probably enough in most cases.
    # Find all users who match the filters specified and then loop through each.
    $ADFilter = {(PasswordNeverExpires -eq $False) -and (PasswordExpired -eq $False) -and (pwdLastSet -ne '0') -and (PasswordLastSet -ne "$Null") -and (Enabled -eq $True) -and (Emailaddress -ne "$Null")}
    $ADUsers = Get-ADUser -Filter $ADFilter -SearchBase $ADSearchBase -SearchScope Subtree -Properties PasswordLastSet,EmailAddress -ErrorVariable +EventLogErrors
    }
    PROCESS {
        ForEach ($ADUser in $ADUsers) {
            # Set a few variables per user for easier usage. This is mostly cosmetic.
            $PwdLastSet = $_.PasswordLastSet
            $PwdExpirationDate = $PwdLastSet + $GpoMaxAge
            $PwdDaysLeft = ($PwdExpirationDate - $Today).days
            
            # Check if the days left until the user's password expired is in the array specified.
            $PwdIntervalHit = $PwdDaysLeft -in $RemindOn
        
            # If the amount of days left until expiration is in the array, send an email to the user.
            If ($PwdIntervalHit) {
        
                $Name = $_.GivenName
                $To = $_.EmailAddress
        
                # Changes a few of the texts used in the mail and stores it in the variable.
                Switch ($PwdDaysLeft) {
                    0 {$InXDays = 'today'}
                    1 {$InXDays = 'in one day'}
                    Default {$InXDay = "in $PwdDaysLeft days"}
                }
                
                # Automatically apply or omit the text about complexity, dependant on if the group policy has been set.
                $ComplexityText = If ($GpoComplexityEnabled) {"<li>The password needs to include 3 of the 4 following categories:</li>
                <ul><li>At least one <em>lower case</em> letter (a-z)</li>
                <li>At least one <em>upper case</em> letter (A-Z)</li>
                <li>At least one <em>number</em> (0-9)</li>
                <li>At least one <em>special character</em> (!,?,*,~, etc.)</li></ul>"}
        
                # Store the subject and body into variables for use in the send-mailmessage cmdlet.
                $Subject = "Reminder: Your password expires $InXDays"
        
                $Body = "<p>Dear $Name,<br></p>
                <p>Your password will expire <em>$InXDays</em>. You can change your password by pressing CTRL+ALT+DEL and then choosing the option to reset your password. If you are not at the HEAD OFFICE or BRANCH OFFICE, you can reset your password via the webmail. Click on the gear icon and subsequently on `"Change password`" to change your password.</p>
                <p>If you have set up your COMPANY email account on other devices such as an iPhone/iPad or an Android device, please change your password on those devices as well.</p>
                Your new password must adhere to the following requirements:
                <ul><li>It must be at least <em>$GpoMinLength</em> characters long</li>
                $ComplexityText
                <li>It may not resemble the last <em>$GpoPasswordHistory</em> passwords</li></ul>
                <p>A good password doesn't need to be very complicated. You can always use multiple random words to create a password. The following examples are as strong as any randomly generated password.</p>
                <ul><li>Correct battery horse staple</li>
                <li>CORRECTbatteryHORSEstaple?</li>
                <li>#correct4battery7horse5staple</li>
                <li>Correct!Battery,horse&staple.</li></ul>
                <p>If you have questions about or issues with changing your password, please to contact the IT department.</p>
                <p>Kind regards,</p>
                <br>
                <p><br>
                COMPANY IT department</p>"
        
                # Splat the attributes for use in the Send-MailMessage cmdlet.
                $MailAttributes = @{
        
                    To = $To
                    From = $From
                    Subject = $Subject
                    Body = $Body
                    Port = $Port
                    BodyAsHTML = $True
        
                }
                
                If ($UseSsl) {$MailAttributes += @{UseSsl = $True}}
                
                If ($MailCredential) {$MailAttributes += @{Credential = $MailCredential}}
        
                # Send the message to the user.
                Send-MailMessage @MailAttributes -ErrorVariable +EventLogErrors
        
                # Add a small description to the messages array if an email has been sent.
                $EventLogMessage += "`n - An email has been sent to $Name as his/her password expires $InXDays."
            }
        }
    }
    END {
        Send-StatusMessage
        Write-StatusToEventLog
    }
}

Function Get-PEUsers {
    [CmdletBinding(DefaultParameterSetName = 'SearchBase')]
    Param (
        [Parameter(ParameterSetName = 'Identity',Mandatory = $true)]
        [String]$Identity,

        [Parameter(ParameterSetName = 'SearchBase')]
        [String]$SearchBase = (Get-ADDomain).DistinguishedName,
        
        [Parameter(ParameterSetName = 'SearchBase')]
        [ValidateSet('Available','Unavailable','Both')]
        [Switch]$EmailAddressAvailability,

        [Parameter(ParameterSetName = 'SearchBase')]
        [ValidateSet('Enabled','Disabled','Both')]
        [Switch]$AccountStatus,

        [Parameter(ParameterSetName = 'SearchBase')]
        [ValidateSet('NotExpired','Expired','Both')]
        [Switch]$PasswordExpiration,

        [Parameter(ParameterSetName = 'SearchBase')]
        [Switch]$PasswordOlderThanMaxAge,

        [Parameter(ParameterSetName = 'SearchBase')]
        [Switch]$PasswordNeverExpires
    )
    BEGIN {
        $ADFilter = {(PasswordNeverExpires -eq $False) -and (PasswordExpired -eq $False) -and (pwdLastSet -ne '0') -and (PasswordLastSet -ne "$Null") -and (Enabled -eq $True) -and (Emailaddress -ne "$Null")}
        $ADUsers = Get-ADUser -Filter $ADFilter -SearchBase $ADSearchBase -SearchScope Subtree -Properties PasswordLastSet,EmailAddress -ErrorVariable +EventLogErrors
    }
    PROCESS{
        ForEach ($ADUser in $ADUsers) {
            $PasswordPolicy = Get-ADUserResultantPasswordPolicy -Identity $ADUser
            
            $Properties = @{
                GivenName = $ADUser.GivenName
                SurName = $ADUser.SurName
                Emailaddress = $ADuser.EmailAddess
            }
                ,Dagentotverlopen(FGPP),MinimaleLengte(FGPP),WachtwoordGeschiedenis(FGPP),Complexiteit(FGPP)
        }
    }
    END{
    }
}

Function Send-PEMailMessage {
    [CmdletBinding(DefaultParameterSetName='ConfigFile')]
    Param (
        [Parameter (ValueByPipeline = $true)]
        $Identity
    )
    BEGIN {
        # Voornaam,Achternaam,Emailadres,Dagentotverlopen(FGPP),MinimaleLengte(FGPP),WachtwoordGeschiedenis(FGPP),Complexiteit(FGPP)
        # Create an empty array for the messages which are to be written to the Event Logs.
        $EventLogMessage = @()
    }
    PROCESS {
        ForEach ($ADUser in $Identity) {
        
            $Name = $_.GivenName
            $To = $_.EmailAddress
        
            # Changes a few of the texts used in the mail and stores it in the variable.
            Switch ($PwdDaysLeft) {
                0 {$InXDays = 'today'}
                1 {$InXDays = 'in one day'}
                Default {$InXDay = "in $PwdDaysLeft days"}
            }
            
            # Automatically apply or omit the text about complexity, dependant on if the group policy has been set.
            $ComplexityText = If ($GpoComplexityEnabled) {"<li>The password needs to include 3 of the 4 following categories:</li>
            <ul><li>At least one <em>lower case</em> letter (a-z)</li>
            <li>At least one <em>upper case</em> letter (A-Z)</li>
            <li>At least one <em>number</em> (0-9)</li>
            <li>At least one <em>special character</em> (!,?,*,~, etc.)</li></ul>"}
        
            # Store the subject and body into variables for use in the send-mailmessage cmdlet.
            $Subject = "Reminder: Your password expires $InXDays"
        
            $Body = "<p>Dear $Name,<br></p>
            <p>Your password will expire <em>$InXDays</em>. You can change your password by pressing CTRL+ALT+DEL and then choosing the option to reset your password. If you are not at the HEAD OFFICE or BRANCH OFFICE, you can reset your password via the webmail. Click on the gear icon and subsequently on `"Change password`" to change your password.</p>
            <p>If you have set up your COMPANY email account on other devices such as an iPhone/iPad or an Android device, please change your password on those devices as well.</p>
            Your new password must adhere to the following requirements:
            <ul><li>It must be at least <em>$GpoMinLength</em> characters long</li>
            $ComplexityText
            <li>It may not resemble the last <em>$GpoPasswordHistory</em> passwords</li></ul>
            <p>A good password doesn't need to be very complicated. You can always use multiple random words to create a password. The following examples are as strong as any randomly generated password.</p>
            <ul><li>Correct battery horse staple</li>
            <li>CORRECTbatteryHORSEstaple?</li>
            <li>#correct4battery7horse5staple</li>
            <li>Correct!Battery,horse&staple.</li></ul>
            <p>If you have questions about or issues with changing your password, please to contact the IT department.</p>
            <p>Kind regards,</p>
            <br>
            <p><br>
            COMPANY IT department</p>"
        
            # Splat the attributes for use in the Send-MailMessage cmdlet.
            $MailAttributes = @{
                To = $To
                From = $From
                Subject = $Subject
                Body = $Body
                Port = $Port
                BodyAsHTML = $True
            }
            
            If ($UseSsl) {$MailAttributes += @{UseSsl = $True}}
            
            If ($MailCredential) {$MailAttributes += @{Credential = $MailCredential}}
        
            # Send the message to the user.
            Send-MailMessage @MailAttributes -ErrorVariable +EventLogErrors
        
            # Add a small description to the messages array if an email has been sent.
            $EventLogMessage += "`n - An email has been sent to $Name as his/her password expires $InXDays."
        }
    }
    END {
        Send-StatusMessage
        Write-StatusToEventLog
    }
}