<#
.SYNOPSIS
    Sends out password expiration notices.
.DESCRIPTION
    The Send-PasswordExpirationMail.ps1 script sends out password expiration notices to users. While it is a bit convoluted it is written to work in most environments without changing much, if anything outside of the parameters.
	To use the EventLog variables, you may need to create your own Logname or Source with the New-EventLog cmdlet if not using a default Windows log and source.
	The AD Filter should suffice for most organizations. There is no smtp server specified in the Send-MailMessage cmdlet as we are using the global $PSEmailServer variable.
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
.PARAMETER Credential
    Specify credentials allowed to send emails from the from address. This is necessary if authentication is needed and the account running the task isn't allowed to send via that address. 
    Can be used with the Get-Credential cmdlet. If not specified it will use the credentials specified in the script. Those can be specified plain text below or pulled from an a file.
If those are not specified either (commented/removed), the emails will be sent anonymously.
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
    Modified: August 10, 2015
.LINK
	New-EventLog
	Write-EventLog
    Get-Credential
    Get-ADUser
    Send-MailMessage
#>

[CmdletBinding(DefaultParameterSetName='None')]
Param (

    [Parameter (Position = 0, Mandatory = $True)]
    [Array]$RemindOn,
    [Parameter (Position = 1, Mandatory = $True)]
    [String]$SmtpServer,
    [Parameter (Position = 2, Mandatory = $True)]
    [String]$From,
    [ValidateRange(1,65535)]
    [Int]$Port = '25',
    [Switch]$UseSsl,
    [String]$SearchBase = (Get-ADDomain).DistinguishedName,
    [Parameter (ParameterSetName = 'EventLog',Mandatory = $True)]
    [String]$EventLogName,
    [Parameter (ParameterSetName = 'EventLog',Mandatory = $True)]
    [String]$EventLogSource,
<<<<<<< HEAD
    [Object]$MailCredential,
    [Parameter (ParameterSetName = 'ConfigFile',Mandatory = $True)]
    [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
    [XML]$ConfigFile = (Get-Content -Path $_)

)

# Creates the log collection
=======
    [Parameter (ParameterSetName = 'EventLog')]
    [Int]$EventWarningID = 1,
    [Parameter (ParameterSetName = 'EventLog')]
    [Int]$EventInformationID = 2,
    [Object]$Credential

)

# Specify an array for the messages that are to be written to the Event Logs.
>>>>>>> origin/development
$EventLogMessage = @()

#Import modules
Import-Module ActiveDirectory -ErrorVariable +EventLogErrors

# Set a few default variables, most of these needn't be changed unless you use fine-grained password policies.
$GpoMaxAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
$GpoMinLength = (Get-ADDefaultDomainPasswordPolicy).MinPasswordLength
$GpoPasswordHistory = (Get-ADDefaultDomainPasswordPolicy).PasswordHistoryCount
$GpoComplexityEnabled = (Get-ADDefaultDomainPasswordPolicy).ComplexityEnabled
$Today = Get-Date 

# The filters to use.
$Filter = {(PasswordNeverExpires -eq $False) -and (PasswordExpired -eq $False) -and (pwdLastSet -ne '0') -and (PasswordLastSet -ne "$Null") -and (Enabled -eq $True) -and (Emailaddress -ne "$Null")}

<<<<<<< HEAD
If ($ConfigFile) {
=======
# Set the $PSEmailServer variable so it doesn't need to be set later.
$PSEmailServer = $SmtpServer
>>>>>>> origin/development

    $SmtpServer = $ConfigFile.Settings.EmailServerSettings.SmtpServer
    $Port = $ConfigFile.Settings.EmailServerSettings.Port
    [Switch]$UseSsl = [Bool]$ConfigFile.Settings.EmailServerSettings.UseSsl
    $From = $ConfigFile.Settings.EmailServerSettings.From

<<<<<<< HEAD
    $SearchBase = $ConfigFile.Settings.DomainSettings.UserSearchBase
    
    $MailCredential = $ConfigFile.Settings.Credentials.MailCredentials

    $RemindOn = $ConfigFile.Settings.ScriptSettings.SendPasswordExpirationMail.RemindOn
}

If ($MailCredential) {
    Import-Module CredentialManager -ErrorVariable +EventLogErrors
    $MailCredential = Get-StoredCredential -Target $MailCredential
}

# Find all users who match the filters set in the previous sections and then loop through each.
=======
# Find all users who match the filters set in the previous variables and then loop through each.
>>>>>>> origin/development
Get-ADUser -Filter $Filter -SearchBase $SearchBase -SearchScope Subtree -Properties PasswordLastSet,EmailAddress -ErrorVariable +EventLogErrors | 
ForEach-Object -Process {

    # Set a few variables per user for easier usage. This is mostly cosmetic.
    $PwdLastSet = $_.PasswordLastSet
    $PwdExpirationDate = $PwdLastSet + $GpoMaxAge
    $PwdDaysLeft = ($PwdExpirationDate - $Today).days
    
    # Check if the days left until the user's password expired is in the array specified earlier.
    $PwdIntervalHit = $PwdDaysLeft -in $RemindOn

    # If the amount of days left until expiration is in the array, send an email to the user.
    If ($PwdIntervalHit) {

        $Name = $_.GivenName
        $To = $_.EmailAddress

        # Changes a few of the texts used in the mail and stores it in the variable.
        $InXDays = Switch ($PwdDaysLeft) {

            0 {'today'}
            1 {'in one day'}
            Default {"in $PwdDaysLeft days"}

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

        # Splat the attributes for use in the send-mailmessage cmdlet.
        $MailAttributes = @{

            To = $To
            From = $From
            Subject = $Subject
            Body = $Body
            Port = $Port
            BodyAsHTML = $True

        }
        
        If ($UseSsl) {$Ssl = @{UseSsl = $True}}
        
        If ($MailCredential) {$Cred = @{Credential = $MailCredential}}

        # Send the message to the user.
        Send-MailMessage @MailAttributes @Ssl @Cred -ErrorVariable +EventLogErrors

        # Add a small description to the messages array if an email has been sent.
        $EventLogMessage += "`n - An email has been sent to $Name as his/her password expires in $InXDays."

    }

} -End {

    If ($EventLogName -ne "$Null" -and $EventLogSource -ne "$Null") {

        # The values here below can be used to have it write to event logs. My knowledge is a bit bad about error handling/outputting errors so improvements are welcome.
        $WriteEventLogWarning = @{
    
            LogName = $EventLogName
            Source = $EventLogSource
            EventId = $EventWarningID
            EntryType = 'Warning'
            Message = "One or more errors occured while running the Send-PasswordExpirationMail Powershell script. See the errors below:`n`n $EventLogErrors"
    
        }
    
        $WriteEventLogInformation = @{
    
            LogName = $EventLogName
            Source = $EventLogSource
            EventId = $EventInformationID
            EntryType = 'Information'
            Message = "The Send-PasswordExpirationMail Powershell script ran succesfully.`n $EventLogMessage"
    
        }
    
        If ($EventLogErrors) {Write-EventLog @WriteEventLogWarning}
        Else {Write-EventLog @WriteEventLogInformation}

    }
    
}