<#
.SYNOPSIS
    Sends out password expiration notices.
.DESCRIPTION
    The Send-PasswordExpirationMail.ps1 script sends out password expiration notices to users. Globally, the script is split into five parts: 
	
		- Setting of 'fixed' variables 
		- AD query/foreach loop 
		- Setting of 'volatile' variables in the loop
		- Sending of the emails (if any)
		- Writing errors or successes to the Event Log
	
	The important variables to change are the EventLog*, AD*, Mail* and PwdReminderDays variables. To use the EventLog variables, you may need to create your own Logname or Source with the New-EventLog cmdlet.
	The ADFilter should suffice for most organizations, but the ADSearchBase variable will need to either be set or removed, depending on your preference.
	Mail variables such as MailServer and MailPassword are used to talk to the mail server. 
	To specify the remaining days on which to send the emails, you can type in an array of integers in the PwdReminderDays variable.
	
	It's important to note is that this script is written with authenticated relays in mind. If it's necessary to use anonymous relays, please change, remove or comment out the following values/variables:
	
		- $MailServer
		- $MailPort
		- $MailUsername
		- $MailPassword
		- $MailSecurePassword
		- $MailCredential
		- In $MailAttributes, 'UseSsl = $True', 'Port = $MailPort', 'Credential = $MailCredential'
.INPUTS
	None. You cannot pipe objects to Send-PasswordExpirationMail.ps1
.OUTPUTS
	None. Send-PasswordExpirationMail.ps1 only outputs to the EventLog
.NOTES
    Author:   Tony Fortes Ramos
    Date:     April 24, 2014
    Modified: August 3, 2015    
.LINK
	New-EventLog
	Write-EventLog
#>

# Specify a logname, source and array for the messages that are to be written to the Event Logs. If they don't exist, create them.
$EventLogName = 'Application'
$EventLogSource = 'MyScripts'
$EventLogMessage = @()

Import-Module ActiveDirectory -ErrorVariable +EventLogErrors

# Set a few default variables, most of these needn't be changed unless you use fine-grained password policies.
$GpoMaxAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
$GpoMinLength = (Get-ADDefaultDomainPasswordPolicy).MinPasswordLength
$GpoPasswordHistory = (Get-ADDefaultDomainPasswordPolicy).PasswordHistoryCount
$GpoComplexityEnabled = (Get-ADDefaultDomainPasswordPolicy).ComplexityEnabled
$Today = Get-Date 

# Specify the filters to use and the OU where the users are located in the domain.
$ADFilter = {(PasswordNeverExpires -eq $False) -and (PasswordExpired -eq $False) -and (Enabled -eq $True) -and (Emailaddress -ne "$Null") -and (PasswordLastSet -ne "$Null")}
$ADSearchBase = 'OU=users,OU=company,DC=domain,DC=com'

# Specify the mail server properties. Note that the from address can be different from the mail account address/username.
$PSEmailServer = 'server.domain.com'
$MailPort = '587'
$MailFrom = 'noreply@domain.com'

# Specify the mail account username and password. Comment the following section or delete it if you are using an anonymous relay.
$MailUsername = 'noreply@domain.com'
$MailPassword = 'P@$$w0rd!'
$MailSecurePassword = $MailPassword | ConvertTo-SecureString -AsPlainText -Force
$MailCredential = New-Object System.Management.Automation.PSCredential ($MailUsername,$MailSecurePassword)

# Specify the days before expiration on which an email should be sent reminding the user to change his password.
$PwdReminderDays = 0,1,2,3,7,14

# Find all users who match the filters set in the previous command and then loop through each.
Get-ADUser -Filter $ADFilter -SearchBase $ADSearchBase -SearchScope Subtree -Properties PasswordLastSet,EmailAddress -ErrorVariable +EventLogErrors | 
ForEach-Object -Process {

    # Set a few variables per user for easier usage. This is mostly cosmetic.
    $PwdLastSet = $_.PasswordLastSet
    $PwdExpirationDate = $PwdLastSet + $GpoMaxAge
    $PwdDaysLeft = ($PwdExpirationDate - $Today).days
    
    # Check if the days left until the user's password expired is in the array specified earlier.
    $PwdIntervalHit = $PwdDaysLeft -in $PwdReminderDays

    # If the amount of days left until expiration is in the array, send an email to the user.
    If ($PwdIntervalHit) {

        $MailName = $_.GivenName
        $MailTo = $_.EmailAddress

        # Changes a few of the texts used in the mail and stores it in the variable.
        $InXDays = Switch ($PwdDaysLeft) {

            0 {'today'}
            1 {'in one day'}
            Default {"in $PwdDaysLeft days"}

        }
        
        # Automatically apply or omit the text about complexity, dependant on if the group policy has been set.
        $MailComplexityText = If ($GpoComplexityEnabled) {"<li>The password needs to include 3 of the 4 following categories:</li>
        <ul><li>At least one <em>lower case</em> letter (a-z)</li>
        <li>At least one <em>upper case</em> letter (A-Z)</li>
        <li>At least one <em>numberr</em> (0-9)</li>
        <li>At least one <em>special character</em> (!,?,*,~, etc.)</li></ul>"}

        # Store the subject and body into variables for use in the send-mailmessage cmdlet.
        $MailSubject = "Reminder: Your password expires $InXDays"

        $MailBody = "<p>Dear $MailName,<br></p>
        <p>Your password will expire <em>$InXDays</em>. You can change your password by pressing CTRL+ALT+DEL and then choosing the option to reset your password. If you are not at the HEAD OFFICE or BRANCH OFFICE, you can reset your password via the webmail. Click on the gear icon and subsequently on `"Change password`" to change your password.</p>
        <p>If you have set up your COMPANY email account on other devices such as an iPhone/iPad or an Android device, please change your password on those devices as well.</p>
        Your new password must adhere to the following requirements:
        <ul><li>It must be at least <em>$GpoMinLength</em> characters long</li>
        $MailComplexityText
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

            To = $MailTo
            From = $MailFrom
            Subject = $MailSubject
            Body = $MailBody
            Credential = $MailCredential
            Port = $MailPort
            UseSsl = $True
            BodyAsHTML = $True

        }

        # Send the message to the user.
        Send-MailMessage @MailAttributes -ErrorVariable +EventLogErrors

        # Add a small description to the messages array if an email has been sent. Need to implement testing to check if an email has actually been sent out.
        $EventLogMessage += "`n - An email has been sent to $MailName as his/her password expires in $InXDays."

    }

} -End {

    # The values here below can be used to have it write to event logs. My knowledge is a bit bad about error handling so improvements are welcome.
    $WriteEventLogWarning = @{
    
        LogName = $EventLogName
        Source = $EventLogSource
        EventId = '2'
        EntryType = 'Warning'
        Message = "One or more errors occured while running the Send-PasswordExpirationMail Powershell script. See the errors below:`n`n $EventLogErrors"
    
    }
    
    $WriteEventLogInformation = @{
    
        LogName = $EventLogName
        Source = $EventLogSource
        EventId = '1'
        EntryType = 'Information'
        Message = "The Send-PasswordExpirationMail Powershell script ran succesfully.`n $EventLogMessage"
    
    }
    
    If ($EventLogErrors) {Write-EventLog @WriteEventLogWarning}
    Else {Write-EventLog @WriteEventLogInformation}
    
}