###############################################################################################
# A basic script to send out emails to users about expiring passwords                       
###############################################################################################

Import-Module ActiveDirectory -ErrorVariable +EventLogErrors


# Specify a source and the array for the messages to be written to the Event Logs.
$EventLogName = 'Application'
$EventLogSource = 'MyScripts'
$EventLogMessage = @()

# Set the mail account username and password. Skip this if you are using an anonymous relay
$User = 'relay@domain.com'
$Password = 'P@$$w0rd!'
$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($User,$SecurePassword)

# Set the mail server properties
$From = 'noreply@domain.com'
$SmtpServer = 'server.domain.com'
$Port = '587'

# Set the days before expiration on which a mail should be sent, but not after their account has been locked
$ReminderDays = 0,1,2,3,7,14

# Set the OU where your users are located
$SearchBase = 'OU=users,OU=company,DC=domain,DC=com'

# Set a few default variables, most of these needen't be changed. 
$MaxAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
$MinLength = (Get-ADDefaultDomainPasswordPolicy).MinPasswordLength
$PasswordHistory = (Get-ADDefaultDomainPasswordPolicy).PasswordHistoryCount
$ComplexityEnabled = (Get-ADDefaultDomainPasswordPolicy).ComplexityEnabled
$ADFilter = {(PasswordNeverExpires -eq $False) -and (PasswordExpired -eq $False) -and (Enabled -eq $True) -and (Emailaddress -ne "$Null") -and (PasswordLastSet -ne "$Null")}


# Search all users who match the aforementioned filters
Get-ADUser -Filter $ADFilter -SearchBase $SearchBase -SearchScope Subtree -Properties PasswordLastSet,EmailAddress -ErrorVariable +EventLogErrors | ForEach-Object {
    

    # Set a few variables per user for easier usage
    $PasswordLastSet = $_.PasswordLastSet
    $Email = $_.EmailAddress
    $ExpirationDate = $PasswordLastSet + $MaxAge
    $Today = Get-Date
    $DaysLeft = ($ExpirationDate - $Today).days
    
    # Check if the amount of days left for the current user is in the previously specified days.
    $IntervalHit = $DaysLeft -in $ReminderDays

    
    # If the password will expire, send a mail to the user
    If ($IntervalHit) {


        # Sets the user's name. Can be changed to full name or omitted if no name should be used        
        $Name = $_.GivenName

        # Changes a few of the texts used in the mail. Need to see if this can be used with the Switch command
        $InXDays = Switch ($DaysLeft) {
            0 {'today'}
            1 {'in one day'}
            Default {"in $DaysLeft days"}
        }
        
        # Automatically apply or omit the text about complexity dependant on group policy
        $ComplexityText = If ($ComplexityEnabled) {"<li>The password needs to include 3 of the 4 following categories:</li>
        <ul><li>At least one <em>lower case</em> letter (a-z)</li>
        <li>At least one <em>upper case</em> letter (A-Z)</li>
        <li>At least one <em>numberr</em> (0-9)</li>
        <li>At least one <em>special character</em> (!,?,*,~, etc.)</li></ul>"}

        $Subject = "Reminder: Your password expires $InXDays"
        $Message = "<p>Dear $Name,<br></p>
        <p>Your password will expire <em>$InXDays</em>. You can change your password by pressing CTRL+ALT+DEL and then choosing the option to reset your password. If you are not at the HEAD OFFICE or BRANCH OFFICE, you can reset your password via the webmail. Click on the gear icon and subsequently on `"Change password`" to change your password.</p>
        <p>If you have set up your COMPANY email account on other devices such as an iPhone/iPad or an Android device, please change your password on those devices as well.</p>
        Your new password must adhere to the following requirements:
        <ul><li>It must be at least <em>$MinLength</em> characters long</li>
        $ComplexityText
        <li>It may not resemble the last <em>$PasswordHistory</em> passwords</li></ul>
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

        $SendMailMessageAttributes = @{

            To = $Email
            From = $From
            Subject = $Subject
            Body = $Message
            SmtpServer = $SmtpServer
            Credential = $Credential
            Port = $Port
            UseSsl = $True
            BodyAsHTML = $True

        }

        Send-MailMessage @SendMailMessageAttributes -ErrorVariable +EventLogErrors

        $EventLogMessage += "`n - A mail has been sent to $Name as his/her password expires in $DaysLeft days."

    }

}


# The values here below can be used to have it write to event logs. My knowledge is a bit bad about error handling so improvements are welcome.
If ($EventLogErrors) {Write-EventLog -LogName $EventLogName -Source $EventLogSource -EventId 2 -EntryType Warning -Message "One or more errors occured while running the Send-PasswordExpirationMail Powershell script. See the errors below:`n`n $EventLogErrors"}
Else {Write-EventLog -LogName $EventLogName -Source $EventLogSource -EventId 1 -EntryType Information -Message "The Send-PasswordExpirationMail Powershell script ran succesfully.`n $EventLogMessage"}