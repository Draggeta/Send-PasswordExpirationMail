# Create an empty array for the messages which are to be written to the Event Logs.
$EventLogMessage = @()

#Import the necessary modules.  
Import-Module ActiveDirectory -ErrorVariable +EventLogErrors

Function Get-PEUser {
    [CmdletBinding(DefaultParameterSetName = 'Filter')]
    Param (
        [Parameter(ParameterSetName = 'Identity', Mandatory = $True)]
        $Identity,

        [Parameter(ParameterSetName = 'Filter')]
        $Filter = '*',

        [Parameter(ParameterSetName = 'Filter')]
        $SearchBase = (Get-ADDomain).DistinguishedName,

        [Parameter(ParameterSetName = 'Filter')]
        [ValidateSet('Subtree','OneLevel','Base')]
        $SearchScope = 'Subtree'
    )
    BEGIN {
        If ($Identity) {
            $ADUsers = Get-ADUser -Identity $Identity
        } Else {
            $ADUsers = Get-ADUser -Filter $Filter -SearchBase $SearchBase -SearchScope $SearchScope -Properties PasswordNeverExpires,PasswordLastSet,EmailAddress -ErrorVariable +EventLogErrors
        }
        $ADPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
        $Today = Get-Date 
    }
    PROCESS {
        ForEach ($ADUser in $ADUsers) {
            $FGPasswordPolicy = Get-ADUserResultantPasswordPolicy $ADUser
            If ($FGPasswordPolicy) {
                If ($ADUser.PasswordNeverExpires -eq $True) {
                    $DaysUntilExpiration = $Null
                } ElseIf ($ADUser.PasswordLastSet -eq $Null) {
                    $DaysUntilExpiration = $Null
                } ElseIf ($ADUser.PasswordNeverExpires -eq $False) {
                    $DaysUntilExpiration = (($ADUser.PasswordLastSet + $FGPasswordPolicy.MaxPasswordAge) - $Today).Days
                }
                $MinPasswordLength = $FGPasswordPolicy.MinPasswordLength
                $PasswordHistoryCount = $FGPasswordPolicy.PasswordHistoryCount
                $ComplexityEnabled = $FGPasswordPolicy.ComplexityEnabled
            }
            ElseIf (-not $FGPasswordPolicy) {
                If ($ADUser.PasswordNeverExpires -eq $True) {
                    $DaysUntilExpiration = $Null
                } ElseIf ($ADUser.PasswordLastSet -eq $Null) {
                    $DaysUntilExpiration = $Null
                } ElseIf ($ADUser.PasswordNeverExpires -eq $False) {
                    $DaysUntilExpiration = (($ADUser.PasswordLastSet + $ADPasswordPolicy.MaxPasswordAge) - $Today).Days
                }
                $MinPasswordLength = $ADPasswordPolicy.MinPasswordLength
                $PasswordHistoryCount = $ADPasswordPolicy.PasswordHistoryCount
                $ComplexityEnabled = $ADPasswordPolicy.ComplexityEnabled
            }
            $PEUserProperties = [Ordered]@{
                'DistinguishedName' = $ADuser.DistinguishedName
                'Enabled' = $ADuser.Enabled
                'Name' = $ADuser.Name
                'GivenName' = $ADuser.GivenName
                'Surname' = $ADuser.Surname
                'EmailAddress' = $ADuser.EmailAddress
                'PasswordNeverExpires' = $ADUser.PasswordNeverExpires
                'DaysUntilExpiration' = $DaysUntilExpiration
                'MinPasswordLength' = $MinPasswordLength
                'PasswordHistoryCount' = $PasswordHistoryCount
                'ComplexityEnabled' = $ComplexityEnabled
            }
            $Object = New-Object -TypeName PSObject -Property $PEUserProperties
            $Object.PSObject.TypeNames.Insert(0,'PE.PEUser')
            Write-Output $Object
        }
    }
    END {
    }
}

Function Send-PEMailMessage {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeLine, Mandatory = $True, Position = 0)]
        [ValidateScript( {
            If (($_ | Get-Member).TypeName -contains 'PE.PEUser') {$True}
            Else { Throw "$_ is not of the object type PE.PEUser" }
        } )]
        [Object[]]$PEObject,

        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True)]
        [Int[]]$RemindOn,

        [Parameter (ParameterSetName = 'Parameter')]
        [ValidateSet('Name','GivenName','Surname')]
        $AddresseeNameType = 'Name',

        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True)]
        [String]$SmtpServer,

        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True)]
        [MailAddress]$From,

        [Parameter (ParameterSetName = 'Parameter')]
        [ValidateRange(1,65535)]
        [Int]$Port = '25',

        [Parameter (ParameterSetName = 'Parameter')]
        [Switch]$UseSsl,

        [Parameter (ParameterSetName = 'Parameter')]
        [PsCredential]$MailCredential,

        [Parameter (ParameterSetName = 'ConfigFile', Mandatory = $True)]
        [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
        [String]$ConfigFile
    )
    BEGIN {
        If ($ConfigFile) {
            [Xml]$ConfigFile = (Get-Content -Path $ConfigFile)
            [String]$PSEmailServer = $ConfigFile.Settings.EmailServerSettings.SmtpServer
            [Int]$Port = $ConfigFile.Settings.EmailServerSettings.Port
            [Bool]$UseSsl = [Int]$ConfigFile.Settings.EmailServerSettings.UseSsl
            [MailAddress]$From = $ConfigFile.Settings.EmailServerSettings.From
            $MailCredential = $ConfigFile.Settings.Credentials.MailCredentials
            [Int[]]$RemindOn = ($ConfigFile.Settings.ScriptSettings.SendPasswordExpirationMail.RemindOn).Split(',')
            If ($MailCredential) {
                Import-Module CredentialManager -ErrorVariable +EventLogErrors
                $MailCredential = Get-StoredCredential -Target $MailCredential
            }
        } Else {
            $PSEmailServer = $SmtpServer
        }
    }
    PROCESS {
        ForEach ($PEUser in $PEObject) {
            If ($PEUser.DaysUntilExpiration -in $RemindOn) {
                Switch ($AddresseeNameType) {
                    Name { $Name = $PEUser.Name }
                    GivenName { $Name = $PEUser.GivenName }
                    SurName { $Name = $PEUser.SurName }
                }
                Switch ($PEUser.DaysUntilExpiration) {
                    0 { $InDays = 'today' }
                    1 { $InDays = 'in one day' }
                    Default { $InDays = "in $($PEUser.DaysUntilExpiration) days" }
                }
                Switch ($PEUser.ComplexityEnabled) {
                    $True { 
                        $ComplexityText = @(
                        '<li>The password needs to include 3 of the 4 following categories:</li>'
                        '<ul><li>At least one <em>lower case</em> letter (a-z)</li>'
                        '<li>At least one <em>upper case</em> letter (A-Z)</li>'
                        '<li>At least one <em>number</em> (0-9)</li>'
                        '<li>At least one <em>special character</em> (!,?,*,~, etc.)</li></ul>'
                        ).ToString()
                    }
                    $False { $ComplexityText = $Null }
                }
                $To = $_.EmailAddress

                # Store the subject and body into variables for use in the send-mailmessage cmdlet.
                $Subject = "Reminder: Your password expires $InDays"

                $Body = 
                "<p>Dear $Name,<br></p>
                <p>Your password will expire <em>$InDays</em>. You can change your password by pressing CTRL+ALT+DEL and then choosing the option to reset your password. If you are not at the HEAD OFFICE or BRANCH OFFICE, you can reset your password via the webmail. Click on the gear icon and subsequently on `"Change password`" to change your password.</p>
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
        # Write to the event log if the logs have been specified
        If ($EventLogName -and $EventLogSource) {
            # The values here below can be used to have it write to event logs.
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
}

Function Send-PEAdminMailMessage {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeLine, Mandatory = $True, Position = 0)]
        [ValidateScript({ If (($_ | Get-Member).TypeName -contains 'PE.PEUser') { $True } Else { Throw "$_ is not of the object type PE.PEUser" }})]
        [Object[]]$PEObject,

        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True)]
        [Int[]]$RemindOn,

        [Parameter (ParameterSetName = 'Parameter')]
        [ValidateSet(Name,GivenName,Surname)]
        $AddresseeNameType = 'Name',

        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True)]
        [String]$SmtpServer,

        [Parameter (ParameterSetName = 'Parameter', Mandatory = $True)]
        [MailAddress]$From,

        [Parameter (ParameterSetName = 'Parameter')]
        [ValidateRange(1,65535)]
        [Int]$Port = '25',

        [Parameter (ParameterSetName = 'Parameter')]
        [Switch]$UseSsl,

        [Parameter (ParameterSetName = 'Parameter')]
        [PsCredential]$MailCredential,

        [Parameter (ParameterSetName = 'ConfigFile', Mandatory = $True)]
        [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
        [String]$ConfigFile
    )
    BEGIN {
        If ($ConfigFile) {
            [Xml]$ConfigFile = (Get-Content -Path $ConfigFile)
            [String]$PSEmailServer = $ConfigFile.Settings.EmailServerSettings.SmtpServer
            [Int]$Port = $ConfigFile.Settings.EmailServerSettings.Port
            [Bool]$UseSsl = [Int]$ConfigFile.Settings.EmailServerSettings.UseSsl
            [MailAddress]$From = $ConfigFile.Settings.EmailServerSettings.From
            $MailCredential = $ConfigFile.Settings.Credentials.MailCredentials
            [Int[]]$RemindOn = ($ConfigFile.Settings.ScriptSettings.SendPasswordExpirationMail.RemindOn).Split(',')
            If ($MailCredential) {
                Import-Module CredentialManager -ErrorVariable +EventLogErrors
                $MailCredential = Get-StoredCredential -Target $MailCredential
            }
        } Else {
            $PSEmailServer = $SmtpServer
        }
    }
    PROCESS {
        ForEach ($PEUser in $PEObject) {
            If ($PEUser.DaysUntilExpiration -in $RemindOn) {
                Switch ($AddresseeNameType) {
                    Name { $Name = $PEUser.Name }
                    GivenName { $Name = $PEUser.GivenName }
                    SurName { $Name = $PEUser.SurName }
                }
                Switch ($PEUser.DaysUntilExpiration) {
                    0 { $InDays = 'today' }
                    1 { $InDays = 'in one day' }
                    Default { $InDays = "in $($PEUser.DaysUntilExpiration) days" }
                }
                Switch ($PEUser.ComplexityEnabled) {
                    $True { 
                        $ComplexityText = 
                        "<li>The password needs to include 3 of the 4 following categories:</li>
                        <ul><li>At least one <em>lower case</em> letter (a-z)</li>
                        <li>At least one <em>upper case</em> letter (A-Z)</li>
                        <li>At least one <em>number</em> (0-9)</li>
                        <li>At least one <em>special character</em> (!,?,*,~, etc.)</li></ul>"
                    }
                    $False { $ComplexityText = $Null }
                }
                $To = $_.EmailAddress

                # Store the subject and body into variables for use in the send-mailmessage cmdlet.
                $Subject = "Reminder: Your password expires $InDays"

                $Body = 
                "<p>Dear $Name,<br></p>
                <p>Your password will expire <em>$InDays</em>. You can change your password by pressing CTRL+ALT+DEL and then choosing the option to reset your password. If you are not at the HEAD OFFICE or BRANCH OFFICE, you can reset your password via the webmail. Click on the gear icon and subsequently on `"Change password`" to change your password.</p>
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
        # Write to the event log if the logs have been specified
        If ($EventLogName -and $EventLogSource) {
            # The values here below can be used to have it write to event logs.
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
}
