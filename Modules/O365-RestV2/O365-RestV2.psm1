$Script:baseUri = "https://outlook.office.com/api/v2.0"


function Get-O365RestV2CalendarItem {
    <#
        .SYNOPSIS
            Gets calendar events from the main user calendar.
        .DESCRIPTION
            This function allows a user or admin to create an event or meeting in his calendar or another user's calendar. 
        .PARAMETER Subject
            Specifies the subject of the meeting or event. This can be any string.
        .PARAMETER Note
            Specifies the text in the body of the meeting or event. This can be any string and in plain-text or as HTML. If HTML is used, use the -AsHtml switch to correctly specify the text type.
        .PARAMETER AsHtml
            Specifies if the Note parameter text is formatted as HTML or not. Defaults to false. 
        .PARAMETER Attendees
            Specifies the attendees for this meeting. While the objects containing the attendees can be created manually, it is easier to use the New-O365RestAttendee to create them. The cmdlet can be used as follows:
        .PARAMETER Location
            Specifies the display name of the location the meeting is planned. If the meeting is planned in a room resource, the display name of the room can be specified here. However, this will only fill in the location field while not booking the room. To book a room, add the it as an attendee of the type 'Resource'.
        .PARAMETER StartDate
            Specifies the start date and time of the event or meeting. This can be specified with the Get-Date cmdlet, but also typed naturally such as '2016-05-24 13:15:00'.
        .PARAMETER StartTimeZone
            Specifies the time zone of the StartDate parameter. By default this will use the computer's current timezone. A list of available time zones can be found by using the .NET command '[System.TimeZoneInfo]::GetSystemTimeZones()'.
        .PARAMETER EndDate
            Specify the end date and time of the event or meeting. This can be specified with the Get-Date cmdlet, but also typed naturally such as '2016-05-24 14:15:00'.
        .PARAMETER EndTimeZone
            Specifies the time zone of the EndDate parameter. By default this will use the computer's current timezone. A list of available time zones can be found by using the .NET command '[System.TimeZoneInfo]::GetSystemTimeZones()'.
        .PARAMETER Reminder
            Specifies when a reminder should be displayed. This value is an integer and specifies the amount of minutes.
        .PARAMETER AllDay
            Specifies if a meeting or event will take place all day. If specified, normally the time of the start and end date need to be '00:00:00', otherwise the command fails. However, this script accounts for that and sets the time to '00:00:00' if this switch has been specified.
        .PARAMETER ShowAs
            Specifies if the meeting or event should show the user as free, busy, working elswhere, tentative or away.
        .PARAMETER UserPrincipalName
            Specifies the user's calendar the meeting is to be created in. This parameter defaults to the user whose credentials are specified in the credential paramter. It should follow the same pattern as an email address or any normal UPN. 
            
            The correct permissions to create events/meetings in the specified user's calendar is required.
        .PARAMETER Credential
            Specifies the user account credentials to use to perform this task. 
            
            To specify this parameter, you can type a user name, such as 'User1@contoso.com' or you can specify a PSCredential object. If you specify a user name for this parameter, the cmdlet prompts for a password.
            
            You can also create a PSCredential object by using a script or by using the Get-Credential cmdlet. You can then set the Credential parameter to the PSCredential object The following example shows how to create credentials.
            
            $AdminCredentials = Get-Credential "User01@contoso.com"
            
            The following shows how to set the Credential parameter to these credentials.
            
            -Credential $AdminCredentials
            
            If the acting credentials do not have the correct permission to perform the task, PowerShell returns a terminating error.
        .EXAMPLE
            New-O365RestCalendarItem -Subject 'Testing the API.' -Note 'Testing the API is a great success!' -StartDate (Get-Date) -EndDate (Get-Date).AddHours(1) -Credential $Credential -ShowAs 'Free' -AsHTML 
            Description
            
            -----------
        
            This command creates a meeting in the logged in user's default calendar with the specified subject and notes, while showing the user as free.
        .EXAMPLE
            $Attendees = New-O365RestAttendee -EmailAddress Mark@domain.com,Sally@contoso.com -Type Required

            The following shows how to set the Attendees parameter to these attendees.

            New-O365RestCalendarItem -Subject 'Testing the API.' -Note 'Testing the API is a great success!' -Attendees $Attendees -StartDate (Get-Date)  -Credential $Credential -AllDay 
            Description
            
            -----------
        
            This command creates a meeting in the logged in user's default calendar with the specified subject and notes, and the specified attendees.
        .INPUTS
        	None. You cannot pipe objects to New-O365RestCalendarItem.
        .OUTPUTS
        	New-O365RestCalendarItem outputs the response from the server.
            Author:   Tony Fortes Ramos
            Created:  May 15, 2016
        .LINK
        	New-O365RestAttendee
        .COMPONENT
            New-O365RestAttendee            
    #>
    [CmdletBinding(
        SupportsShouldProcess = $True,
        ConfirmImpact= 'Low'
    )]
    param(
        [Parameter()]
        [string]$EventId,

        [Parameter()]
        [datetime]$StartDate,

        [Parameter()]
        [datetime]$EndDate,

        [Parameter()]
        [ValidateScript({ [System.TimeZoneInfo]::FindSystemTimeZoneById($_) })]
        [string]$TimeZone = [System.Timezone]::CurrentTimeZone.StandardName,

        [Parameter(ParameterSetName = 'ResultSize')]
        [ValidateRange(1,499)]
        [int]$ResultSize,          

        [Parameter(ParameterSetName = 'All')]
        [switch]$All,

        [Parameter()]
        [mailaddress]$UserPrincipalName,

        [Parameter(Mandatory = $True)]
        [string]$AccessToken
    )

    begin {
        $contentType = "application/json"
        $headers = @{
            Authorization = "Bearer $AccessToken"
            Accept = 'application/json'
            OData = 'verbose'
            Prefer = "outlook.timezone = `"$TimeZone`""
        }
        $timeFormat = "yyyy-MM-ddTHH:mm:ss"
    }
    
    process {
        switch ($UserPrincipalName) {
            { $UserPrincipalName }      { $uri =  "$Script:baseUri/users/$($UserPrincipalName.Address)/events" }
            { -not $UserPrincipalName } { $uri =  "$Script:baseUri/me/events" }
            { $EventId }                { $uri += "/$EventId" }
            { $ResultSize }             { $uri += "?`$top=$ResultSize" }
            { -not $ResultSize }        { $uri += "?`$top=10" }
        }
        $filter = @()
        switch ($filter) {
            { $StartDate }  {
                                $startString = Get-Date $StartDate -Format $timeFormat 
                                $filter += "Start/DateTime ge '$startString'" 
                            }
            { $EndDate }    {
                                $endString = Get-Date $EndDate -Format $timeFormat
                                $filter += "End/DateTime le '$endString'" 
                            }
        }
        if ($filter) {
            $uri += "&`$filter=$($filter -join '&')"
        }   
        if ($PSCmdlet.ShouldProcess($UserPrincipalName)) { 
            Invoke-RestMethod -Uri $uri -Method Get -ContentType $contentType -Headers $headers
        }
    }
    
    end {
    }
}


function New-O365RestV2Attendee {
    <#
        .SYNOPSIS
            Create an Attendee object.
        .DESCRIPTION
            This function creates an attendee object for use with the New-O365RestCalendarItem Attendee parameter. 
        .PARAMETER EmailAddress
            Specifies an array of emailaddresses which should have the same type applied to them.
        .PARAMETER Note
            Specifies the type of attendee. Required means that the user is required to be there. Optional specifies that the attendee isn't required to be there. Resource specifies that the user is a resource. This can be an equipment or room resource.
        .EXAMPLE
            $Attendees = New-O365RestAttendee -EmailAddress Mark@domain.com,Sally@contoso.com -Type Required
            $Attendees += New-O365RestAttendee -EmailAddress MeetingRoom1@domain.com -Type Resource
            Description
            -----------
        
            This command creates an array of attendees, both with the type required and resource. These can be passed down to the Attendee parameter of the New-O365RestCalendarItem function.
        .INPUTS
        	None. You cannot pipe objects to New-O365RestAttendee.
        .OUTPUTS
        	Object. New-O365RestAttendee outputs a PSObject.
        .NOTES
            Author:   Tony Fortes Ramos
            Created:  May 15, 2016
        .LINK
        	New-O365RestCalendarItem
        .COMPONENT
            New-O365RestCalendarItem            
    #>
    [CmdletBinding(
        SupportsShouldProcess = $True,
        ConfirmImpact= 'Low'
    )]
    param(
        [Parameter(Position = 0, Mandatory = $True)]
        [mailaddress[]]$EmailAddress,

        [Parameter()]
        [ValidateSet('Required','Optional','Resource')]
        $Type = 'Required'
    )

    begin {
    }
    
    process {
        foreach ($Address in $EmailAddress) {
            $properties = @{
                Name = $Address.Address
                EmailAddress = $Address.Address
                Type = $Type
            }
            if ($PSCmdlet.ShouldProcess($Address.Address)) { 
                $object = New-Object -TypeName PSObject -Property $properties
                return $object
            }
        }
    }
    
    end {
    }
}


function New-O365RestV2CalendarItem {
    <#
        .SYNOPSIS
            Creates a calendar item via the Office 365 REST API.
        .DESCRIPTION
            This function allows a user or admin to create an event or meeting in his calendar or another user's calendar. 
        .PARAMETER Subject
            Specifies the subject of the meeting or event. This can be any string.
        .PARAMETER Note
            Specifies the text in the body of the meeting or event. This can be any string and in plain-text or as HTML. If HTML is used, use the -AsHtml switch to correctly specify the text type.
        .PARAMETER AsHtml
            Specifies if the Note parameter text is formatted as HTML or not. Defaults to false. 
        .PARAMETER Attendees
            Specifies the attendees for this meeting. While the objects containing the attendees can be created manually, it is easier to use the New-O365RestAttendee to create them. The cmdlet can be used as follows:
        .PARAMETER Location
            Specifies the display name of the location the meeting is planned. If the meeting is planned in a room resource, the display name of the room can be specified here. However, this will only fill in the location field while not booking the room. To book a room, add the it as an attendee of the type 'Resource'.
        .PARAMETER StartDate
            Specifies the start date and time of the event or meeting. This can be specified with the Get-Date cmdlet, but also typed naturally such as '2016-05-24 13:15:00'.
        .PARAMETER StartTimeZone
            Specifies the time zone of the StartDate parameter. By default this will use the computer's current timezone. A list of available time zones can be found by using the .NET command '[System.TimeZoneInfo]::GetSystemTimeZones()'.
        .PARAMETER EndDate
            Specify the end date and time of the event or meeting. This can be specified with the Get-Date cmdlet, but also typed naturally such as '2016-05-24 14:15:00'.
        .PARAMETER EndTimeZone
            Specifies the time zone of the EndDate parameter. By default this will use the computer's current timezone. A list of available time zones can be found by using the .NET command '[System.TimeZoneInfo]::GetSystemTimeZones()'.
        .PARAMETER Reminder
            Specifies when a reminder should be displayed. This value is an integer and specifies the amount of minutes.
        .PARAMETER AllDay
            Specifies if a meeting or event will take place all day. If specified, normally the time of the start and end date need to be '00:00:00', otherwise the command fails. However, this script accounts for that and sets the time to '00:00:00' if this switch has been specified.
        .PARAMETER ShowAs
            Specifies if the meeting or event should show the user as free, busy, working elswhere, tentative or away.
        .PARAMETER UserPrincipalName
            Specifies the user's calendar the meeting is to be created in. This parameter defaults to the user whose credentials are specified in the credential paramter. It should follow the same pattern as an email address or any normal UPN. 
            
            The correct permissions to create events/meetings in the specified user's calendar is required.
        .PARAMETER Credential
            Specifies the user account credentials to use to perform this task. 
            
            To specify this parameter, you can type a user name, such as 'User1@contoso.com' or you can specify a PSCredential object. If you specify a user name for this parameter, the cmdlet prompts for a password.
            
            You can also create a PSCredential object by using a script or by using the Get-Credential cmdlet. You can then set the Credential parameter to the PSCredential object The following example shows how to create credentials.
            
            $AdminCredentials = Get-Credential "User01@contoso.com"
            
            The following shows how to set the Credential parameter to these credentials.
            
            -Credential $AdminCredentials
            
            If the acting credentials do not have the correct permission to perform the task, PowerShell returns a terminating error.
        .EXAMPLE
            New-O365RestCalendarItem -Subject 'Testing the API.' -Note 'Testing the API is a great success!' -StartDate (Get-Date) -EndDate (Get-Date).AddHours(1) -Credential $Credential -ShowAs 'Free' -AsHTML 
            Description
            
            -----------
        
            This command creates a meeting in the logged in user's default calendar with the specified subject and notes, while showing the user as free.
        .EXAMPLE
            $Attendees = New-O365RestAttendee -EmailAddress Mark@domain.com,Sally@contoso.com -Type Required

            The following shows how to set the Attendees parameter to these attendees.

            New-O365RestCalendarItem -Subject 'Testing the API.' -Note 'Testing the API is a great success!' -Attendees $Attendees -StartDate (Get-Date)  -Credential $Credential -AllDay 
            Description
            
            -----------
        
            This command creates a meeting in the logged in user's default calendar with the specified subject and notes, and the specified attendees.
        .INPUTS
        	None. You cannot pipe objects to New-O365RestCalendarItem.
        .OUTPUTS
        	New-O365RestCalendarItem outputs the response from the server.
            Author:   Tony Fortes Ramos
            Created:  May 15, 2016
        .LINK
        	New-O365RestAttendee
        .COMPONENT
            New-O365RestAttendee            
    #>
    [CmdletBinding(
        SupportsShouldProcess = $True,
        ConfirmImpact= 'Medium'
    )]
    param(
        [Parameter(Mandatory = $True)]
        [string]$Subject,
        
        [Parameter()]
        [string]$Note,
        
        [Parameter()]
        [switch]$AsHtml,
        
        [Parameter()]
        $Attendees,

        [Parameter()]
        $Location,

        [Parameter()]
        [datetime]$StartDate = (Get-Date),
        
        [Parameter()]
        [ValidateScript({ [System.TimeZoneInfo]::FindSystemTimeZoneById($_) })]
        [string]$StartTimeZone = [System.Timezone]::CurrentTimeZone.StandardName,
        
        [Parameter()]
        [datetime]$EndDate = (Get-Date).AddMinutes(30),
        
        [Parameter()]
        [ValidateScript({ [System.TimeZoneInfo]::FindSystemTimeZoneById($_) })]
        [string]$EndTimeZone = [System.Timezone]::CurrentTimeZone.StandardName,
        
        [Parameter()]
        [int]$Reminder,

        [Parameter()]
        [switch]$AllDay,
        
        [Parameter()]
        [ValidateSet('Free','WorkingElsewhere','Tentative','Busy','Away')]
        [string]$ShowAs = 'Busy',
        
        [Parameter()]
        [mailaddress]$UserPrincipalName,

        [Parameter(Mandatory = $True)]
        [string]$AccessToken
    )

    begin {
        $contentType = "application/json"
        $headers = @{
            Authorization = "Bearer $AccessToken"
            Accept = 'application/json';
            OData = 'verbose'
        }
        $timeFormat = "yyyy-MM-ddTHH:mm:ss"
    }
    
    process {
        switch ($UserPrincipalName) {
            { $UserPrincipalName }      { $uri = "$Script:baseUri/users/$($UserPrincipalName.Address)/events" }
            { -not $UserPrincipalName } { $uri = "$Script:baseUri/me/events" }
        }
        switch ($AsHTML) {
            $False { $NoteContentType = 'Text' }
            $True { $NoteContentType = 'HTML' }
        }
        switch ($AllDay.IsPresent) {
            { $StartDate -and ($_ -eq $False) } { $start = Get-Date $StartDate -Format $timeFormat }
            { $EndDate -and ($_ -eq $False) }   { $end = Get-Date $EndDate -Format $timeFormat }
            { $StartDate -and ($_ -eq $True) }  { $start = Get-Date $StartDate.Date -Format $timeFormat }
            { $EndDate -and ($_ -eq $True) }    { $end = Get-Date $EndDate.Date -Format $timeFormat }
        }
        $attendeesProperties = foreach ($Attendee in $Attendees) {
            @{
                EmailAddress = @{
                    Address = $Attendee.EmailAddress
                    Name = $Attendee.Name
                }
                Type = $Attendee.Type
            }
        }
        $body = @{}
        switch ($body) {
            { $Subject }        { $body.Subject = $Subject }
            { $Note }           { 
                                  $body.Body = @{}
                                  $body.Body.Content = $Note
                                  $body.Body.ContentType = $NoteContentType
                                }
            { $start }          { 
                                  $body.Start = @{}
                                  $body.Start.DateTime = $start
                                  $body.Start.TimeZone = $StartTimeZone 
                                }
            { $end }            { 
                                  $body.End = @{}
                                  $body.End.DateTime = $end
                                  $body.End.TimeZone = $EndTimeZone
                                }
            { $Reminder } { 
                                  $body.ReminderMinutesBeforeStart = $Reminder
                                  $body.IsReminderOn = ($Reminder -ne 0)
                                }
            { $Attendees }      { $body.Attendees = $attendeesProperties }
            { $Location }       { 
                                  $body.Location = @{}
                                  $body.Location.DisplayName = $Location
                                  $body.Location.Address = $null
                                  $body.Location.Coordinates = $null
                                }
            { $ShowAs }         { $body.ShowAs = $ShowAs }
            { $AllDay }         { $body.IsAllDay = $AllDay.IsPresent }
        }
        if ($PSCmdlet.ShouldProcess("$Subject with a start time of $start", "create an appointment")) { 
            Invoke-RestMethod -Uri $uri -Method Post -ContentType $contentType -Headers $headers -Body (ConvertTo-Json $body -Depth 10)
        }
    }
    
    end {
    }
}


function Remove-O365RestV2CalendarItem {
    <#
        .SYNOPSIS
            Retrieves an Azure authorization code.
        .DESCRIPTION
            This cmdlet retrieves an Azure REST API authorization code by displaying a pop up browser window where you log in. 
        .PARAMETER ClientId
            The client/application ID that identifies this application.
        .PARAMETER TenantId
            The tenant ID that identifies the organization. Can be 'common' or a specific tenant ID. If version 2.0 of the API is used, 'consumer' and 'organization' can be specified as well. Consumer specifies that only Microsoft Accounts can authenticate. Organization allows only Organizational Accounts to log in.
        .PARAMETER RedirectUri
            The URI to where you should be redirected after authenticating. Native apps should use 'urn:ietf:wg:oauth:2.0:oob' as their Redirect URI in version 1.0. In version 2.0 'https://login.microsoftonline.com/common/oauth2/nativeclient' should be specified for native apps, however 'urn:ietf:wg:oauth:2.0:oob' also works.
        .PARAMETER Scope
            An array of the permissions you require from this application. Required when using the v2.0 API.
            In version 1.0 specify the scopes as 'calendars.read' or 'user.readwrite'.
            When using version 2.0, specify the scopes in the format 'http://graph.microsoft.com/user.readbasic.all' and 'https://outlook.office.com/mail.read'.
        .PARAMETER Prompt
            Specifies what type of login is needed.None specifies single sign-on. Login specifies that credentials must be entered and SSO is negated. Consent specifies that the user must give consent. Not available with the v2.0 authentication API, Admin_Consent specifies that an admin automatically approves the application for all users.
        .PARAMETER ApiV2
            Enables the use of version 2.0 of the authentication API. Version 2.0 apps can be registered at https://apps.dev.microsoft.com/.
        .EXAMPLE
            Get-OAuth2AzureAuthorization -ClientId $appId -TenantId contoso.com
        
            Code         : O2tTBPNzSgjnjaZWCoBial92z4c6QpoOzM-M8qy16_IGif6NQz-TGF_Z3AenDL1fffUB5JyBHpB0mKylnDIdikaibRIuiWfUdH...
            SessionState : fed8744b-c5cf-4935-b836-142756485e48
            State        : 031d3567-25c3-123f-a4d4-8a7e7fb2343e

            Opens a browser window to login.microsoftonline.com and retrieve an authorization code using version 1.0 of the API.
        .EXAMPLE
            Get-OAuth2AzureAuthorization -ClientId $apiv2ClientId -Scope $Scope -Prompt Consent -RedirectUri 'https://login.microsoftonline.com/common/oauth2/nativeclient' -ApiV2
        
            Code         : GYkA6Ses3jm62gaJTFrt0tlrPBMMPWBM_BXG2hciutILnTAMGOReRfZZ3OXBNqLDl5tD24dTeMosol9eIVlTXXfAkGekWWgkci...
            SessionState : 9c4b9ec2-3dd9-4762-939a-e0bf877a4ac4
            State        : a52c08af-8b94-434e-878e-793f4e66a62b

            Opens a browser window to login.microsoftonline.com and retrieve a version 2.0 authorization code. The aurthorization code grants only the access specified in the scope.
        .INPUTS
        	This command does not accept pipeline input.
        .OUTPUTS
        	This command outputs the returned authorization code, state and session state.
        .LINK
        	Get-OAuth2AzureToken
        .COMPONENT
            OAuth2OpenWindow 
    #>
    [CmdletBinding(
        SupportsShouldProcess = $True,
        ConfirmImpact= 'High'
    )]
    param(
        [Parameter(Mandatory = $True)]
        [string]$EventId,

        [Parameter()]
        [mailaddress]$UserPrincipalName,

        [Parameter(Mandatory = $True)]
        [string]$AccessToken,

        [Parameter()]
        [switch]$Force
    )

    begin {
        $contentType = "application/json"
        $headers = @{
            Authorization = "Bearer $AccessToken"
            Accept = 'application/json'
            OData = 'verbose'
        }
    }
    
    process {
        switch ($UserPrincipalName) {
            { $UserPrincipalName }      { $uri = "$Script:baseUri/users/$($UserPrincipalName.Address)/events/$EventId" }
            { -not $UserPrincipalName } { $uri = "$Script:baseUri/me/events/$EventId" }
        }
        if ($Force -or $PSCmdlet.ShouldProcess($EventId)) { 
            Invoke-RestMethod -Uri $uri -Method Delete -ContentType $contentType -Headers $headers
        }
    }
    
    end {
    }
}


function Set-O365RestV2CalendarItem {
    <#
        .SYNOPSIS
            Retrieves an Azure authorization code.
        .DESCRIPTION
            This cmdlet retrieves an Azure REST API authorization code by displaying a pop up browser window where you log in. 
        .PARAMETER ClientId
            The client/application ID that identifies this application.
        .PARAMETER TenantId
            The tenant ID that identifies the organization. Can be 'common' or a specific tenant ID. If version 2.0 of the API is used, 'consumer' and 'organization' can be specified as well. Consumer specifies that only Microsoft Accounts can authenticate. Organization allows only Organizational Accounts to log in.
        .PARAMETER RedirectUri
            The URI to where you should be redirected after authenticating. Native apps should use 'urn:ietf:wg:oauth:2.0:oob' as their Redirect URI in version 1.0. In version 2.0 'https://login.microsoftonline.com/common/oauth2/nativeclient' should be specified for native apps, however 'urn:ietf:wg:oauth:2.0:oob' also works.
        .PARAMETER Scope
            An array of the permissions you require from this application. Required when using the v2.0 API.
            In version 1.0 specify the scopes as 'calendars.read' or 'user.readwrite'.
            When using version 2.0, specify the scopes in the format 'http://graph.microsoft.com/user.readbasic.all' and 'https://outlook.office.com/mail.read'.
        .PARAMETER Prompt
            Specifies what type of login is needed.None specifies single sign-on. Login specifies that credentials must be entered and SSO is negated. Consent specifies that the user must give consent. Not available with the v2.0 authentication API, Admin_Consent specifies that an admin automatically approves the application for all users.
        .PARAMETER ApiV2
            Enables the use of version 2.0 of the authentication API. Version 2.0 apps can be registered at https://apps.dev.microsoft.com/.
        .EXAMPLE
            Get-OAuth2AzureAuthorization -ClientId $appId -TenantId contoso.com
        
            Code         : O2tTBPNzSgjnjaZWCoBial92z4c6QpoOzM-M8qy16_IGif6NQz-TGF_Z3AenDL1fffUB5JyBHpB0mKylnDIdikaibRIuiWfUdH...
            SessionState : fed8744b-c5cf-4935-b836-142756485e48
            State        : 031d3567-25c3-123f-a4d4-8a7e7fb2343e

            Opens a browser window to login.microsoftonline.com and retrieve an authorization code using version 1.0 of the API.
        .EXAMPLE
            Get-OAuth2AzureAuthorization -ClientId $apiv2ClientId -Scope $Scope -Prompt Consent -RedirectUri 'https://login.microsoftonline.com/common/oauth2/nativeclient' -ApiV2
        
            Code         : GYkA6Ses3jm62gaJTFrt0tlrPBMMPWBM_BXG2hciutILnTAMGOReRfZZ3OXBNqLDl5tD24dTeMosol9eIVlTXXfAkGekWWgkci...
            SessionState : 9c4b9ec2-3dd9-4762-939a-e0bf877a4ac4
            State        : a52c08af-8b94-434e-878e-793f4e66a62b

            Opens a browser window to login.microsoftonline.com and retrieve a version 2.0 authorization code. The aurthorization code grants only the access specified in the scope.
        .INPUTS
        	This command does not accept pipeline input.
        .OUTPUTS
        	This command outputs the returned authorization code, state and session state.
        .LINK
        	Get-OAuth2AzureToken
        .COMPONENT
            OAuth2OpenWindow 
    #>
    [CmdletBinding(
        SupportsShouldProcess = $True,
        ConfirmImpact= 'Medium'
    )]
    param(
        [Parameter(Mandatory = $True)]
        [string]$EventId,

        [Parameter()]
        [string]$Subject,
        
        [Parameter()]
        [string]$Note,
        
        [Parameter()]
        [switch]$AsHtml,
        
        [Parameter()]
        $Attendees,

        [Parameter()]
        $Location,

        [Parameter()]
        [datetime]$StartDate,
        
        [Parameter()]
        [ValidateScript({ [System.TimeZoneInfo]::FindSystemTimeZoneById($_) })]
        [string]$StartTimeZone = [System.Timezone]::CurrentTimeZone.StandardName,
        
        [Parameter()]
        [datetime]$EndDate,
        
        [Parameter()]
        [ValidateScript({ [System.TimeZoneInfo]::FindSystemTimeZoneById($_) })]
        [string]$EndTimeZone = [System.Timezone]::CurrentTimeZone.StandardName,
        
        [Parameter()]
        [int]$Reminder,

        [Parameter()]
        [switch]$AllDay,
        
        [Parameter()]
        [ValidateSet('Free','WorkingElsewhere','Tentative','Busy','Away')]
        [string]$ShowAs,
        
        [Parameter()]
        [mailaddress]$UserPrincipalName,

        [Parameter(Mandatory = $True)]
        [string]$AccessToken,

        [Parameter()]
        [switch]$Force
    )

    begin {
        $contentType = "application/json"
        $headers = @{
            Authorization = "Bearer $AccessToken"
            Accept = 'application/json'
            OData = 'verbose'
        }
        $TimeFormat = "yyyy-MM-ddTHH:mm:ss"
    }
    
    process {
        switch ($UserPrincipalName) {
            { $UserPrincipalName }      { $uri = "$Script:baseUri/users/$($UserPrincipalName.Address)/events/$EventId" }
            { -not $UserPrincipalName } { $uri = "$Script:baseUri/me/events/$EventId" }
        }
        switch ($AsHTML) {
            $False { $NoteContentType = 'Text' }
            $True  { $NoteContentType = 'HTML' }
        }
        switch ($AllDay.IsPresent) {
            { $StartDate -and ($_ -eq $False) } { $start = Get-Date $StartDate -Format $TimeFormat }
            { $EndDate -and ($_ -eq $False) }   { $end = Get-Date $EndDate -Format $TimeFormat }
            { $StartDate -and ($_ -eq $True) }  { $start = Get-Date $StartDate.Date -Format $TimeFormat }
            { $EndDate -and ($_ -eq $True) }    { $end = Get-Date $EndDate.Date -Format $TimeFormat }
        }
        $attendeesProperties = foreach ($Attendee in $Attendees) {
            @{
                EmailAddress = @{
                    Address = $Attendee.EmailAddress
                    Name = $Attendee.Name
                }
                Type = $Attendee.Type
            }
        }
        $body = @{}
        switch ($body) {
            { $Subject }        { $body.Subject = $Subject }
            { $Note }    { 
                                  $body.Body = @{}
                                  $body.Body.Content = $Note
                                  $body.Body.ContentType = $NoteContentType
                                }
            { $start }          { 
                                  $body.Start = @{}
                                  $body.Start.DateTime = $start
                                  $body.Start.TimeZone = $StartTimeZone 
                                }
            { $end }            { 
                                  $body.End = @{}
                                  $body.End.DateTime = $end
                                  $body.End.TimeZone = $EndTimeZone
                                }
            { $Reminder } { 
                                  $body.ReminderMinutesBeforeStart = $Reminder
                                  $body.IsReminderOn = ($Reminder -ne 0)
                                }
            { $Attendees }      { $body.Attendees = $attendeesProperties }
            { $Location }       { 
                                  $body.Location = @{}
                                  $body.Location.DisplayName = $Location
                                  $body.Location.Address = $null
                                  $body.Location.Coordinates = $null
                                }
            { $ShowAs }         { $body.ShowAs = $ShowAs }
            { $AllDay }         { $body.IsAllDay = $AllDay.IsPresent }
        }
        if ($Force -or $PSCmdlet.ShouldProcess($EventId)) { 
            Invoke-RestMethod -Uri $uri -Method Patch -ContentType $contentType -Headers $headers -Body (ConvertTo-Json $body -Depth 10)
        }
    }
    
    end {
    }
}