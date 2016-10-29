$Script:baseUri = "https://outlook.office365.com/api/beta"

function Get-O365RestCalendarItem {
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
        [PSCredential]
        [System.Management.Automation.Credential()]$Credential = (Get-Credential)
    )

    begin {
        $contentType = "application/json"
        $headers = @{
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
            Invoke-RestMethod -Uri $uri -Credential $Credential -Method Get -ContentType $contentType -Headers $headers
        }
    }
    
    end {
    }
}


function New-O365RestAttendee {
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
    Param (
        [Parameter(Position = 0, Mandatory = $True)]
        [MailAddress[]]$EmailAddress,

        [Parameter()]
        [ValidateSet('Required','Optional','Resource')]
        $Type = 'Required'
    )
    BEGIN {
    }
    PROCESS {

        ForEach ($Address in $EmailAddress){
            $Properties = @{
                Name = $Address.Address
                EmailAddress = $Address.Address
                Type = $Type
            }
            If ($PSCmdlet.ShouldProcess($Address.Address, "create new attendee object")) { 
                $Object = New-Object -TypeName PSObject -Property $Properties
                Write-Output $Object
            }
        }

    }
    END {
    }
}


function New-O365RestCalendarItem {
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
        [String]$Subject,
        
        [Parameter()]
        [String]$Note,
        
        [Parameter()]
        [Switch]$AsHtml,
        
        [Parameter()]
        $Attendees,

        [Parameter()]
        $Location,

        [Parameter()]
        [DateTime]$StartDate = (Get-Date),
        
        [Parameter()]
        [ValidateScript({ [System.TimeZoneInfo]::FindSystemTimeZoneById($_) })]
        [String]$StartTimeZone = [System.Timezone]::CurrentTimeZone.StandardName,
        
        [Parameter()]
        [DateTime]$EndDate = (Get-Date).AddMinutes(30),
        
        [Parameter()]
        [ValidateScript({ [System.TimeZoneInfo]::FindSystemTimeZoneById($_) })]
        [String]$EndTimeZone = [System.Timezone]::CurrentTimeZone.StandardName,
        
        [Parameter()]
        [Int]$Reminder,

        [Parameter()]
        [Switch]$AllDay,
        
        [Parameter()]
        [ValidateSet('Free','WorkingElsewhere','Tentative','Busy','Away')]
        [String]$ShowAs = 'Busy',
        
        [Parameter()]
        [MailAddress]$UserPrincipalName,

        [Parameter(Mandatory = $True)]
        [PSCredential]
        [System.Management.Automation.Credential()]$Credential = (Get-Credential)
    )
    begin {

        $ContentType = "application/json"
        $Headers = @{
            Accept = 'application/json';
            OData = 'verbose'
        }
        $TimeFormat = "yyyy-MM-ddTHH:mm:ss"

    }
    process {
        switch ($UserPrincipalName) {
            { $UserPrincipalName }      { $Uri = "$Script:baseUri/users/$($UserPrincipalName.Address)/events"; break }
            { -not $UserPrincipalName } { $Uri = "$Script:baseUri/me/events" }
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
        If ($PSCmdlet.ShouldProcess("$Subject with a start time of $Start", "create an appointment")) { 
            Invoke-RestMethod -Uri $Uri -Credential $Credential -Method Post -ContentType $ContentType -Headers $Headers -Body (ConvertTo-Json $Body -Depth 10)
        }
    }
    end {
    }
}


function Remove-O365RestCalendarItem {
    <#
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
        [PSCredential]
        [System.Management.Automation.Credential()]$Credential = (Get-Credential),

        [Parameter()]
        [switch]$Force
    )

    begin {
        $contentType = "application/json"
        $headers = @{
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
            Invoke-RestMethod -Uri $uri -Credential $Credential -Method Delete -ContentType $contentType -Headers $headers
        }
    }
    
    end {
    }
}


function Set-O365RestCalendarItem {
    <#
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
        [PSCredential]
        [System.Management.Automation.Credential()]$Credential = (Get-Credential),

        [Parameter()]
        [switch]$Force
    )

    begin {
        $contentType = "application/json"
        $headers = @{
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
            Invoke-RestMethod -Uri $uri -Credential $Credential -Method Patch -ContentType $contentType -Headers $headers -Body (ConvertTo-Json $body -Depth 10)
        }
    }
    
    end {
    }
}