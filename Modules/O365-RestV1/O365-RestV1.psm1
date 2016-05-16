Function New-O365RestCalendarItem {
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
            
            $Attendees = New-O365RestAttendee -EmailAddress Mark@domain.com,Sally@contoso.com -Type Required
            $Attendees += New-O365RestAttendee -EmailAddress MeetingRoom1@domain.com -Type Resource

            The following shows how to set the Attendees parameter to these attendees.

            -Attendees $Attendees
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
        .INPUTS
        	None. You cannot pipe objects to New-O365RestCalendarItem.
        .OUTPUTS
        	None. New-O365RestCalendarItem does not output.
        .NOTES
            Author:   Tony Fortes Ramos
            Created:  May 15, 2016
        .LINK
        	New-O365RestAttendee
        .COMPONENT
            New-O365RestAttendee            
    #>
    [CmdletBinding()]
    Param(
        [Parameter()]
        [String]$Subject,
        
        [Parameter()]
        [String]$Note,
        
        [Parameter()]
        [Switch]$AsHtml,
        
        [Parameter()]
        $Attendees,

        [Parameter()]
        $Location,

        [Parameter(Mandatory = $True)]
        $StartDate,
        
        [Parameter()]
        [ValidateScript({ $_ -in ([System.TimeZoneInfo]::GetSystemTimeZones()).id })]
        [String]$StartTimeZone = (Get-WmiObject -Class win32_timezone).StandardName,
        
        [Parameter(Mandatory = $True)]
        $EndDate,
        
        [Parameter()]
        [ValidateScript({ $_ -in ([System.TimeZoneInfo]::GetSystemTimeZones()).id })]
        [String]$EndTimeZone = (Get-WmiObject -Class win32_timezone).StandardName,
        
        [Parameter()]
        [Switch]$AllDay,
        
        [Parameter()]
        [ValidateSet('Free','WorkingElsewhere','Tentative','Busy','Away')]
        $ShowAs,
        
        [Parameter()]
        [MailAddress]$UserPrincipalName,

        [Parameter(Mandatory = $True)]
        [PSObject]$Credential = (Get-Credential)
    )
    BEGIN {
        If (-not $UserPrincipalName) {
            $User = $Credential.UserName
        }
        ElseIf ($UserPrincipalName) {
            $User = $UserPrincipalName.Address
        }
        $Uri = "https://outlook.office365.com/api/v1.0/users('$User')/events"
        $ContentType = "application/json"
        $Headers = @{
            accept = 'application/json';
            odata = 'verbose'
        }
    }
    PROCESS {
        Switch ($AsHTML) {
            $False { $NoteContentType = 'Text' }
            $True { $NoteContentType = 'HTML' }
            Default { $NoteContentType = 'Text' }
        }
        Switch ($AllDay) {
            $False { 
                $StartDate = Get-Date $StartDate -Format o
                $EndDate = Get-Date $EndDate -Format o
            }
            $True {
                $StartDate = Get-Date $StartDate.Date -Format o
                $EndDate = Get-Date $EndDate.Date -Format o
            }
            Default {
                $StartDate = Get-Date $StartDate -Format o
                $EndDate = Get-Date $EndDate -Format o
            }
        }
        $BodyContent = If ($Note) {
            @{
                ContentType = $NoteContentType
                Content = $Note
            }
        }
        $AttendeesProperties = Foreach ($Attendee in $Attendees) {
            @{
                EmailAddress = @{
                    Address = $Attendee.EmailAddress
                    Name = $Attendee.Name
                }
                Type = $Attendee.Type
            }
        }
        $Body = @{
            Subject = $Subject
            Body = $BodyContent
            Start = $StartDate
            StartTimeZone = $StartTimeZone
            End = $EndDate
            EndTimeZone = $EndTimeZone
            Attendees = @(
                $AttendeesProperties
            )
            Location = @{
                DisplayName = $Location
                Address = $null
                Coordinates = $null
            }
            ShowAs = $ShowAs
            IsAllDay = $AllDay.IsPresent
        }
        (ConvertTo-Json $Body -Depth 10)
        Invoke-RestMethod -Uri $Uri -Credential $Credential -Method Post -ContentType $ContentType -Headers $Headers -Body (ConvertTo-Json $Body -Depth 10)
    }
    END {
    }
}

Function New-O365RestAttendee {
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
    [CmdletBinding()]
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
                Name = $Address
                EmailAddress = $Address
                Type = $Type
            }
            $Object = New-Object -TypeName PSObject -Property $Properties
            Write-Output $Object
        }
    }
    END {
    }
}












$body = @{
    Message = @{
        Subject = "Meet for lunch?"
        Body = @{
            ContentType = "Text"
            Content = "The new cafeteria is open."
        }
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = "tony.fortesramos@ogd.nl"
                }
            }
        )
        Attachments = @(
        )
    }
    SaveToSentItems = $false
}

Invoke-RestMethod -Uri "https://outlook.office365.com/api/v1.0/me/sendmail" -Credential $cred `
-Method Post -ContentType $contentType -Headers $headers  `
-Body (ConvertTo-Json $body -Depth 10)

$body= @{ 
    IsRead = $false 
}

Invoke-RestMethod -Uri "https://outlook.office365.com/api/v1.0/Users('ogdadmin@webster.nl')/Messages('AAMkAGJkMGI2OTk5LTI0Y2EtNGJjMS04N2JhLTc1ZjVkM2ViNmZiZABGAAAAAACvp4Afb7WjS6fnU1giKmuJBwCV5FnB3-AASKaaYis7O-b7AAAAqOtSAACzQt-LRvQZQ5c5jf_UxK6iAAB8T_wPAAA=')" -Credential $cred `
-Method Patch -ContentType $contentType `
-Body (ConvertTo-Json $body)

Invoke-RestMethod -Uri "https://outlook.office365.com/api/v1.0/me/Events('AAMkAGJkMGI2OTk5LTI0Y2EtNGJjMS04N2JhLTc1ZjVkM2ViNmZiZABGAAAAAACvp4Afb7WjS6fnU1giKmuJBwCV5FnB3-AASKaaYis7O-b7AAAAqOtbAACzQt-LRvQZQ5c5jf_UxK6iAAB-Vr18AAA=')" -Credential $cred `
-Method Delete -ContentType $contentType

        Start = $(Get-Date -Format o)
        StartTimeZone = "W. Europe Standard Time"
        End = $(Get-Date).AddHours(2).ToString("yyyy-MM-dd\THH:mm:ss\+02:00")

$test = Invoke-RestMethod -Uri "https://outlook.office365.com/api/v1.0/me/calendarview?startDateTime=2016-05-15T01:00:00Z&endDateTime=2016-05-16T23:00:00Z" -Credential $credential
