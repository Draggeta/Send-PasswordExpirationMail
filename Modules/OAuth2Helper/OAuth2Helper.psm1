$Script:authenticationUrl = 'https://login.microsoftonline.com'


function OAuth2OpenWindow {
    <#
        .SYNOPSIS
            Opens a browser window to allow authentication.
        .DESCRIPTION
            Opens the specified URL in a browser window to facilitate OAuth2 authentication. Can be used for more than just Microsoft APIs.
        .PARAMETER Url
            The Url that should be opened.
        .EXAMPLE
            OAuth2OpenWindow -Url $requestUrl
            
            -----------
        
            Opens a browser window with the specified page.
        .INPUTS
        	This command does not accept pipeline input.
        .OUTPUTS
        	This command outputs the returned data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Url
    )

    begin {
    }

    process {
        #Load the System.Windows.Forms assembly required to open the WebPage to authenticate
        Add-Type -AssemblyName System.Windows.Forms

        #Create a form and a webbrowser object to display the authentication page.
        $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
        $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($Url -f ($Scope -join "%20")) }

        #Specify what the browser should do once a message matching the regex specified below has been returned.
        $docComp  = {
            $uri = $web.Url.AbsoluteUri        
            if ($uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
        }
        #Disable dialog boxes such as script error messages
        $web.ScriptErrorsSuppressed = $true
        #Add the DocumentsCompleted action to the webbrowser. This way it knows to close itself when a response has been received
        $web.Add_DocumentCompleted($docComp)
        $form.Controls.Add($web)
        #Use the method to display the form on the foreground when visible.
        $form.Add_Shown({$form.Activate()})
        #Show the form. Will pop up to the foreground due to the activate method.
        $form.ShowDialog() | Out-Null
    }

    end {
        return $web
    }
}


function Get-OAuth2AzureAuthorization {
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
    [CmdletBinding(DefaultParameterSetName = 'ApiV1')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter()]
        [string]$TenantId = 'common',

        [Parameter()]
        [string]$RedirectUri = 'urn:ietf:wg:oauth:2.0:oob',

        [Parameter(ParameterSetName = 'ApiV2', Mandatory = $true)]
        [Parameter(ParameterSetName = 'ApiV1')]
        [string[]]$Scope,

        [Parameter()]
        [ValidateSet('Login','Consent','Admin_Consent','None')]
        [string]$Prompt = 'Login',

        [Parameter(ParameterSetName = 'ApiV2', Mandatory = $true)]
        [switch]$ApiV2
    )

    begin {
    }

    process {
        #Load the System.Web assembly required to encode the values that will be entered into the URL.
        Add-Type -AssemblyName System.Web
        
        #UrlEncode the redirect URI, resource and scope for special characters 
        $state = New-Guid
        $scopeEncoded = [System.Web.HttpUtility]::UrlEncode($Scope)
        switch ($RedirectUri) {
            'urn:ietf:wg:oauth:2.0:oob' { $redirectUriEncoded = 'urn:ietf:wg:oauth:2.0:oob' }
            Default                     { $redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($RedirectUri) }
        }
        switch ($ApiV2.IsPresent) {
            $false { $url = "$Script:authenticationUrl/$TenantId/oauth2/authorize?response_type=code&client_id=$ClientId&redirect_uri=$redirectUriEncoded&state=$state&prompt=$($Prompt.ToLower())" }
            $true  { $url = "$Script:authenticationUrl/$TenantId/oauth2/v2.0/authorize?response_type=code&client_id=$ClientId&redirect_uri=$redirectUriEncoded&state=$state&prompt=$($Prompt.ToLower())&response_mode=query" }
        }
        if ($Scope) {
            $url += "&scope=$scopeEncoded" 
        }
        #Open a window to the specific url and authenticate with your credentials.
        $query = OAuth2OpenWindow -Url $url
        #Parse the query so the code and session state can be found.
        $output = [System.Web.HttpUtility]::ParseQueryString($query.Url.Query)
        $properties = @{
            Code = $output['code']
            SessionState = $output['session_state']
            State = $output['state']
        }
        $object = New-Object -TypeName PSObject -Property $properties
    }
    
    end {
        if ($object.State -eq $state) {
            return $object
        }
        else {
            Write-Warning "The returned state '$($object.State)' isn't equal to generated state '$state'. Reply cannot be trusted."
            break
        }
    }
}


function Get-OAuth2AzureToken {
    <#
        .SYNOPSIS
            Retrieves an access token.
        .DESCRIPTION
            Uses the authorization code to request an access token for a specific resource. 
        .PARAMETER ClientId
            The client/application ID that identifies this application.
        .PARAMETER ClientSecret
            The secret key used to authenticate your application. In version 2.0 of the API it cannot be used for native apps.
        .PARAMETER TenantId
            The tenant ID that identifies the organization. Can be 'common' or a specific tenant ID. If version 2.0 of the API is used, 'consumer' and 'organization' can be specified as well. Consumer specifies that only Microsoft Accounts can authenticate. Organization allows only Organizational Accounts to log in.
        .PARAMETER RedirectUri
            The URI to where you should be redirected after authenticating. Native apps should use 'urn:ietf:wg:oauth:2.0:oob' as their Redirect URI in version 1.0. In version 2.0 'https://login.microsoftonline.com/common/oauth2/nativeclient' should be specified for native apps, however 'urn:ietf:wg:oauth:2.0:oob' also works.
        .PARAMETER ResourceUri
            The URI of the resource you're trying to access. Only used by version 1.0 of the API. Specify this parameter in the format 'https://outlook.office.com'.
        .PARAMETER Scope
            An array of the permissions you require from this application. Can only be the same or a superset of the scope defined in the authorization request. Mandatory for version 2.0 of the API.
            In version 1.0 specify the scopes as 'calendars.read' or 'user.readwrite'.
            When using version 2.0, specify the scopes in the format 'http://graph.microsoft.com/user.readbasic.all' and 'https://outlook.office.com/mail.read'.
        .PARAMETER AuthorizationCode
            The authorization code necessary to request an access token. Can be attained by using Get-OAuth2AzureAuthorization.
        .PARAMETER ApiV2
            Enables the use of version 2.0 of the authentication API. Version 2.0 apps can be registered at https://apps.dev.microsoft.com/.
        .EXAMPLE
            Get-OAuth2AzureAccessToken -ClientId $appId -ClientSecret $key -RedirectUri https://app.fabrikam.com/endpoint -ResourceUri https://graph.microsoft.com -AuthorizationCode $authCode
            
            RefreshToken : JTMjU2IiwieDV0IjoiSTZvQnc0VnpCSE9xbGVHclYyQUpkQTVFbVhjIiwia2lkIjoiSTZvQnc0VnpCSE9xbGVHclYyQUpkQTVF...
            AccessToken  : eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEUk5ZUlEzZGhSU3JtLTRLLWFkcENKYUVOeFFJUHlXLVRieUVwWllzSG...
            IdToken      : uUmVhZCIsInN1YiI6IjByMEJyTl9GOGxRck5aeHJBS0RFNHhTQzFzbWJIRDRvOXU0dkVHTG9Kb00iLCJ0aWQiOiJlZmQzYzc2Y...

            Returns an access token, refresh token and ID token for use in API requests.
        .EXAMPLE
            Get-OAuth2AzureAccessToken -ClientId $appId -Scope 'http://graph.microsoft.com/user.readbasic.all' -AuthorizationCode $authCode -ApiV2
            
            RefreshToken : 
            AccessToken  : AQABAAAAAADRNYRQ3dhRSrmU7elRk-QCim_YGkzAatTUt0ESeVdvfpNlx6P-ec7mO9HiVu2BJ6Z0RbCagK4tCEPK1g0yNf3i-P...
            IdToken      : 

            Returns an access token for use in API requests. As security cannot be guaranteed on native apps, the 2.0 version does not return refresh and ID tokens.
        .INPUTS
        	This command does not accept pipeline input.
        .OUTPUTS
        	This command outputs the returned access token, refresh token and ID token.
        .LINK
        	Get-OAuth2AzureAuthorization
        .COMPONENT
            OAuth2OpenWindow 
    #>
    [CmdletBinding(DefaultParameterSetName = 'ApiV1')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(ParameterSetName = 'ApiV2')]
        [Parameter(ParameterSetName = 'ApiV1', Mandatory = $true)]
        [string]$ClientSecret,

        [Parameter()]
        [string]$TenantId = 'common',

        [Parameter()]
        [string]$RedirectUri = 'urn:ietf:wg:oauth:2.0:oob',

        [Parameter(ParameterSetName = 'ApiV1', Mandatory = $true)]
        [string]$ResourceUri,

        [Parameter(ParameterSetName = 'ApiV1')]
        [Parameter(ParameterSetName = 'ApiV2', Mandatory = $true)]
        [string[]]$Scope,

        [Parameter(Mandatory = $true)]
        [string]$AuthorizationCode,

        [Parameter(ParameterSetName = 'ApiV2', Mandatory = $true)]
        [switch]$ApiV2
    )

    begin {    
    }

    process {
        #Load the System.Web assembly required to encode the values that will be entered into the URL.
        Add-Type -AssemblyName System.Web
        
        #UrlEncode the ClientSecret for special characters.
        $clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($ClientSecret)
        $scopeEncoded = [System.Web.HttpUtility]::UrlEncode($Scope)
        #prepare the parameters that should be passed to the Invoke-Restmethod cmdlet.
        switch ($ApiV2.IsPresent) {
            $false { $url = "$Script:authenticationUrl/$TenantId/oauth2/token" }
            $true  { $url = "$Script:authenticationUrl/$TenantId/oauth2/v2.0/token" }
        }
        $body = "grant_type=authorization_code&client_id=$ClientId&client_secret=$clientSecretEncoded&redirect_uri=$RedirectUri&code=$AuthorizationCode"
        switch ($body) {
            { $ResourceUri } { $body += "&resource=$ResourceUri" }
            { $Scope }       { $body += "&scope=$scopeEncoded" }
        }
        $invokeRestMethodParams = @{
            Uri = $url
            Method = 'Post'
            ContentType = 'application/x-www-form-urlencoded'
            Body = $body
            ErrorAction = 'Stop'    
        }
        #Perform the request
        $authorization = Invoke-RestMethod @invokeRestMethodParams
        #Get the access token
        $properties = @{
            AccessToken = $authorization.access_token
            RefreshToken = $authorization.refresh_token
            IdToken = $authorization.id_token
        }
        $object = New-Object -TypeName PSObject -Property $properties
    }

    end {
        return $object
    }
}