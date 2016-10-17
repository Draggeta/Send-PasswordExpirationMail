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


function Get-OAuth2AzureAuthorizationCode {
    <#
        .SYNOPSIS
            Retrieves an Azure authorization code.
        .DESCRIPTION
            This cmdlet retrieves an Azure REST API authorization code by displaying a pop up browser window where you log in. 
        .PARAMETER ClientId
            The client/application ID that identifies this application.
        .PARAMETER TenantId
            The tenant ID that identifies your organization. Can be a GUID or one of your verified domain names.
        .PARAMETER RedirectUri
            The URI to where you should be redirected after authenticating. Native Apps should use 'urn:ietf:wg:oauth:2.0:oob' as their Redirect URI.
        .PARAMETER ResourceUri
            The URI of the resource you're trying to access.
        .PARAMETER Scope
            An array of the permissions you require from this application. Necessary for the V2.0 API. 
        .PARAMETER AdminConsent
            Specifies if administrative consent is necessary to run this application.
        .EXAMPLE
            Get-OAuth2AzureAuthorizationCode -ClientId $appId
            
            -----------
        
            Opens a browser window to login.microsoftonline.com so you can log in and retrieve an authorization code.
        .INPUTS
        	This command does not accept pipeline input.
        .OUTPUTS
        	This command outputs the returned authorization code.
        .LINK
        	Get-OAuth2AzureAccessToken
        .COMPONENT
            OAuth2OpenWindow 
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter()]
        [string]$TenantId = 'common',

        [Parameter()]
        $RedirectUri = 'urn:ietf:wg:oauth:2.0:oob',

        [Parameter()]
        [string[]]$ResourceUri,

        [Parameter()]
        [string]$Scope,

        [Parameter()]
        [switch]$AdminConsent
    )

    begin {
    }

    process {
        #Load the System.Web assembly required to encode the values that will be entered into the URL.
        Add-Type -AssemblyName System.Web
        
        #UrlEncode the redirect URI, resource and scope for special characters 
        $redirectUriEncoded =  [System.Web.HttpUtility]::UrlEncode($redirectUri)
        $resourceEncoded = [System.Web.HttpUtility]::UrlEncode($resourceUri)
        $scopeEncoded = [System.Web.HttpUtility]::UrlEncode($Scope)

        $url = "$Script:authenticationUrl/$TenantId/oauth2/authorize?response_type=code&prompt=login"
        switch ($url) {
            { $RedirectUri -eq 'urn:ietf:wg:oauth:2.0:oob' }    { $url += "&redirect_uri=$RedirectUri" }
            { $RedirectUri -ne 'urn:ietf:wg:oauth:2.0:oob' }    { $url += "&redirect_uri=$redirectUriEncoded" }
            { $ClientId }       { $url += "&client_id=$ClientId" }
            { $ResourceUri }    { $url += "&resource=$resourceEncoded" }
            { $AdminConsent }   { $url += "&prompt=admin_consent" }
            { $Scope }          { $url += "&scope=$scopeEncoded" }
        }
        #Open a window to the specific url and authenticate with your credentials.
        $query = OAuth2OpenWindow -Url $url
        #Parse the query so the code and session state can be found.
        $queryOutput = [System.Web.HttpUtility]::ParseQueryString($query.Url.Query)
        $output = @{}
        foreach($key in $queryOutput.Keys){
            $output["$key"] = $queryOutput[$key]
        }
        #Combine the code and session state to form the authorization code.
        $authorizationCode = "$($output.code)&session_state=$($output.session_state)"
    }
    
    end {
        return $authorizationCode
    }
}


function Get-OAuth2AzureAccessToken {
    <#
        .SYNOPSIS
            Retrieves an access token.
        .DESCRIPTION
            Uses the authorization code to request an access token for a specific resource. 
        .PARAMETER ClientId
            The client/application ID that identifies this application.
        .PARAMETER ClientSecret
            The secret key used to authenticate your application.
        .PARAMETER TenantId
            The tenant ID that identifies your organization. Can be a GUID or one of your verified domain names.
        .PARAMETER RedirectUri
            The URI to where you should be redirected after authenticating. Native Apps should use 'urn:ietf:wg:oauth:2.0:oob' as their Redirect URI.
        .PARAMETER ResourceUri
            The URI of the resource you're trying to access.
        .PARAMETER AuthorizationCode
            The authorization code necessary to request an access token.
        .EXAMPLE
            Get-OAuth2AzureAccessToken -ClientId $appId -ClientSecret $key -ResourceUri https://graph.microsoft.com -AuthorizationCode $authCode
            
            -----------
        
            Returns an access token for use in API requests.
        .INPUTS
        	This command does not accept pipeline input.
        .OUTPUTS
        	This command outputs the returned access token.
        .LINK
        	Get-OAuth2AzureAuthorizationCode
        .COMPONENT
            OAuth2OpenWindow 
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,

        [Parameter()]
        [string]$TenantId = 'common',

        [Parameter()]
        $RedirectUri = 'urn:ietf:wg:oauth:2.0:oob',

        [Parameter(Mandatory = $true)]
        [string[]]$ResourceUri,

        [Parameter(Mandatory = $true)]
        [string]$AuthorizationCode
    )

    begin {    
    }

    process {
        #Load the System.Web assembly required to encode the values that will be entered into the URL.
        Add-Type -AssemblyName System.Web
        
        #UrlEncode the ClientSecret for special characters.
        $clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($ClientSecret)
        #prepare the parameters that should be passed to the Invoke-Restmethod cmdlet.
        $body = "grant_type=authorization_code&redirect_uri=$RedirectUri&client_id=$ClientId&client_secret=$clientSecretEncoded&code=$AuthorizationCode&resource=$ResourceUri"
        $invokeRestMethodParams = @{
            Uri = "$Script:authenticationUrl/common/oauth2/token"
            Method = 'Post'
            ContentType = 'application/x-www-form-urlencoded'
            Body = $body
            ErrorAction = 'Stop'    
        }
        #Perform the request
        $authorization = Invoke-RestMethod @invokeRestMethodParams
        #Get the access token
        $accessToken = $authorization.access_token
    }

    end {
        return $accessToken
    }
}