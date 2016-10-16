$Script:authenticationUrl = 'https://login.microsoftonline.com'


function OAuth2OpenWindow {
    <#
        .SYNOPSIS
            Brief description.
        .DESCRIPTION
            Longer description. 
        .PARAMETER Name1
            Information about a parameter.
        .PARAMETER Name2
            Information about a parameter.
        .EXAMPLE
            Cmdlet
            
            -----------
        
            Explanation about the example.
        .INPUTS
        	Information about inputs and piping
        .OUTPUTS
        	Information abou the outputs.
        .NOTES
            Extra notes.
        .LINK
        	Related Cmdlets.
        .COMPONENT
            Used Cmdlets. 
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        $Url
    )

    begin {
    }

    process {
        #Load the System.Windows.Forms assembly required to open the WebPage to authenticate
        Add-Type -AssemblyName System.Windows.Forms

        $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
        $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($Url -f ($Scope -join "%20")) }

        #Specify what to do when the browser notices that an action has been completed.
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
            Brief description.
        .DESCRIPTION
            Longer description. 
        .PARAMETER Name1
            Information about a parameter.
        .PARAMETER Name2
            Information about a parameter.
        .EXAMPLE
            Cmdlet
            
            -----------
        
            Explanation about the example.
        .INPUTS
        	Information about inputs and piping
        .OUTPUTS
        	Information abou the outputs.
        .NOTES
            Extra notes.
        .LINK
        	Related Cmdlets.
        .COMPONENT
            Used Cmdlets. 
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
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
        
        # UrlEncode the ClientID and ClientSecret and URL's for special characters 
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
        # Get Authorization Code
        $query = OAuth2OpenWindow -Url $url

        $queryOutput = [System.Web.HttpUtility]::ParseQueryString($query.Url.Query)
        $output = @{}
        foreach($key in $queryOutput.Keys){
            $output["$key"] = $queryOutput[$key]
        }
        # Extract Access token from the returned URI
        $authorizationCode = "$($output.code)&session_state=$($output.session_state)"
    }
    
    end {
        return $authorizationCode
    }
}


function Get-OAuth2AzureAccessToken {
    <#
        .SYNOPSIS
            Brief description.
        .DESCRIPTION
            Longer description. 
        .PARAMETER Name1
            Information about a parameter.
        .PARAMETER Name2
            Information about a parameter.
        .EXAMPLE
            Cmdlet
            
            -----------
        
            Explanation about the example.
        .INPUTS
        	Information about inputs and piping
        .OUTPUTS
        	Information abou the outputs.
        .NOTES
            Extra notes.
        .LINK
        	Related Cmdlets.
        .COMPONENT
            Used Cmdlets. 
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ClientId,
        [Parameter()]
        [string]$ClientSecret,
        [Parameter()]
        [string]$TenantId = 'common',
        [Parameter()]
        $RedirectUri = 'urn:ietf:wg:oauth:2.0:oob',
        [Parameter()]
        [string]$AuthorizationCode,
        [Parameter()]
        [string[]]$ResourceUri
    )

    begin {    
    }

    process {
        Add-Type -AssemblyName System.Web
        
        # UrlEncode the ClientID and ClientSecret and URL's for special characters 
        $clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($ClientSecret)

        #get Access Token
        $body = "grant_type=authorization_code&redirect_uri=$RedirectUri&client_id=$ClientId&client_secret=$clientSecretEncoded&code=$AuthorizationCode&resource=$ResourceUri"
        $invokeRestMethodParams = @{
            Uri = "$Script:authenticationUrl/common/oauth2/token"
            Method = 'Post'
            ContentType = 'application/x-www-form-urlencoded'
            Body = $body
            ErrorAction = 'Stop'    
        }
        $authorization = Invoke-RestMethod @invokeRestMethodParams
        $accessToken = $authorization.access_token
    }

    end {
        return $accessToken
    }
}