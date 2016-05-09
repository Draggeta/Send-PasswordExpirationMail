function Connect-Office365 {

    <# 
    .SYNOPSIS 
        Log in to Office 365 services. 
    .DESCRIPTION 
        Log in to your selected Office 365 services.
    .PARAMETER Credential
        Credentials to log in to Office 365 with.
    .PARAMETER MsOnline
        Logs in to Microsoft Azure AD.
    .PARAMETER ComplianceCenter
        Logs in to Microsoft Security and Compliance Center.
    .PARAMETER ExchangeOnline
        Logs in to Microsoft Exchange Online.
    .PARAMETER SkypeForBusinessOnline
        Logs in to Microsoft Skype for Business Online.
    .PARAMETER SharePointOnline
        Logs in to Microsoft SharePoint Online. Requires the tenant name to be specified.
    .PARAMETER TenantName
        Tenant name. Required to connect to SharePoint Online.
    .EXAMPLE 
        Login-Office365 -Credential $Credential -MsOnline -ExchangeOnline
        Description 
         
        ----------- 
     
        Logs in to Azure AD and Exchange Online.
    .INPUTS 
    	None. You cannot pipe objects to Login-Office365 
    .OUTPUTS 
    	None.
    .NOTES 
        Author:   Tony Fortes Ramos 
        Created:  May 02, 2016
        Created:  May 02, 2016 
    .LINK 
    	Logoff-Office365
        Get-PSSession
        New-PSSession
        Import-PSSession 
    #>

    [CmdletBinding(DefaultParameterSetName = 'None')]
    param(

        [Parameter(Mandatory = $true)]
        [PsCredential]$Credential = (Get-Credential),
        [Parameter()]
        [Switch]$MsOnline,
        [Parameter()]
        [Switch]$ComplianceCenter,
        [Parameter()]
        [Switch]$ExchangeOnline,
        [Parameter(Mandatory = $false, ParameterSetName = 'SharePointOnline')]
        [Switch]$SharePointOnline,
        [Parameter()]
        [Switch]$SkypeForBusinessOnline,
        [Parameter(Mandatory = $true, ParameterSetName = 'SharePointOnline')]
        [String]$TenantName
        
    )

    if ($MsOnline) {

     Connect-MsolService -Credential $credential   
     #'https://msdn.microsoft.com/en-us/library/jj151815.aspx'

    }
    if ($ComplianceCenter) {

        $testSession = Get-PSSession
        if ($testSession.ComputerName -like "*.compliance.protection.outlook.com") {
            
            Write-Verbose 'There is already a connection with Microsoft Compliance Center.'

        }        

        $CcSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
        Import-PSSession $CcSession -Prefix CC

    }
    if ($ExchangeOnline) {

        $testSession = Get-PSSession
        if ($testSession.ComputerName -contains 'outlook.office365.com') {
            
            Write-Verbose 'There is already a connection with Microsoft Exchange Online.'

        }
        else {

            #Login to Exchange Online
            $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
            
            #Import the sessions
            Import-PSSession $exchangeSession -DisableNameChecking

        }

    }
    if ($SharePointOnline) {
        
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
        Connect-SPOService -Url "https://$TenantName-admin.sharepoint.com" -Credential $Credential
        
        #'https://www.microsoft.com/en-us/download/details.aspx?id=35588'

    }
    if ($SkypeForBusinessOnline) {
        
        $testSession = Get-PSSession
        if ($testSession.ComputerName -like "*.online.lync.com") {
            
            Write-Verbose 'There is already a connection with Microsoft Skype for Business Online.'

        }
        else {

            Import-Module SkypeOnlineConnector
            $SfboSession = New-CsOnlineSession -Credential $Credential
            Import-PSSession $sfboSession

            #'https://www.microsoft.com/en-us/download/details.aspx?id=39366'

        }

    }

}
function Disconnect-Office365 {
    
    <# 
    .SYNOPSIS 
        Log out from Office 365 services. 
    .DESCRIPTION 
        Log out from selected Office 365 services.
    .PARAMETER ComplianceCenter
        Logs out from Microsoft Security and Compliance Center.
    .PARAMETER ExchangeOnline
        Logs out from Microsoft Exchange Online.
    .PARAMETER SkypeForBusinessOnline
        Logs out from Microsoft Skype for Business Online.
    .PARAMETER SharePointOnline
        Logs out from Microsoft SharePoint Online.
    .PARAMETER All
        Logs out from all Microsoft Office 365 services.
    .EXAMPLE 
        Logoff-Office365 -ComplianceCenter -SkypeForBusinessOnline
        Description 
         
        ----------- 
     
        Logs off from Security and Compliance Center and Skype for Business Online.
    .INPUTS 
    	None. You cannot pipe objects to Logoff-Office365 
    .OUTPUTS 
    	None.
    .NOTES 
        Author:   Tony Fortes Ramos 
        Created:  May 02, 2016
        Created:  May 02, 2016 
    .LINK 
    	Login-Office365
        Get-PSSession
        Remove-PSSession 
    #>
    
    [CmdletBinding(DefaultParameterSetName = 'Specific')]
    Param (

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$ComplianceCenter,
        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$ExchangeOnline,
        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$SharePointOnline,
        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$SkypeForBusinessOnline,
        [Parameter(ParameterSetName = 'All')]
        [Switch]$All

    )

    if ($ComplianceCenter) {

        (Get-PSSession).where{ $_.ComputerName -like "*.compliance.protection.outlook.com" } | Remove-PSSession
        Write-Verbose 'Removed sessions to Microsoft Compliance Center.'

    }
    if ($ExchangeOnline) {

        (Get-PSSession).where{ $_.ComputerName -eq 'outlook.office365.com' } | Remove-PSSession
        Write-Verbose 'Removed sessions to Microsoft Exchange Online.'

    }
    if ($SharePointOnline) {

        Disconnect-SPOService
        Write-Verbose 'Removed sessions to Microsoft Sharepoint Online.'

    }
    if ($SkypeForBusinessOnline) {

        (Get-PSSession).where{ $_.ComputerName -like "*.online.lync.com" } | Remove-PSSession
        Write-Verbose 'Removed sessions to Microsoft Skype for Business Online.'

    }
    if ($All) {

        (Get-PSSession).where{ $_.ComputerName -like "*.compliance.protection.outlook.com" } | Remove-PSSession
        Write-Verbose 'Removed sessions to Microsoft Compliance Center.'

        (Get-PSSession).where{ $_.ComputerName -eq 'outlook.office365.com' } | Remove-PSSession
        Write-Verbose 'Removed sessions to Microsoft Exchange Online.'

        Disconnect-SPOService -ErrorAction SilentlyContinue
        Write-Verbose 'Removed sessions to Microsoft Sharepoint Online.'

        (Get-PSSession).where{ $_.ComputerName -like "*.online.lync.com" } | Remove-PSSession
        Write-Verbose 'Removed sessions to Microsoft Skype for Business Online.'

    }

}