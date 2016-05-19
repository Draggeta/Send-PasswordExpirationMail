Function Connect-Office365Session {
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
        Tenant name without the '.onmicrosoft.com' part. Required to connect to SharePoint Online and Skype for Business Online.
    .EXAMPLE 
        Connect-Office365Session -Credential $Credential -MsOnline -ExchangeOnline
        Description 
         
        ----------- 
     
        Logs in to Azure AD and Exchange Online.
    .INPUTS 
    	None. You cannot pipe objects to Connect-Office365Session
    .OUTPUTS 
    	None.
    .NOTES 
        Author:   Tony Fortes Ramos 
        Created:  May 02, 2016
    .LINK 
    	Disconnect-Office365Session
        Get-PSSession
        New-PSSession
        Import-PSSession 
    #>
    [CmdletBinding(DefaultParameterSetName = 'None')]
    Param(
        [Parameter(Mandatory = $true)]
        [PsCredential]$Credential = (Get-Credential),

        [Parameter()]
        [Switch]$AzureAD,
        
        [Parameter()]
        [Switch]$Azure,
        
        [Parameter()]
        [Switch]$AzureRM,
        
        [Parameter()]
        [Switch]$AzureRMS,

        [Parameter()]
        [Switch]$ComplianceCenter,

        [Parameter()]
        [Switch]$ExchangeOnline,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'SkypeSharePoint')]
        [Switch]$SharePointDevPNP,

        [Parameter(Mandatory = $false, ParameterSetName = 'SkypeSharePoint')]
        [Switch]$SharePointOnline,

        [Parameter(Mandatory = $false, ParameterSetName = 'SkypeSharePoint')]
        [Switch]$SkypeForBusinessOnline,

        [Parameter(Mandatory = $true, ParameterSetName = 'SkypeSharePoint')]
        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [String]$TenantName,

        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [Switch]$All
    )
    BEGIN {
        Switch ($All) {
            $True { 
                $MSOnline = $true
                $Azure = $true
                $AzureRM = $true
                $AzureRMS = $true
                $ComplianceCenter = $true
                $ExchangeOnline = $true
                $SharePointDevPNP = $true
                $SharePointOnline = $true
                $SkypeForBusinessOnline = $true
            }
        }
        $testModule = Get-Module -ListAvailable
    }
    PROCESS {
        If ($AzureAD) {
            $Name = 'Azure Active Directory'
            $Module = 'MSOnline'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Verbose "The $Name module is not installed."
            }
            Else {
                Import-Module $Module -DisableNameChecking
                Try {
                    Connect-MsolService -Credential $Credential
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }
            }
        }
        If ($Azure) {
            $Name = 'Azure Classic'
            $Module = 'Azure'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Verbose "The $Name module is not installed."
            }
            Else {
                Import-Module $Module -DisableNameChecking
                Try {
                    Add-AzureAccount -Credential $Credential
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }
            }
        }
        If ($AzureRM) {
            $Name = 'Azure Resource Manager'
            $Module = 'AzureRM'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Verbose "The $Name module is not installed."
            }
            Else {
                Import-Module $Module -DisableNameChecking
                Try {
                    Login-AzureRmAccount -Credential $Credential
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }
            }
        }
        If ($AzureRMS) {
            $Name = 'Azure Rights Management Service'
            $Module = 'AADRM'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Verbose "The $Name module is not installed."
            }
            Else {
                Import-Module $Module -DisableNameChecking
                Try {
                    Connect-AadrmService -Credential $Credential
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }
            }
        }
        If ($ComplianceCenter) {
            $Name = 'Security and Compliance Center'
            $Module = 'ComplianceCenter'
            $ConnectionUri = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
            If (Get-PSSession -Name $Module -ErrorAction SilentlyContinue) {
                Write-Verbose "There is already a connection with $Name. Skipping the creation of a new session."
            } 
            Else {
                Try {
                    $CcSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Name $Module -Credential $Credential -Authentication Basic -AllowRedirection
                    Import-PSSession $CcSession -DisableNameChecking -Prefix CC
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }        
            }
        }
        If ($ExchangeOnline) {
            $Name = 'Exchange Online'
            $Module = 'ExchangeOnline'
            $ConnectionUri = 'https://outlook.office365.com/powershell-liveid/'
            If (Get-PSSession -Name $Module -ErrorAction SilentlyContinue) {
                Write-Verbose "There is already a connection with $Name. Skipping the creation of a new session."
            } 
            Else {
                Try {
                    $EoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Name $Module -Credential $Credential -Authentication Basic -AllowRedirection
                    Import-PSSession $EoSession -DisableNameChecking -Prefix CC
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }        
            }
        }
        If ($SharePointDevPNP) {
            $Name = 'SharePoint Office Dev PNP'
            $Module = 'OfficeDevPnP.PowerShell.V16.Commands'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Verbose "The $Name module is not installed."
            }
            Else {
                Import-Module $Module -DisableNameChecking
                Try {
                    Connect-SPOnline -Url "https://$TenantName-admin.sharepoint.com" -Credential $Credential
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }
            }
        }
        If ($SharePointOnline) {
            $Name = 'SharePoint Online'
            $Module = 'Microsoft.Online.SharePoint.PowerShell'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Verbose "The $Name module is not installed."
            }
            Else {
                Import-Module $Module -DisableNameChecking
                Try {
                    Connect-SPOnline -Url "https://$TenantName-admin.sharepoint.com" -Credential $Credential
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }
            }
        }
        If ($SkypeForBusinessOnline) {
            $Name = 'Skype for Business Online'
            $Module = 'SkypeOnlineConnector'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Verbose "The $Name module is not installed."
            }
            ElseIf (Get-PSSession -Name $Module -ErrorAction SilentlyContinue) {
                Write-Verbose "There is already a connection with $Name. Skipping the creation of a new session."
            }
            Else {
                Import-Module $Module -DisableNameChecking
                Try {
                    $SfboSession = New-CsOnlineSession -Credential $Credential -OverrideAdminDomain "$TenantName.onmicrosoft.com"
                    $SfboSession.Name = $Module
                    Import-PSSession $SfboSession
                }
                Catch {
                    Write-Verbose "There was an error connecting to $Name."
                }
            }
        }
    }
    END {
    }
}

Function Disconnect-Office365Session {
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
        Disconnect-Office365Session -ComplianceCenter -SkypeForBusinessOnline
        Description 
         
        ----------- 
     
        Logs off from Security and Compliance Center and Skype for Business Online.
    .INPUTS 
    	None. You cannot pipe objects to Disconnect-Office365Session
    .OUTPUTS 
    	None.
    .NOTES 
        Author:   Tony Fortes Ramos 
        Created:  May 02, 2016
    .LINK 
    	Connect-Office365Session
        Get-PSSession
        Remove-PSSession 
    #>
    [CmdletBinding(DefaultParameterSetName = 'Specific')]
    Param (
        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$Azure,
        
        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$AzureRMS,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$ComplianceCenter,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$ExchangeOnline,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$SharePointDevPNP,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$SharePointOnline,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$SkypeForBusinessOnline,

        [Parameter(ParameterSetName = 'All')]
        [Switch]$All
    )
    BEGIN {
        Switch ($All) {
            $True {
                $Azure = $true
                $AzureRMS = $true
                $ComplianceCenter = $true
                $ExchangeOnline = $true
                $SharePointDevPNP = $true
                $SharePointOnline = $true
                $SkypeForBusinessOnline = $true
            }
        }
    }
    PROCESS {
        If ($Azure) {
            $Name = 'Azure Classic'
            If (Get-AzureAccount) {
                Try {
                    Remove-AzureAccount -Name (Get-AzureAccount).Id -Confirm $false
                }
                Catch {
                    Write-Verbose "Could not remove sessions to $Name."
                }
            }
        }
        If ($AzureRMS) {
            $Name = 'Azure Rights Management Service'
            Try {
                Disconnect-AadrmService
            }
            Catch {
                Write-Verbose "Could not remove sessions to $Name."
            }
        }
        If ($ComplianceCenter) {
            $Name = 'Security and Compliance Center'
            $Module = 'ComplianceCenter'
            Try {
                Get-PSSession -Name $Module -ErrorAction SilentlyContinue | Remove-PSSession
            }
            Catch {
                Write-Verbose "Could not remove sessions to $Name."
            }
        } 
        If ($ExchangeOnline) {
            $Name = 'Exchange Online'
            $Module = 'ExchangeOnline'
            Try {
                Get-PSSession -Name $Module -ErrorAction SilentlyContinue | Remove-PSSession
            }
            Catch {
                Write-Verbose "Could not remove sessions to $Name."
            }
        } 
        If ($SharePointDevPNP) {
            $Name = 'SharePoint Office Dev PNP'
            Try {
                Disconnect-SPOnline
            }
            Catch {
                Write-Verbose "Could not remove sessions to $Name."
            }
        }
        If ($SharePointOnline) {
            $Name = 'SharePoint Online'
            $Module = 'Microsoft.Online.SharePoint.PowerShell'
            Try {
                Disconnect-SPOService
            }
            Catch {
                Write-Verbose "Could not remove sessions to $Name."
            }
        }
        If ($SkypeForBusinessOnline) {
            $Name = 'Skype for Business Online'
            $Module = 'SkypeOnlineConnector'
            Try {
                Get-PSSession -Name $Module -ErrorAction SilentlyContinue | Remove-PSSession
            }
            Catch {
                Write-Verbose "Could not remove sessions to $Name."
            }
        }
    }
    END {
    }
}

Function Install-Office365Prerequisites {
    [CmdletBinding()]
    Param (
        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$MSOnline,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$Azure,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$AzureRM,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$AzureRMS,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$OfficeDevPNPSharePointOnline,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$SharePointOnline,

        [Parameter(ParameterSetName = 'Specific')]
        [Switch]$SkypeForBusinessOnline,

        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [Switch]$All,

        [Parameter()]
        [Switch]$Update,
        
        [Parameter()]
        [Switch]$Force
    )
    BEGIN {
        Switch ($All) {
            $True { 
                $MSOnline = $true
                $Azure = $true
                $AzureRM = $true
                $AzureRMS = $true
                $OfficeDevPNPSharePointOnline = $true
                $SharePointOnline = $true
                $SkypeForBusinessOnline = $true
            }
        }
        $Modules = Get-Module -ListAvailable
    }
    PROCESS {
        If ($MSOnline) {
            $MSOnlineAssistantUri = 'https://download.microsoft.com/download/5/0/1/5017D39B-8E29-48C8-91A8-8D0E4968E6D4/en/msoidcli_64.msi'
            $MSOnlineAssistantFile = $MSOnlineAssistantUri.Split("/")[-1]
            $MSOnlineModuleUri = 'https://bposast.vo.msecnd.net/MSOPMW/Current/amd64/AdministrationConfig-en.msi'
            $MSOnlineModuleFile = $MSOnline

            Start-BitsTransfer -Source $MSOnlineAssistantUri,$MSOnlineModuleUri -Destination $env:TEMP,$env:TEMP
        }
        If ($Azure) {
            Install-Module Azure
        }
        If ($AzureRM) {
            Install-Module AzureRM
        }
        If ($AzureRMS) {
            https://download.microsoft.com/download/1/6/6/166A2668-2FA6-4C8C-BBC5-93409D47B339/WindowsAzureADRightsManagementAdministration_x64.exe
        }
        If ($OfficeDevPNPSharePointOnline) {
            Install-Module -Name OfficeDevPnP.PowerShell.V16.Commands -Force -Confirm $false
        }
        If ($SharePointOnline) {
            https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/sharepointonlinemanagementshell_5214-1200_x64_en-us.msi
        }
        If ($SkypeForBusinessOnline) {
            https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowershell.exe
        }
    }
    END {
    }
}