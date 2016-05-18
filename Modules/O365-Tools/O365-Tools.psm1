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
        Connect+-Office365Session -Credential $Credential -MsOnline -ExchangeOnline
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
        [Switch]$MSOnline,
        
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
        If ($MSOnline) {
            Write-Verbose 'Testing if the Azure Active Directory (MSOnline) module is available.'
            If ($testModule.Name -notcontains 'MSOnline') {
                Write-Verbose 'The Azure AD module is not installed. Please download and install it from http://go.microsoft.com/fwlink/p/?linkid=236297 before trying again. Also make sure that the Microsoft Online Services Sign-In Assistant for IT Professionals is installed. Download the Sign-In Assistant from https://www.microsoft.com/en-US/download/details.aspx?id=41950.'
            }
            ElseIf ($testModule.Name -contains 'MSOnline') {
                Write-Verbose 'The Azure AD module is installed on the system. Importing.'
                If (-not (Get-Module MSOnline)) {
                    Try {
                        Import-Module MSOnline -DisableNameChecking
                        Write-Verbose 'The Azure AD module has been imported'
                    }
                    Catch {
                        Write-Verbose 'Could not import the Azure AD Module.'
                    }
                }
                ElseIf (Get-Module MSOnline) {
                    Write-Verbose 'The Microsoft Azure AD module has already been imported'
                }
                Write-Verbose 'Connecting to the Microsoft Office 365 services.'
                Try {
                    Connect-MsolService -Credential $Credential
                }
                Catch {
                    Write-Warning 'Could not connect to Microsoft Office 365.'
                }
            }
        }
        If ($Azure) {
            Write-Verbose 'Testing if the Azure module is available.'
            If ($testModule.Name -notcontains 'Azure') {
                Write-Verbose 'The Azure module is not installed. Please download it from https://www.microsoft.com/en-us/download/details.aspx?id=35588 before trying again.'
            }
            ElseIf ($testModule.Name -contains 'Azure') {
                Write-Verbose 'The Azure module is installed on the system. Importing.'
                If (-not (Get-Module Azure)) {
                    Try {
                        Import-Module Azure -DisableNameChecking
                        Write-Verbose 'The Azure module has been imported'
                    }
                    Catch {
                        Write-Verbose 'Could not import the Azure Module. Check if the shell is run as an administrator.'
                    }
                }
                ElseIf (Get-Module Azure) {
                    Write-Verbose 'The Azure module has already been imported'
                }
                Write-Verbose 'Testing if there is a connection with the Azure Classic services.'
                $testSession = Get-AzureAccount
                If ($testSession) {
                    Write-Verbose 'There is already a connection with Azure Classic services.'
                }
                ElseIf (-not $testSession) {
                    Write-Verbose 'There is no existing session. Creating a session to the Azure Classic services.'
                    Try {
                        Add-AzureAccount -Credential $Credential
                        Write-Verbose 'Successfully connected to Azure Classic.'
                    }
                    Catch {
                        Write-Verbose 'There was an error connecting to the Azure Classic services.'
                    }
                }
            }
        }
        If ($AzureRM) {
            Write-Verbose 'Testing if the Azure RM module is available.'
            If ($testModule.Name -notcontains 'AzureRM') {
                Write-Verbose 'The Azure RM module is not installed. Please download it from https://www.microsoft.com/en-us/download/details.aspx?id=35588 before trying again.'
            }
            ElseIf ($testModule.Name -contains 'AzureRM') {
                Write-Verbose 'The Azure RM module is installed on the system. Importing.'
                If (-not (Get-Module AzureRM)) {
                    Try {
                        Import-Module AzureRM -DisableNameChecking
                        Write-Verbose 'The Azure RM module has been imported'
                    }
                    Catch {
                        Write-Verbose 'Could not import the Azure RM Module. Check if the shell is run as an administrator.'
                    }
                }
                ElseIf (Get-Module AzureRM) {
                    Write-Verbose 'The Azure RM module has already been imported'
                }
                $testSession = Get-AzureRmContext
                If ($testSession) {
                    Write-Verbose 'There is already a connection with Azure RM services.'
                }
                ElseIf (-not $testSession) {
                    Write-Verbose 'There is no existing session. Creating a session to the Azure RM services.'
                    Try {
                        Login-AzureRmAccount -Credential $Credential
                        Write-Verbose 'Successfully connected to Azure RM.'
                    }
                    Catch {
                        Write-Verbose 'There was an error connecting to the Azure RM services.'
                    }
                }
            }
        }
        If ($AzureRMS) {
            Write-Verbose 'Testing if the Azure RMS module is available.'
            If ($testModule.Name -notcontains 'AADRM') {
                Write-Verbose 'The Azure RMS module is not installed. Please download it from https://www.microsoft.com/en-us/download/details.aspx?id=35588 before trying again.'
            }
            ElseIf ($testModule.Name -contains 'AADRM') {
                Write-Verbose 'The Azure RMS module is installed on the system. Importing.'
                If (-not (Get-Module AADRM)) {
                    Try {
                        Import-Module AADRM -DisableNameChecking
                        Write-Verbose 'The Azure RMS module has been imported'
                    }
                    Catch {
                        Write-Verbose 'Could not import the Azure RMS Module. Check if the shell is run as an administrator.'
                    }
                }
                ElseIf (Get-Module AADRM) {
                    Write-Verbose 'The Azure RMS module has already been imported'
                }
                Write-Verbose 'Connecting to the Azure RMS services.'
                Try {
                    Connect-AadrmService -Credential $Credential
                    Write-Verbose 'Successfully connected to Azure RMS.'
                }
                Catch {
                    Write-Verbose 'There was an error connecting to the Azure RMS services.'
                }
            }
        }
        If ($ComplianceCenter) {
            Write-Verbose 'Testing if there is a connection with the Microsoft Security and Compliance Center.' 
            $testSession = (Get-PSSession).where{ $_.ComputerName -like "*.compliance.protection.outlook.com" }
            If ($testSession) {
                Write-Verbose 'There is already a connection with the Microsoft Security Compliance Center. Skipping the creation of a new session.'
            } 
            ElseIf (-not $testSession) {
                Write-Verbose 'There is no existing session. Creating a session to the Microsoft Security and Compliance Center.'
                Try {
                    $ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
                    Import-PSSession $ccSession -Prefix CC
                    Write-Verbose 'Successfully connected to the Microsoft Security and Compliance Center.'
                }
                Catch {
                    Write-Verbose 'There was an error connecting to the Microsoft Security and Compliance Center.'
                }        
            }
        }
        If ($ExchangeOnline) {
            Write-Verbose 'Testing if there is a connection with the Microsoft Exchange Online services.'
            $testSession = (Get-PSSession).where{ $_.ComputerName -like 'outlook.office365.com' }
            If ($testSession) {
                Write-Verbose 'There is already a connection with Microsoft Exchange Online services.'
            }
            ElseIf (-not $testSession) {
                Write-Verbose 'There is no existing session. Creating a session to the Microsoft Exchange Online services.'
                Try {
                    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
                    Import-PSSession $exchangeSession -DisableNameChecking
                    Write-Verbose 'Successfully connected to Microsoft Exchange Online.'
                }
                Catch {
                    Write-Verbose 'There was an error connecting to the Microsoft Exchange Online services.'
                }
            }
        }
        If ($SharePointDevPNP) {
            Write-Verbose 'Testing if the Microsoft SharePoint Office Dev PNP module is available.'
            If ($testModule.Name -notcontains 'OfficeDevPnP.PowerShell.V16.Commands') {
                Write-Verbose 'The Microsoft SharePoint Office Dev PNP module is not installed. Please download it from https://www.microsoft.com/en-us/download/details.aspx?id=35588 before trying again.'
            }
            ElseIf ($testModule.Name -contains 'OfficeDevPnP.PowerShell.V16.Commands') {
                Write-Verbose 'The Microsoft SharePoint Office Dev PNP module is installed on the system. Importing.'
                If (-not (Get-Module OfficeDevPnP.PowerShell.V16.Commands)) {
                    Try {
                        Import-Module OfficeDevPnP.PowerShell.V16.Commands -DisableNameChecking
                        Write-Verbose 'The Microsoft SharePoint Office Dev PNP module has been imported'
                    }
                    Catch {
                        Write-Verbose 'Could not import the Microsoft SharePoint Office Dev PNP Module. Check if the shell is run as an administrator.'
                    }
                }
                ElseIf (Get-Module OfficeDevPnP.PowerShell.V16.Commands) {
                    Write-Verbose 'The Microsoft SharePoint Office Dev PNP module has already been imported'
                }
                Write-Verbose 'Connecting to the Microsoft SharePoint Office Dev PNP services.'
                Try {
                    Connect-SPOnline -Url "https://$TenantName-admin.sharepoint.com" -Credential $Credential
                    Write-Verbose 'Connected to the Microsoft SharePoint Office Dev PNP services.'
                }
                Catch {
                    Write-Verbose 'Could not connect to Microsoft SharePoint Office Dev PNP.'
                }
            }
        }
        If ($SharePointOnline) {
            Write-Verbose 'Testing if the Microsoft SharePoint Online module is available.'
            If ($testModule.Name -notcontains 'Microsoft.Online.SharePoint.PowerShell') {
                Write-Verbose 'The Microsoft SharePoint Online module is not installed. Please download it from https://www.microsoft.com/en-us/download/details.aspx?id=35588 before trying again.'
            }
            ElseIf ($testModule.Name -contains 'Microsoft.Online.SharePoint.PowerShell') {
                Write-Verbose 'The Microsoft SharePoint Online module is installed on the system. Importing.'
                If (-not (Get-Module Microsoft.Online.SharePoint.PowerShell)) {
                    Try {
                        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
                        Write-Verbose 'The Microsoft SharePoint Online module has been imported'
                    }
                    Catch {
                        Write-Verbose 'Could not import the Microsoft SharePoint Online Module. Check if the shell is run as an administrator.'
                    }
                }
                ElseIf (Get-Module Microsoft.Online.SharePoint.PowerShell) {
                    Write-Verbose 'The Microsoft SharePoint Online module has already been imported'
                }
                Write-Verbose 'Connecting to the Microsoft SharePoint Online services.'
                Try {
                    Connect-SPOService -Url "https://$TenantName-admin.sharepoint.com" -Credential $Credential
                    Write-Verbose 'Connected to the Microsoft SharePoint Online services.'
                }
                Catch {
                    Write-Verbose 'Could not connect to Microsoft SharePoint Online.'
                }
            }
        }
        If ($SkypeForBusinessOnline) {
            Write-Verbose 'Testing if the Microsoft Skype for Business Online module is available.'
            If ($testModule.Name -notcontains 'SkypeOnlineConnector') {
                Write-Verbose 'The Microsoft Skype for Business Online module is not installed. Please download it from https://www.microsoft.com/en-us/download/details.aspx?id=39366 before trying again.'
            }
            ElseIf ($testModule.Name -contains 'SkypeOnlineConnector') {
                Write-Verbose 'The Microsoft Skype for Business Online module is installed on the system. Importing.'
                If (-not (Get-Module SkypeOnlineConnector)) {
                    Try {
                        Import-Module SkypeOnlineConnector -DisableNameChecking
                        Write-Verbose 'The Microsoft Skype for Business Online module has been imported'
                    }
                    Catch {
                        Write-Verbose 'Could not import the Microsoft Skype for Business Online Module.'
                    }
                }
                ElseIf (Get-Module SkypeOnlineConnector) {
                    Write-Verbose 'The Microsoft Skype for Business Online module has already been imported'
                }
                Write-Verbose 'Connecting to the Microsoft Skype for Business Online services.'
                $testSession = (Get-PSSession).where{ $_.ComputerName -like "*.online.lync.com" }
                If ($testSession) {
                    Write-Verbose 'There is already a connection with Microsoft Skype for Business Online.'
                }
                ElseIf (-not $testSession) {
                    Write-Verbose 'There is no existing session. Creating a session to the Microsoft Skype for Business Online services.'
                    Try {
                        $SfboSession = New-CsOnlineSession -Credential $Credential -OverrideAdminDomain "$TenantName.onmicrosoft.com"
                        Import-PSSession $sfboSession
                        Write-Verbose 'Successfully connected to Microsoft Skype for Business Online.'
                    }
                    Catch {
                        Write-Verbose 'There was an error connecting to the Microsoft Skype for Business Online services.'
                    }
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
            Write-Verbose 'Testing if there is a session active to Microsoft Security and Compliance Center.'
            $testSession = Get-AzureAccount
            If ($testSession) {
                Write-Verbose 'There are active sessions to the Microsoft Security and Compliance Center. Starting the removal of the active sessions.'
                Try {
                    $testSession | Remove-AzureAccount
                    Write-Verbose 'Removed all sessions to the Microsoft Security and Compliance Center.'
                }
                Catch {
                    Write-Verbose 'Could not remove sessions to the Microsoft Security and Compliance Center.'
                }
            }
            ElseIf (-not $testSession) {
                Write-Verbose 'There are no active sessions to the Microsoft Security and Compliance Center.'
            }
        }
        If ($AzureRMS) {
            Write-Verbose 'Removing all sessions to Microsoft Sharepoint Online.'
            Try {
                Disconnect-AadrmService
                Write-Verbose 'Removed all sessions to Microsoft Sharepoint Online.'
            }
            Catch {
                Write-Verbose 'There were no connections to SharePoint Online, or the sessions could not be disconnected.'
            }
        }
        If ($ComplianceCenter) {
            Write-Verbose 'Testing if there is a session active to Microsoft Security and Compliance Center.'
            $testSession = (Get-PSSession).where{ $_.ComputerName -like "*.compliance.protection.outlook.com" } 
            If ($testSession) {
                Write-Verbose 'There are active sessions to the Microsoft Security and Compliance Center. Starting the removal of the active sessions.'
                Try {
                    $testSession | Remove-PSSession
                    Write-Verbose 'Removed all sessions to the Microsoft Security and Compliance Center.'
                }
                Catch {
                    Write-Verbose 'Could not remove sessions to the Microsoft Security and Compliance Center.'
                }
            }
            ElseIf (-not $testSession) {
                Write-Verbose 'There are no active sessions to the Microsoft Security and Compliance Center.'
            }
        }
        If ($ExchangeOnline) {
            Write-Verbose 'Testing if there is a session active to Microsoft Exchange Online.'
            $testSession = (Get-PSSession).where{ $_.ComputerName -like 'outlook.office365.com' } 
            If ($testSession) {
                Write-Verbose 'There are active sessions to the Microsoft Exchange Online services. Starting the removal of the active sessions.'
                Try {
                    $testSession | Remove-PSSession
                    Write-Verbose 'Removed all sessions to the Microsoft Exchange Online services.'
                }
                Catch {
                    Write-Verbose 'Could not remove sessions to the Microsoft Exchange Online services.'
                }
            }
            ElseIf (-not $testSession) {
                Write-Verbose 'There are no active sessions to the Microsoft Exchange Online services.'
            }
        }
        If ($SharePointDevPNP) {
            Write-Verbose 'Removing all sessions to Microsoft SharePoint Office Dev PNP.'
            Try {
                Disconnect-SPOnline
                Write-Verbose 'Removed all sessions to Microsoft SharePoint Office Dev PNP.'
            }
            Catch {
                Write-Verbose 'There were no connections to SharePoint Office Dev PNP, or the sessions could not be disconnected.'
            }
        }
        If ($SharePointOnline) {
            Write-Verbose 'Removing all sessions to Microsoft Sharepoint Online.'
            Try {
                Disconnect-SPOService
                Write-Verbose 'Removed all sessions to Microsoft Sharepoint Online.'
            }
            Catch {
                Write-Verbose 'There were no connections to SharePoint Online, or the sessions could not be disconnected.'
            }
        }
        if ($SkypeForBusinessOnline) {
            Write-Verbose 'Testing if there is a session active to Microsoft Skype for Business Online.'
            $testSession = (Get-PSSession).where{ $_.ComputerName -like "*.online.lync.com" }
            If ($testSession) {
                Write-Verbose 'There are active sessions to the Microsoft Skype for Business Online services. Starting the removal of the active sessions.'
                Try {
                    $testSession | Remove-PSSession
                    Write-Verbose 'Removed all sessions to the Microsoft Skype for Business Online services.'
                }
                Catch {
                    Write-Verbose 'Could not remove sessions to the Microsoft Skype for Business Online services.'
                }
            }
            ElseIf (-not $testSession) {
                Write-Verbose 'There are no active sessions to Microsoft Skype for Business Online.'
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
            Install-Module -Name OfficeDevPnP.PowerShell.V16.Commands -Force -Confirm:$false
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