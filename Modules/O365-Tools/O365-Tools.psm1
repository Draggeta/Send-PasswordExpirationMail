Function Connect-O365Session {
    <# 
    .SYNOPSIS 
        Log in to Azure/Office 365 services. 
    .DESCRIPTION 
        Log in to selected Azure/Office 365 services. Some of these services may require prerequisites to be installed first.
    .PARAMETER Credential
        Credentials to log in to Office 365 with.
    .PARAMETER AzureAD
        Logs in to Microsoft Azure AD.
    .PARAMETER Azure
        Logs in to Microsoft Azure Classic services.
    .PARAMETER AzureRM
        Logs in to Microsoft Azure Resource Model services.
    .PARAMETER AzureRMS
        Logs in to Microsoft Azure Rights Management Services.
    .PARAMETER ComplianceCenter
        Logs in to Microsoft Security and Compliance Center.
    .PARAMETER ExchangeOnline
        Logs in to Microsoft Exchange Online.
    .PARAMETER SharePointDevPnP
        Logs in to Microsoft SharePoint Online with the Office Developers PnP module. Requires the tenant name to be specified.
    .PARAMETER SharePointOnline
        Logs in to Microsoft SharePoint Online. Requires the tenant name to be specified.
    .PARAMETER SkypeForBusinessOnline
        Logs in to Microsoft Skype for Business Online. Requires the tenant name to be specified.
    .PARAMETER TenantName
        Tenant name without the '.onmicrosoft.com' part. Required to connect to SharePoint Online and Skype for Business Online.
    .EXAMPLE 
        Connect-O365Session -Credential $Credential -MsOnline -ExchangeOnline
        Description 
         
        ----------- 
     
        Logs in to Azure AD and Exchange Online.
    .EXAMPLE 
        Connect-O365Session -Credential $Credential -All -TenantName Contoso
        Description 
         
        ----------- 
     
        Logs in to all Azure/Office 365 services. As SharePoint and Skype for Business may require a tenant domain, this must be specified.
    .INPUTS 
    	None. You cannot pipe objects to Connect-O365Session
    .OUTPUTS 
    	None.
    .NOTES 
        Author:   Tony Fortes Ramos 
        Created:  May 02, 2016
    .LINK 
    	Disconnect-O365Session
        Install-O365Module
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
        [Switch]$SharePointDevPnP,

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
                $AzureAD = $true
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
    }
    PROCESS {
        If ($AzureAD) {
            $Name = 'Azure Active Directory'
            $Module = 'MSOnline'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Warning "The $Name module is not installed. Please install the module with the Install-O365Module cmdlet."
            }
            Else {
                Write-Verbose "Importing module $Module."
                Import-Module $Module -DisableNameChecking
                Try {
                    Write-Verbose "Connecting to $Name."
                    Connect-MsolService -Credential $Credential
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }
            }
        }
        If ($Azure) {
            $Name = 'Azure Classic'
            $Module = 'Azure'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Warning "The $Name module is not installed. Please install the module with the Install-O365Module cmdlet."
            }
            Else {
                Write-Verbose "Importing module $Module."
                Import-Module $Module -DisableNameChecking
                Try {
                    Write-Verbose "Connecting to $Name."
                    Add-AzureAccount -Credential $Credential
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }
            }
        }
        If ($AzureRM) {
            $Name = 'Azure Resource Manager'
            $Module = 'AzureRM'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Warning "The $Name module is not installed. Please install the module with the Install-O365Module cmdlet."
            }
            Else {
                Write-Verbose "Importing module $Module."
                Import-Module $Module -DisableNameChecking
                Try {
                    Write-Verbose "Connecting to $Name."
                    Login-AzureRmAccount -Credential $Credential
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }
            }
        }
        If ($AzureRMS) {
            $Name = 'Azure Rights Management Service'
            $Module = 'AADRM'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Warning "The $Name module is not installed. Please install the module with the Install-O365Module cmdlet."
            }
            Else {
                Write-Verbose "Importing module $Module."
                Import-Module $Module -DisableNameChecking
                Try {
                    Write-Verbose "Connecting to $Name."
                    Connect-AadrmService -Credential $Credential
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }
            }
        }
        If ($ComplianceCenter) {
            $Name = 'Security and Compliance Center'
            $Module = 'ComplianceCenter'
            $ConnectionUri = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
            If (Get-PSSession -Name $Module -ErrorAction SilentlyContinue) {
                Write-Warning "There is already a connection with $Name. Skipping the creation of a new session."
            } 
            Else {
                Try {
                    Write-Verbose "Connecting to $Name."
                    $CcSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Name $Module -Credential $Credential -Authentication Basic -AllowRedirection
                    Import-PSSession $CcSession -DisableNameChecking -Prefix CC
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }        
            }
        }
        If ($ExchangeOnline) {
            $Name = 'Exchange Online'
            $Module = 'ExchangeOnline'
            $ConnectionUri = 'https://outlook.office365.com/powershell-liveid/'
            If (Get-PSSession -Name $Module -ErrorAction SilentlyContinue) {
                Write-Warning "There is already a connection with $Name. Skipping the creation of a new session."
            } 
            Else {
                Try {
                    Write-Verbose "Connecting to $Name."
                    $EoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Name $Module -Credential $Credential -Authentication Basic -AllowRedirection
                    Import-PSSession $EoSession -DisableNameChecking
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }        
            }
        }
        If ($SharePointDevPNP) {
            $Name = 'SharePoint Office Dev PNP'
            $Module = 'OfficeDevPnP.PowerShell.V16.Commands'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Warning "The $Name module is not installed. Please install the module with the Install-O365Module cmdlet."
            }
            Else {
                Write-Verbose "Importing module $Module."
                Import-Module $Module -DisableNameChecking
                Try {
                    Write-Verbose "Connecting to $Name."
                    Connect-SPOnline -Url "https://$TenantName-admin.sharepoint.com" -Credential $Credential
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }
            }
        }
        If ($SharePointOnline) {
            $Name = 'SharePoint Online'
            $Module = 'Microsoft.Online.SharePoint.PowerShell'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Warning "The $Name module is not installed. Please install the module with the Install-O365Module cmdlet."
            }
            Else {
                Write-Verbose "Importing module $Module."
                Import-Module $Module -DisableNameChecking
                Try {
                    Write-Verbose "Connecting to $Name."
                    Connect-SPOnline -Url "https://$TenantName-admin.sharepoint.com" -Credential $Credential
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }
            }
        }
        If ($SkypeForBusinessOnline) {
            $Name = 'Skype for Business Online'
            $Module = 'SkypeOnlineConnector'
            If (-not (Get-Module -Name $Module -ListAvailable)) {
                Write-Warning "The $Name module is not installed. Please install the module with the Install-O365Module cmdlet."
            }
            ElseIf (Get-PSSession -Name $Module -ErrorAction SilentlyContinue) {
                Write-Verbose "There is already a connection with $Name. Skipping the creation of a new session."
            }
            Else {
                Write-Verbose "Importing module $Module."
                Import-Module $Module -DisableNameChecking
                Try {
                    Write-Verbose "Connecting to $Name."
                    $SfboSession = New-CsOnlineSession -Credential $Credential -OverrideAdminDomain "$TenantName.onmicrosoft.com"
                    $SfboSession.Name = $Module
                    Import-PSSession $SfboSession
                    Write-Verbose "Successfully connected to $Name."
                }
                Catch {
                    Write-Warning "There was an error connecting to $Name."
                }
            }
        }
    }
    END {
    }
}

Function Disconnect-O365Session {
    <# 
    .SYNOPSIS 
        Log out from Office 365 services. 
    .DESCRIPTION 
        Log out from selected Office 365 services. Some services such as Azure AD and Azure RM don't have an option to close an open connection
    .PARAMETER Azure
        Logs out from Microsoft Azure Classic services.
    .PARAMETER AzureRMS
        Logs out from Microsoft Azure Rights Management Services.
    .PARAMETER ComplianceCenter
        Logs out from Microsoft Security and Compliance Center.
    .PARAMETER ExchangeOnline
        Logs out from Microsoft Exchange Online.
    .PARAMETER SharePointDevPnP
        Logs out from Microsoft SharePoint Online when logged in via the Office Developer PnP module.
    .PARAMETER SharePointOnline
        Logs out from Microsoft SharePoint Online.
    .PARAMETER SkypeForBusinessOnline
        Logs out from Microsoft Skype for Business Online.
    .PARAMETER All
        Logs out from all Microsoft Office 365 services that support logging out.
    .EXAMPLE 
        Disconnect-O365Session -ComplianceCenter -SkypeForBusinessOnline
        Description 
         
        ----------- 
     
        Logs off from Security and Compliance Center and Skype for Business Online.
    .INPUTS 
    	None. You cannot pipe objects to Disconnect-O365Session
    .OUTPUTS 
    	None.
    .NOTES 
        Author:   Tony Fortes Ramos 
        Created:  May 02, 2016
    .LINK 
    	Connect-O365Session
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
        [Switch]$SharePointDevPnP,

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
                    Write-Verbose "Disconnecting from $Name."
                    Remove-AzureAccount -Name (Get-AzureAccount).Id -Confirm $false
                    Write-Verbose "Disconnected from $Name."
                }
                Catch {
                    Write-Warning "Could not remove connection to $Name."
                }
            }
        }
        If ($AzureRMS) {
            $Name = 'Azure Rights Management Service'
            Try {
                Write-Verbose "Disconnecting from $Name."
                Disconnect-AadrmService
                Write-Verbose "Disconnected from $Name."
            }
            Catch {
                Write-Warning "Could not remove connection to $Name."
            }
        }
        If ($ComplianceCenter) {
            $Name = 'Security and Compliance Center'
            $Module = 'ComplianceCenter'
            Try {
                Write-Verbose "Disconnecting from $Name."
                Get-PSSession -Name $Module -ErrorAction SilentlyContinue | Remove-PSSession
                Write-Verbose "Disconnected from $Name."
            }
            Catch {
                Write-Warning "Could not remove session(s) to $Name."
            }
        } 
        If ($ExchangeOnline) {
            $Name = 'Exchange Online'
            $Module = 'ExchangeOnline'
            Try {
                Write-Verbose "Disconnecting from $Name."
                Get-PSSession -Name $Module -ErrorAction SilentlyContinue | Remove-PSSession
                Write-Verbose "Disconnected from $Name."
            }
            Catch {
                Write-Warning "Could not remove session(s) to $Name."
            }
        } 
        If ($SharePointDevPNP) {
            $Name = 'SharePoint Office Dev PNP'
            Try {
                Write-Verbose "Disconnecting from $Name."
                Disconnect-SPOnline
                Write-Verbose "Disconnected from $Name."
            }
            Catch {
                Write-Warning "Could not remove connection to $Name."
            }
        }
        If ($SharePointOnline) {
            $Name = 'SharePoint Online'
            $Module = 'Microsoft.Online.SharePoint.PowerShell'
            Try {
                Write-Verbose "Disconnecting from $Name."
                Disconnect-SPOService
                Write-Verbose "Disconnected from $Name."
            }
            Catch {
                Write-Warning "Could not remove connection to $Name."
            }
        }
        If ($SkypeForBusinessOnline) {
            $Name = 'Skype for Business Online'
            $Module = 'SkypeOnlineConnector'
            Try {
                Write-Verbose "Disconnecting from $Name."
                Get-PSSession -Name $Module -ErrorAction SilentlyContinue | Remove-PSSession
                Write-Verbose "Disconnected from $Name."
            }
            Catch {
                Write-Warning "Could not remove session(s) to $Name."
            }
        }
    }
    END {
    }
}

Function Install-O365Module {
    <# 
    .SYNOPSIS 
        Installs Azure and Office 365 modules. 
    .DESCRIPTION 
        Installs the software/modules required to connect to specific or all Azure/Office 365 services. Some services such as Exchange Online and Security and Compliance Center don't require separate software.
    .PARAMETER Name
        Specifies which modules/software should be installed on the system.
    .PARAMETER All
        Specifies that all required modules/software should be installed on the system.
    .PARAMETER Force
        When used, all selected prerequisites will be first removed and then reinstalled on the system.
    .EXAMPLE 
        Install-O365Module -Name AzureAD,SharePointDevPnP
        Description 
         
        ----------- 
     
        Installs the prerequisites required to connect to the Azure AD, and SharePoint services via the Office Developer PnP module.
    .EXAMPLE 
        Install-O365Module -All -Force
        Description 
         
        ----------- 
     
        Reinstalls all prerequisites required to connect to all Azure/Office 365 services. This command will first uninstall the software and modules and then reinstall them.
    .INPUTS 
    	None. You cannot pipe objects to Disconnect-O365Session
    .OUTPUTS 
    	None.
    .NOTES 
        Author:   Tony Fortes Ramos 
        Created:  May 02, 2016
    .LINK 
    	Connect-O365Session
        Install-Module
        Start-Process
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ParameterSetName = 'Specific')]
        [ValidateSet('AzureAD','Azure','AzureRM','AzureRMS','SharePointDevPnP','SharePointOnline','SkypeForBusinessOnline')]
        [String[]]$Name,

        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [Switch]$All,
        
        [Parameter()]
        [Switch]$Force
    )
    BEGIN {
        If ($All) {
            $Name = 'AzureAD','Azure','AzureRM','AzureRMS','SharePointDevPnP','SharePointOnline','SkypeForBusinessOnline'
        }
        $XmlUri = 'https://raw.githubusercontent.com/draggeta/sysadmin-scripts/development/Modules/O365-Tools/O365-Tools.xml'
        $XmlName = [System.IO.Path]::GetFileName($XmlUri)
        $XmlPath = "$env:TEMP\$XmlName"
        Remove-Item -Path $XmlPath -Confirm:$false -Force
        Start-BitsTransfer -Source $XmlUri -Destination $XmlPath
        $XMLConfig = Select-Xml -Path "$env:TEMP\$XmlName" -XPath '/'
    }
    PROCESS {
        ForEach ($Service in $Name) {            
            $ConfigData = $XmlConfig.Node.Settings.Module | Where Service -eq $Service
            ForEach ($Config in $ConfigData) {
                $InstallUri = $Config.URL
                $InstallName = [System.IO.Path]::GetFileName($InstallUri)
                $InstallPath = "$env:TEMP\$InstallName"
                $InstallID = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{$($Config.ProductID)}" -ErrorAction SilentlyContinue
                If ($Config.Type -eq 'MSI') {
                    If ($Force) {
                        Try {
                            Write-Verbose "Uninstalling $($Config.Name)."
                            Start-Process -FilePath msiexec.exe -ArgumentList "/uninstall $($Config.ProductID) /passive /norestart" -Wait
                            Write-Verbose "Uninstalled $($Config.Name)."
                        }
                        Catch {
                            Write-Warning "An error occured during the uninstallation of $($Config.Name)."
                        }
                    }
                    If (-not $InstallID -or $Force) {
                        Try {
                            Write-Verbose "Downloading $($Config.Name)."
                            Start-BitsTransfer -Source $InstallUri -Destination $InstallPath
                            If (Test-Path $InstallPath) {
                                Write-Verbose "Downloaded $($Config.Name). Validating publisher certificate."
                                $Certificate = Get-AuthenticodeSignature $InstallPath
                                If ($Certificate.Status -eq 'Valid' -and $Certificate.SignerCertificate.DnsNameList.Unicode -contains 'Microsoft Corporation') {
                                    Write-Verbose "Certificate validated. Installing $($Config.Name)."
                                    Start-Process -FilePath msiexec.exe -ArgumentList "/package $InstallPath /passive" -Wait
                                    Write-Verbose "Installed $($Config.Name). Removing installation file."
                                    Remove-Item -Path $InstallPath -Confirm:$false -Force
                                    Write-Verbose "Removed installation file."
                                }
                            }
                        }
                        Catch {
                            Write-Warning "An error occured during the installation of $($Config.Name)."
                        }
                    }
                    ElseIf ($InstallID) {
                        Write-Verbose "$($Config.Name) is already installed."
                    }
                }
                ElseIf ($Config.Type -eq 'EXE') {
                    If ($Force) {
                        Try {
                            Write-Verbose "Uninstalling $($Config.Name)."
                            Start-Process -FilePath msiexec.exe -ArgumentList "/uninstall $($Config.ProductID) /passive /norestart" -Wait
                            Write-Verbose "Uninstalled $($Config.Name)."
                        }
                        Catch {
                            Write-Warning "An error occured during the uninstallation of $($Config.Name)."
                        }
                    }
                    If (-not $InstallID -or $Force) {
                        Try {
                            Write-Verbose "Downloading $($Config.Name)."
                            Start-BitsTransfer -Source $InstallUri -Destination $InstallPath
                            If (Test-Path $InstallPath) {
                                Write-Verbose "Downloaded $($Config.Name). Validating publisher certificate."
                                $Certificate = Get-AuthenticodeSignature $InstallPath
                                If ($Certificate.Status -eq 'Valid' -and $Certificate.SignerCertificate.DnsNameList.Unicode -contains 'Microsoft Corporation') {
                                    Write-Verbose "Certificate validated. Installing $($Config.Name)."
                                    Start-Process $InstallPath -ArgumentList $Config.Argument -Wait
                                    Write-Verbose "Installed $($Config.Name). Removing installation file."
                                    Remove-Item -Path $InstallPath -Confirm:$false -Force
                                    Write-Verbose "Removed installation file."
                                }
                            }
                        }
                        Catch {
                            Write-Warning "An error occured during the installation of $($Config.Name)"
                        }
                    }
                    ElseIf ($InstallID) {
                        Write-Information "$($Config.Name) is already installed."
                    }
                }
                ElseIf ($Config.Type -eq 'MOD') {
                    $InstalledModule = Get-Module -Name $Config.ModuleName -ListAvailable
                    $CurrentModule = Find-Module -Name $Config.ModuleName
                    If ($Force) {
                        Try {
                            Write-Verbose "Uninstalling module $($Config.Name)."
                            Uninstall-Module -Name $Config.ModuleName -AllVersions
                            Write-Verbose "Uninstalled module $($Config.Name)."
                        }
                        Catch {
                            Write-Warning "Could not uninstall module $($Config.Name)."
                        }
                    }
                    If ($InstalledModule.Version -lt $CurrentModule.Version) {
                        Try {
                            Write-Verbose "Installing module $($Config.Name)."
                            Install-Module $Config.ModuleName
                            Write-Verbose "Installed module $($Config.Name)."
                        }
                        Catch {
                            Write-Warning "Could not install module $($Config.Name)."
                        }
                    }
                    ElseIf ($InstalledModule.Version -ge $CurrentModule.Version) {
                        Write-Information "Module $($Config.Name) is already installed."
                    }
                }
            }
        }
    }
    END {
    }
}