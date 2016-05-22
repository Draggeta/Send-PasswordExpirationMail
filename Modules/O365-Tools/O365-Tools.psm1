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

Function Install-Office365Prerequisite {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ParameterSetName = 'Specific')]
        [ValidateSet('MSOnline','Azure','AzureRM','AzureRMS','SharePointDevPnP','SharePointOnline','SkypeForBusinessOnline')]
        [String[]]$ServiceName,

        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [Switch]$All,
        
        [Parameter()]
        [Switch]$Force
    )
    BEGIN {
        If ($All) {
            $ServiceName = 'MSOnline','Azure','AzureRM','AzureRMS','SharePointDevPnP','SharePointOnline','SkypeForBusinessOnline'
        }
        $Modules = Get-Module -ListAvailable
        [Xml]$ConfigXml = Get-Content 'C:\GitRepositories\sysadmin-scripts\Modules\O365-Tools\O365-ToolsConfig.xml'
    }
    PROCESS {
        ForEach ($Service in $ServiceName) {            
            $ConfigData = $ConfigXml.Settings.Module | Where Service -eq $Service
            ForEach ($Config in $ConfigData) {
                $DownloadUri = $Config.URL
                $Installer = [System.IO.Path]::GetFileName($DownloadUri)
                $InstallerPath = "$env:TEMP\$Installer"
                $InstallerID = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$($Config.ProductID)"
                If ($Config.Type -eq 'MSI') {
                    If ($Force) {
                        Start-Process -FilePath msiexec.exe -ArgumentList "/uninstall $($Config.ProductID) /passive /norestart" -Wait
                    }
                    If (-not $InstallerID) {
                        Try {
                            Start-BitsTransfer -Source $DownloadUri -Destination $InstallerPath
                            $Cert = Get-AuthenticodeSignature $InstallerPath
                            If ($Cert.Status -eq 'Valid' -and $Cert.SignerCertificate.DnsNameList.unicode -contains 'Microsoft Corporation') {
                                Start-Process -FilePath msiexec.exe -ArgumentList "/package $InstallerPath /passive" -Wait
                                Remove-Item -Path $InstallerPath -Confirm:$false -Force
                            }
                        }
                        Catch {
                        }
                    }
                }
                ElseIf ($Config.Type -eq 'EXE') {
                    If ($Force) {
                        Start-Process -FilePath msiexec.exe -ArgumentList "/uninstall $($Config.ProductID) /passive /norestart" -Wait
                    }
                    If (-not $InstallerID) {
                        Try {
                            Start-BitsTransfer -Source $DownloadUri -Destination $InstallerPath
                            $Cert = Get-AuthenticodeSignature $InstallerPath
                            If ($cert.Status -eq 'Valid' -and $Cert.SignerCertificate.DnsNameList.unicode -contains 'Microsoft Corporation') {
                                Start-Process $InstallerPath -ArgumentList $Config.Argument -Wait
                                Remove-Item -Path $InstallerPath -Confirm:$false -Force
                            }
                        }
                        Catch {
                        }
                    }
                }
                ElseIf ($Config.Type -eq 'MOD') {
                    $InstalledModule = Get-Module -Name $Config.ModuleName -ListAvailable
                    $CurrentModule = Find-Module -Name $Config.ModuleName
                    If ($Force) {
                        Try {
                            Uninstall-Module -Name $Config.ModuleName -AllVersions
                        }
                        Catch {
                        }
                    }
                    If ($InstalledModule.Version -lt $CurrentModule.Version) {
                        Try {
                            Install-Module $Config.ModuleName
                        }
                        Catch {
                        }
                    }
                }
            }
        }
    }
    END {
    }
}