#Sysadmin Scripts

This repository contains various public scripts and modules I use. These are free to use and modify. If you have any comments or issues, don't hesitate to tell me or create an issue ticket. While there most likely are better scripts out there, I develop these to further my skills.

##Requirements

Most of these items should run in PowerShell 3.0 and higher. As I don't have systems with lower versions of PowerShell installed, I've not tested these on version 2.0. Most if not all of these will not run on PowerShell v1.0.

##Resources

###Modules
Modules currently available or in development.

####O365-RestV1
Module to use the O365 REST API. It is currently lagging behind the v2.0 API version.

####O365-RestV2
Module to use the O365 REST version 2.0 API. 

####O365-Tools
Module to connect with less difficulty to Microsoft cloud services.

####OAuth2Helper
Module to authenticate to the Microsoft API. Supports v1.0 and v2.0.

###Scripts

####Send-PasswordExpirationMail
Currently functions. Trying to rewrite it to make it better.

####Set-O365LIcense
Module to set up Office 365 Licenses. Works with a JSON configuration file and a simple HTML e-mail template.
