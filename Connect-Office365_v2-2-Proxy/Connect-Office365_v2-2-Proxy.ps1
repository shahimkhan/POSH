#################################################################################################################
###                                                                                                           ###
###  	Script by Terry Munro -                                                                               ###
###     Technical Blog -               http://365admin.com.au                                                 ###
###     Webpage -                      https://www.linkedin.com/in/terry-munro/                               ###
###     TechNet Gallery Scripts -      http://tinyurl.com/TerryMunroTechNet                                   ###
###                                                                                                           ###            
###     Version 1.0        - 07/02/2017                                                                       ###
###     Version 2.0        - 15/4/2017                                                                        ###
###     Version 2.1        - 05/05/2017                                                                       ### 
###     Version 2.2        - 02/06/2017                                                                       ###
###     Version 2.2-Proxy  - 20/01/2018                                                                       ###
###     Revision -                                                                                            ###
###               v2.0        - Added connection to other services and load modules                           ###
###               v2.1        - Added Azure AD Rights Management connection                                   ###
###               v2.2        - Added variables to simplify editing the script                                ###
###               v2.2-Proxy  - Second script added with proxy settings                                       ###
###                                                                                                           ###  	
###     Created with the follow links as reference                                                            ###
###     - http://powershellblogger.com/2016/02/connect-to-all-office-365-services-with-powershell/            ###
###     - https://technet.microsoft.com/en-us/library/dn568015.aspx                                           ###
###                                                                                                           ###
###                                                                                                           ###
#################################################################################################################

####  Notes for Usage  ######################################################################
#                                                                                           #
#  Ensure you update the script with your tenant name and username                          #
#  Your username is in the Exchange Online section for Get-Credential                       #
#  The tenant name is used in the Exchange Online section for Get-Credential                #
#  The tenant name is used in the SharePoint Online section for SharePoint connection URL   # 
#                                                                                           #
#  Support Guides -                                                                         #
#   - Pre-Requisites -                                                                      #
#   - - - http://www.365admin.com.au/2017/01/how-to-configure-your-desktop-pc-for.html      #      
#   - Usage Guide -                                                                         # 
#   - - - http://www.365admin.com.au/2017/01/how-to-connect-to-office-365-via.html          #
#                                                                                           #
#############################################################################################


#####################################################################################################

###                      Edit the two variables below with your details                           ###


$Tenant = "TenantName"

$Cred = Get-credential "admin@tenant.onmicrosoft.com"


#####################################################################################################

$ProxySettings = New-PSSessionOption -ProxyAccessType IEConfig


###   Exchange Online
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -Authentication "Basic" -AllowRedirection -SessionOption $ProxySettings -ErrorAction Stop
Import-PSSession $ExchangeSession -DisableNameChecking -AllowClobber -ErrorAction Stop


### Exchange Online Protection
$EOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $EOPSession –AllowClobber


### Compliance Center
$ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
Import-PSSession $ccSession –AllowClobber


### Azure Active Directory Rights Management
Import-Module AADRM
Connect-AadrmService -Credential $cred
    

### Azure Resource Manager
Login-AzureRmAccount -Credential $cred


###   Azure Active Directory v1.0
Import-Module MsOnline
Connect-MsolService -Credential $cred


###  SharePoint Online
Import-Module Microsoft.Online.SharePoint.PowerShell
Connect-SPOService -Url "https://$($Tenant)-admin.sharepoint.com" -Credential $cred


### Skype Online
Import-Module SkypeOnlineConnector
$SkypeSession = New-CsOnlineSession -Credential $cred
Import-PSSession $SkypeSession 


### Azure AD v2.0
Connect-AzureAD -Credential $cred

