############################################################################################################################
###                                                                                                                      ###
###  	Script by Terry Munro -                                                                                          ###
###     Technical Blog -               http://365admin.com.au                                                            ###
###     Webpage -                      https://www.linkedin.com/in/terry-munro/                                          ###
###     TechNet Gallery Scripts -      http://tinyurl.com/TerryMunroTechNet                                              ###
###                                                                                                                      ###
###     TechNet Download link -        https://gallery.technet.microsoft.com/Mailbox-Permission-Pack-9e0f2ace            ###
###                                                                                                                      ###
###     Support -                      http://www.365admin.com.au/2018/01/powershell-scripts-to-report-on-mailbox.html   ###
###                                                                                                                      ###
###     Version 1.0 - 03/01/2018                                                                                         ### 
###                                                                                                                      ###
###                                                                                                                      ###
############################################################################################################################

###   Notes
###
###   Script Use - Exchange On-Premises - List all mailboxes to which a user or email enabled security group has Send As Access
###
###   Parameters will prompt for input
###   Prompt - $AliasOfMailboxOrGroup - enter the ALIAS of the mailbox that you are checking which user or email enabled security group have access

param (
    [Parameter(mandatory=$true)]
    [string] $AliasOfMailboxOrGroup
)

Import-Module ActiveDirectory
Get-Mailbox -ResultSize Unlimited | Get-ADPermission | where {($_.ExtendedRights -like "*Send-As*") -and ($_.User -like "*$AliasOfMailboxOrGroup*")} | ft identity,user,extendedrights