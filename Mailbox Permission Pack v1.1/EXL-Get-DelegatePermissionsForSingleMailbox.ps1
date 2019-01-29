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
###     Version 1.0 - 18/03/2018                                                                                         ### 
###                                                                                                                      ###
###                                                                                                                      ###
############################################################################################################################

###   Notes
###
###   Script Use - Exchange On-Premises - Generate reports on a single mailbox - Reports which users have Full Access, Send-As, Send on Behalf and Default Folder permissions to the mailbox you specify
###



param (

    [Parameter(mandatory=$true)]

    [string] $Mailbox


)


$Log = "c:\reports"
$FormatEnumerationLimit=-1



Get-MailboxFolderPermission -Identity "${mailbox}:\Calendar" | Select Identity,FolderName,User,AccessRights | Format-List | out-file "$Log\$Mailbox-Calendar.txt"

Get-MailboxFolderPermission -Identity "${mailbox}:\Tasks" | Select Identity,FolderName,User,AccessRights | Format-List | out-file "$Log\$Mailbox-Tasks.txt"

Get-MailboxFolderPermission -Identity "${mailbox}:\Inbox" | Select Identity,FolderName,User,AccessRights | Format-List | out-file "$Log\$Mailbox-Inbox.txt"

Get-MailboxFolderPermission -Identity "${mailbox}:\Contacts" | Select Identity,FolderName,User,AccessRights | Format-List | out-file "$Log\$Mailbox-Contacts.txt"

Get-MailboxFolderPermission -Identity "${mailbox}:\Notes" | Select Identity,FolderName,User,AccessRights | Format-List | out-file "$Log\$Mailbox-Notes.txt"

Get-MailboxPermission $Mailbox | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false} | Select Identity,User,AccessRights,InheritanceType,IsInherited,IsValid | Format-List | out-file "$Log\$Mailbox-FullAccess.txt"

Get-Mailbox $Mailbox | select DisplayName,GrantSendOnBehalfTo | Format-List | out-file "$Log\$Mailbox-SendOnBehalf.txt"

Get-ADPermission $Mailbox | where {($_.ExtendedRights -like “*Send-As*”) -and ($_.IsInherited -eq $false) -and -not ($_.User -like “NT AUTHORITY\SELF”)} | select Identity,User,AccessRights | Format-List | out-file "$Log\$Mailbox-SendAs.txt"

