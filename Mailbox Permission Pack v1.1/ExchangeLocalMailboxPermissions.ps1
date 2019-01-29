#

###   Notes
###
###   Script Use - Exchange On-Premises - Generate reports on ALL mailboxes with Full Access, Send As, Send on Behalf permissions and who has that permission
###



########################################################

$Mailboxes = get-mailbox -ResultSize Unlimited

$logpath = "c:\reports"

########################################################

Import-Module ActiveDirectory

$Mailboxes | Get-ADPermission | where {($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")} | Select Identity,User | Export-Csv -NoTypeInformation "$logpath\MailboxSendAsAccess-LocalExchange.csv"

$Mailboxes | Where-Object {$_.GrantSendOnBehalfTo} | select Name,@{Name='GrantSendOnBehalfTo';Expression={($_ | Select -ExpandProperty GrantSendOnBehalfTo | Select -ExpandProperty Name) -join ","}} | export-csv -notypeinformation "$logpath\MailboxSendOnBehalf-LocalExchange.csv"

$Mailboxes | Get-MailboxPermission | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like “NT AUTHORITY\SELF”) -and -not ($_.User -like '*Discovery Management*') } | Select Identity, user | Export-Csv -NoTypeInformation "$logpath\MailboxFullAccess-LocalExchange.csv"

