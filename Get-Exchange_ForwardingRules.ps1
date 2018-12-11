#====================================================
# Get-Exchange_ForwardingRules.ps1
# Author: Shahim Khan
# This scripts gets all the forwarding and redirectTo Rules and `
# Create a Report in a presentable format.
#====================================================
$Report = @()

$Report += foreach ($i in (Get-Mailbox -ResultSize unlimited)) { Get-InboxRule -Mailbox $i.DistinguishedName | 
				Where {($_.RedirectTo) -or ($_.ForwardTo) -or ($_.ForwardAsAttachmentTo) } | 
				Select @{Name="Rule Name";Expression={$_.Name}},`
				       @{Name="Rule Description";Expression={$_.Description}},`
				       @{Name="PrimarySmtpAddress";Expression={$i.PrimarySmtpAddress.Address}},`
				       @{Name="RedirectTo";Expression={$_.RedirectTo -join ','}},`
				       @{Name="ForwardTo";Expression={$_.ForwardTo -join ','}},`
				       @{Name="ForwardAsAttachmentTo";Expression={$_.ForwardAsAttachmentTo -join ','}} }
				
$Report | Export-CSV C:\temp\InboxRules_Forwards_12Dec18.csv
