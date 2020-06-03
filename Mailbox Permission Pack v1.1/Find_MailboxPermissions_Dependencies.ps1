<#
Name  : Find_MailboxPermissons_Dependencies.ps1
Author: nivleshc@yahoo.com

SCRIPT IS PROVIDED 'AS IS' WITHOUT ANY WARRANTY OF ANY KIND.
Before running any script, please read the entire script to ensure there is nothing in it that will harm your environment

Requirements:
- Run Roman Zarkar's Export-MailboxPermissions.ps1 (this script is part of the bundle at https://gallery.technet.microsoft.com/scriptcenter/Migrate-Mailbox-Permissions-2f262f8b)
and provide locations of the output files MailboxAccess.csv, MailboxFolderDelegate.csv, MaiboxSendAs.csv, MailboxSendOnBehalf.csv
 
- Office365 Mailboxes Details
Run the following command against your Office 365 environment to get details of all your Office 365 mailboxes
Get-Mailbox -ResultSize unlimited | Select DisplayName,UserPrincipalName,EmailAddresses,WindowsEmailAddress,RecipientTypeDetails | Export-Csv -NoTypeInformation -Path O365_Mbx_Details.csv

If you don't have any mailboxes in Office 365, then just create a file with the following header
"DisplayName","UserPrincipalName","EmailAddresses","WindowsEmailAddress","RecipientTypeDetails"

- OnPrem Mailboxes Details
Run a script to get the following information for all your on-premises mailboxes and put it into a csv file
DisplayName,UserPrincipalName,PrimarySmtpAddress,RecipientTypeDetails,Department,Title,Office,State,OrganizationalUnit

Change History
Date        Author    VersionApplicableTo     Comments
08/11/2017  Nivlesh   1.1PR                  Line 443 was referring to WindowsEmailAddress when it is supposed to use primarySmtpAddress.Fixed
#>

#Declare variables ---start
$version = "1.1PR"
$decimal_places = 2 #number of decimal places to round off the mailbox sizes to
$datetime_format = "dd/MM/yyyy hh:mm:ss"
$datetime_now = Get-Date -UFormat "%d%m%yT%H%M%S"
$verbose = $false  #if TRUE then everything is logged to file else only what is displayed on screen is logged (minus the output written to output_file).
                   #if TRUE this can cause some significant performance degredation.

$root_dir = "C:\PermissionAnalysis\"
$MailboxAccess_file = -join ($root_dir,"MailboxAccess.csv")
$MailboxFolderDelegate_file = -join ($root_dir,"MailboxFolderDelegate.csv")
$MailboxSendAs_file = -join ($root_dir,"MailboxSendAs.csv")
$MailboxSendOnBehalf_file = -join ($root_dir,"MailboxSendOnBehalf.csv")
$O365Mbx_file  = -join ($root_dir,"O365_Mbx_Details.csv")
$OnPremMbx_file = -join ($root_dir,"OnPrem_Mbx_Details.csv")

$output_file = -join ($root_dir,"Find_MailboxPermissions_Dependencies_" + $datetime_now + "_csv.csv")
$log_File = -join ($root_dir,"Find_MailboxPermissions_Dependencies_Log_" + $datetime_now + ".log")

$header =  "PermTo_OtherMbx_Or_FromOtherMbx?;PermTo_Or_PermFrom_O365Mbx?;Migration Readiness;DisplayName;UserPrincipalName;PrimarySmtp;MailboxType;Department;Title;"
$header += "SendOnBehalf_GivenTo;SendOnBehalf_GivenOn;SendAs_GivenTo;SendAs_GivenOn;MailboxFolderDelegate_GivenTo;MailboxFolderDelegate_GivenTo_FolderLocation;MailboxFolderDelegate_GivenTo_DelegateAccess;"
$header += "MailboxFolderDelegate_GivenOn;MailboxFolderDelegate_GivenOn_FolderLocation;MailboxFolderDelegate_GivenOn_DelegateAccess;MailboxAccess_GivenTo;MailboxAccess_GivenTo_DelegateAccess;MailboxAccess_GivenOn;"
$header += "MailboxAccess_GivenOn_DelegateAccess;OrganizationalUnit"

#counters
$total_MailboxAccess_records = 0
$total_MailboxFolderDelegates_records = 0
$total_MailboxSendAs_records = 0
$total_MailboxSendOnBehalf_records = 0
$total_O365Mbx_records  = 0
$total_OnPremMbx_records = 0

#colorcodes for the different permission dependancies
$cc_NoDependancy   = "LightBlue"
$cc_MbxAccess      = "DarkGreen"
$cc_OtherMbxinO365 = "LightGreen"
$cc_SendAs         = "Orange"
$cc_FolderDelegate = "Pink"
$cc_SendOnBehalf   = "Red"


#initialise output files
$header | Out-File $output_file
Out-File $log_File

cls
Write-Host -ForegroundColor Yellow "Find_MailboxPermissions_Dependencies.ps1" -NoNewLine; Write-Host -ForegroundColor Green  "   version $version"
"Map_MailboxPermissions.ps1 version $version" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow "`nChecking Input Files"
"Checking Input Files" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow "Mailbox Access File  Path   : " -NoNewLine; Write-Host -ForegroundColor Green $MailboxAccess_file -NoNewline
"Mailbox Access File  Path   :$MailboxAccess_file" | Out-File -Append $log_File

if (!(Test-Path $MailboxAccess_file)){ 
    Write-Host -ForegroundColor Red " File Not Found! Exiting. Please ensure file is in specified location and start script again"
    " File Not Found! Exiting. Please ensure file is in specified location and start script again" | Out-File -Append $log_File
    exit
}else{
 Write-Host -ForegroundColor Cyan " <=Found"
 " <=Found" | Out-File -Append $log_File
}

Write-Host -ForegroundColor Yellow "Mailbox Folder Delegate File: " -NoNewLine; Write-Host -ForegroundColor Green $MailboxFolderDelegate_file -NoNewline
"Mailbox Folder Delegate File: $MailboxFolderDelegate_file" | Out-File -Append $log_File

if (!(Test-Path $MailboxFolderDelegate_file)){
    Write-Host -ForegroundColor Red " File Not Found! Exiting. Please ensure file is in specified location and start script again";
    " File Not Found! Exiting. Please ensure file is in specified location and start script again" | Out-File -Append $log_File
    exit
}else{ 
    Write-Host -ForegroundColor Cyan " <=Found"
    " <=Found" | Out-File -Append $log_File
}

Write-Host -ForegroundColor Yellow "Mailbox SendAs File         : " -NoNewLine; Write-Host -ForegroundColor Green $MailboxSendAs_file -NoNewline
"Mailbox SendAs File         :$MailboxSendAs_file " | Out-File -Append $log_File

if (!(Test-Path $MailboxSendAs_file)){
    Write-Host -ForegroundColor Red " File Not Found! Exiting. Please ensure file is in specified location and start script again"
    " File Not Found! Exiting. Please ensure file is in specified location and start script again" | Out-File -Append $log_File
    exit
}else{
    Write-Host -ForegroundColor Cyan " <=Found"
    " <=Found" | Out-File -Append $log_File
}

Write-Host -ForegroundColor Yellow "Mailbox SendOnBehalf File   : " -NoNewLine; Write-Host -ForegroundColor Green $MailboxSendOnBehalf_file -NoNewline
"Mailbox SendOnBehalf File   :$MailboxSendOnBehalf_file"  | Out-File -Append $log_File

if (!(Test-Path $MailboxSendOnBehalf_file)){
    Write-Host -ForegroundColor Red " File Not Found! Exiting. Please ensure file is in specified location and start script again"
    " File Not Found! Exiting. Please ensure file is in specified location and start script again" | Out-File -Append $log_File
    exit
}else{ 
    Write-Host -ForegroundColor Cyan " <=Found"
    " <=Found" | Out-File -Append $log_File
}

Write-Host -ForegroundColor Yellow "O365 Mailboxes File         : " -NoNewLine; Write-Host -ForegroundColor Green $O365Mbx_file -NoNewline
"O365 Mailboxes File         :$O365Mbx_file" | Out-File -Append $log_File

if (!(Test-Path $O365Mbx_file)){
    Write-Host -ForegroundColor Red " File Not Found! Exiting. Please ensure file is in specified location and start script again"
    " File Not Found! Exiting. Please ensure file is in specified location and start script again" | Out-File -Append $log_File
    exit
}else{
    Write-Host -ForegroundColor Cyan " <=Found"
    " <=Found" | Out-File -Append $log_File
}

Write-Host -ForegroundColor Yellow "OnPremises Mailboxes File   : " -NoNewLine; Write-Host -ForegroundColor Green $OnPremMbx_file -NoNewline
"OnPremises Mailboxes File   :$OnPremMbx_file" | Out-File -Append $log_File

if (!(Test-Path $OnPremMbx_file)){
    Write-Host -ForegroundColor Red " File Not Found! Exiting. Please ensure file is in specified location and start script again"
    " File Not Found! Exiting. Please ensure file is in specified location and start script again" | Out-File -Append $log_File
    exit
}else{
    Write-Host -ForegroundColor Cyan " <=Found"
    " <=Found" | Out-File -Append $log_File
}

Write-Host -ForegroundColor Yellow "Output File                 : " -NoNewLine; Write-Host -ForegroundColor Green $output_file
"Output File                 :$output_file" | Out-File -Append $log_File
Write-Host -ForegroundColor Yellow "Log File                    : " -NoNewLine; Write-Host -ForegroundColor Green $log_file
"Log File                    :$log_file " | Out-File -Append $log_File
Write-Host -ForegroundColor Yellow "Verbose                     : " -NoNewLine; Write-Host -ForegroundColor Green $verbose
"Verbose                     :$verbose" | Out-File -Append $log_File

Write-Host -ForegroundColor White "`n========================================================================================================================================"
Write-Host -ForegroundColor Red "SCRIPT IS PROVIDED 'AS IS' WITHOUT ANY WARRANTY OF ANY KIND."
"SCRIPT IS PROVIDED 'AS IS' WITHOUT ANY WARRANTY OF ANY KIND." | Out-File -Append $log_File
Write-Host -ForegroundColor Red "Before running any script, please read the entire to script to ensure there is nothing in it that will adversely affect your environment"
Write-Host -ForegroundColor White "========================================================================================================================================"
"Before running any script, please read the entire script to ensure there is nothing in it that will harm your environment" | Out-File -Append $log_File


Read-Host "Continue? [Press Enter to Continue or Ctrl+C to quit"

"User pressed a key to Confirm to continue"| Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">Reading File"
 ">Reading File" | Out-File -Append $log_File
Write-Host -ForegroundColor Yellow ">>>>Reading Mailbox SendOnBehalf file" -NoNewline
">>>>Reading Mailbox SendOnBehalf file" | Out-File -Append $log_File
$MailboxSendOnBehalf = Import-Csv -Path $MailboxSendOnBehalf_file
$total_MailboxSendOnBehalf_records = $MailboxSendOnBehalf.count
Write-Host -ForegroundColor Green "..Done. Read $total_MailboxSendOnBehalf_records records"
"..Done. Read $total_MailboxSendOnBehalf_records records" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Reading Mailbox SendAs file" -NoNewline
">>>>Reading Mailbox SendAs file" | Out-File -Append $log_File
$MailboxSendAs = Import-Csv -Path $MailboxSendAs_file
$total_MailboxSendAs_records = $MailboxSendAs.count
Write-Host -ForegroundColor Green "..Done. Read $total_MailboxSendAs_records records"
"..Done. Read $total_MailboxSendAs_records records" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Reading Mailbox Folder Delegate file" -NoNewline
">>>>Reading Mailbox Folder Delegate file" | Out-File -Append $log_File
$MailboxFolderDelegate = Import-Csv -Path $MailboxFolderDelegate_file
$total_MailboxFolderDelegate_records = $MailboxFolderDelegate.count
Write-Host -ForegroundColor Green "..Done. Read $total_MailboxFolderDelegate_records records"
"..Done. Read $total_MailboxFolderDelegate_records records" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Reading Mailbox Access file" -NoNewline
">>>>Reading Mailbox Access file" | Out-File -Append $log_File
$MailboxAccess = Import-Csv -Path $MailboxAccess_file
$total_MailboxAccess_records = $MailboxAccess.count
Write-Host -ForegroundColor Green "..Done. Read $total_MailboxAccess_records records"
">>>>Reading Mailbox Access file" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Reading O365 Mbxs file" -NoNewline
">>>>Reading O365 Mbxs file" | Out-File -Append $log_File
$O365Mbx = Import-Csv -Path $O365Mbx_file
$total_O365Mbx_records = $O365Mbx.count
Write-Host -ForegroundColor Green "..Done. Read $total_O365Mbx_records records"
"..Done. Read $total_O365Mbx_records records" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Reading OnPrem Mbxs file" -NoNewline
">>>>Reading OnPrem Mbxs file" | Out-File -Append $log_File
$OnPremMbx = Import-Csv -Path $OnPremMbx_file
$total_OnPremMbx_records = $OnPremMbx.count
Write-Host -ForegroundColor Green "..Done. Read $total_OnPremMbx_records records"
"..Done. Read $total_OnPremMbx_records records" | Out-File -Append $log_File

#Lets make a hash table for all the files that have been read. The hash table will allow us to quickly find records for mailboxes
#format of hash table will be {mbx email address,contatenated value of indexes in the respective array where the records for that mbx are}

$hash_MailboxSendOnBehalf_GivenTo = @{}
$hash_MailboxSendOnBehalf_GivenOn = @{}
$hash_MailboxSendAs_GivenTo = @{}
$hash_MailboxSendAs_GivenOn = @{}
$hash_MailboxFolderDelegate_GivenTo = @{}
$hash_MailboxFolderDelegate_GivenOn = @{}
$hash_MailboxAccess_GivenTo = @{}
$hash_MailboxAccess_GivenOn = @{}
$hash_O365Mbx               = @{}
$hash_OnPremMbx             = @{}

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for MailboxSendOnBehalf_GivenTo" -NoNewline
">>>>Creating HashTable for MailboxSendOnBehalf_GivenTo" | Out-File -Append $log_File

$index = 0
foreach ($record in $MailboxSendOnBehalf){
    try{
        $hash_MailboxSendOnBehalf_GivenTo.Add($record.MailboxEmail,$index)
        if ($verbose){
            "Added [$($record.MailboxEmail),$index] to hash_MailboxSendOnBehalf_GivenTo" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        $existing_value = $hash_MailboxSendOnBehalf_GivenTo.Item($record.MailboxEmail).toString()
        $new_value      = $existing_value + ",$index"
        $hash_MailboxSendOnBehalf_GivenTo.Set_Item($record.MailboxEmail,$new_value)
        if ($verbose){
            "Existing record for [$($record.MailboxEmail)] with value:[$existing_value] found in hash_MailboxSendOnBehalf_GivenTo. Updated record with $index to [$($record.MailboxEmail),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for MailboxSendOnBehalf_GivenTo..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for MailboxSendOnBehalf_GivenOn" -NoNewline
">>>>Creating HashTable for MailboxSendOnBehalf_GivenOn" | Out-File -Append $log_File
$index = 0
foreach ($record in $MailboxSendOnBehalf){
    try{
        $hash_MailboxSendOnBehalf_GivenOn.Add($record.DelegateEmail,$index)
        if ($verbose){
            "Added [$($record.DelegateEmail),$index] to hash_MailboxSendOnBehalf_GivenOn" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        $existing_value = $hash_MailboxSendOnBehalf_GivenOn.Item($record.DelegateEmail).toString()
        $new_value      = $existing_value + ",$index"
        $hash_MailboxSendOnBehalf_GivenOn.Set_Item($record.DelegateEmail,$new_value)
        if ($verbose){
            "Existing record for [$($record.DelegateEmail)] with value:[$existing_value] found in hash_MailboxSendOnBehalf_GivenOn. Updated record with $index to [$($record.DelegateEmail),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for MailboxSendOnBehalf_GivenOn..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for MailboxSendAs_GivenTo" -NoNewline
">>>>Creating HashTable for MailboxSendAs_GivenTo" | Out-File -Append $log_File
$index = 0
foreach ($record in $MailboxSendAs){
    try{
        $hash_MailboxSendAs_GivenTo.Add($record.MailboxEmail,$index)
        if ($verbose){
            "Added [$($record.MailboxEmail),$index] to hash_MailboxSendAs_GivenTo" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        $existing_value = $hash_MailboxSendAs_GivenTo.Item($record.MailboxEmail).toString()
        $new_value      = $existing_value + ",$index"
        $hash_MailboxSendAs_GivenTo.Set_Item($record.MailboxEmail,$new_value)
        if ($verbose){
            "Existing record for [$($record.MailboxEmail)] with value[$existing_value] found in hash_MailboxSendAs_GivenTo. Updated record with $index to [$($record.MailboxEmail),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for MailboxSendAs_GivenTo..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for MailboxSendAs_GivenOn" -NoNewline
">>>>Creating HashTable for MailboxSendAs_GivenOn" | Out-File -Append $log_File
$index = 0
foreach ($record in $MailboxSendAs){
    try{
        $hash_MailboxSendAs_GivenOn.Add($record.DelegateEmail,$index)
        if ($verbose){
            "Added [$($record.DelegateEmail),$index] to hash_MailboxSendAs_GivenOn" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        $existing_value = $hash_MailboxSendAs_GivenOn.Item($record.DelegateEmail).toString()
        $new_value      = $existing_value + ",$index"
        $hash_MailboxSendAs_GivenOn.Set_Item($record.DelegateEmail,$new_value)
        if ($verbose){
            "Existing record for [$($record.DelegateEmail)] with value:[$existing_value] found in hash_MailboxSendAs_GivenOn. Updated record with $index to [$($record.DelegateEmail),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for MailboxSendAs_GivenOn..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for MailboxFolderDelegate_GivenTo" -NoNewline
">>>>Creating HashTable for MailboxFolderDelegate_GivenTo" | Out-File -Append $log_File
$index = 0
foreach ($record in $MailboxFolderDelegate){
    try{
        $hash_MailboxFolderDelegate_GivenTo.Add($record.MailboxEmail,$index)
        if ($verbose){
            "Added [$($record.MailboxEmail),$index] to hash_MailboxFolderDelegate_GivenTo" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        $existing_value = $hash_MailboxFolderDelegate_GivenTo.Item($record.MailboxEmail).toString()
        $new_value      = $existing_value + ",$index"
        $hash_MailboxFolderDelegate_GivenTo.Set_Item($record.MailboxEmail,$new_value)
        if ($verbose){
            "Existing record for [$($record.MailboxEmail)] with value:[$existing_value] found in hash_MailboxFolderDelegate_GivenTo. Updated record with $index to [$($record.MailboxEmail),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for MailboxFolderDelegate_GivenTo..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for MailboxFolderDelegate_GivenOn" -NoNewline
">>>>Creating HashTable for MailboxFolderDelegate_GivenOn" | Out-File -Append $log_File
$index = 0
foreach ($record in $MailboxFolderDelegate){
    try{
        $hash_MailboxFolderDelegate_GivenOn.Add($record.DelegateEmail,$index)
        if ($verbose){
            "Added [$($record.DelegateEmail),$index] to hash_MailboxFolderDelegate_GivenOn" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        $existing_value = $hash_MailboxFolderDelegate_GivenOn.Item($record.DelegateEmail).toString()
        $new_value      = $existing_value + ",$index"
        $hash_MailboxFolderDelegate_GivenOn.Set_Item($record.DelegateEmail,$new_value)
        if ($verbose){
            "Existing record for [$($record.DelegateEmail)] with value:[$existing_value] found in hash_MailboxFolderDelegate_GivenOn. Updated record with $index to [$($record.DelegateEmail),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for MailboxFolderDelegate_GivenOn..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for MailboxAccess_GivenTo" -NoNewline
">>>>Creating HashTable for MailboxAccess_GivenTo" | Out-File -Append $log_File
$index = 0
foreach ($record in $MailboxAccess){
    try{
        $hash_MailboxAccess_GivenTo.Add($record.MailboxEmail,$index)
        if ($verbose){
            "Added [$($record.MailboxEmail),$index] to hash_MailboxAccess_GivenTo" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        $existing_value = $hash_MailboxAccess_GivenTo.Item($record.MailboxEmail).toString()
        $new_value      = $existing_value + ",$index"
        $hash_MailboxAccess_GivenTo.Set_Item($record.MailboxEmail,$new_value)
        if ($verbose){
            "Existing record for [$($record.MailboxEmail)] with value:[$existing_value] found in hash_MailboxAccess_GivenTo. Updated record with $index to [$($record.MailboxEmail),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for MailboxAccess_GivenTo..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for MailboxAccess_GivenOn" -NoNewline
">>>>Creating HashTable for MailboxAccess_GivenOn" | Out-File -Append $log_File
$index = 0
foreach ($record in $MailboxAccess){
    try{
        $hash_MailboxAccess_GivenOn.Add($record.DelegateEmail,$index)
        if ($verbose){
            "Added [$($record.DelegateEmail),$index] to hash_MailboxAccess_GivenOn" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        $existing_value = $hash_MailboxAccess_GivenOn.Item($record.DelegateEmail).toString()
        $new_value      = $existing_value + ",$index"
        $hash_MailboxAccess_GivenOn.Set_Item($record.DelegateEmail,$new_value)
        if ($verbose){
            "Existing record for [$($record.DelegateEmail)] with value:[$existing_value] found in hash_MailboxAccess_GivenOn. Updated record with $index to [$($record.DelegateEmail),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for MailboxAccess_GivenOn..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for O365 Mbxs" -NoNewline
">>>>Creating HashTable for O365 Mbxs" | Out-File -Append $log_File
$index = 0
foreach ($record in $O365Mbx){
    try{
        $hash_O365Mbx.Add($record.WindowsEmailAddress,$index)
        if ($verbose){
            "Added [$($record.WindowsEmailAddress),$index] to hash_O365Mbx" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        #this shouldn't happen for O365 Users but catch anyways
        $existing_value = $hash_O365Mbx.Item($record.WindowsEmailAddress).toString()
        $new_value      = $existing_value + ",$index"
        $hash_O365Mbx.Set_Item($record.WindowsEmailAddress,$new_value)
        if ($verbose){
            "Existing record for [$($record.WindowsEmailAddress)] with value:[$existing_value] found in hash_O365Mbx. Updated record with $index to [$($record.WindowsEmailAddress),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for O365 Mbxs..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Creating HashTable for OnPrem Mbxs" -NoNewline
">>>>Creating HashTable for OnPrem Mbxs" | Out-File -Append $log_File
$index = 0
foreach ($record in $OnPremMbx){
    try{
        $hash_OnPremMbx.Add($record.PrimarySmtpAddress,$index)
        if ($verbose){
            "Added [$($record.PrimarySmtpAddress),$index] to hash_OnPremMbx" | Out-File -Append $log_File
        }
    }catch{
        #an error occurs during add if there is already an item with same key. in that case concatenate to existing value
        #this shouldn't happen for OnPrem Users but catch anyways
        $existing_value = $hash_OnPremMbx.Item($record.PrimarySmtpAddress).toString()
        $new_value      = $existing_value + ",$index"
        $hash_OnPremMbx.Set_Item($record.PrimarySmtpAddress,$new_value)
        if ($verbose){
            "Existing record for [$($record.PrimarySmtpAddress)] with value:[$existing_value] found in hash_OnPremMbx. Updated record with $index to [$($record.PrimarySmtpAddress),$new_value]" | Out-File -Append $log_File
        }
    }
    $index++
}
Write-Host -ForegroundColor Green "..Done."
">>>>Creating HashTable for OnPrem Mbxs..Done" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow ">>>>Finding Mailbox Permissions Dependancies" -NoNewLine; Write-Host -ForegroundColor Green "..Start"
">>>>Finding Mailbox Permissions Dependancies ..Start" | Out-File -Append $log_File

<#Now that everything has been created and populated, lets go through all the onprem mailboxes and for each mbx do the following
Find 
     what mbx has SendOnBehalfOf Rights to it
     what mbx's this mbx has SendOnBehalfOf rights to
     what mbx has SendAs rights to it
     what mbx this mbx has SendAs rights to
     what mbx has Mailbox Folder Delegates rights to this mbx
     what mbx this mbx has Mailbox Folder Delegates rights to
     what mbx has Mailbox Access to this mbx
     what mbx's this mbx has Mailbox Access to
#>

$processed = 0
$count_No_PermTo_OtherMbx_Or_FromOtherMbx = 0 #counter to find how many mbx do not have any permissions to other mbx and no other mbx has permissions to them
$count_PermTo_Or_PermFrom_O365Mbx         = 0 #counter for all permissions that are to or from mailboxes that have already been migrated to O365
 

foreach ($mbx in $OnPremMbx){
    $mbx_displayName = $mbx.DisplayName
    $mbx_UPN         = $mbx.UserPrincipalName
    $mbx_primarySmtp = $mbx.PrimarySmtpAddress
    $mbx_type        = $mbx.RecipientTypeDetails
    $mbx_OU          = $mbx.OrganizationalUnit
    $mbx_Department  = $mbx.Department
    $mbx_title       = $mbx.Title

    $lines_output = 0 #this counter is used to track how many permissions we have output for this mbx. If this is still zero at the end then
                      #this mbx doesn't have any permission to any other mbx and other mbxs do not have any permissions to this mbx

    $line         = "$mbx_displayname;$mbx_UPN;$mbx_primarySmtp;$mbx_type;$mbx_Department;$mbx_Title"


    #find all mbx that have SendOnBehalfOf rights to this mailbox
    $mbx_MailboxSendOnBehalf_GivenTo = $hash_MailboxSendOnBehalf_GivenTo.Item($mbx_primarySmtp)

    if ($mbx_MailboxSendOnBehalf_GivenTo){
        #for each mailbox that has MailboxSendOnBehalf on this mbx, output the details
        foreach ($index in ($mbx_MailboxSendOnBehalf_GivenTo.toString().split(","))){
            $delegate_mbx_email = ($MailboxSendOnBehalf[$index]).DelegateEmail

            $delegate_mbx_division = ""

            if ($hash_OnPremMbx.ContainsKey($delegate_mbx_email)){
                #$onprem_index = $hash_OnPremMbx.Item($delegate_mbx_email)
                $delegate_mbx_division = $OnPremMbx[($hash_OnPremMbx.Item($delegate_mbx_email))].Division
            }
            
            
            $PermTo_Or_PermFrom_O365Mbx = $false

            if ($delegate_mbx_email -ne $null){
                if ($hash_O365Mbx.Item($delegate_mbx_email)){
                    $PermTo_Or_PermFrom_O365Mbx = $true
                    $count_PermTo_Or_PermFrom_O365Mbx++
                }
            }

            $output = "Y;$PermTo_Or_PermFrom_O365Mbx;$cc_SendOnBehalf;" + $line + ";$delegate_mbx_email;;;;;;;;;;;;;;$mbx_OU"
            $lines_output++
            $output
            $output | Out-File -Append $output_file
            if ($verbose){ #writing to file is quite a drain on performance, so only output to log_file if verbose enabled since this data is in output_file anyways
                $output | Out-File -Append $log_file
            }
        }
    }

    #find all mbx that this mbx has SendOnBehalfOf rights to
    $mbx_MailboxSendOnBehalf_GivenOn = $hash_MailboxSendOnBehalf_GivenOn.Item($mbx_primarySmtp)

    if ($mbx_MailboxSendOnBehalf_GivenOn){
        #for each mailbox that this mbx has MailboxSendOnBehalf permissions, output the details
        foreach ($index in ($mbx_MailboxSendOnBehalf_GivenOn.toString().split(","))){
            $mbx_email = ($MailboxSendOnBehalf[$index]).MailboxEmail
      
            $mbx_email_division = ""

            if ($hash_OnPremMbx.ContainsKey($mbx_email)){
                $mbx_email_division = $OnPremMbx[($hash_OnPremMbx.Item($mbx_email))].Division
            }

            $PermTo_Or_PermFrom_O365Mbx = $false

            if ($mbx_email -ne $null){
                if ($hash_O365Mbx.Item($mbx_email)){
                    $PermTo_Or_PermFrom_O365Mbx = $true
                    $count_PermTo_Or_PermFrom_O365Mbx++
                }
            }

            $output = "Y;$PermTo_Or_PermFrom_O365Mbx;$cc_SendOnBehalf;" + $line + ";;$mbx_email;;;;;;;;;;;;;$mbx_OU"
            $lines_output++
            $output
            $output | Out-File -Append $output_file
            if ($verbose){ #writing to file is quite a drain on performance, so only output to log_file if verbose enabled since this data is in output_file anyways
                $output | Out-File -Append $log_file
            }
        }
    }

    #find all mbx that have SendAs rights to this mailbox
    $mbx_MailboxSendAs_GivenTo = $hash_MailboxSendAs_GivenTo.Item($mbx_primarySmtp)

    if ($mbx_MailboxSendAs_GivenTo){
        #for each mailbox that has been given SendAs permission to this mailbox, output the details
        foreach ($index in ($mbx_MailboxSendAs_GivenTo.toString().split(","))) {
            $delegate_mbx_email = ($MailboxSendAs[$index]).DelegateEmail
            
            $delegate_mbx_division = ""

            if ($hash_OnPremMbx.ContainsKey($delegate_mbx_email)){
                $delegate_mbx_division = $OnPremMbx[($hash_OnPremMbx.Item($delegate_mbx_email))].Division
            }

            $PermTo_Or_PermFrom_O365Mbx = $false

            if ($delegate_mbx_email -ne $null){
                if ($hash_O365Mbx.Item($delegate_mbx_email)){
                    $PermTo_Or_PermFrom_O365Mbx = $true
                    $count_PermTo_Or_PermFrom_O365Mbx++
                }
            }
             
            $output = "Y;$PermTo_Or_PermFrom_O365Mbx;$cc_SendAs;" + $line + ";;;$delegate_mbx_email;;;;;;;;;;;;$mbx_OU"
            $lines_output++
            $output
            $output | Out-File -Append $output_file
            if ($verbose){ #writing to file is quite a drain on performance, so only output to log_file if verbose enabled since this data is in output_file anyways
                $output | Out-File -Append $log_file
            }
        }
    }

    #find all mbx that this mbx has SendAs rights to
    $mbx_MailboxSendAs_GivenOn = $hash_MailboxSendAs_GivenOn.Item($mbx_primarySmtp)
    
    if ($mbx_MailboxSendAs_GivenOn){
        #for each mailbox that this mbx has SendAs permission on, output the details
        foreach ($index in ($mbx_MailboxSendAs_GivenOn.toString().split(","))){
            $mbx_email = ($MailboxSendAs[$index]).MailboxEmail

            $mbx_email_division = ""

            if ($hash_OnPremMbx.ContainsKey($mbx_email)){
                $mbx_email_division = $OnPremMbx[($hash_OnPremMbx.Item($mbx_email))].Division
            }

            $PermTo_Or_PermFrom_O365Mbx = $false

            if ($mbx_email -ne $null){
                if ($hash_O365Mbx.Item($mbx_email)){
                    $PermTo_Or_PermFrom_O365Mbx = $true
                    $count_PermTo_Or_PermFrom_O365Mbx++
                }
            }
            $output = "Y;$PermTo_Or_PermFrom_O365Mbx;$cc_SendAs;" + $line + ";;;;$mbx_email;;;;;;;;;;;$mbx_OU"
            $lines_output++
            $output
            $output | Out-File -Append $output_file
            if ($verbose){ #writing to file is quite a drain on performance, so only output to log_file if verbose enabled since this data is in output_file anyways
                $output | Out-File -Append $log_file
            }
        }
    }

    #find all mbx that have Mailbox Folder Delegate Permission to this mailbox
    $mbx_MailboxFolderDelegate_GivenTo = $hash_MailboxFolderDelegate_GivenTo.Item($mbx_primarySmtp)

    if ($mbx_MailboxFolderDelegate_GivenTo){
        #for each mailbox that has been given Mailbox Folder Delegate Permissions to this mbx, output the details
        foreach ($index in ($mbx_MailboxFolderDelegate_GivenTo.toString().split(","))){
            $delegate_mbx_email      = ($MailboxFolderDelegate[$index]).DelegateEmail
            $delegate_FolderLocation = ($MailboxFolderDelegate[$index]).FolderLocation
            $delegate_DelegateAccess = ($MailboxFolderDelegate[$index]).DelegateAccess

            $delegate_mbx_division = ""

            if ($hash_OnPremMbx.ContainsKey($delegate_mbx_email)){
                $delegate_mbx_division = $OnPremMbx[($hash_OnPremMbx.Item($delegate_mbx_email))].Division
            }

            $PermTo_Or_PermFrom_O365Mbx = $false

            if ($delegate_mbx_email -ne $null){
                if ($hash_O365Mbx.Item($delegate_mbx_email)){
                    $PermTo_Or_PermFrom_O365Mbx = $true
                    $count_PermTo_Or_PermFrom_O365Mbx++
                }
            }
            
            $output = "Y;$PermTo_Or_PermFrom_O365Mbx;$cc_FolderDelegate;" + $line + ";;;;;$delegate_mbx_email;$delegate_FolderLocation;$delegate_DelegateAccess;;;;;;;;$mbx_OU"
            $lines_output++
            $output
            $output | Out-File -Append $output_file
            if ($verbose){ #writing to file is quite a drain on performance, so only output to log_file if verbose enabled since this data is in output_file anyways
                $output | Out-File -Append $log_file
            }
        }
    }

    #find all mbx that this mbx has Mailbox Folder Delegate Permission to
    $mbx_MailboxFolderDelegate_GivenOn = $hash_MailboxFolderDelegate_GivenOn.Item($mbx_primarySmtp)

    if ($mbx_MailboxFolderDelegate_GivenOn){
        #for each mailbox that this mailbox has Mailbox Folder Delegate Permissions on, output the details
        foreach ($index in ($mbx_MailboxFolderDelegate_GivenOn.toString().split(","))){
            $mbx_email = ($MailboxFolderDelegate[$index]).MailboxEmail
            $mbx_FolderLocation = ($MailboxFolderDelegate[$index]).FolderLocation
            $mbx_DelegateAccess = ($MailboxFolderDelegate[$index]).DelegateAccess

            $mbx_email_division = ""

            if ($hash_OnPremMbx.ContainsKey($mbx_email)){
                $mbx_email_division = $OnPremMbx[($hash_OnPremMbx.Item($mbx_email))].Division
            }

            $PermTo_Or_PermFrom_O365Mbx = $false

            if ($mbx_email -ne $null){
                if ($hash_O365Mbx.Item($mbx_email)){
                    $PermTo_Or_PermFrom_O365Mbx = $true
                    $count_PermTo_Or_PermFrom_O365Mbx++
                }
            }

            $output = "Y;$PermTo_Or_PermFrom_O365Mbx;$cc_FolderDelegate;" + $line + ";;;;;;;;$mbx_email;$mbx_FolderLocation;$mbx_DelegateAccess;;;;;$mbx_OU"
            $lines_output++
            $output
            $output | Out-File -Append $output_file
            if ($verbose){ #writing to file is quite a drain on performance, so only output to log_file if verbose enabled since this data is in output_file anyways
                $output | Out-File -Append $log_file
            }
        }
    }

    #find all mbx that have Mailbox Access Permission to this mailbox
    $mbx_MailboxAccess_GivenTo = $hash_MailboxAccess_GivenTo.Item($mbx_primarySmtp)

    if ($mbx_MailboxAccess_GivenTo){
        #for each mailbox that has been given Mailbox Access to this mailbox, output the details
        foreach ($index in ($mbx_MailboxAccess_GivenTo.toString().split(","))){
            $delegate_mbx_email = ($MailboxAccess[$index]).DelegateEmail
            $delegate_mbx_DelegateAccess = ($MailboxAccess[$index]).DelegateAccess

            $delegate_mbx_division = ""

            if ($hash_OnPremMbx.ContainsKey($delegate_mbx_email)){
                $delegate_mbx_division = $OnPremMbx[($hash_OnPremMbx.Item($delegate_mbx_email))].Division
            }

            $PermTo_Or_PermFrom_O365Mbx = $false

            if ($delegate_mbx_email -ne $null){
                if ($hash_O365Mbx.Item($delegate_mbx_email)){
                    $PermTo_Or_PermFrom_O365Mbx = $true
                    $count_PermTo_Or_PermFrom_O365Mbx++
                }
            }
            
            $output = "Y;$PermTo_Or_PermFrom_O365Mbx;$cc_MbxAccess;" + $line + ";;;;;;;;;;;$delegate_mbx_email;$delegate_mbx_DelegateAccess;;;$mbx_OU"
            $lines_output++
            $output
            $output | Out-File -Append $output_file
            if ($verbose){ #writing to file is quite a drain on performance, so only output to log_file if verbose enabled since this data is in output_file anyways
                $output | Out-File -Append $log_file
            }
        }
    }

    #find all mbx that this mbx has Mailbox Access Permission to
    $mbx_MailboxAccess_GivenOn = $hash_MailboxAccess_GivenOn.Item($mbx_primarySmtp)

    if ($mbx_MailboxAccess_GivenOn){
        #for each mailbox that this mailbox has been given Mailbox Access on, output the details
        foreach ($index in ($mbx_MailboxAccess_GivenOn.toString().split(","))){
            $mbx_email = ($MailboxAccess[$index]).MailboxEmail
            $mbx_email_DelegateAccess = ($MailboxAccess[$index]).DelegateAccess

            $mbx_email_division = ""

            if ($hash_OnPremMbx.ContainsKey($mbx_email)){
                $mbx_email_division = $OnPremMbx[($hash_OnPremMbx.Item($mbx_email))].Division
            }

            $PermTo_Or_PermFrom_O365Mbx = $false

            if ($mbx_email -ne $null){
                if ($hash_O365Mbx.Item($mbx_email)){
                    $PermTo_Or_PermFrom_O365Mbx = $true
                    $count_PermTo_Or_PermFrom_O365Mbx++
                }
            }

            $output = "Y;$PermTo_Or_PermFrom_O365Mbx;$cc_MbxAccess;" + $line + ";;;;;;;;;;;;;$mbx_email;$mbx_email_DelegateAccess;$mbx_OU"
            $lines_output++
            $output
            $output | Out-File -Append $output_file
            if ($verbose){
                $output | Out-File -Append $log_file
            }
        }
    }

    if ($lines_output -eq 0){ #there was no line output for this mailbox, that means there are no mbx with permissions to this mbx
                              #and this mbx doesn't have permissions to any other mbx

        $PermTo_Or_PermFrom_O365Mbx = $false

        $output = "N;$PermTo_Or_PermFrom_O365Mbx;$cc_NoDependancy;" + $line + ";;;;;;;;;;;;;;;$mbx_OU"
        $count_No_PermTo_OtherMbx_Or_FromOtherMbx++
        $output
        $output | Out-File -Append $output_file
        if ($verbose){ #writing to file is quite a drain on performance, so only output to log_file if verbose enabled since this data is in output_file anyways
            $output | Out-File -Append $log_file
        }
    }

    $processed++
   
    Write-Progress -Activity "Finding Mailbox Permissions Dependancies" -Status "[Number of Mbx with No_PermTo_OtherMbx_Or_FromOtherMbx:$count_No_PermTo_OtherMbx_Or_FromOtherMbx|Number Of PermTo Or PermFrom O365 Mbx:$count_PermTo_Or_PermFrom_O365Mbx|Processed:$processed|Total:$total_OnPremMbx_records]Email of Mbx Being Analysed:$mbx_primarySmtp" -PercentComplete ([math]::round(($processed / $total_OnPremMbx_records*100),2))

}

Write-Host -ForegroundColor Yellow ">>>>Finding Mailbox Permissions Dependancies" -NoNewLine; Write-Host -ForegroundColor Green "..Finished"
">>>>Finding Mailbox Permissions Dependancies ..Finished" | Out-File -Append $log_File

Write-Host -ForegroundColor Yellow "`nSummary"
"Summary" | Out-File -Append $log_file
Write-Host -ForegroundColor Yellow "Total OnPrem Mbx:$total_OnPremMbx_records"
"Total OnPrem Mbx:$total_OnPremMbx_records" | Out-File -Append $log_file
Write-Host -ForegroundColor Yellow "Total Processed:$processed"
"Total Processed:$processed" | Out-File -Append $log_file
Write-Host -ForegroundColor Yellow "Total Mbx with No PermsTo other mbx or other mbxs with perms on the mbx:$count_No_PermTo_OtherMbx_Or_FromOtherMbx"
"Total Mbx with No PermsTo other mbx or other mbxs with perms on the mbx:$count_No_PermTo_OtherMbx_Or_FromOtherMbx"  | Out-File -Append $log_file
Write-Host -ForegroundColor Yellow "Total Number of PermsTo or Perms from O365 Mbx:$count_PermTo_Or_PermFrom_O365Mbx"
"Total Number of PermsTo or Perms from O365 Mbx:$count_PermTo_Or_PermFrom_O365Mbx"  | Out-File -Append $log_file

Write-Host -ForegroundColor Yellow "`nFile Paths"
Write-Host -ForegroundColor Yellow "Mailbox Access File         : " -NoNewLine; Write-Host -ForegroundColor Green $MailboxAccess_file
Write-Host -ForegroundColor Yellow "Mailbox Folder Delegate File: " -NoNewLine; Write-Host -ForegroundColor Green $MailboxFolderDelegate_file
Write-Host -ForegroundColor Yellow "Mailbox SendAs File         : " -NoNewLine; Write-Host -ForegroundColor Green $MailboxSendAs_file
Write-Host -ForegroundColor Yellow "Mailbox SendOnBehalf File   : " -NoNewLine; Write-Host -ForegroundColor Green $MailboxSendOnBehalf_file
Write-Host -ForegroundColor Yellow "O365 Mailboxes File         : " -NoNewLine; Write-Host -ForegroundColor Green $O365Mbx_file
Write-Host -ForegroundColor Yellow "OnPremises Mailboxes File   : " -NoNewLine; Write-Host -ForegroundColor Green $OnPremMbx_file

Write-Host -ForegroundColor Yellow "Output File                 : " -NoNewLine; Write-Host -ForegroundColor Green $output_file
Write-Host -ForegroundColor Yellow "Log File                    : " -NoNewLine; Write-Host -ForegroundColor Green $log_file