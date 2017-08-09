$computerlist = Get-Content C:\Temp\MEL_machines.txt

Foreach ($computer in $computerlist) 
    { 
    if (Test-Connection -Computername $computer -BufferSize 16 -Count 1 -Quiet) 
        {  
            $updatesession =  [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$computer))  
            $updatesearcher = $updatesession.CreateUpdateSearcher()  
            $searchresult = $updatesearcher.Search("IsInstalled=0")  
            #Write-Output "$((Get-Date).ToShortTimeString()): PATCHING MESSAGE - There are $($searchresult.Updates.count) updates via WSUS to be processed on $($computer)"  
            $output= "$((Get-Date).ToString()): PATCHING MESSAGE - There are $($searchresult.Updates.count) updates via WSUS to be processed on $($computer)"  
            $output | add-content c:\temp\MEL_patches.txt 
        }
    }

