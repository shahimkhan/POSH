#Author: Adnan Amin
#Blog: http://MsTechtalk.com
#LinkedIn: https://www.linkedin.com/in/adnanamin/

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"



function GetListItemCount($siteUrl)
{
    #*** you can also move below line outside the function to get rid of login again if you need to call the function multiple time. ***
    $Cred= Get-Credential
    
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL) 
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.Username, $Cred.Password)
    $ctx.Credentials = $credentials 

    $web = $ctx.Web  

    $lists = $web.Lists
    $ctx.Load($lists)
    $ctx.ExecuteQuery()
    Write-Host -ForegroundColor Yellow "The site URL is" $siteUrl

    #output the list item count
    $tableListNames = foreach ($list in $lists)
    {
        $objList = @{
        "List Name" = $list.Title
        "No. of Items" = $list.ItemCount
        }
        New-Object psobject -Property $objList
    }

    Write-Host -ForegroundColor Green "List item count completed successfully"
    return $tableListNames;
}

#GetListItemCount "https://MSTalk.sharepoint.com/"| Out-GridView

#GetListItemCount "https://MSTalk.sharepoint.com"| ExportCsv -Path "C:\itemcount.csv"