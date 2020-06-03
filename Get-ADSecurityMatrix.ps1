$ADGroups = @{
    DomainLocal = @{}
    Global = @{}
    Universal = @{}
}

#region Get Members security Groups in AD
Get-ADGroup -Filter * |
ForEach-Object{
    $SamAccountName = $_.SamAccountName
    switch ($_.GroupScope) {
        'DomainLocal' {
            $ADGroups.DomainLocal.$SamAccountName = Get-ADGroupMember -Identity $SamAccountName | Select-Object -ExpandProperty SamAccountName
        }
        'Global'{
            $ADGroups.Global.$SamAccountName = Get-ADGroupMember -Identity $SamAccountName | Select-Object -ExpandProperty SamAccountName
        }
        'Universal' {
            $ADGroups.Universal.$SamAccountName = Get-ADGroupMember -Identity $SamAccountName | Select-Object -ExpandProperty SamAccountName
        }
    }
}
#endregion

#region Get Member Count AD Groups
$ADGroups.Keys |
ForEach-Object{
    $GroupScope = $_

    $ADGroups.$GroupScope.Keys |
    ForEach-Object{
        [PSCustomObject]@{
            Group = $_
            GroupScope = $GroupScope
            Count = @($ADGroups.$GroupScope.$_).Count
        }
    }
}
#endregion

#region Get Security Matrix for Global AD Groups
$htUsers = @{}
$htProps = @{}
$ADGroups.Global.Keys | ForEach-Object {$htProps.$_ = $null}

foreach ($group in $ADGroups.Global.keys){
   foreach ($user in $ADGroups.Global.$($group)){
    
      if (!$htUsers.ContainsKey($user)){
         $htProps.SamAccountName = $user
         $htUsers.$user = $htProps.Clone()
      }
      ($htUsers.$user).$($group) = 'x'
   }
}

$htUsers.GetEnumerator() | 
ForEach-Object{
      [PSCustomObject]$($_.Value)
} |
#Out-GridView |
 export-csv c:\temp\GroupMatrix_25Oct18.csv -notypeinformation
#endregion