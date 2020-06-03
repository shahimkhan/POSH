function Get-MappedDrives($ComputerName){
  $output = @()
  if(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet){
    $Hive = [long]$HIVE_HKU = 2147483651
    $sessions = Get-WmiObject -ComputerName $ComputerName -Class win32_process | ?{$_.name -eq "explorer.exe"}
    if($sessions){
      foreach($explorer in $sessions){
        $sid = ($explorer.GetOwnerSid()).sid
        $owner  = $explorer.GetOwner()
        $RegProv = get-WmiObject -List -Namespace "root\default" -ComputerName $ComputerName | Where-Object {$_.Name -eq "StdRegProv"}
        $DriveList = $RegProv.EnumKey($Hive, "$($sid)\Network")
        if($DriveList.sNames.count -gt 0){
          foreach($drive in $DriveList.sNames){
          $output += "$($drive)`t$(($RegProv.GetStringValue($Hive, "$($sid)\Network\$($drive)", "RemotePath")).sValue)`t$($owner.Domain)`t$($owner.user)`t$($ComputerName)"
          }
        }else{write-debug "No mapped drives on $($ComputerName)"}
      }
    }else{write-debug "explorer.exe not running on $($ComputerName)"}
  }else{write-debug "Can't connect to $($ComputerName)"}
  return $output
}

<#
#Enable if you want to see the write-debug messages
$DebugPreference = "Continue"

$list = "Server01", "Server02"
$report = $(foreach($ComputerName in $list){Get-MappedDrives $ComputerName}) | ConvertFrom-Csv -Delimiter `t -Header Drive, Path, Domain, User, Computer
#>