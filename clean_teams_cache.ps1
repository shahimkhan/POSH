Write-Host "Clearing Teams Disk Cache" -ForegroundColor Yellow
try{
Stop-Process -Name Teams -WhatIf;
Write-Host "Teams process killed" -ForegroundColor Green

Get-ChildItem "C:\Users\*\AppData\Roaming\Microsoft\Teams\*" -directory | Where name -in ('application cache','blob storage','databases','GPUcache','IndexedDB','Local Storage','tmp') | ForEach{Remove-Item $_.FullName -Recurse -Force -WhatIf}
Write-Host "Teams Disk Cache Cleaned" -ForegroundColor Green
}catch{
echo $_
}
