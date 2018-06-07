(gci C*.spf | measure Length -s).Sum /1GB

$daily=(gci Rec*cd*.spi | measure Length -a).Average /1GB
$monthly=(gci Rec*cm*.spi | measure Length -a).Average /1GB
write-host -ForegroundColor yellow "Average Daily change for Recovery Drive"=$daily 
write-host -ForegroundColor green "Average Monthly change for Recovery Drive"=$monthly
Clear-Variable -Name daily
Clear-Variable -Name monthly


$daily=(gci C*cd*.spi | measure Length -a).Average /1GB
$monthly=(gci C*cm*.spi | measure Length -a).Average /1GB
write-host -ForegroundColor yellow "Average Daily change for C Drive"=$daily 
write-host -ForegroundColor green "Average Monthly change for C Drive"=$monthly
Clear-Variable -Name daily
Clear-Variable -Name monthly

$daily=(gci F*cd*.spi | measure Length -a).Average /1GB
$monthly=(gci F*cm*.spi | measure Length -a).Average /1GB
write-host -ForegroundColor yellow "Average Daily change for F Drive"=$daily 
write-host -ForegroundColor green "Average Monthly change for F Drive"=$monthly
Clear-Variable -Name daily
Clear-Variable -Name monthly


$daily=(gci S*cd*.spi | measure Length -a).Average /1GB
$monthly=(gci S*cm*.spi | measure Length -a).Average /1GB
write-host -ForegroundColor yellow "Average Daily change for S Drive"=$daily 
write-host -ForegroundColor green "Average Monthly change for S Drive"=$monthly
Clear-Variable -Name daily
Clear-Variable -Name monthly

"{0:N2} GB" -f ((gci C*cd*.spi | measure Length -s).Sum /1GB)