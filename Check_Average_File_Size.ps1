$foo = (Get-ChildItem Sys*cm.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem Sys*cm.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "green" "Avg Monthly Sys Drive = "$avrg

$foo = (Get-ChildItem C*cm.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem C*cm.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "green" "Avg Monthly C Drive = "$avrg

$foo = (Get-ChildItem D*cm.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem D*cm.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "green" "Avg Monthly D Drive = "$avrg

$foo = (Get-ChildItem E*cm.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem E*cm.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "green" "Avg Monthly E Drive = "$avrg

$foo = (Get-ChildItem F*cm.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem F*cm.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "green" "Avg Monthly F Drive = "$avrg



$foo = (Get-ChildItem Sys*cd.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem Sys*cd.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "magenta" "Avg Daily Sys Drive = "$avrg

$foo = (Get-ChildItem C*cd.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem C*cd.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "magenta" "Avg Daily C Drive = "$avrg

$foo = (Get-ChildItem D*cd.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem D*cd.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "magenta" "Avg Daily D Drive = "$avrg

$foo = (Get-ChildItem E*cd.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem E*cd.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "magenta" "Avg Daily E Drive = "$avrg

$foo = (Get-ChildItem F*cd.spi -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem F*cd.spi -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "magenta" "Avg Daily F Drive = "$avrg

## Calculate Base image sizes

$foo = (Get-ChildItem Sys*.spf -recurse | measure-object | select -expand Count)
$bar = ((Get-ChildItem Sys*cm.spf -recurse | Measure-Object -property length -sum).sum /1MB)
$avrg = $bar / $foo
write-host -ForegroundColor "green" "Avg Monthly Sys Drive = "$avrg

