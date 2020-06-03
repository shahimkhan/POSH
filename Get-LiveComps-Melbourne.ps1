$outputfile="c:\temp\machines.txt"
if (Test-Path $outputfile) { Clear-Content -path $outputfile }
#Get-ADComputer -Filter * -SearchBase "OU=Client Machines,OU=New York Office,DC=mesoblastltd,DC=local" | sort-object DNSHostName |
Get-ADComputer -Filter * | sort-object DNSHostName |
ForEach-Object {
                $rtn = Test-Connection -CN $_.dnshostname -Count 1 -BufferSize 16 -Quiet
                #IF($rtn -match 'True') {write-host -ForegroundColor green $_.dnshostname}
                IF($rtn -match 'True') {add-content $outputfile $_.dnshostname}
                #ELSE { Write-host -ForegroundColor red $_.dnshostname }
                }
#Write-host $machines | out-file -FilePath c:\temp\machines.txt