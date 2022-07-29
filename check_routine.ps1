$sep = "=" * 120
Write-Host $sep
Write-Host "Display routine works log"
Write-Host $sep
$d = Read-Host -Prompt "Please input the report date [format = YYYYMMDD]"

Write-Host $sep
Get-Content E:\Works\Download_Program\log\*$d*
Write-Host $sep
Get-Content E:\Works\Download_Program_Security_THS\log\*$d*
Get-Content E:\Works\TradeFuturesAux\log\*$d*

Write-Host $sep
$ans = 47*5
Write-Host "Check Utility Funs for Futures, Normal result should be $ans"
Get-Content E:\Works\UtilityFutures\log\$d.log | Measure-Object
Write-Host "If the count value displayed above is not $ans, Please check E:\Works\UtilityFutures\ for more details"

Write-Host $sep
Pause
