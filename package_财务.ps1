$trade_date = Read-Host -Prompt 'Please input report date [format = YYYYMMDD]'
$trade_year = $trade_date.substring(0, 4)
Copy-Item E:\Works\Trade\Reports\output\$trade_year\$trade_date\08* .
# Copy-Item E:\Works\Trade\Reports_Equity2\output\$trade_year\$trade_date\07* .
7z a C:\Users\Administrator\Desktop\$trade_date-����-����Ȼ *.xlsx
Remove-Item *.xlsx
pause