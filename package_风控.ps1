$trade_date = Read-Host -Prompt 'Please input report date [format = YYYYMMDD]'
$trade_year = $trade_date.substring(0, 4)
Copy-Item E:\Works\Trade\Reports\output\$trade_year\$trade_date\0[6-7]* .
# Copy-Item E:\Works\Trade\Reports_Equity2\output\$trade_year\$trade_date\0[6-7]* .
7z a C:\Users\Administrator\Desktop\$trade_date-·ç¿Ø-ÖÜâùÉ¼ *.xlsx
Remove-Item *.xlsx
pause