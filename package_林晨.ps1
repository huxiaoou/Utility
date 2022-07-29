$trade_date = Read-Host -Prompt 'Please input report date [format = YYYYMMDD]'
$trade_year = $trade_date.substring(0, 4)
python gen_text_report.V3.py $trade_date
Copy-Item E:\Works\Trade\*\output\$trade_year\$trade_date\0[3-5]_* .
7z a C:\Users\Administrator\Desktop\$trade_date *.xlsx
Remove-Item *.xlsx
pause
