# this script MUST be opened and edited in GB18030 format since some Chinese characters are included

$trade_date = Read-Host -Prompt "Please input the report date to pack, format = [YYYYMMDD]`nOr you can hit ENTER key directly to use the default value, which is today"
if (-not($trade_date)) {
    $trade_date = Get-Date -Format yyyyMMdd
}
$trade_year = $trade_date.substring(0, 4)
$sep = "=" * 120

Write-Host $sep
python gen_text_report.V3.py $trade_date

# # for zengtongyi
# Write-Host $sep
# Copy-Item \Works\Trade\Reports\output\$trade_year\$trade_date\0[3-5]_* .
# Copy-Item \Works\Trade\Reports_Merge\output\$trade_year\$trade_date\0[3-5]_* .
# 7z a C:\Users\Administrator\Desktop\$trade_date *.xlsx
# Remove-Item *.xlsx

# for qiuyue
Write-Host $sep
Copy-Item \Works\Trade\Reports\output\$trade_year\$trade_date\08* .
Copy-Item \Works\Trade\Reports_Equity2\output\$trade_year\$trade_date\07* .
7z a C:\Users\Administrator\Desktop\$trade_date-����-��h *.xlsx
Remove-Item *.xlsx

# # for zhouyishan
# Write-Host $sep
# Copy-Item \Works\Trade\Reports\output\$trade_year\$trade_date\0[6]* .
# Copy-Item \Works\Trade\Reports_Equity2\output\$trade_year\$trade_date\0[6]* .
# 7z a C:\Users\Administrator\Desktop\$trade_date-����ɼ *.xlsx
# Remove-Item *.xlsx

Write-Host $sep
# for qianyelan
Copy-Item \Works\Trade\Reports\output\$trade_year\$trade_date\06* .
Copy-Item \Works\Trade\Reports_Merge\output\$trade_year\$trade_date\04* .
7z a C:\Users\Administrator\Desktop\$trade_date-ǮҶ�� *.xlsx
Remove-Item *.xlsx

Write-Host $sep
# for fanyabin
Copy-Item \Works\Trade\Reports\output\$trade_year\$trade_date\0[4-6]* .
Copy-Item \Works\Trade\Reports\intermediary\��Ͼ�ֵ.xlsx .
Copy-Item \Works\Trade\Reports\intermediary\*.png .
7z a C:\Users\Administrator\Desktop\$trade_date-���Ǳ� *.xlsx
7z a C:\Users\Administrator\Desktop\$trade_date-���Ǳ� *.png
Remove-Item *.xlsx
Remove-Item *.png

Write-Host $sep
# for daily report
Copy-Item \Works\Trade\Reports\templates\clean\�ձ�_YYYYMMDD.docx C:\Users\Administrator\Desktop\�ձ�_$trade_date.docx
Start-Process C:\Users\Administrator\Desktop\�ձ�_$trade_date.docx
Start-Process \Works\Trade\Reports\output\$trade_year\$trade_date\04_����Ʒ�ֲ������ϸ��_������Ʒ_$trade_date.xlsx
Start-Process \Works\Trade\Reports_Merge\output\$trade_year\$trade_date\04_�ֲ������ϸ��_�����й���Ŀ_$trade_date.xlsx
Pause
