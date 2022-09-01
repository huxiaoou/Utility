# this script MUST be opened and edited in GB18030 format since some Chinese characters are included

$trade_date = Read-Host -Prompt "Please input the report date to pack, format = [YYYYMMDD]`nOr you can hit ENTER key directly to use the default value, which is today"
if (-not($trade_date)) {
    $trade_date = Get-Date -Format yyyyMMdd
}
$trade_year = $trade_date.substring(0, 4)
$sep = "=" * 120

# for Linchen
Write-Host $sep
python gen_text_report.V4.py $trade_date

# for Qiuyue
Write-Host $sep
7z a C:\Users\Administrator\Desktop\$trade_date-����-��h \Works\Trade\Reports\output\$trade_year\$trade_date\08*
# 7z a C:\Users\Administrator\Desktop\$trade_date-����-��h \Works\Trade\Reports_Equity2\output\$trade_year\$trade_date\07*

# for Qianyelan
Write-Host $sep
7z a C:\Users\Administrator\Desktop\$trade_date-ǮҶ�� \Works\Trade\Reports\output\$trade_year\$trade_date\06*
# 7z a C:\Users\Administrator\Desktop\$trade_date-ǮҶ�� \Works\Trade\Reports_Merge\output\$trade_year\$trade_date\04*

# for Fanyabin
Write-Host $sep
7z a C:\Users\Administrator\Desktop\$trade_date-���Ǳ� \Works\Trade\Reports\output\$trade_year\$trade_date\04*
7z a C:\Users\Administrator\Desktop\$trade_date-���Ǳ� \Works\Trade\Reports\output\$trade_year\$trade_date\05*
7z a C:\Users\Administrator\Desktop\$trade_date-���Ǳ� \Works\Trade\Reports\output\$trade_year\$trade_date\06*
7z a C:\Users\Administrator\Desktop\$trade_date-���Ǳ� \Works\Trade\Reports\intermediary\��Ͼ�ֵ.xlsx
7z a C:\Users\Administrator\Desktop\$trade_date-���Ǳ� \Works\Trade\Reports\intermediary\*.png

## for daily report
#Write-Host $sep
#Copy-Item \Works\Trade\Reports\templates\clean\�ձ�_YYYYMMDD.docx C:\Users\Administrator\Desktop\�ձ�_$trade_date.docx
#Start-Process C:\Users\Administrator\Desktop\�ձ�_$trade_date.docx
#Start-Process \Works\Trade\Reports\output\$trade_year\$trade_date\04_����Ʒ�ֲ������ϸ��_������Ʒ_$trade_date.xlsx
#Start-Process \Works\Trade\Reports_Merge\output\$trade_year\$trade_date\04_�ֲ������ϸ��_�����й���Ŀ_$trade_date.xlsx

Write-Host $sep
Pause
