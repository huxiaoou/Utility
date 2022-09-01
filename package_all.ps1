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
7z a C:\Users\Administrator\Desktop\$trade_date-财务-邱玥 \Works\Trade\Reports\output\$trade_year\$trade_date\08*
# 7z a C:\Users\Administrator\Desktop\$trade_date-财务-邱玥 \Works\Trade\Reports_Equity2\output\$trade_year\$trade_date\07*

# for Qianyelan
Write-Host $sep
7z a C:\Users\Administrator\Desktop\$trade_date-钱叶兰 \Works\Trade\Reports\output\$trade_year\$trade_date\06*
# 7z a C:\Users\Administrator\Desktop\$trade_date-钱叶兰 \Works\Trade\Reports_Merge\output\$trade_year\$trade_date\04*

# for Fanyabin
Write-Host $sep
7z a C:\Users\Administrator\Desktop\$trade_date-樊亚彬 \Works\Trade\Reports\output\$trade_year\$trade_date\04*
7z a C:\Users\Administrator\Desktop\$trade_date-樊亚彬 \Works\Trade\Reports\output\$trade_year\$trade_date\05*
7z a C:\Users\Administrator\Desktop\$trade_date-樊亚彬 \Works\Trade\Reports\output\$trade_year\$trade_date\06*
7z a C:\Users\Administrator\Desktop\$trade_date-樊亚彬 \Works\Trade\Reports\intermediary\组合净值.xlsx
7z a C:\Users\Administrator\Desktop\$trade_date-樊亚彬 \Works\Trade\Reports\intermediary\*.png

## for daily report
#Write-Host $sep
#Copy-Item \Works\Trade\Reports\templates\clean\日报_YYYYMMDD.docx C:\Users\Administrator\Desktop\日报_$trade_date.docx
#Start-Process C:\Users\Administrator\Desktop\日报_$trade_date.docx
#Start-Process \Works\Trade\Reports\output\$trade_year\$trade_date\04_衍生品持仓情况明细表_大宗商品_$trade_date.xlsx
#Start-Process \Works\Trade\Reports_Merge\output\$trade_year\$trade_date\04_持仓情况明细表_固收托管项目_$trade_date.xlsx

Write-Host $sep
Pause
