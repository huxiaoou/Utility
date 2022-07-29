# encoding = gb18030
Write-Host "Update global/by_portfolio position"
$d = Read-Host -Prompt "Please input the report date to update, format = [YYYYMMDD]`nOr you can hit ENTER key directly to use the default value, which is today"
if (-not($report_date)) {
    $d = Get-Date -Format yyyyMMdd
}
$y = $d.substring(0, 4)
Set-Location \Works\Monitor\
python \Works\Monitor\convert_04_format.py $d
Start-Process excel \Works\Monitor\monitor.V3.xlsx
Start-Process excel \Works\Monitor\pos.copy_and_paste_to_monitor.sheet_portfolio.csv
Start-Process excel \Works\TradeClearing\data\pos.by_portfolio\$y\pos.by_portfolio.$d.csv
Start-Process excel \Works\Trade\Reports\intermediary\×éºÏ¾»Öµ.xlsx;
