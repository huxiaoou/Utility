# encoding = gb18030
Write-Host "copy_trades_by_portfolio"
$d = Read-Host -Prompt "Please input the report date to of which trades are used, format = [YYYYMMDD]`nOr you can hit ENTER key directly to use the default value, which is today"
if (-not($d)) {
    $d = Get-Date -Format yyyyMMdd
}
$y = $d.substring(0, 4)
Copy-Item \Works\Trade\Reports\output\$y\$d\02_衍生品当日成交_大宗商品_$d.xlsx \Works\TradeClearing\data\trades.by_portfolio\$y\02_衍生品当日成交_大宗商品_$d.by_portfolio.xlsx
Start-Process excel \Works\TradeClearing\data\trades.by_portfolio\$y\02_衍生品当日成交_大宗商品_$d.by_portfolio.xlsx
