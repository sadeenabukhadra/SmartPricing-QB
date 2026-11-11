# probe.ps1 - READ-ONLY connection test for the "Last Price" tool.
# It only READS from QuickBooks: version info, company file path, price levels, 3 recent invoices.
# It does NOT add, change or delete anything.
# Run with 32-bit PowerShell:
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File probe.ps1

$ErrorActionPreference = "Stop"
$out = Join-Path ([Environment]::GetFolderPath("Desktop")) "lastprice-probe.txt"
function Log($t) { $t | Out-File -FilePath $out -Append -Encoding UTF8; Write-Host $t }
if (Test-Path $out) { Remove-Item $out }

Log "=== Last Price probe  $(Get-Date -Format 'yyyy-MM-dd HH:mm') ==="
Log "PowerShell 64-bit process: $([Environment]::Is64BitProcess)  (must be False)"

$rp = New-Object -ComObject QBXMLRP2.RequestProcessor
$rp.OpenConnection2("", "Last Price Tool (read-only test)", 1)
$ticket = $rp.BeginSession("", 2)   # 2 = use the company file that is open now
try {
    Log "Company file: $($rp.GetCurrentCompanyFileName($ticket))"

    $host_rq = '<?xml version="1.0"?><?qbxml version="2.0"?><QBXML><QBXMLMsgsRq onError="stopOnError"><HostQueryRq/></QBXMLMsgsRq></QBXML>'
    [xml]$h = $rp.ProcessRequest($ticket, $host_rq)
    $hr = $h.QBXML.QBXMLMsgsRs.HostQueryRs.HostRet
    Log "Product: $($hr.ProductName)  Major: $($hr.MajorVersion)  Minor: $($hr.MinorVersion)"
    $vers = @($hr.SupportedQBXMLVersion)
    Log "Supported qbXML versions: $($vers -join ', ')"
    $v = ($vers | Sort-Object { [double]$_ } | Select-Object -Last 1)
    Log "Using qbXML $v"

    $pl_rq = "<?xml version=`"1.0`"?><?qbxml version=`"$v`"?><QBXML><QBXMLMsgsRq onError=`"continueOnError`"><PriceLevelQueryRq/></QBXMLMsgsRq></QBXML>"
    [xml]$p = $rp.ProcessRequest($ticket, $pl_rq)
    $prs = $p.QBXML.QBXMLMsgsRs.PriceLevelQueryRs
    Log "PriceLevelQuery status: $($prs.statusCode) $($prs.statusMessage)"
    foreach ($pl in @($prs.PriceLevelRet)) { if ($pl) { Log ("  Price level: {0} | items: {1}" -f $pl.Name, @($pl.PriceLevelPerItemRet).Count) } }

    $inv_rq = "<?xml version=`"1.0`"?><?qbxml version=`"$v`"?><QBXML><QBXMLMsgsRq onError=`"continueOnError`"><InvoiceQueryRq><MaxReturned>3</MaxReturned><TxnDateRangeFilter><FromTxnDate>2026-10-01</FromTxnDate></TxnDateRangeFilter><IncludeLineItems>true</IncludeLineItems></InvoiceQueryRq></QBXMLMsgsRq></QBXML>"
    [xml]$i = $rp.ProcessRequest($ticket, $inv_rq)
    $irs = $i.QBXML.QBXMLMsgsRs.InvoiceQueryRs
    Log "InvoiceQuery status: $($irs.statusCode) $($irs.statusMessage)"
    foreach ($inv in @($irs.InvoiceRet)) {
        if (-not $inv) { continue }
        Log ("  Invoice {0} {1} {2}" -f $inv.RefNumber, $inv.TxnDate, $inv.CustomerRef.FullName)
        foreach ($ln in @($inv.InvoiceLineRet)) { if ($ln) { Log ("     {0} | qty {1} {2} | rate {3} | ratePct {4} | amt {5}" -f $ln.ItemRef.FullName, $ln.Quantity, $ln.UnitOfMeasure, $ln.Rate, $ln.RatePercent, $ln.Amount) } }
    }
    Log "=== DONE. Nothing was changed. ==="
}
finally {
    $rp.EndSession($ticket)
    $rp.CloseConnection()
}
