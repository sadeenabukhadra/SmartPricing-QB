# lastprice.ps1 - "Last sold price per customer" tool for QuickBooks Desktop (Enterprise 11.0)
#
# What it does:
#   1. READS Invoices + Sales Receipts (from -FromDate to today), Items, Customers, Price Levels.
#   2. Computes, for every customer + item, the LAST unit price BEFORE discount (the Rate on the line)
#      and the discount % that was given (Discount line(s) spread over the lines). -NetPrice = old net mode.
#   3. DRY-RUN (default): writes a plan CSV to Desktop\LastPrice\ and changes NOTHING in QuickBooks.
#   4. Only with -Apply (and -AllowFile = exact path of the open company file, and typing YES):
#      creates/updates one Per-Item price level "LP <customer>" per active customer and links it
#      to the customer (only if the customer's Price Level field is empty or already an "LP " level).
#
# Run with 32-bit PowerShell (QuickBooks open on the TEST company file, logged in as Admin):
#   Dry run :  C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File lastprice.ps1
#   Write   :  ... -File lastprice.ps1 -Apply -AllowFile "C:\Users\Lenovo\Desktop\petragate sadeen.qbw" [-OnlyCustomer "Prime Cinema / Irbid"]

param(
    [switch]$Apply,
    [string]$AllowFile = "",
    [string]$OnlyCustomer = "",
    [string]$FromDate = "2021-01-01",
    [int]$ActiveDays = 365,
    [int]$MaxLevels = 99,      # QuickBooks 2011 refused the 100th price level (error 3180)
    [int]$Rotating = 12,       # "LPX nn" levels kept free for prepare.ps1 (smaller customers, on demand)
    [string[]]$ExcludeCustomers = @("Cash Van", "Cash Sales", "Samples Marketing"),
    [string]$AppName = "Last Price Tool",
    [switch]$NetPrice          # old mode: price AFTER discount. Default now: price BEFORE discount (the line Rate);
                               # the last discount % is listed separately (DiscountPct) for the invoice Memo.
)

$ErrorActionPreference = "Stop"
$inv = [Globalization.CultureInfo]::InvariantCulture
$stamp = Get-Date -Format "yyyyMMdd-HHmm"
$outDir = Join-Path ([Environment]::GetFolderPath("Desktop")) "LastPrice"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$logFile = Join-Path $outDir "run-$stamp.txt"
function Log($t) { $t | Out-File -FilePath $logFile -Append -Encoding UTF8; Write-Host $t }
function Esc($s) { [Security.SecurityElement]::Escape([string]$s) }
function T($node, $path) { if (-not $node) { return $null }; $n = $node.SelectSingleNode($path); if ($n) { return $n.InnerText } else { return $null } }
function Num($s) { if ([string]::IsNullOrEmpty($s)) { return $null }; return [double]::Parse($s, $inv) }
function Fmt($d) { return ([math]::Round($d, 3)).ToString("0.###", $inv) }

if ([Environment]::Is64BitProcess) { Write-Host "ERROR: run with 32-bit PowerShell (C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe)" -ForegroundColor Red; exit 1 }

Log "=== Last Price tool  $(Get-Date -Format 'yyyy-MM-dd HH:mm')  mode: $(if ($Apply) {'APPLY (write)'} else {'DRY-RUN (no changes)'}) ==="
Log "Price: $(if ($NetPrice) {'NET (after discount)'} else {'BEFORE discount (discount % goes in the Memo)'})"

$rp = New-Object -ComObject QBXMLRP2.RequestProcessor
$rp.OpenConnection2("", $AppName, 1)
$ticket = $rp.BeginSession("", 2)
$V = "10.0"

function Send($body) {
    $x = "<?xml version=`"1.0`"?><?qbxml version=`"$V`"?><QBXML><QBXMLMsgsRq onError=`"continueOnError`">$body</QBXMLMsgsRq></QBXML>"
    [xml]$r = $rp.ProcessRequest($ticket, $x)
    return $r.QBXML.QBXMLMsgsRs.FirstChild
}
function CheckRs($rs, $what) {
    if (-not $rs) { throw "$what failed: empty response" }
    if ($rs.statusSeverity -eq "Error") { throw "$what failed: $($rs.statusCode) $($rs.statusMessage)" }
}

try {
    $companyFile = $rp.GetCurrentCompanyFileName($ticket)
    Log "Company file: $companyFile"

    # ---------- Items: classify by type ----------
    $rs = Send "<ItemQueryRq><ActiveStatus>All</ActiveStatus></ItemQueryRq>"
    CheckRs $rs "ItemQuery"
    $items = @{}
    $priceable = @("ItemInventoryRet", "ItemInventoryAssemblyRet", "ItemNonInventoryRet", "ItemServiceRet")
    foreach ($n in $rs.ChildNodes) {
        if ($n.NodeType -ne "Element") { continue }
        $items[(T $n "ListID")] = [pscustomobject]@{ Type = $n.LocalName; Name = (T $n "FullName"); Active = ((T $n "IsActive") -eq "true") }
    }
    Log "Items: $($items.Count)"

    # ---------- Customers ----------
    $rs = Send "<CustomerQueryRq><ActiveStatus>All</ActiveStatus></CustomerQueryRq>"
    CheckRs $rs "CustomerQuery"
    $custs = @{}
    foreach ($n in $rs.SelectNodes("CustomerRet")) {
        $custs[(T $n "ListID")] = [pscustomobject]@{
            ListID = (T $n "ListID"); Name = (T $n "FullName"); Active = ((T $n "IsActive") -eq "true")
            EditSeq = (T $n "EditSequence"); PLId = (T $n "PriceLevelRef/ListID"); PLName = (T $n "PriceLevelRef/FullName")
            CurId = (T $n "CurrencyRef/ListID"); CurName = (T $n "CurrencyRef/FullName")
        }
    }
    Log "Customers: $($custs.Count)"

    # ---------- Price levels ----------
    $rs = Send "<PriceLevelQueryRq><ActiveStatus>All</ActiveStatus></PriceLevelQueryRq>"
    CheckRs $rs "PriceLevelQuery"
    $levels = @{}; $levelByName = @{}
    foreach ($n in $rs.SelectNodes("PriceLevelRet")) {
        $its = @{}
        foreach ($pi in $n.SelectNodes("PriceLevelPerItemRet")) {
            $p = T $pi "CustomPrice"; if ($p) { $its[(T $pi "ItemRef/ListID")] = (Num $p) }
        }
        $lv = [pscustomobject]@{ ListID = (T $n "ListID"); Name = (T $n "Name"); EditSeq = (T $n "EditSequence"); Items = $its }
        $levels[$lv.ListID] = $lv; $levelByName[$lv.Name.ToLower()] = $lv
    }
    Log "Price levels in file: $($levels.Count)"

    # ---------- Sales documents, month by month ----------
    $last = @{}   # key custId|itemId -> record
    $docCount = 0; $lineCount = 0
    $start = [datetime]::ParseExact($FromDate, "yyyy-MM-dd", $inv)
    $today = (Get-Date).Date
    $custLastSale = @{}
    $custValue = @{}   # net sales value in the last $ActiveDays days (used to rank customers)
    $cutoff = $today.AddDays(-$ActiveDays).ToString("yyyy-MM-dd")

    function Process-Doc($doc, $docType, $lineTag, $groupTag) {
        $custId = T $doc "CustomerRef/ListID"
        if (-not $custId) { return }
        $date = T $doc "TxnDate"; $created = T $doc "TimeCreated"; $ref = T $doc "RefNumber"
        # flatten lines in order (group components included)
        $lines = New-Object System.Collections.ArrayList
        foreach ($c in $doc.ChildNodes) {
            if ($c.LocalName -eq $lineTag) { [void]$lines.Add($c) }
            elseif ($c.LocalName -eq $groupTag) { foreach ($g in $c.SelectNodes($lineTag)) { [void]$lines.Add($g) } }
        }
        $L = @()
        $i = 0
        foreach ($ln in $lines) {
            $itemId = T $ln "ItemRef/ListID"
            $type = if ($itemId -and $items.ContainsKey($itemId)) { $items[$itemId].Type } else { "None" }
            $amt = Num (T $ln "Amount"); if ($amt -eq $null) { $amt = 0 }
            $kind = switch ($type) { "ItemDiscountRet" { "disc" } "ItemSubtotalRet" { "sub" } default { if ($priceable -contains $type) { "sale" } else { "other" } } }
            $L += [pscustomobject]@{ Idx = $i; Kind = $kind; ItemId = $itemId; Qty = (Num (T $ln "Quantity")); UM = (T $ln "UnitOfMeasure"); Rate = (Num (T $ln "Rate")); RatePct = (Num (T $ln "RatePercent")); Amt = $amt; Net = $amt; Pct = 0.0; Cover = @() }
            $i++
        }
        # Discount rules (checked against real invoices 3620, 416, 1168, 5512, 8584):
        #  - % discount item (RatePercent set): QuickBooks rule = applies to the line above, or to the
        #    lines covered by a Sub Total directly above it.
        #  - fixed-amount discount (staff type the amount): it is a % of the invoice total.
        #    One such line -> ALL lines of the document, or only the lines above it, whichever gives a round %.
        #    Several -> each one covers the lines since the previous discount line (or doc start).
        $nFixed = @($L | Where-Object { $_.Kind -eq "disc" -and $_.RatePct -eq $null -and $_.Amt -lt 0 }).Count
        $lastSub = -1; $lastDisc = -1
        for ($k = 0; $k -lt $L.Count; $k++) {
            $x = $L[$k]
            if ($x.Kind -eq "sub") { if ($k -gt 0) { $x.Cover = @($L[($lastSub + 1)..($k - 1)] | Where-Object { $_.Kind -ne "sub" -and $_.Kind -ne "disc" }) }; $lastSub = $k; continue }
            if ($x.Kind -ne "disc") { continue }
            $cover = @()
            if ($x.RatePct -ne $null) {
                if ($k -gt 0) { $prev = $L[$k - 1]; if ($prev.Kind -eq "sub" -or $prev.Kind -eq "disc") { $cover = @($prev.Cover) } else { $cover = @($prev) } }
            } elseif ($nFixed -eq 1) {
                # whole invoice, or only the lines above it? pick the one that gives a "round" % (e.g. 15%, 12.5%)
                $all = @($L | Where-Object { $_.Kind -ne "sub" -and $_.Kind -ne "disc" -and $_.Net -gt 0 })
                $above = @($L[0..($k - 1)] | Where-Object { $_.Kind -ne "sub" -and $_.Kind -ne "disc" -and $_.Net -gt 0 })
                $cover = $all
                if ($k -gt 0 -and $above.Count -gt 0 -and $above.Count -lt $all.Count) {
                    $sA = 0.0; foreach ($c in $all) { $sA += $c.Net }; $sB = 0.0; foreach ($c in $above) { $sB += $c.Net }
                    $pA = - $x.Amt / $sA * 200; $pB = - $x.Amt / $sB * 200
                    if ([math]::Abs($pB - [math]::Round($pB)) -le [math]::Abs($pA - [math]::Round($pA))) { $cover = $above }
                }
            } elseif ($k -gt 0) {
                $cover = @($L[($lastDisc + 1)..($k - 1)] | Where-Object { $_.Kind -ne "sub" -and $_.Kind -ne "disc" })
            }
            $lastDisc = $k
            $cover = @($cover | Where-Object { $_.Net -gt 0 })
            $x.Cover = $cover
            $base = 0.0; foreach ($c in $cover) { $base += $c.Net }
            if ($base -le 0 -or $x.Amt -ge 0) { continue }
            $pct = - $x.Amt / $base
            # staff round the discount amount: 14.96% / 14.999% -> 15%, 12.47% -> 12.5%
            $r2 = [math]::Round($pct * 200) / 200; if ([math]::Abs($pct - $r2) -lt 0.0006) { $pct = $r2 }
            if ($pct -ge 1) { continue }
            foreach ($c in $cover) { $c.Net = $c.Net * (1 - $pct); $c.Pct = 1 - (1 - $c.Pct) * (1 - $pct) }
        }
        if (-not $custLastSale.ContainsKey($custId) -or $custLastSale[$custId] -lt $date) { $custLastSale[$custId] = $date }
        foreach ($x in $L) {
            if ($x.Kind -ne "sale" -or -not $x.Qty -or $x.Qty -le 0 -or $x.Net -le 0) { continue }
            $script:lineCount++
            if ($date -ge $cutoff) { $custValue[$custId] = [double]$custValue[$custId] + $x.Net }
            $netUnit = if ($x.Rate -ne $null -and $x.Rate -gt 0) { [math]::Round($x.Rate * (1 - $x.Pct), 3) } else { [math]::Round($x.Net / $x.Qty, 3) }
            $grossUnit = if ($x.Rate -ne $null -and $x.Rate -gt 0) { [math]::Round($x.Rate, 3) } elseif ($x.Pct -lt 1) { [math]::Round($x.Net / $x.Qty / (1 - $x.Pct), 3) } else { $netUnit }
            $unit = if ($NetPrice) { $netUnit } else { $grossUnit }
            $key = "$custId|$($x.ItemId)"
            $sortKey = "$date|$created|{0:D4}" -f $x.Idx
            $r = $last[$key]
            if (-not $r) { $r = [pscustomobject]@{ CustId = $custId; ItemId = $x.ItemId; Sort = ""; Price = 0; NetUnit = 0; Rate = $null; Pct = 0; Date = ""; DocType = ""; Ref = ""; UM = ""; Count = 0; Prev = $null; UMs = @{} }; $last[$key] = $r }
            $r.Count++
            if ($x.UM) { $r.UMs[$x.UM] = 1 }
            if ($sortKey -gt $r.Sort) {
                if ($r.Sort -ne "" -and $r.Price -ne $unit) { $r.Prev = $r.Price }
                $r.Sort = $sortKey; $r.Price = $unit; $r.NetUnit = $netUnit; $r.Rate = $x.Rate; $r.Pct = $x.Pct; $r.Date = $date; $r.DocType = $docType; $r.Ref = $ref; $r.UM = $x.UM
            }
        }
    }

    $m = Get-Date -Year $start.Year -Month $start.Month -Day 1
    while ($m -le $today) {
        $from = $m.ToString("yyyy-MM-dd"); $to = $m.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
        foreach ($t in @(@("Invoice", "InvoiceLineRet", "InvoiceLineGroupRet"), @("SalesReceipt", "SalesReceiptLineRet", "SalesReceiptLineGroupRet"))) {
            $rs = Send "<$($t[0])QueryRq><TxnDateRangeFilter><FromTxnDate>$from</FromTxnDate><ToTxnDate>$to</ToTxnDate></TxnDateRangeFilter><IncludeLineItems>true</IncludeLineItems></$($t[0])QueryRq>"
            CheckRs $rs "$($t[0])Query $from"
            foreach ($d in $rs.SelectNodes("$($t[0])Ret")) { Process-Doc $d $t[0] $t[1] $t[2]; $docCount++ }
        }
        Write-Host ("  read {0}  docs so far: {1}" -f $m.ToString("yyyy-MM"), $docCount)
        $m = $m.AddMonths(1)
    }
    Log "Documents read: $docCount   sale lines used: $lineCount   customer+item pairs: $($last.Count)"

    # ---------- Build plan ----------
    $byCust = $last.Values | Group-Object CustId
    # rank eligible customers by sales value; only the top ones get a price level (QuickBooks list limit)
    $nonLP = @($levels.Values | Where-Object { -not $_.Name.StartsWith("LP ") -and -not $_.Name.StartsWith("LPX ") }).Count
    $slots = $MaxLevels - $nonLP - $Rotating
    $elig = foreach ($g in $byCust) {
        $c = $custs[$g.Name]
        if (-not $c -or ($ExcludeCustomers -contains $c.Name) -or -not $c.Active -or $custLastSale[$c.ListID] -lt $cutoff) { continue }
        if ($c.PLName -and -not $c.PLName.StartsWith("LP ") -and -not $c.PLName.StartsWith("LPX ")) { continue }
        [pscustomobject]@{ Id = $c.ListID; V = [double]$custValue[$c.ListID] }
    }
    $topSet = @{}; $rank = 0
    foreach ($e in ($elig | Sort-Object V -Descending)) { $rank++; if ($rank -le $slots) { $topSet[$e.Id] = $rank } }
    Log "Fixed price level slots: $slots (limit $MaxLevels, other levels $nonLP, rotating LPX $Rotating). Eligible customers: $(@($elig).Count)"
    $plan = New-Object System.Collections.ArrayList
    $custPlan = New-Object System.Collections.ArrayList
    $usedNames = @{}; foreach ($k in $levelByName.Keys) { $usedNames[$k] = 1 }

    foreach ($g in $byCust) {
        $c = $custs[$g.Name]; if (-not $c) { continue }
        if ($OnlyCustomer -and $c.Name -ne $OnlyCustomer) { continue }
        $status = "OK"
        if ($ExcludeCustomers -contains $c.Name) { $status = "SKIP: excluded (generic customer)" }
        elseif (-not $c.Active) { $status = "SKIP: customer inactive" }
        elseif ($c.PLName -and -not $c.PLName.StartsWith("LP ") -and -not $c.PLName.StartsWith("LPX ")) { $status = "SKIP: already has price level '$($c.PLName)'" }
        elseif ($custLastSale[$c.ListID] -lt $cutoff) { $status = "LIST ONLY: no sale in last $ActiveDays days (use prepare)" }
        elseif (-not $topSet.ContainsKey($c.ListID)) { $status = "LIST ONLY: not in top $slots" }

        # target level
        $lv = $null; $lvName = ""
        if ($c.PLId -and $levels.ContainsKey($c.PLId) -and $c.PLName.StartsWith("LP ")) { $lv = $levels[$c.PLId]; $lvName = $lv.Name }
        else {
            $base = ("LP " + ($c.Name -replace ":", "-")).Trim(); if ($base.Length -gt 31) { $base = $base.Substring(0, 31).Trim() }
            $lvName = $base; $n = 2
            while ($usedNames.ContainsKey($lvName.ToLower())) { $s = " ~$n"; $lvName = $base.Substring(0, [math]::Min($base.Length, 31 - $s.Length)).Trim() + $s; $n++ }
            if ($status -eq "OK") { $usedNames[$lvName.ToLower()] = 1 }
        }

        $changes = 0
        foreach ($r in ($g.Group | Sort-Object { $items[$_.ItemId].Name })) {
            $it = $items[$r.ItemId]
            $cur = $null; if ($lv -and $lv.Items.ContainsKey($r.ItemId)) { $cur = $lv.Items[$r.ItemId] }
            $act = if ($status -ne "OK") { "-" } elseif (-not $it.Active) { "SKIP item inactive" } elseif ($cur -eq $null) { "NEW" } elseif ([math]::Abs($cur - $r.Price) -lt 0.0005) { "SAME" } else { "CHANGE" }
            if ($act -eq "NEW" -or $act -eq "CHANGE") { $changes++ }
            [void]$plan.Add([pscustomobject]@{
                Customer = $c.Name; CustomerStatus = $status; PriceLevel = $lvName; Action = $act; Item = $it.Name; UM = $r.UM
                MixedUM = $(if ($r.UMs.Count -gt 1) { "YES" } else { "" }); LastPrice = (Fmt $r.Price); LastNetPrice = (Fmt $r.NetUnit); LineRate = $(if ($r.Rate -ne $null) { Fmt $r.Rate } else { "" })
                DiscountPct = (Fmt ($r.Pct * 100)); Date = $r.Date; DocType = $r.DocType; RefNumber = $r.Ref; PrevPrice = $(if ($r.Prev -ne $null) { Fmt $r.Prev } else { "" })
                TimesSold = $r.Count; CurrentLevelPrice = $(if ($cur -ne $null) { Fmt $cur } else { "" }); Currency = $c.CurName
            })
        }
        [void]$custPlan.Add([pscustomobject]@{ Cust = $c; Status = $status; Level = $lv; LevelName = $lvName; Changes = $changes; Recs = $g.Group })
    }

    $planFile = Join-Path $outDir "plan-$stamp.csv"
    $plan | Export-Csv -Path $planFile -NoTypeInformation -Encoding UTF8
    $ok = @($custPlan | Where-Object { $_.Status -eq "OK" })
    $newLv = @($ok | Where-Object { -not $_.Level }).Count
    Log "Plan file: $planFile"
    Log ("Customers in plan: {0}   to process: {1}   new price levels: {2}   price entries NEW/CHANGE: {3}" -f $custPlan.Count, $ok.Count, $newLv, (@($plan | Where-Object { $_.Action -eq "NEW" -or $_.Action -eq "CHANGE" }).Count))
    foreach ($s in ($custPlan | Group-Object Status)) { Log ("  {0,5}  {1}" -f $s.Count, $s.Name) }
    # LP levels to release: linked to a customer that is no longer in the top list, or linked to nobody
    $keep = @{}; foreach ($cp in $ok) { if ($cp.Level) { $keep[$cp.Level.ListID] = 1 } }
    $release = @($levels.Values | Where-Object { $_.Name.StartsWith("LP ") -and -not $keep.ContainsKey($_.ListID) })
    $releaseCust = @($custs.Values | Where-Object { $_.PLId -and $_.PLName.StartsWith("LP ") -and -not $keep.ContainsKey($_.PLId) })
    if ($OnlyCustomer) { $release = @(); $releaseCust = @() }   # single-customer test: never touch other customers
    Log "LP price levels to release (customer left the top list / unused): $($release.Count)   customers to unlink: $($releaseCust.Count)"
    $lpxNow = @($levels.Values | Where-Object { $_.Name.StartsWith("LPX ") }).Count
    $finalTotal = $levels.Count - $release.Count + $newLv + [math]::Max(0, $Rotating - $lpxNow)
    if ($finalTotal -gt $MaxLevels) { Log "WARNING: total price levels would be $finalTotal (> $MaxLevels)." }

    # lookup list for the accountant (all active customers, incl. those without a price level)
    $lookup = $plan | Where-Object { $_.CustomerStatus -eq "OK" -or $_.CustomerStatus -like "LIST ONLY*" } |
        Select-Object Customer, Item, UM, LastPrice, LastNetPrice, LineRate, DiscountPct, Date, DocType, RefNumber, PrevPrice, TimesSold, @{ n = "AutoInInvoice"; e = { if ($_.CustomerStatus -eq "OK") { "YES" } else { "" } } }
    # cache for prepare.ps1 (IDs, one row per customer+item) - only on full runs
    if (-not $OnlyCustomer) {
    $cache = foreach ($cp in $custPlan) {
        if ($cp.Status -ne "OK" -and $cp.Status -notlike "LIST ONLY*") { continue }
        foreach ($r in $cp.Recs) { if ($items[$r.ItemId].Active) { [pscustomobject]@{ CustId = $cp.Cust.ListID; Customer = $cp.Cust.Name; CurId = $cp.Cust.CurId; Fixed = $(if ($cp.Status -eq "OK") { "YES" } else { "" }); ItemId = $r.ItemId; Item = $items[$r.ItemId].Name; Price = (Fmt $r.Price); DiscPct = (Fmt ($r.Pct * 100)); Date = $r.Date } } }
    }
    try { $cache | Export-Csv -Path (Join-Path $outDir "cache-lastprices.csv") -NoTypeInformation -Encoding UTF8 } catch { Log "WARNING: could not write cache-lastprices.csv" }
    $lookupFile = Join-Path $outDir "last-prices.csv"
    try { $lookup | Export-Csv -Path $lookupFile -NoTypeInformation -Encoding UTF8; Log "Lookup list: $lookupFile  ($(@($lookup).Count) rows)" } catch { Log "WARNING: could not write $lookupFile (is it open in Excel?)" }
    }

    if (-not $Apply) { Log "=== DRY-RUN DONE. Nothing was changed in QuickBooks. ==="; return }

    # ---------- APPLY ----------
    if ($AllowFile -eq "" -or $companyFile -ne $AllowFile) { Log "STOP: -AllowFile does not match the open company file. Nothing written."; return }
    if ($finalTotal -gt $MaxLevels) { Log "STOP: price level limit. Nothing written."; return }
    $ans = Read-Host "Release $($release.Count) levels and write $($ok.Count) customers' price levels into '$companyFile'? Type YES to continue"
    if ($ans -ne "YES") { Log "Cancelled by user. Nothing written."; return }

    # backup of current LP levels + customer links (for undo)
    $bk = foreach ($lv in $levels.Values) { if ($lv.Name.StartsWith("LP ")) { foreach ($k in $lv.Items.Keys) { [pscustomobject]@{ PriceLevel = $lv.Name; Item = $items[$k].Name; Price = (Fmt $lv.Items[$k]) } } } }
    $bk | Export-Csv -Path (Join-Path $outDir "backup-levels-$stamp.csv") -NoTypeInformation -Encoding UTF8
    $custs.Values | Where-Object { $_.PLName } | Select-Object Name, PLName | Export-Csv -Path (Join-Path $outDir "backup-customer-links-$stamp.csv") -NoTypeInformation -Encoding UTF8

    $done = 0; $fail = 0
    # 1) release levels of customers that left the top list (frees slots first)
    foreach ($c in $releaseCust) {
        $rs2 = Send "<CustomerModRq><CustomerMod><ListID>$($c.ListID)</ListID><EditSequence>$($c.EditSeq)</EditSequence><PriceLevelRef><FullName></FullName></PriceLevelRef></CustomerMod></CustomerModRq>"
        if ([int]$rs2.statusCode -ne 0) { Log "WARN unlink customer '$($c.Name)': $($rs2.statusCode) $($rs2.statusMessage)" } else { Log "OK   unlinked customer '$($c.Name)' from '$($c.PLName)'" }
    }
    foreach ($lv in $release) {
        $rs2 = Send "<ListDelRq><ListDelType>PriceLevel</ListDelType><ListID>$($lv.ListID)</ListID></ListDelRq>"
        if ([int]$rs2.statusCode -ne 0) { Log "FAIL delete level '$($lv.Name)': $($rs2.statusCode) $($rs2.statusMessage)"; $fail++ } else { Log "OK   deleted level '$($lv.Name)'" }
    }
    # 2) create / update levels of the top customers
    foreach ($cp in $ok) {
        $c = $cp.Cust
        if ($cp.Level -and $cp.Changes -eq 0 -and $c.PLId -eq $cp.Level.ListID) { continue }
        # full item list = existing prices + new last prices (never removes an item)
        $prices = @{}
        if ($cp.Level) { foreach ($k in $cp.Level.Items.Keys) { $prices[$k] = $cp.Level.Items[$k] } }
        foreach ($r in $cp.Recs) { if ($items[$r.ItemId].Active) { $prices[$r.ItemId] = $r.Price } }
        $per = ($prices.Keys | ForEach-Object { "<PriceLevelPerItem><ItemRef><ListID>$_</ListID></ItemRef><CustomPrice>$(Fmt $prices[$_])</CustomPrice></PriceLevelPerItem>" }) -join ""
        $cur = if ($c.CurId) { "<CurrencyRef><ListID>$($c.CurId)</ListID></CurrencyRef>" } else { "" }
        if ($cp.Level) {
            $rs = Send "<PriceLevelModRq><PriceLevelMod><ListID>$($cp.Level.ListID)</ListID><EditSequence>$($cp.Level.EditSeq)</EditSequence>$per</PriceLevelMod></PriceLevelModRq>"
        } else {
            $rs = Send "<PriceLevelAddRq><PriceLevelAdd><Name>$(Esc $cp.LevelName)</Name><IsActive>true</IsActive>$per$cur</PriceLevelAdd></PriceLevelAddRq>"
        }
        if ([int]$rs.statusCode -ne 0) { Log "FAIL level '$($cp.LevelName)' ($($c.Name)): $($rs.statusCode) $($rs.statusMessage)"; $fail++; continue }
        $plId = T $rs "PriceLevelRet/ListID"
        Log ("OK   level '{0}'  items: {1}  changed: {2}" -f $cp.LevelName, $prices.Count, $cp.Changes)
        if ($c.PLId -ne $plId) {
            $rs2 = Send "<CustomerModRq><CustomerMod><ListID>$($c.ListID)</ListID><EditSequence>$($c.EditSeq)</EditSequence><PriceLevelRef><ListID>$plId</ListID></PriceLevelRef></CustomerMod></CustomerModRq>"
            if ([int]$rs2.statusCode -ne 0) { Log "FAIL link customer '$($c.Name)': $($rs2.statusCode) $($rs2.statusMessage)"; $fail++; continue }
            Log "OK   customer '$($c.Name)' -> '$($cp.LevelName)'"
        }
        $done++
    }
    Log "=== APPLY DONE. customers updated: $done   failures: $fail ==="
}
catch { Log "ERROR: $($_.Exception.Message)" }
finally {
    $rp.EndSession($ticket)
    $rp.CloseConnection()
}
