# prepare.ps1 - "Prepare customer" button for customers WITHOUT a fixed price level.
# The accountant picks a customer -> the tool puts that customer's last prices into one of the
# rotating price levels ("LPX 01".."LPX 12", least recently used) and links it to the customer.
# Then a new invoice for that customer auto-fills the last prices (BEFORE discount).
# The window also shows the customer's last discount % and copies "Discount xx%" to the clipboard,
# so the accountant pastes it into the invoice Memo (Ctrl+V) and adds the Discount line.
# Needs: QuickBooks open on the company file, and Desktop\LastPrice\cache-lastprices.csv
# (written every night by lastprice.ps1).
# Run with 32-bit PowerShell:  C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -STA -ExecutionPolicy Bypass -File prepare.ps1

param(
    [int]$Rotating = 12,
    [string]$AppName = "Last Price Tool",
    [string]$AllowFile = ""          # optional: refuse to write if another company file is open
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$dir = Join-Path ([Environment]::GetFolderPath("Desktop")) "LastPrice"
$cacheFile = Join-Path $dir "cache-lastprices.csv"
$logFile = Join-Path $dir "prepare-log.txt"
function Log($t) { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $t" | Out-File -FilePath $logFile -Append -Encoding UTF8 }
function Esc($s) { [Security.SecurityElement]::Escape([string]$s) }
function T($node, $path) { if (-not $node) { return $null }; $n = $node.SelectSingleNode($path); if ($n) { return $n.InnerText } else { return $null } }

if (-not (Test-Path $cacheFile)) { [Windows.Forms.MessageBox]::Show("Missing $cacheFile`nRun the nightly tool (lastprice) first.", "Last Price") | Out-Null; exit 1 }
# Search helper: makes Arabic/English names comparable (no case, no tashkeel/tatweel,
# alef forms -> alef, ta marbuta -> ha, alef maqsura -> ya, hamza on waw/ya -> waw/ya, Arabic digits -> 0-9).
# Only \u escapes are used so the file works in any encoding.
$script:AR_ALEF = [string][char]0x0627; $script:AR_HA = [string][char]0x0647; $script:AR_YA = [string][char]0x064A; $script:AR_WAW = [string][char]0x0648
function Norm($s) {
    if (-not $s) { return "" }
    $s = ([string]$s).ToLowerInvariant()
    $s = $s -replace '[\u064B-\u065F\u0670\u0640]', ''
    $s = $s -replace '[\u0622\u0623\u0625\u0671]', $script:AR_ALEF
    $s = $s -replace '\u0629', $script:AR_HA
    $s = $s -replace '[\u0649\u0626]', $script:AR_YA
    $s = $s -replace '\u0624', $script:AR_WAW
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) {
        $c = [int]$ch
        if ($c -ge 0x0660 -and $c -le 0x0669) { [void]$sb.Append([char](48 + $c - 0x0660)) }
        elseif ($c -ge 0x06F0 -and $c -le 0x06F9) { [void]$sb.Append([char](48 + $c - 0x06F0)) }
        else { [void]$sb.Append($ch) }
    }
    return (($sb.ToString() -replace '[\s\-_.,/\\]+', ' ').Trim())
}

$cache = Import-Csv $cacheFile
$byCust = @{}; foreach ($r in $cache) { if (-not $byCust.ContainsKey($r.CustId)) { $byCust[$r.CustId] = New-Object System.Collections.ArrayList }; [void]$byCust[$r.CustId].Add($r) }
$names = $byCust.Keys | ForEach-Object {
    $r = $byCust[$_][0]
    # last discount = discount on the customer's most recent sale (if the cache has the column)
    $lastRow = $byCust[$_] | Sort-Object { $_.Date } | Select-Object -Last 1
    $d = if ($lastRow.PSObject.Properties["DiscPct"]) { [string]$lastRow.DiscPct } else { "" }
    [pscustomobject]@{ Id = $_; Name = $r.Customer; Fixed = $r.Fixed; CurId = $r.CurId; Disc = $d; DiscDate = $(if ($lastRow.PSObject.Properties["Date"]) { $lastRow.Date } else { "" }); Items = $byCust[$_].Count; Key = (Norm $r.Customer); Other = $(if ($r.PSObject.Properties["OtherLevel"]) { $r.OtherLevel } else { "" }) } } | Sort-Object Name
function DiscText($c) {
    if (-not $c.Disc -or $c.Disc -eq "0") { return "Last discount: none." }
    $memo = "Discount $($c.Disc)%"
    try { [Windows.Forms.Clipboard]::SetText($memo) } catch {}
    return "Last discount: $($c.Disc)% ($($c.DiscDate)). '$memo' copied - paste it in the Memo (Ctrl+V) and add the Discount line."
}

function Prepare-Customer($cust) {
    $rp = New-Object -ComObject QBXMLRP2.RequestProcessor
    $rp.OpenConnection2("", $AppName, 1)
    $ticket = $rp.BeginSession("", 2)
    function Send($body) {
        $x = "<?xml version=`"1.0`"?><?qbxml version=`"10.0`"?><QBXML><QBXMLMsgsRq onError=`"continueOnError`">$body</QBXMLMsgsRq></QBXML>"
        [xml]$r = $rp.ProcessRequest($ticket, $x); return $r.QBXML.QBXMLMsgsRs.FirstChild
    }
    try {
        $file = $rp.GetCurrentCompanyFileName($ticket)
        if ($AllowFile -and $file -ne $AllowFile) { return "STOP: open company file is '$file', not '$AllowFile'. Nothing changed." }

        # rotating levels
        $rs = Send "<PriceLevelQueryRq><ActiveStatus>All</ActiveStatus><NameFilter><MatchCriterion>StartsWith</MatchCriterion><Name>LPX </Name></NameFilter></PriceLevelQueryRq>"
        $lpx = @($rs.SelectNodes("PriceLevelRet") | ForEach-Object { [pscustomobject]@{ ListID = (T $_ "ListID"); Name = (T $_ "Name"); Modified = (T $_ "TimeModified") } })

        # all customers' links (to find who uses which LPX level)
        $rs = Send "<CustomerQueryRq><ActiveStatus>All</ActiveStatus><IncludeRetElement>ListID</IncludeRetElement><IncludeRetElement>FullName</IncludeRetElement><IncludeRetElement>EditSequence</IncludeRetElement><IncludeRetElement>PriceLevelRef</IncludeRetElement></CustomerQueryRq>"
        $all = @($rs.SelectNodes("CustomerRet") | ForEach-Object { [pscustomobject]@{ ListID = (T $_ "ListID"); Name = (T $_ "FullName"); EditSeq = (T $_ "EditSequence"); PLId = (T $_ "PriceLevelRef/ListID"); PLName = (T $_ "PriceLevelRef/FullName") } })
        $me = $all | Where-Object { $_.ListID -eq $cust.Id } | Select-Object -First 1
        if (-not $me) { return "Customer not found in QuickBooks." }
        if ($me.PLName -and $me.PLName.StartsWith("LP ")) { return "'$($me.Name)' already has a fixed price level ($($me.PLName)). Nothing to do - just open the invoice." }
        if ($me.PLName -and -not $me.PLName.StartsWith("LPX ")) { $ans = [Windows.Forms.MessageBox]::Show($f, "'$($me.Name)' has the price level '$($me.PLName)'.`nReplace it with this customer's last prices?", "Last Price", [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Question); if ($ans -ne [Windows.Forms.DialogResult]::Yes) { return "'$($me.Name)' keeps its price level ($($me.PLName)). Not changed." }; Log "replacing level '$($me.PLName)' on '$($me.Name)'"; $me.PLName = ""; $me.PLId = "" }

        # choose slot: own LPX -> new LPX (if fewer than $Rotating) -> least recently used
        $slot = $null; $slotName = ""
        if ($me.PLName) { $slot = $lpx | Where-Object { $_.ListID -eq $me.PLId } | Select-Object -First 1 }
        if (-not $slot -and $lpx.Count -lt $Rotating) {
            for ($n = 1; $n -le $Rotating; $n++) { $nm = "LPX {0:D2}" -f $n; if (-not ($lpx | Where-Object { $_.Name -eq $nm })) { $slotName = $nm; break } }
        }
        if (-not $slot -and -not $slotName) { $slot = $lpx | Sort-Object Modified | Select-Object -First 1 }
        if ($slot) { $slotName = $slot.Name }

        if ($slot) {
            # unlink every customer that uses this slot, then delete it (so no old prices stay in it)
            foreach ($u in ($all | Where-Object { $_.PLId -eq $slot.ListID })) {
                $ok = $false
                foreach ($empty in @("<FullName></FullName>", "<ListID></ListID>")) {
                    $r2 = Send "<CustomerModRq><CustomerMod><ListID>$($u.ListID)</ListID><EditSequence>$($u.EditSeq)</EditSequence><PriceLevelRef>$empty</PriceLevelRef></CustomerMod></CustomerModRq>"
                    if ([int]$r2.statusCode -eq 0) { $ok = $true; if ($u.ListID -eq $me.ListID) { $me.EditSeq = T $r2 "CustomerRet/EditSequence" }; break }
                }
                if (-not $ok) { Log "FAIL unlink $($u.Name) from $slotName : $($r2.statusMessage)"; return "Could not free $slotName from '$($u.Name)': $($r2.statusMessage). Nothing else changed." }
                Log "unlinked '$($u.Name)' from $slotName"
            }
            $r2 = Send "<ListDelRq><ListDelType>PriceLevel</ListDelType><ListID>$($slot.ListID)</ListID></ListDelRq>"
            if ([int]$r2.statusCode -ne 0) { Log "FAIL delete $slotName : $($r2.statusMessage)"; return "Could not reset $slotName : $($r2.statusMessage)" }
        }

        # create the slot with this customer's last prices and link it
        $rows = $byCust[$cust.Id]
        $per = ($rows | ForEach-Object { "<PriceLevelPerItem><ItemRef><ListID>$($_.ItemId)</ListID></ItemRef><CustomPrice>$($_.Price)</CustomPrice></PriceLevelPerItem>" }) -join ""
        $cur = if ($cust.CurId) { "<CurrencyRef><ListID>$($cust.CurId)</ListID></CurrencyRef>" } else { "" }
        $r2 = Send "<PriceLevelAddRq><PriceLevelAdd><Name>$(Esc $slotName)</Name><IsActive>true</IsActive>$per$cur</PriceLevelAdd></PriceLevelAddRq>"
        if ([int]$r2.statusCode -ne 0) { Log "FAIL add $slotName : $($r2.statusMessage)"; return "Could not create $slotName : $($r2.statusMessage)" }
        $plId = T $r2 "PriceLevelRet/ListID"
        $r3 = Send "<CustomerModRq><CustomerMod><ListID>$($me.ListID)</ListID><EditSequence>$($me.EditSeq)</EditSequence><PriceLevelRef><ListID>$plId</ListID></PriceLevelRef></CustomerMod></CustomerModRq>"
        if ([int]$r3.statusCode -ne 0) { Log "FAIL link $($me.Name) : $($r3.statusMessage)"; return "Prices loaded but could not link the customer: $($r3.statusMessage)" }
        Log "prepared '$($me.Name)' in $slotName ($($rows.Count) items)"
        return "READY: '$($me.Name)' -> $slotName ($($rows.Count) items). Open a new invoice now."
    }
    finally { $rp.EndSession($ticket); $rp.CloseConnection() }
}

# ---------- window ----------
$font = New-Object Drawing.Font("Tahoma", 11)
$f = New-Object Windows.Forms.Form
$f.Text = "Last Price - prepare customer"; $f.Size = New-Object Drawing.Size(560, 650); $f.StartPosition = "CenterScreen"; $f.TopMost = $true; $f.KeyPreview = $true
$lbl = New-Object Windows.Forms.Label; $lbl.Text = "Search customer (any part of the name, Arabic or English, words in any order):"; $lbl.Location = "10,10"; $lbl.AutoSize = $true
$tb = New-Object Windows.Forms.TextBox; $tb.Location = "10,32"; $tb.Width = 420; $tb.Font = $font
$clr = New-Object Windows.Forms.Button; $clr.Text = "Clear (Esc)"; $clr.Location = "440,31"; $clr.Width = 90; $clr.Height = 28
$cnt = New-Object Windows.Forms.Label; $cnt.Location = "10,66"; $cnt.AutoSize = $true; $cnt.ForeColor = [Drawing.Color]::DimGray
$lb = New-Object Windows.Forms.ListBox; $lb.Location = "10,88"; $lb.Size = New-Object Drawing.Size(520, 380); $lb.Font = $font; $lb.IntegralHeight = $false
$info = New-Object Windows.Forms.Label; $info.Location = "10,474"; $info.Size = New-Object Drawing.Size(520, 22); $info.ForeColor = [Drawing.Color]::Navy
$btn = New-Object Windows.Forms.Button; $btn.Text = "Prepare (Enter)"; $btn.Location = "10,500"; $btn.Width = 140; $btn.Height = 32
$st = New-Object Windows.Forms.Label; $st.Location = "10,540"; $st.Size = New-Object Drawing.Size(520, 66)
$f.Controls.AddRange(@($lbl, $tb, $clr, $cnt, $lb, $info, $btn, $st))
$script:shown = @()

function Fill() {
    $words = @((Norm $tb.Text).Split(' ') | Where-Object { $_ })
    $first = @(); $rest = @()
    foreach ($n in $names) {
        $ok = $true
        foreach ($w in $words) { if (-not $n.Key.Contains($w)) { $ok = $false; break } }
        if (-not $ok) { continue }
        # names that START with the first word come first
        if ($words.Count -gt 0 -and $n.Key.StartsWith($words[0])) { $first += $n } else { $rest += $n }
    }
    $script:shown = @($first + $rest)
    $lb.BeginUpdate(); $lb.Items.Clear()
    foreach ($n in $script:shown) { $tag = switch ($n.Fixed) { "YES" { "   [automatic]" } "OTHER" { "   [has level: $($n.Other)]" } "INACTIVE" { "   [inactive]" } default { "" } }; [void]$lb.Items.Add("$($n.Name)$tag") }
    $lb.EndUpdate()
    if ($script:shown.Count -gt 0) { $lb.SelectedIndex = 0 } else { $info.Text = "" }
    $cnt.Text = if ($words.Count -eq 0) { "$($names.Count) customers" } elseif ($script:shown.Count -eq 0) { "No customer matches - try fewer letters or one word." } else { "$($script:shown.Count) of $($names.Count) customers" }
}
$lb.Add_SelectedIndexChanged({
    if ($lb.SelectedIndex -lt 0) { $info.Text = ""; return }
    $c = $script:shown[$lb.SelectedIndex]
    $d = if ($c.Disc -and $c.Disc -ne "0") { "last discount $($c.Disc)%" } else { "no discount" }
    $k = switch ($c.Fixed) { "YES" { "automatic (fixed level)" } "OTHER" { "has price level $($c.Other) - Prepare will ask to replace it" } "INACTIVE" { "INACTIVE customer - needs Prepare" } default { "needs Prepare" } }
    $info.Text = "$($c.Items) items  |  $d  |  $k"
})
$tb.Add_TextChanged({ Fill })
$go = {
    if ($lb.SelectedIndex -lt 0) { $st.Text = "Pick a customer first."; return }
    $c = $script:shown[$lb.SelectedIndex]
    if ($c.Fixed -eq "YES") { $st.Text = "'$($c.Name)' is automatic already. Just open the invoice.`n" + (DiscText $c); return }
    $st.Text = "Working..."; $f.Refresh()
    try { $st.Text = (Prepare-Customer $c) + "`n" + (DiscText $c); if ($c.Fixed -eq "INACTIVE") { $st.Text += "`nNote: this customer is INACTIVE in QuickBooks - make it active before invoicing." } } catch { $st.Text = "ERROR: $($_.Exception.Message)"; Log "ERROR $($_.Exception.Message)" }
}
$clr.Add_Click({ $tb.Text = ""; $tb.Focus() })
# keyboard: type to search, Up/Down to move in the list, Enter = Prepare, Esc = clear search
$tb.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq "Down" -and $lb.Items.Count -gt 0) { if ($lb.SelectedIndex -lt $lb.Items.Count - 1) { $lb.SelectedIndex++ }; $e.Handled = $true }
    elseif ($e.KeyCode -eq "Up" -and $lb.SelectedIndex -gt 0) { $lb.SelectedIndex--; $e.Handled = $true }
    elseif ($e.KeyCode -eq "Enter") { $e.SuppressKeyPress = $true; & $go }
    elseif ($e.KeyCode -eq "Escape") { $e.SuppressKeyPress = $true; $tb.Text = "" }
})
$lb.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq "Enter") { $e.SuppressKeyPress = $true; & $go }
    elseif ($e.KeyCode -eq "Escape") { $e.SuppressKeyPress = $true; $tb.Text = ""; $tb.Focus() }
})
$btn.Add_Click($go); $lb.Add_DoubleClick($go)
$f.Add_Shown({ $tb.Focus() })
Fill
[void]$f.ShowDialog()
