<div align="center">

# PriceRecall

**Auto-fill every customer's last sold price in QuickBooks Desktop invoices**

![Platform](https://img.shields.io/badge/platform-Windows-0078D6)
![QuickBooks](https://img.shields.io/badge/QuickBooks-Enterprise%2011.0%20(2011)-2CA01C)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20(32--bit)-5391FE)
![qbXML](https://img.shields.io/badge/SDK-qbXML%2010.0-orange)
![Version](https://img.shields.io/badge/version-5.3-blue)

</div>

---

## Table of contents

1. [The problem](#the-problem)
2. [The solution](#the-solution)
3. [How it works](#how-it-works)
4. [Requirements](#requirements)
5. [Installation](#installation)
6. [Usage](#usage)
7. [Daily workflow for the accountant](#daily-workflow-for-the-accountant)
8. [How discounts are detected](#how-discounts-are-detected)
9. [Output files](#output-files)
10. [Parameters](#parameters)
11. [Safety features](#safety-features)
12. [Scheduling the nightly run](#scheduling-the-nightly-run)
13. [Undo and recovery](#undo-and-recovery)
14. [Troubleshooting](#troubleshooting)
15. [Known limitations](#known-limitations)
16. [Repository layout](#repository-layout)
17. [Changelog](#changelog)

---

## The problem

QuickBooks Desktop has no "last price" feature. When the accountant creates an invoice, each line fills in the item's **default sales price**, not the price this specific customer paid last time. In a business where every customer has negotiated prices, staff have to look up old invoices by hand for every line. That is slow, and a single wrong price means a customer is over- or under-charged.

## The solution

PriceRecall reads your sales history through the official QuickBooks SDK and turns it into **QuickBooks Price Levels**, a feature QuickBooks already has:

- Every important customer gets a **personal per-item price level** named `LP <customer name>` holding the last price they paid for each item.
- That level is **linked to the customer**, so when you pick the customer on a new invoice, QuickBooks itself fills every item with that customer's last price. There is no add-on to open and nothing to copy.
- **Every other customer with at least one sale** (customers with another price level and generic cash customers included; inactive customers are hidden) is served on demand from a small **"Prepare customer" window**, which loads their prices into one of 12 rotating levels in one click.
- The window also shows the customer's **last discount %** and copies `Discount 15%` to the clipboard, ready to paste into the invoice Memo.

Prices are stored **before discount** (the Rate on the invoice line). The discount is added as a separate Discount line, exactly as the accountant already does.

---

## How it works

```mermaid
flowchart LR
    A[(QuickBooks<br/>company file)] -- "read Invoices,<br/>Sales Receipts,<br/>Items, Customers" --> B[lastprice.ps1<br/>nightly]
    B -- "dry run" --> C[plan CSV<br/>review only]
    B -- "-Apply" --> D["LP &lt;customer&gt;<br/>price levels<br/>(top customers)"]
    B --> E[cache-lastprices.csv]
    E --> F[prepare.ps1<br/>window]
    F -- "on demand" --> G["LPX 01..12<br/>rotating levels<br/>(other customers)"]
    D --> H[New invoice<br/>auto-fills last price]
    G --> H
```

### Step by step

1. **Read.** `lastprice.ps1` connects to the open company file and reads all Items, Customers and existing Price Levels. It then reads every Invoice and Sales Receipt from `-FromDate` (default `2021-01-01`) to today, **one month at a time** so large files don't time out.
2. **Compute.** For every *customer + item* pair it keeps the most recent sale line (by transaction date, then creation time, then line order). It records the unit price before discount, the net price after discount, the discount %, the previous price, how often the item was sold and the unit of measure.
3. **Rank.** QuickBooks 2011 allows a **maximum of 99 price levels** (the 100th fails with error 3180). The tool counts the levels you already use for other purposes, reserves 12 for the rotating `LPX` slots, and gives the remaining slots to the customers with the **highest net sales in the last 365 days**. The ranking is redone on every run, so the list follows your business as it changes.
4. **Plan.** Every customer and item gets an action: `NEW`, `CHANGE`, `SAME`, or a reason it is skipped. Everything goes into a plan CSV you can open in Excel.
5. **Apply** (only with `-Apply`). The tool:
   - unlinks and deletes `LP` levels of customers who **dropped out** of the top list (this frees slots first),
   - creates or updates the `LP` level for each top customer, adding new prices and never removing an existing item,
   - links the level to the customer.
6. **Prepare** (during the day). For a customer without a fixed level, the accountant picks the name in `prepare.bat`. The tool takes a free `LPX` slot (or the least recently used one), clears it, fills it with this customer's last prices and links it. If QuickBooks refuses a new slot because its list is full, the tool recycles the oldest `LPX` slot instead, so Prepare always works. The next invoice for that customer fills automatically.

---

## Requirements

| Requirement | Details |
|---|---|
| Operating system | Windows (any version that runs QuickBooks Desktop) |
| QuickBooks | Desktop **Enterprise Solutions 11.0 (2011)**. The scripts use qbXML 10.0, so later Desktop versions should also work. |
| QuickBooks features | **Price Levels** enabled, with the *per item* type (Edit > Preferences > Sales & Customers) |
| QuickBooks SDK | `QBXMLRP2.RequestProcessor`, which is installed with QuickBooks Desktop |
| PowerShell | Windows PowerShell 5.1, **32-bit**: `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe` (the `.bat` files already use it) |
| User | QuickBooks open on the company file, logged in as **Admin** for the first authorisation |

> **Why 32-bit?** The QuickBooks 2011 SDK is a 32-bit COM component. A 64-bit PowerShell cannot load it, and the script stops with a clear message if you try.

---

## Installation

1. **Download** this repository (green **Code** button > **Download ZIP**) and extract it to a folder, for example `C:\PriceRecall\`. Keep all files together.
2. **Set the company file path.** All `.bat` files read it from **`settings.bat`**, the only file you edit. Open QuickBooks on the company file and double-click **`set-path.bat`**: it asks QuickBooks for the open file and writes it into `settings.bat` (if that fails, it takes the `Company file:` line from the latest dry-run log). You can also edit it by hand:
   ```bat
   set "QBFILE=D:\QuickBooksData\YourCompany.qbw"
   set "NIGHTLY_TIME=23:30"
   ```
   Use the path **as seen on that PC** (a server path like `D:\...` on the server, a network path like `\\server\share\...` on a workstation).
3. **Authorise the tool once.** Open QuickBooks as Admin, then run in the 32-bit PowerShell:
   ```powershell
   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File probe.ps1
   ```
   QuickBooks asks whether to allow **"Last Price Tool"** to access the file. Choose **"Yes, always; allow access even if QuickBooks is not running"** and confirm. `probe.ps1` only reads, and writes a short report to `Desktop\lastprice-probe.txt`.

> ⚠️ **Always test on a backup copy of the company file first**, and take a fresh QuickBooks backup before the first run on the live file.

---

## Usage

### 1. Dry run (no changes)

Double-click **`dryrun.bat`**.

The tool reads everything, builds the plan and writes it to `Desktop\LastPrice\plan-YYYYMMDD-HHMM.csv`. **Nothing is changed in QuickBooks.** Open the CSV in Excel and check a few customers you know well.

### 2. Try one customer

Run **`apply-one.bat`** and type the customer's full name when asked. Type `YES` when asked. Only that customer is touched. Create a test invoice for the customer and check that the prices fill in.

### 3. Apply to all customers

Run **`apply-all.bat`** and type `YES`. On the first run this creates up to ~84 `LP` levels. Later runs only send what changed.

### 4. Prepare a customer without a fixed level

Run **`prepare.bat`**. A small window opens:

- Type **any part of the name**: one letter, one word, or several words in any order. Search ignores upper/lower case and Arabic spelling variants (أ/إ/آ = ا, ة = ه, ى = ي, ؤ = و, ئ = ي, tashkeel, tatweel, Arabic digits). Names that **start** with the first word are listed first.
- The list shows every **active** customer with at least one sale (inactive customers are hidden), with a tag:

  | Tag | Meaning |
  |---|---|
  | `[automatic]` | Has a fixed `LP` level. Go straight to the invoice. |
  | `[has level: X]` | Uses another price level. Prepare asks **Yes/No** before replacing it. |
  | *(no tag)* | Needs Prepare. |

- The line under the list shows the number of items, the last discount and the status of the selected customer.
- Keyboard: **Up/Down** move in the list while typing, **Enter** = Prepare, **Esc** = clear the search. Double-click also prepares.
- The status line shows `READY: '<customer>' -> LPX 03 (57 items). Open a new invoice now.`
- If the customer had a discount last time, `Discount 15%` is **already on the clipboard**. Paste it into the invoice Memo with **Ctrl+V** and add the Discount line.

---

## Daily workflow for the accountant

| When | What | Who |
|---|---|---|
| Every night | `apply-all.bat` refreshes all `LP` levels and the cache | Scheduled (see below) |
| Before an invoice for a big customer | Nothing. Just open the invoice. | Accountant |
| Before an invoice for any other customer | Open `prepare.bat`, pick the customer, click **Prepare** | Accountant |
| On the invoice | Paste `Discount xx%` into the Memo (Ctrl+V) and add the Discount line | Accountant |
| Price lookup without invoicing | Open `Desktop\LastPrice\last-prices.csv` in Excel and filter by customer | Anyone |

---

## How discounts are detected

Discounts in QuickBooks are separate lines, so the tool has to work out which sale lines each discount line applies to. These rules were checked against real invoices:

| Discount line type | What it applies to |
|---|---|
| **Percentage** discount item (e.g. `-10%`) | The line directly above it, or all lines covered by a **Subtotal** directly above it. This is QuickBooks' own rule. |
| **One fixed-amount** discount line | Either the whole invoice or only the lines above it. The tool picks whichever gives a "round" percentage (e.g. 15% or 12.5%). |
| **Several fixed-amount** lines | Each covers the lines since the previous discount line. |

Staff often type a rounded amount, so a near-round result such as **14.96% or 14.999% is snapped to 15%** (and 12.47% to 12.5%), within 0.06 percentage points.

Discounts combine: a 10% line discount followed by a 5% invoice discount gives an effective 14.5%.

---

## Output files

Everything is written to **`Desktop\LastPrice\`**. These files contain customer prices, so they are **excluded from Git** by `.gitignore`.

| File | Written by | Contents |
|---|---|---|
| `plan-YYYYMMDD-HHMM.csv` | every run | Full plan: customer, status, price level, action, item, unit of measure, last price, net price, discount %, date, document number, previous price, times sold |
| `last-prices.csv` | full runs | Lookup list for every active customer, including those without a level (`AutoInInvoice = YES/blank`) |
| `cache-lastprices.csv` | full runs | Data used by the `prepare` window: every customer with sales, with `Fixed` = `YES` / `OTHER` / `INACTIVE` / blank and `OtherLevel` |
| `run-YYYYMMDD-HHMM.txt` | every run | Log of what was read, planned and written |
| `backup-levels-*.csv` | `-Apply` | Every `LP` level price **before** the change |
| `backup-customer-links-*.csv` | `-Apply` | Which customer was linked to which level **before** the change |
| `prepare-log.txt` | `prepare` | Every customer prepared and any error |

---

## Parameters

### `lastprice.ps1`

| Parameter | Default | Meaning |
|---|---|---|
| `-Apply` | off | Write to QuickBooks. Without it the run is a dry run. |
| `-AllowFile` | *(empty)* | Exact path of the company file allowed to be written. Required with `-Apply`. |
| `-OnlyCustomer` | *(empty)* | Process only this customer (full name). Never touches other customers. |
| `-FromDate` | `2021-01-01` | Read sales from this date (`yyyy-MM-dd`) |
| `-ActiveDays` | `365` | A customer needs a sale in this many days to get a fixed level. Also the window used for ranking. |
| `-MaxLevels` | `100` | QuickBooks price level limit (warning only) |
| `-Rotating` | `12` | Number of `LPX` slots reserved for `prepare` |
| `-ExcludeCustomers` | `Cash Van`, `Cash Sales`, `Samples Marketing` | Generic customers that never get a level |
| `-NetPrice` | off | Old mode: store the price **after** discount instead of before |

### `prepare.ps1`

| Parameter | Default | Meaning |
|---|---|---|
| `-AllowFile` | *(empty)* | Refuse to write if another company file is open |
| `-Rotating` | `12` | Number of `LPX` slots (must match `lastprice.ps1`) |

---

## Safety features

- **Dry run by default.** Nothing is written unless you pass `-Apply`.
- **File lock.** With `-Apply`, the open company file must match `-AllowFile` exactly, or nothing is written.
- **Explicit confirmation.** The tool asks you to type `YES` before writing.
- **Your own price levels are never touched automatically.** The nightly run skips customers who have a price level that does not start with `LP ` or `LPX `. Only `prepare` can replace it, and only after the accountant answers **Yes**.
- **Prices are never removed.** Updating a level adds and changes prices but never deletes an item from it.
- **Backups before every write**, with the previous prices and customer links saved to CSV.
- **Limit check.** If the result would exceed 99 levels, the run stops before writing. If QuickBooks still refuses a level, fixed customers win: one `LPX` slot is freed and the add is retried once.
- **Inactive items are skipped.** Inactive and generic cash customers never get a fixed `LP` level; they can only be prepared on demand.

---

## Scheduling the nightly run

Run **`install-nightly.bat`** once on the PC (or server) that keeps QuickBooks open on the company file. It creates the Windows task **"PriceRecall Nightly"**, which starts `nightly.bat` every day at `NIGHTLY_TIME` from `settings.bat`. `nightly.bat` runs `lastprice.ps1 -Apply -Yes`, so no `YES` prompt is needed.

- Do **not** run it as Administrator ("Run as administrator"): an elevated task cannot talk to the normal QuickBooks window and fails with *multiple instances of QuickBooks*.
- The Windows user must stay **logged in** (disconnect Remote Desktop, never sign out).
- Recommended: schedule the QuickBooks backup half an hour before the nightly run (e.g. 23:00 backup, 23:30 nightly).
- Remove the task with **`uninstall-nightly.bat`**.
- Each run writes `Desktop\LastPrice\run-YYYYMMDD-HHMM.txt`, and copies the cache to `<company file folder>\LastPrice\` so `prepare` on other PCs gets fresh data.

QuickBooks must be open, and no one should be editing a customer or a price level at that moment.

---

## Undo and recovery

- **Restore a QuickBooks backup.** This is the cleanest option and the reason to back up before the first live run.
- **Remove everything the tool created.** Delete the price levels whose names start with `LP ` and `LPX ` (Lists > Price Level List). QuickBooks clears the link on the customers automatically. Note: deleted levels keep counting against the QuickBooks 2011 limit, so new ones may not be creatable afterwards.
- **Restore old prices.** The `backup-levels-*.csv` and `backup-customer-links-*.csv` files record the state before each run.

---

## Troubleshooting

| Message / symptom | Cause | Fix |
|---|---|---|
| `ERROR: run with 32-bit PowerShell` | Started from the normal 64-bit PowerShell | Use the `.bat` files or the `SysWOW64` path |
| `Could not start QuickBooks` / access denied | Tool not authorised, or QuickBooks not open | Open QuickBooks as Admin and run `probe.ps1` again. Check *Edit > Preferences > Integrated Applications*. |
| `STOP: -AllowFile does not match` | Another company file is open, or the path differs | Copy the exact *Location* from F2 into the `.bat` file |
| `You have exceeded the list limit` / error **3180** when adding a level | The price level list is full (100). "Deleted" levels are only inactive and still count. | Since v5 the tool reuses existing levels instead of creating new ones. **Do not delete price levels by hand**: it does not free space. |
| `multiple instances of QuickBooks` | The script was started as Administrator, or an extra `QBW32.EXE` is running | End the extra `QBW32.EXE` in Task Manager and start the `.bat` by double-click (not as admin) |
| Invoice says *"associated with an inactive Price Level"* | The level was reused while inactive | Click **Make it active**. v5.3 activates reused levels automatically. |
| `could not write last-prices.csv` | The file is open in Excel | Close Excel and run again |
| `Missing ...cache-lastprices.csv` in prepare | The nightly run has not run yet on this PC | Run `dryrun.bat` or `apply-all.bat` once |
| Invoice doesn't fill the last price | Customer not linked, or the item was not sold to them before | Run `prepare` for the customer and check the *Price Level* field on the customer |
| `SKIP: already has price level '...'` in the plan | The customer uses one of your own levels | Intended. In `prepare` the customer shows `[has level: ...]` and can be prepared after answering Yes. |
| `could not link the customer ... The list element is in use` in prepare | The customer is open in an *Edit Customer* window in QuickBooks | Close that window (Cancel) and press Prepare again |
| A customer is missing from prepare | Customer is inactive (hidden), no sale since `-FromDate`, or the nightly run has not run since the change | Run `apply-all.bat` (or `dryrun.bat`) once |

---

## Known limitations

- **100 price levels** in QuickBooks 2011 (inactive ones included), which is why only the top customers get a fixed level and the rest use the 12 rotating slots.
- **One price per item per level.** If a customer buys the same item in different units of measure, the plan marks `MixedUM = YES` and the price of the most recent unit is used.
- **Discounts are not filled automatically.** QuickBooks price levels hold prices only. The discount % is shown and copied for the accountant.
- **Level names** are limited to 31 characters, so long customer names are shortened (and `~2`, `~3` is added if two names collide).
- The prices are as fresh as the **last nightly run**. Prepare uses the same nightly cache, so the tags in the list are also from that run (Prepare itself always checks the customer live in QuickBooks).
- Customers who **never bought anything** do not appear in `prepare`, because they have no last price.

---

## Repository layout

```
PriceRecall/
├── lastprice.ps1          Main engine: read, compute, plan, apply
├── prepare.ps1            "Prepare customer" window (rotating LPX levels)
├── probe.ps1              Read-only connection test and first authorisation
├── settings.bat           Company file path + nightly time (the only file to edit)
├── set-path.bat / .ps1    Writes the open company file path into settings.bat
├── dryrun.bat             Dry run, writes the plan only
├── apply-one.bat          Apply for one customer (asks for the name)
├── apply-all.bat          Apply for all customers (asks YES)
├── nightly.bat            Unattended apply-all (used by the scheduled task)
├── install-nightly.bat    Creates the "PriceRecall Nightly" Windows task
├── uninstall-nightly.bat  Removes that task
├── prepare.bat            Opens the prepare window (no console)
├── .gitignore             Keeps customer data out of Git
└── README.md
```

> 🔒 **Never commit** the `LastPrice` folder, CSV or Excel exports, or company files (`.qbw`, `.qbb`). They contain customer prices. The `.gitignore` blocks them, but check before you push.

---

## Changelog

**v5.3 (2026-11-22)**
- Reused levels are made active again with a separate `PriceLevelMod` (QuickBooks 2011 rejects `IsActive` together with price changes, error 3210).
- A released level that already carries the customer's name is reused first (no `~2` names).
- No STOP at the limit; `-MaxLevels` default is now 100.

**v5 (2026-11-22)**
- Found on the live file: in QuickBooks 2011, deleting a price level through the SDK only **makes it inactive**, and inactive levels count against the 100-level limit.
- Price levels are now **never deleted**. `lastprice.ps1` reuses released `LP` levels by renaming them for new top customers; unused ones are kept as `LP ~spare nn`. If a new level is still refused, an `LPX` level is taken over (renamed), not deleted.
- `prepare.ps1` refills the customer's own `LPX`, tries a new one, and otherwise renames and refills a `LP ~spare` or the least recently used `LPX`. Items of the previous owner are set back to the normal price (0%).

**v4 (2026-11-22)**
- `prepare.ps1`: when QuickBooks refuses a new `LPX` level (*You have exceeded the list limit*), the oldest `LPX` slot is recycled instead of failing. The log records how many price levels the file has.
- `lastprice.ps1`: if adding a fixed `LP` level hits the limit, one `LPX` slot is freed and the add is retried once.

**v3 (2026-11-14)**
- `settings.bat` + `set-path.bat`: one place for the company file path and nightly time.
- Unattended nightly run (`nightly.bat`, `-Yes`), `install-nightly.bat` / `uninstall-nightly.bat`.
- Shared cache next to the company file for `prepare` on other PCs.
- Price before discount; last discount copied as `Discount xx%` for the Memo.

**v2 (2026-11-11)**
- Arabic-aware search in `prepare`, all customers listed with tags, inactive customers hidden.
