# أداة آخر سعر للعميل — QuickBooks Last Price Tool

أداة لـ QuickBooks Enterprise 11.0 (2011) تجعل الفاتورة الجديدة تعبّي تلقائياً **آخر سعر بيع لكل عميل ولكل صنف** (السعر قبل الخصم)، وتعرض آخر نسبة خصم أُعطيت للعميل.

A tool for QuickBooks Desktop Enterprise 11.0 that makes new invoices auto-fill **the last price each customer paid for each item** (price before discount), and shows the customer's last discount %.

---

## كيف تعمل / How it works

1. `lastprice.ps1` يقرأ الفواتير وإيصالات البيع، ويحسب آخر سعر لكل عميل وصنف وآخر نسبة خصم.
2. أكبر العملاء (حسب مبيعات آخر 12 شهر) يأخذ كل واحد منهم Price Level ثابت باسم `LP <اسم العميل>` (حد QuickBooks 2011 هو 99 مستوى سعر).
3. باقي العملاء: 12 مستوى متنقل `LPX 01..12` يُعطى للعميل عند الطلب من نافذة `prepare`.
4. النافذة تنسخ نص `Discount xx%` للصق في Memo الفاتورة، والمحاسب يضيف سطر الخصم كالمعتاد.

1. `lastprice.ps1` reads Invoices + Sales Receipts and computes the last price and discount % per customer/item.
2. Top customers by 12‑month sales get a fixed per‑item price level `LP <customer>` (QB 2011 allows max 99 levels).
3. Other customers share 12 rotating levels `LPX 01..12`, assigned on demand from the `prepare` window.
4. The window copies `Discount xx%` to the clipboard for the invoice Memo; the accountant adds the Discount line as usual.

## الملفات / Files

| الملف / File | الوظيفة / Purpose |
|---|---|
| `probe.ps1` | فحص اتصال للقراءة فقط / Read‑only connection test |
| `dryrun.bat` | تجربة بدون أي تغيير، تكتب خطة CSV / Dry run, writes a plan CSV, changes nothing |
| `apply-one.bat` | يطبّق على عميل واحد (للتجربة) / Apply for one customer (test) |
| `apply-all.bat` | يطبّق على كل العملاء (للتشغيل الليلي) / Apply for all customers (nightly) |
| `prepare.bat` / `prepare.ps1` | نافذة "تجهيز عميل" للمحاسب / "Prepare customer" window for the accountant |
| `lastprice.ps1` | السكربت الرئيسي / Main script |

## طريقة الاستخدام / How to use

**المتطلبات / Requirements:** Windows، QuickBooks مفتوح على ملف الشركة وداخل كـ Admin، و PowerShell 32‑bit:
`C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe` (ملفات .bat تستخدمه تلقائياً).

1. ضع كل الملفات في مجلد واحد. / Put all files in one folder.
2. شغّل `probe.ps1` أول مرة ووافق على طلب QuickBooks للسماح للتطبيق "Last Price Tool". / Run `probe.ps1` once and allow the app in QuickBooks.
3. شغّل `dryrun.bat` وراجع الخطة في `Desktop\LastPrice\`. / Run `dryrun.bat` and review the plan in `Desktop\LastPrice\`.
4. عدّل مسار ملف الشركة بعد `-AllowFile` في ملفات `.bat`. الأداة **ترفض الكتابة** على أي ملف شركة غير هذا المسار. / Edit the company file path after `-AllowFile` in the `.bat` files. The tool **refuses to write** to any other company file.
5. شغّل `apply-all.bat` (واكتب `YES` للتأكيد). يفضّل جدولته ليلياً في Task Scheduler. / Run `apply-all.bat` (type `YES`). Schedule it nightly with Task Scheduler.
6. المحاسب يشغّل `prepare.bat` للعميل الذي ليس له مستوى ثابت قبل عمل الفاتورة. / The accountant runs `prepare.bat` for customers without a fixed level before invoicing.

> ⚠️ جرّب دائماً على نسخة احتياطية أولاً، وخذ Backup قبل أول تشغيل على الملف الأساسي.
> Always test on a backup copy first, and back up before the first run on the live file.

## ملاحظات / Notes

- لا ترفع مجلد `LastPrice` أو أي ملف CSV/Excel: فيها أسعار العملاء. (`.gitignore` يمنع ذلك.)
- Never commit the `LastPrice` folder or CSV/Excel files: they contain customer prices (blocked by `.gitignore`).
- لا تغيّر الأداة مستوى سعر لعميل عنده مستوى غير `LP`/`LPX`. / It never overwrites a customer's non‑LP price level.
