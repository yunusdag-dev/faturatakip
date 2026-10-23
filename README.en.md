<div align="center">

<img src="docs/ikon.png" width="96" alt="Fatura Takip icon">

# Purchase Invoice Registration and Tracking

**A desktop app for small teams, running on a shared network folder, that reads PDF invoices and fills in the form for you.**

[Türkçe](README.md) · **English**

![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)
![SQLite](https://img.shields.io/badge/SQLite-shared%20network%20folder-003B57?logo=sqlite&logoColor=white)
![Platform](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?logo=windows&logoColor=white)
![UI](https://img.shields.io/badge/tkinter-Sun%20Valley%20theme-4B8BBE)

<img src="docs/ekran-acik.png" width="880" alt="Light theme">

</div>

> **Note:** the application's user interface is in **Turkish**. This page describes it in English and gives the Turkish label of each screen element in parentheses where it helps.
>
> *In short:* a desktop app for a small team to register and track purchase invoices on a shared network folder. It extracts invoice fields (date, number, supplier, tax ID, amounts) from PDFs, including multi-page scans, using a text-layer parser with a Windows OCR fallback. Python + tkinter + SQLite, packaged as a standalone `.exe`.

## Purpose

To record purchase invoices, to cut the effort and error risk of typing supplier / number / date / amount by hand for every invoice, and to keep all the documents belonging to an invoice in one place. The application:

- **reads the invoice PDF** and fills in the fields itself (you check and save),
- opens **all documents** of an invoice (request, letter, delivery note…) from a single place,
- lets a team of 4-5 people use **the same database** from a shared network folder,
- exports month by month to **a tidy Excel table**.

## Features

| | |
|---|---|
| **Per-row Browse filling** | In the New Record form (*Yeni Kayıt*) every row has its own *Gözat* (Browse) button; you pick a PDF, DOCX or XLSX. The invoice PDF supplies date, invoice number, supplier and the amount excluding VAT; a Request Form (*Talep Formu*) PDF supplies the request number and date; in a Needs Form (*İhtiyaç Formu*) PDF the text under the "Name and characteristics of the service/material" heading becomes the job name; the Direct Procurement Approval document (DOCX) supplies the procurement number (and, if wanted, the job name). Invoice date, number, supplier and amount are filled together from a single invoice PDF. Results are highlighted in yellow, the user checks and saves; nothing is saved without confirmation. |
| **Category** | Every record belongs to a category (truck rental, fuel, toll passes, etc.); the Excel output gets one sheet per category. |
| **Multi-page document support** | In a 10-20 page scan it **finds the invoice page automatically** (page scoring); if wrong, type `7` or `7-8` into the "Invoice page" box. |
| **Scanned (image) PDFs** | Read with the Windows built-in OCR engine; no extra software needed. |
| **Document management** | Attachments such as request, letter, quotation and delivery note are linked to an invoice; double-click the invoice to list them all and open any with one click. |
| **Search and filter** | Year / month / supplier / date range / free-text search; a monthly summary window. |
| **Excel output** | Formatted `.xlsx` with one sheet per category: real date and number cells, a totals row, filter, frozen header, print setup, clickable PDF links. CSV is supported too. |
| **Duplicate protection** | The same supplier + invoice number, or the same PDF, cannot be saved twice. |
| **Modern interface** | Light / Dark / Blue / Green / Purple / Orange / Red themes (*Açık / Koyu / Mavi / Yeşil / Mor / Turuncu / Kırmızı*), zebra rows, application icon. |

<div align="center">
<img src="docs/ekran-koyu.png" width="880" alt="Dark theme">
</div>

## Engineering notes

Small as it looks, the project solves several real problems with deliberate decisions:

- **Amounts are stored as `int` in kuruş (hundredths of a lira).** No rounding errors from floating point; parsing uses `Decimal`, and spellings such as `1.234,56` / `1234.56` / `1.234` are told apart correctly.
- **Its own PDF text parser** (standard library only): content streams, Flate/ASCII85 filters, `ToUnicode` CMaps and line merging by page position. If PyMuPDF is installed it is used; otherwise the built-in reader takes over. So the program **also runs without any extra library**.
- **Tiered reading:** PyMuPDF → built-in text-layer reader → Windows OCR. Each stage takes over what the previous one could not do; pages without a text layer go through OCR only when needed.
- **Invoice page detection:** every page is scored on "invoice hints" (invoice number, ETTN, total including taxes…); request and letter pages do not leak into the fields.
- **SQLite on a shared network folder:** errors such as locking are shown to the user as clear Turkish messages; UNC paths (`\\server\share`) are supported.
- **Rules drawn from real invoice layouts:** telling a VKN (10 digits) from a TCKN (11 digits), excluding your own organisation's number, supplier names split across line ends (`… SAN. VE` / `TİC.LTD.ŞTİ`), and similar cases.
- **The interface does not freeze:** long PDF/OCR jobs run in a separate thread.

The code is in a single file (about 2,500 lines) and layered: settings → helpers → database → PDF reading → interface.

## Installation

### For users (no Python needed)

1. From the [Releases](https://github.com/yunusdag-dev/faturatakip/releases/latest) section download `FaturaTakip-...-windows.zip` and extract it to a permanent folder.
2. Create a desktop shortcut for `FaturaTakip.exe`.

> The exe does not run on its own; it must stay together with the `_internal` folder next to it.

### Installing on a shared network (for a team)

1. On the shared computer create the `C:\FaturaTakip\Program`, `Veri` (data) and `Evraklar` (documents) folders; put the program in `Program`.
2. Share the `C:\FaturaTakip` folder; give users **Read + Change** permission (SQLite needs write permission).
3. In `Program\config.ini` write the paths as UNC paths:
   ```ini
   [ayarlar]
   veritabani = \\SERVER-PC\FaturaTakip\Veri\fatura.db
   tarama_klasoru = \\SERVER-PC\FaturaTakip\Evraklar
   ```
4. On the other computers create a shortcut to `\\SERVER-PC\FaturaTakip\Program\FaturaTakip.exe`.
5. **Updating:** after everyone has closed the program, copy the new files into `Program`; do not touch `config.ini` or `Veri`.
6. Back up `Veri\fatura.db` regularly.
7. **First launch (administrator account):** before distributing the shortcuts, **open the application once from your own computer**. If the database has no accounts, the "First Setup" (*İlk Kurulum*) screen appears; create the administrator account (this screen never appears again). Add the other people under **File > Users** (*Dosya > Kullanıcılar*): you give a user name and a temporary password, and each person sets their own password at first sign-in. Accounts are kept inside `fatura.db` (passwords as salted hashes).

**Accounts and permissions**

| | Administrator | User |
|---|---|---|
| Adding / editing records, documents, Excel output | ✔ | ✔ |
| Deleting invoices | ✔ | ✘ |
| Adding users, resetting passwords, deactivating, granting administrator rights | ✔ | ✘ |

- "Remember me on this computer" (valid for 90 days) only stores a random key in that person's own Windows profile; do not tick it on someone else's computer. **File > Sign Out** (*Dosya > Oturumu Kapat*) deletes this record.
- **If the administrator password is forgotten:** anyone with access to the database file can run `python fatura_takip.py --yonetici-sifirla` with Python and set a new administrator password.
- This arrangement provides **accountability, not security** (you can see who added what). Because the database is just a file in a shared folder, someone with access to that folder can bypass the application; for real protection, restrict the folder's Windows permissions.

For scanned PDFs the **Turkish OCR language pack** must be installed in Windows.

### For developers

```bash
git clone https://github.com/yunusdag-dev/faturatakip.git
cd faturatakip
pip install pymupdf openpyxl sv_ttk      # all optional
python fatura_takip.py
```

To run the tests (no GUI needed, no real data used): `python -m unittest discover -s tests -v`

To build the exe, run `pip install pyinstaller` once, then double-click [derle.bat](derle.bat). The output is in `dist\FaturaTakip\`.

## Settings (`config.ini`)

Example: [config.ornek.ini](config.ornek.ini)

| Setting | Description |
|---|---|
| `veritabani` | Database path (UNC for a network; do not use drive letters) |
| `tarama_klasoru` | First folder shown by the PDF "Browse" dialog |
| `tema` | Açık / Koyu / Mavi / Yeşil / Mor / Turuncu / Kırmızı (Light / Dark / Blue / Green / Purple / Orange / Red) |
| `kurum_anahtar` | Words from your own organisation's name, so that the recipient is not mistaken for the supplier in a PDF |
| `kurum_vkn` | Your own organisation's tax number |

## Roadmap

Ideas planned according to the team's needs, **not yet done**:

- [ ] Payment status and due-date tracking (paid / unpaid, upcoming-due warning)
- [ ] Yearly summary per supplier
- [ ] Bulk PDF upload: read the invoices in a folder one after another and check them on a single screen
- [ ] Automatic, dated database backup
- [ ] Change history (who changed what, and when)
- [ ] Amount consistency warning (amounts that do not match the VAT rate)
- [ ] Attaching PDFs by drag and drop

## Known limitations

- Scans read with OCR have an error margin (for example `L` instead of `I`); fields highlighted in yellow must always be checked.
- The database is a single SQLite file; at this scale (4-5 users, wired network) it is sufficient, but it is not suitable for many-user or remote access.
- Files are not copied, only their locations are recorded; documents must stay in a shared folder the team can reach.
- Fully works only on Windows (OCR and file-opening functions depend on Windows).

## License

The source code is **MIT licensed** ([LICENSE](LICENSE)): organisations included, you may use, modify and distribute it.
Because the compiled package contains AGPL-licensed PyMuPDF, see [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) for the component licenses when **distributing** it. Internal use is not restricted.

## Privacy

`fatura.db`, `config.ini` and invoice PDFs are kept out of the repository by `.gitignore`. The companies and amounts shown in the screenshots in this repository are **fictitious**.

## Developer

**Yunus Dağ** · [@yunusdag-dev](https://github.com/yunusdag-dev)
