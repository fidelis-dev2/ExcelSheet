# Web Excel Spreadsheet

Web Excel Spreadsheet ni spreadsheet application inayofanya kazi moja kwa moja ndani ya web browser. Imetengenezwa kwa HTML, CSS na JavaScript katika faili moja, bila kuhitaji database, web server au installation ya programu maalumu.

Application ina muonekano unaofanana na Microsoft Excel na inaruhusu mtumiaji kuandika data, kutumia formulas, ku-format cells, kutengeneza sheets nyingi, ku-upload Excel workbook na ku-export data.

## Main Features

- Excel-style navigation header na ribbon menu.
- Editable spreadsheet cells zenye rows na columns.
- New Workbook popup form.
- Upload Excel popup form.
- Add Sheet popup form.
- Quick Formula popup form.
- Multiple worksheets ndani ya workbook moja.
- Import ya `.xlsx`, `.xls`, `.csv` na `.webexcel.json`.
- Export ya `.xlsx`, `.csv` na Web Excel project JSON.
- Browser autosave kwa kutumia `localStorage`.
- Undo na redo history.
- Copy, cut na paste ya cell moja au range.
- Cell formatting: font, size, bold, italic, underline, text color na background color.
- Text alignment, wrap text na merge cells.
- Number, TZS currency, USD currency, percentage na date formats.
- Insert na delete rows au columns.
- Sorting, find and replace, remove duplicates, text-to-columns na fill series.
- Bar chart, line chart na pie chart kutoka kwenye selected data.
- Image insertion.
- Print support.
- Fullscreen mode na adjustable zoom.

## Project Files

```text
web_excel.html    Main spreadsheet application
README.md         Project documentation
```

Application yote ipo ndani ya `web_excel.html`; hakuna build process inayohitajika.

## Requirements

- Google Chrome, Microsoft Edge, Firefox au browser nyingine ya kisasa.
- JavaScript lazima iwe enabled.
- Internet connection inahitajika wakati wa kutumia Excel `.xlsx`/`.xls` import na export kwa sababu SheetJS inapakiwa kupitia CDN.

Core spreadsheet features na project JSON/CSV zinaweza kufanya kazi bila backend.

## Running the Application

### Option 1: Open Directly

1. Weka `web_excel.html` kwenye folder unayotaka.
2. Double-click faili hilo.
3. Litafunguka kwenye default web browser.

### Option 2: Run Through a Local Server

Unaweza pia kutumia local server kama Tomcat, Apache, VS Code Live Server au Python HTTP server.

Mfano kwa Python:

```bash
python -m http.server 8080
```

Kisha fungua:

```text
http://localhost:8080/web_excel.html
```

## Quick Start

1. Fungua `web_excel.html` kwenye browser.
2. Double-click cell ili kuandika au kuhariri data.
3. Chagua cells kwa kuburuta mouse.
4. Tumia ribbon menu kufanya formatting, sorting au kuingiza formula.
5. Tumia `Hifadhi` kupakua project JSON, au `Excel` kupakua `.xlsx` workbook.

## Top Navigation Header

Header ya juu ina quick actions zifuatazo:

| Button | Kazi |
| --- | --- |
| New | Hufungua popup ya kutengeneza workbook mpya na kuchagua rows na columns. |
| Upload Excel | Hufungua popup ya kuchagua `.xlsx`, `.xls`, `.csv` au project JSON. |
| Add Sheet | Hufungua popup ya kuweka jina, rows na columns za sheet mpya. |
| Formula | Hufungua popup ya kuchagua function, data range na result cell. |
| Undo/Redo | Hurudisha au kurejesha mabadiliko yaliyofanyika. |
| Hifadhi | Hupakua workbook katika format ya `.webexcel.json`. |
| Excel | Hupakua workbook katika format ya `.xlsx`. |

## Working with Cells

- Single-click cell ili kuichagua.
- Double-click cell ili kuhariri content yake.
- Drag mouse kuchagua range ya cells.
- Tumia `Delete` au `Backspace` kufuta content ya selection.
- Tumia formula bar kuona au kubadilisha raw value/formula ya active cell.
- Andika address kama `A1` au range kama `A1:C10` kwenye name box kwenda moja kwa moja kwenye selection hiyo.

## Supported Formulas

Formula yoyote inaanza kwa alama ya `=`.

| Function | Example | Maelezo |
| --- | --- | --- |
| Arithmetic | `=A1+B1*2` | Hesabu za kawaida. |
| SUM | `=SUM(A1:A10)` | Jumla ya values. |
| AVERAGE | `=AVERAGE(B2:B20)` | Wastani wa values. |
| MIN | `=MIN(C1:C8)` | Value ndogo zaidi. |
| MAX | `=MAX(C1:C8)` | Value kubwa zaidi. |
| COUNT | `=COUNT(D1:D30)` | Idadi ya cells zenye namba. |
| ROUND | `=ROUND(A1/3,2)` | Hupunguza decimal places. |
| IF | `=IF(A1>=50,1,0)` | Conditional numeric result. |

Absolute cell references kama `$A$1` pia zinatambuliwa.

> `IF` kwenye version hii imeundwa hasa kwa numeric results. Complex Excel functions, named ranges, macros na VBA bado hazijajumuishwa.

## Uploading an Excel Workbook

1. Bonyeza `Upload Excel` kwenye header.
2. Chagua Excel file kwenye popup.
3. Bonyeza `Upload & Open`.
4. Kama workbook ina sheets nyingi, kila sheet itaingizwa kama worksheet tofauti.

Upload hufanyika ndani ya browser; application hii haitumi faili kwenda kwenye backend server.

### Information Preserved During Import

- Cell values.
- Basic formulas.
- Multiple worksheets.
- Merged ranges.

### Information That May Not Be Preserved

- Advanced Excel formatting na themes.
- Conditional formatting.
- Pivot tables.
- Macros/VBA.
- External workbook links.
- Embedded Excel charts na objects.

## Saving and Exporting

### Save Project

`Hifadhi` au `Ctrl + S` hupakua `.webexcel.json`. Format hii huhifadhi workbook structure, cell values, internal styles, sheets, merges na settings za Web Excel kwa usahihi zaidi.

### Export Excel

`Excel` hutengeneza `.xlsx` workbook inayoweza kufunguliwa kwenye Microsoft Excel, LibreOffice Calc au Google Sheets.

### Export CSV

CSV export inahifadhi active sheet pekee. CSV haihifadhi formulas kama formulas, sheets nyingi wala cell formatting.

## Autosave

Application huhifadhi workbook automatically kwenye browser kwa kutumia `localStorage`.

- Autosave inabaki kwenye browser na device iliyotumika.
- Kufuta browser data kunaweza kufuta autosave.
- Autosave si mbadala wa kupakua project JSON au Excel file.
- `Restore` inarudisha autosave ya mwisho.
- `Clear data` inafuta autosave bila kufuta workbook inayoonekana kwa wakati huo.

## Keyboard Shortcuts

| Shortcut | Action |
| --- | --- |
| `Ctrl + Z` | Undo |
| `Ctrl + Y` | Redo |
| `Ctrl + S` | Save project JSON |
| `Ctrl + C` | Copy selection |
| `Ctrl + X` | Cut selection |
| `Ctrl + V` | Paste selection |
| `Ctrl + B` | Bold |
| `Ctrl + I` | Italic |
| `Ctrl + U` | Underline |
| `Ctrl + F` | Find and replace |
| `F2` | Edit active cell |
| `Enter` | Edit/confirm cell na kwenda row inayofuata |
| `Tab` | Kwenda cell inayofuata |
| Arrow keys | Kutembea kwenye grid |
| `Delete` / `Backspace` | Clear selected cells |

## Charts

1. Chagua data range.
2. Nenda `Insert`.
3. Chagua Bar Chart, Line Chart au Pie Chart.

Kwa range yenye columns mbili, column ya kwanza hutumika kama labels na column ya pili kama numeric values. Chart inaweza kuburuzwa ndani ya workspace na kufungwa kwa button ya `×`.

## Data Tools

- Ascending na descending sort hutumia column ya kwanza ya selected range kama sort key.
- Remove Duplicates huondoa rows zinazofanana ndani ya selected range.
- Find and Replace hutafuta raw cell content.
- Text to Columns hugawanya text kwa delimiter iliyochaguliwa.
- Fill Series hujaza sequence ya namba ndani ya selection.

## Browser Data and Privacy

- Spreadsheet data inabaki kwenye browser isipokuwa mtumiaji mwenyewe apakue au ahamishe faili.
- Hakuna database connection iliyojumuishwa.
- Hakuna user account, login au cloud synchronization.
- Usihifadhi taarifa nyeti kwenye shared computer bila kufuta autosave baada ya matumizi.

## Customization

Kwa sababu project ipo katika HTML file moja, unaweza kuibadilisha kwa text editor yoyote.

Main sections ndani ya `web_excel.html` ni:

1. `<head>` — metadata, SheetJS CDN na styles.
2. `.titlebar` — top navigation header.
3. `.tabs` na `.ribbon` — Excel-style menus na tools.
4. `.workspace` — spreadsheet grid.
5. `.modal-backdrop` — popup forms.
6. JavaScript IIFE — workbook state, formulas, import/export na interactions.

Theme colors zinaweza kubadilishwa kwenye CSS variables:

```css
:root {
  --excel: #107c41;
  --excel-dark: #0b5d30;
  --excel-soft: #e8f3ed;
}
```

## Troubleshooting

### Excel file halifunguki

- Hakikisha internet ipo ili SheetJS CDN iweze kupakiwa.
- Hakikisha file ni `.xlsx`, `.xls`, `.csv` au `.webexcel.json`.
- Jaribu ku-refresh page na upload tena.

### Export Excel haifanyi kazi

- Angalia internet connection.
- Ruhusu downloads kwenye browser.
- Kama SheetJS haijapakiwa, tumia `Save Project` au `CSV`.

### Data ya autosave haionekani

- Hakikisha unatumia browser na device ileile.
- Bonyeza `File` kisha `Restore`.
- Browser private/incognito mode inaweza kufuta storage baada ya window kufungwa.

### Formula inaonyesha `#ERROR!`

- Hakikisha formula inaanza na `=`.
- Tumia function inayosapotiwa.
- Hakikisha cell references ni sahihi.
- Usitumie result cell ndani ya range yake; hilo linaweza kuleta circular reference.

### Formula inaonyesha `#CIRC!`

Cell inajirejea yenyewe moja kwa moja au kupitia cell nyingine. Badilisha formula au result cell ili kuondoa circular reference.

## Current Limitations

- Si replacement kamili ya Microsoft Excel.
- Hakuna real-time multi-user collaboration.
- Hakuna backend database au cloud synchronization.
- Hakuna Excel macros/VBA.
- Advanced formulas na complex nested functions hazijasapotiwa zote.
- Inserted charts na images ni browser workspace objects na hazihifadhiwi kikamilifu kwenye project/Excel export.
- Advanced formatting kutoka uploaded Excel workbook inaweza kupotea.

## Recommended Backup Practice

Kwa data muhimu:

1. Tumia autosave wakati unaendelea kufanya kazi.
2. Pakua `.webexcel.json` mara kwa mara ili kuhifadhi Web Excel structure.
3. Pakua `.xlsx` kwa matumizi ya Microsoft Excel au applications nyingine.
4. Hifadhi copies kwenye folder salama au external backup.
