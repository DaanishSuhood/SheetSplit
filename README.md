# SheetSplit

Splits an Excel workbook into separate files — either one file per tab, or one sheet cut into blocks
of rows with the header repeated at the top of every block. The source workbook is opened read-only
and is never changed.

Script: [`SheetSplit.ahk`](SheetSplit.ahk), also compiled to a standalone `SheetSplit.exe` (available from
[Releases](https://github.com/DaanishSuhood/SheetSplit/releases/latest)).

---

## Requirements

| Requirement | Notes |
| --- | --- |
| [AutoHotkey v2.0+](https://www.autohotkey.com/) | Only needed to run the `.ahk` script; `SheetSplit.exe` runs without it. Declares `#Requires AutoHotkey v2.0` |
| Microsoft Excel | The script drives Excel over COM, in its own hidden instance |

## Getting started

1. Run `SheetSplit.exe` from the [latest release](https://github.com/DaanishSuhood/SheetSplit/releases/latest)
   (no AutoHotkey install needed), or run `SheetSplit.ahk` with AutoHotkey v2.
2. Drag a workbook onto the window, or click **Browse…**. Accepted: `.xlsx`, `.xlsm`, `.xlsb`, `.xls`,
   and the `.xlt*` template formats.
3. Choose **Split by: Tabs** or **Rows** and set the options.
4. Check the preview at the bottom of the window — it lists exactly what will be written.
5. Click **Split**.

The **Output folder** is filled in for you as `<workbook name>_split` beside the source, and created
if it does not exist.

You can also pass a path on the command line, which makes it usable as a *Send to* target or a drop
target on a shortcut:

```
SheetSplit.exe "C:\path\to\book.xlsx"
```

### Rebuilding the exe

After editing the script, recompile with [Ahk2Exe](https://github.com/AutoHotkey/Ahk2Exe/releases)
(right-click `SheetSplit.ahk` → *Compile Script*, or from the command line):

```
Ahk2Exe.exe /in SheetSplit.ahk /out SheetSplit.exe /base "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
```

The icon (`SheetSplit.ico`) and version info are set by the `;@Ahk2Exe-…` directives at the top of
the script, so keep the `.ico` next to the script when compiling.

## Split by tabs

Every tab is written out as its own `.xlsx`, named after the tab: `Sheet1.xlsx`, `Sheet2.xlsx`, …

| Option | Default | Effect |
| --- | --- | --- |
| **Include hidden sheets** | off | Hidden and very-hidden tabs are skipped unless this is ticked. An exported hidden tab is visible in its own file, since a one-sheet workbook cannot be saved with its only sheet hidden |
| **Skip empty sheets** | **on** | A sheet counts as empty when its used range is a single blank cell. Chart sheets are never treated as empty |
| **Prefix filenames with workbook name** | off | `Book - Sheet1.xlsx` instead of `Sheet1.xlsx` |

Chart sheets are exported too, as a one-sheet workbook holding the chart.

## Split by rows

One sheet is cut into blocks of rows. Each block gets the header rows on top, and keeps the sheet's
formatting, column widths, frozen panes, and page setup.

| Option | Default | Effect |
| --- | --- | --- |
| **Sheet** | first visible sheet | Which worksheet to split. Hidden sheets are listed with `(hidden)` |
| **Header rows** | 1 | How many rows at the top are repeated on every block. `0` for none |
| **Split into _n_ equal blocks** | **4** | Blocks differ in size by at most one row — 10 rows into 4 blocks gives 3, 3, 2, 2 |
| **Blocks of _n_ rows each** | 1000 | Fixed-size blocks; the last one holds whatever is left over |
| **Save each block as its own file** | off | See below |

The row count is found automatically: the data runs from just below the header to the **last row
holding a value or formula**. Rows further down that only carry formatting are ignored. Blank rows
*inside* the data are kept and counted.

Where the blocks go:

| Save each block as its own file | Output |
| --- | --- |
| off | One new workbook, `<workbook>_blocks.xlsx`, with tabs `Block_1`, `Block_2`, … in order |
| on | One file per block: `<workbook>_Block_1.xlsx`, `<workbook>_Block_2.xlsx`, … |

Block numbers are zero-padded to the same width (`Block_01` … `Block_12`) so they sort correctly.
At most 1000 blocks are made in one run. Asking for more blocks than there are data rows gives one
row per block, and the preview says so.

## Output options

These apply to both modes.

| Option | Default | Effect |
| --- | --- | --- |
| **Convert formulas to values** | **on** | Replaces formulas with their current results, keeping formatting. See below |
| **Overwrite existing files** | off | Off: a clashing name becomes ` (2)`, ` (3)`, … up to 199, then that file is skipped. On: the existing file is deleted and replaced |
| **Open output folder when finished** | **on** | Opens the folder in Explorer if at least one file was written |

**Leave _Convert formulas to values_ on unless you specifically need the formulas.** Once a sheet is
on its own, formulas that point at other tabs become links back to the original workbook. In row mode
it is worse: a formula that uses a row outside its block — a running total, a `SUM` over the whole
column, a lookup — shows `#REF!`. With the option on, every cell keeps the value it had in the
original workbook.

## The window

| Control | Purpose |
| --- | --- |
| **Source workbook** | Read-only path box; use **Browse…** or drag and drop |
| **Output folder** | Editable. Auto-fills from the source until you type in it or browse, after which your choice sticks |
| **Split by** | **Tabs** or **Rows** — swaps the options box below it |
| **Split** | Runs the split |
| **Stop** | Stops after the tab or block currently being written |
| Progress bar + status line | `Tab n of total: <name>` / `Block n of total: rows X-Y` |
| Preview / log | Before a run: what will be written with the current settings, updated as you change them. During and after a run: a timestamped log with one line per tab or block — `saved`, `skip (hidden)`, `skip (empty)`, `skip (exists)`, or `FAILED` with the error text |

The window is resizable — the boxes, progress bar, and log follow the width, and the log takes the
extra height.

## How it works

A hidden Excel instance is started (`Visible`, `DisplayAlerts`, `ScreenUpdating`, `EnableEvents` all
off, plus `AskToUpdateLinks` and `AlertBeforeOverwriting` where available) and the workbook is opened
read-only with links not updated. Excel is opened once when you pick the workbook, to read the tab
list for the preview, and again for the run. It is always quit at the end, including when the window
is closed mid-run.

The tab list is **snapshotted up front** — name, visibility, emptiness, and last used row — because
new workbooks get created as the run proceeds and would otherwise shift the indexes. The last row
comes from `Find("*")` searching formulas backwards from `A1`, which, unlike `UsedRange`, skips rows
that only carry formatting.

**Tabs:** each exported tab is copied with `Sheets(i).Copy()` (no arguments, which creates a new
one-sheet workbook), optionally flattened, and saved with `SaveAs(target, 51)` — `51` is
`xlOpenXMLWorkbook`.

**Rows:** the chosen sheet is flattened **once, in the source, in memory** — the source was opened
read-only and is closed without saving, so the file on disk is untouched. Doing it before cutting is
what lets running totals and whole-column formulas keep their original values. Each block then starts
as a full copy of the sheet; the rows below the block are deleted, then the rows between the header
and the block. Copying the whole sheet is what carries column widths, frozen panes, page setup,
conditional formatting, and data validation across. For the single-workbook output, the first copy
creates the workbook and the rest are copied in after the last tab, so there is no leftover blank
`Sheet1`.

Flattening copies the used range over itself with `PasteSpecial(xlPasteValues)`, which keeps text
such as `00123` as text; it falls back to `ur.Value2 := ur.Value2` if that fails.

Excel will not copy a hidden sheet, so a hidden tab is made visible in the in-memory source just
before it is copied.

`Sleep(-1)` calls between tabs and blocks keep the window responsive so **Stop** and the log stay
live during a long run.

Names are made filename-safe: newlines and tabs collapse to spaces, `\ / : * ? " < > |` become `_`,
leading/trailing spaces and dots are trimmed, reserved DOS names (`CON`, `PRN`, `NUL`, `COM1`…) get an
underscore prefix, and names longer than 100 characters are truncated.

## Limitations

- Output is always `.xlsx`. Macros are dropped (a `.xlsm` source gives macro-free files), and defined
  names, pivot caches, and data connections do not survive the copy intact.
- Row mode splits **by position only**. It cannot split by the value in a column (one file per
  region, say).
- The header is always the top rows of the sheet, and the data must start directly beneath it.
- Every block starts as a copy of the whole sheet, so many blocks from a very large sheet take a
  while. 16,000 rows into 4 blocks takes about two seconds.
- **Stop** finishes the current tab or block first. When the blocks are going into one workbook,
  stopping saves nothing rather than leaving a half-finished file.
- **Password-protected workbooks** cannot be opened — remove the password in Excel first.
- A sheet **protected with a password** cannot be split by rows, since rows cannot be deleted from it.
  A sheet protected *without* a password is split, and its blocks are saved unprotected. In tab
  mode, a protected sheet is exported with its formulas, and the log says so.
- Converting formulas to values goes through the Windows clipboard (`PasteSpecial`), so avoid
  copying anything else while a run is going.
- `#SingleInstance Force`, so launching the script again replaces the running copy and abandons a run
  in progress.
- Windows only, and the source is a local path — a workbook open in Excel elsewhere can block the
  read-only open.
