#Requires AutoHotkey v2.0
#SingleInstance Force
;@Ahk2Exe-SetMainIcon SheetSplit.ico
;@Ahk2Exe-SetName SheetSplit
;@Ahk2Exe-SetDescription SheetSplit
;@Ahk2Exe-SetVersion 1.0.0

if !A_IsCompiled && FileExist(A_ScriptDir "\SheetSplit.ico")
    TraySetIcon(A_ScriptDir "\SheetSplit.ico")

; ============================================================================
;  SheetSplit
;  Splits an Excel workbook into separate files - one file per tab, or one
;  sheet cut into blocks of rows with the header repeated on every block.
;
;  Drag a workbook onto the window, or pass a path on the command line.
;  Needs Microsoft Excel installed (the script drives it over COM).
; ============================================================================

; ---- Excel constants -------------------------------------------------------
XLSX             := 51      ; xlOpenXMLWorkbook
XL_SHEET_VISIBLE := -1      ; xlSheetVisible
XL_PASTE_VALUES  := -4163   ; xlPasteValues
XL_FORMULAS      := -4123   ; xlFormulas  (Find: LookIn)
XL_PART          := 2       ; xlPart      (Find: LookAt)
XL_BY_ROWS       := 1       ; xlByRows    (Find: SearchOrder)
XL_PREVIOUS      := 2       ; xlPrevious  (Find: SearchDirection)

MAX_BLOCKS := 1000

S := { src: "", outAuto: true, busy: false, cancel: false, xl: "", tabs: [], wsNames: [] }

BuildGui()
OnExit(QuitOnExit)
if (A_Args.Length)
    SetSource(A_Args[1])


; ============================================================================
;  GUI
; ============================================================================
BuildGui() {
    global S

    g := Gui("+Resize +MinSize600x620", "SheetSplit")
    g.MarginX := 14, g.MarginY := 12
    g.SetFont("s10", "Segoe UI")
    S.gui := g

    g.Add("Text", "x14 y12 w572", "Source workbook  (.xlsx  .xlsm  .xlsb  .xls)")
    S.edSrc  := g.Add("Edit", "x14 y32 w470 h24 ReadOnly", "Drag a workbook onto this window, or click Browse...")
    S.btnSrc := g.Add("Button", "x492 y31 w94 h26", "Browse...")

    g.Add("Text", "x14 y66 w572", "Output folder")
    S.edOut  := g.Add("Edit", "x14 y86 w470 h24")
    S.btnOut := g.Add("Button", "x492 y85 w94 h26", "Browse...")

    g.Add("Text", "x14 y124 w62", "Split by")
    S.rbTabs := g.Add("Radio", "x80 y124 w200 Group Checked", "Tabs - one file per tab")
    S.rbRows := g.Add("Radio", "x290 y124 w290", "Rows - cut one sheet into blocks")

    S.grpMode := g.Add("GroupBox", "x14 y150 w572 h104", "Tab options")

    ; -- Tab mode --
    S.cbHidden    := g.Add("CheckBox", "x28 y174 w280", "Include hidden sheets")
    S.cbSkipEmpty := g.Add("CheckBox", "x28 y200 w280 Checked", "Skip empty sheets")
    S.cbPrefix    := g.Add("CheckBox", "x28 y226 w280", "Prefix filenames with workbook name")
    S.tabCtls := [S.cbHidden, S.cbSkipEmpty, S.cbPrefix]

    ; -- Row mode --
    lbSheet      := g.Add("Text", "x28 y175 w50", "Sheet")
    S.ddSheet    := g.Add("DropDownList", "x80 y171 w222")
    lbHeader     := g.Add("Text", "x320 y175 w88", "Header rows")
    S.edHeader   := g.Add("Edit", "x410 y171 w56 h24 Number", "1")
    udHeader     := g.Add("UpDown", "Range0-50 0x80", 1)
    S.rbCount    := g.Add("Radio", "x28 y203 w90 Group Checked", "Split into")
    S.edCount    := g.Add("Edit", "x120 y200 w70 h24 Number", "4")
    udCount      := g.Add("UpDown", "Range1-" MAX_BLOCKS " 0x80", 4)
    lbCount      := g.Add("Text", "x198 y203 w104", "equal blocks")
    S.rbSize     := g.Add("Radio", "x28 y229 w90", "Blocks of")
    S.edSize     := g.Add("Edit", "x120 y226 w70 h24 Number", "1000")
    lbSize       := g.Add("Text", "x198 y229 w104", "rows each")
    S.cbSeparate := g.Add("CheckBox", "x320 y203 w260", "Save each block as its own file")
    S.lbSeparate := g.Add("Text", "x338 y229 w244 c666666")
    S.rowCtls := [lbSheet, S.ddSheet, lbHeader, S.edHeader, udHeader, S.rbCount, S.edCount, udCount,
        lbCount, S.rbSize, S.edSize, lbSize, S.cbSeparate, S.lbSeparate]

    S.grpOut := g.Add("GroupBox", "x14 y262 w572 h80", "Output options")
    S.cbValues    := g.Add("CheckBox", "x28  y286 w260 Checked", "Convert formulas to values")
    S.cbOverwrite := g.Add("CheckBox", "x300 y286 w270", "Overwrite existing files")
    S.cbOpen      := g.Add("CheckBox", "x28  y312 w260 Checked", "Open output folder when finished")

    S.btnGo    := g.Add("Button", "x14  y354 w130 h32 Default", "Split")
    S.btnStop  := g.Add("Button", "x152 y354 w90  h32 Disabled", "Stop")
    S.btnClose := g.Add("Button", "x486 y354 w100 h32", "Close")

    S.pb     := g.Add("Progress", "x14 y396 w572 h14 Range0-100")
    S.status := g.Add("Text", "x14 y418 w572 h18", "Ready.")

    g.SetFont("s9", "Consolas")
    S.log := g.Add("Edit", "x14 y442 w572 h164 Multi ReadOnly VScroll")
    g.SetFont("s10", "Segoe UI")

    ; Everything the user can change, so a run can lock it all in one go.
    S.inputs := [S.btnSrc, S.btnOut, S.edOut, S.rbTabs, S.rbRows, S.cbValues, S.cbOverwrite, S.cbOpen]
    for c in S.tabCtls
        S.inputs.Push(c)
    for c in S.rowCtls
        S.inputs.Push(c)

    S.btnSrc.OnEvent("Click", PickSource)
    S.btnOut.OnEvent("Click", PickOutput)
    S.btnGo.OnEvent("Click", RunSplit)
    S.btnStop.OnEvent("Click", OnStop)
    S.btnClose.OnEvent("Click", (*) => ExitApp())
    S.edOut.OnEvent("Change", OnOutEdited)
    for c in [S.rbTabs, S.rbRows, S.rbCount, S.rbSize, S.cbHidden, S.cbSkipEmpty, S.cbPrefix, S.cbSeparate]
        c.OnEvent("Click", UpdatePreview)
    for c in [S.ddSheet, S.edHeader, S.edCount, S.edSize]
        c.OnEvent("Change", UpdatePreview)
    g.OnEvent("DropFiles", OnDrop)
    g.OnEvent("Size", OnSize)
    g.OnEvent("Close", (*) => ExitApp())

    g.Show("w600 h620")
    UpdatePreview()
}

OnSize(g, minMax, w, h) {
    global S
    if (minMax = -1)
        return
    m := 14, cw := w - 2 * m
    S.edSrc.Move(, , cw - 102)
    S.btnSrc.Move(m + cw - 94)
    S.edOut.Move(, , cw - 102)
    S.btnOut.Move(m + cw - 94)
    S.grpMode.Move(, , cw)
    S.grpOut.Move(, , cw)
    S.btnClose.Move(m + cw - 100)
    S.pb.Move(, , cw)
    S.status.Move(, , cw)
    S.log.Move(, , cw, h - 442 - m)
    try S.grpMode.Redraw()
    try S.grpOut.Redraw()
    try S.log.Redraw()
}

; Show the controls for the chosen mode and enable the row-count box that
; matches the chosen "Split into" / "Blocks of" option.
ApplyMode() {
    global S
    rows := S.rbRows.Value
    S.grpMode.Text := rows ? "Row options" : "Tab options"
    for c in S.tabCtls
        c.Visible := !rows
    for c in S.rowCtls
        c.Visible := rows
    if (!S.busy) {
        S.edCount.Enabled := S.rbCount.Value
        S.edSize.Enabled  := S.rbSize.Value
    }
    base := S.src = "" ? "<workbook>" : SafeName(SourceBase())
    S.lbSeparate.Value := S.cbSeparate.Value
        ? base "_Block_1.xlsx, _Block_2.xlsx, ..."
        : "Off: one workbook, " base "_blocks.xlsx"
}

OnStop(*) {
    global S
    S.cancel := true
    Status("Stopping after the current " (S.rbRows.Value ? "block" : "tab") "...")
}

OnOutEdited(*) {
    global S
    S.outAuto := false
}

OnDrop(g, ctrl, files, x, y) {
    global S
    if (S.busy)
        return
    if (files.Length > 1)
        Log("Note: " files.Length " files dropped - using the first one only.")
    SetSource(files[1])
}

PickSource(*) {
    global S
    f := FileSelect(3, S.src, "Choose a workbook to split",
        "Excel workbooks (*.xlsx; *.xlsm; *.xlsb; *.xls)")
    if (f != "")
        SetSource(f)
}

PickOutput(*) {
    global S
    d := DirSelect("*" S.edOut.Value, 3, "Choose an output folder")
    if (d != "") {
        S.edOut.Value := d
        S.outAuto := false
    }
}

SetSource(path) {
    global S
    path := Trim(path, ' "')
    if !FileExist(path) {
        Alert("That file does not exist:`n" path)
        return
    }
    if InStr(FileExist(path), "D") {
        Alert("That is a folder, not a workbook:`n" path)
        return
    }
    SplitPath(path, &fname, &dir, &ext, &base)
    if !RegExMatch(ext, "i)^(xlsx|xlsm|xlsb|xls|xltx|xltm|xlt)$") {
        Alert("Not an Excel workbook:`n" fname "`n`nSheetSplit opens .xlsx, .xlsm, .xlsb and .xls files.")
        return
    }

    ; Read the tab list now so the preview and the sheet picker can show it.
    SetBusy(true)
    Status("Reading " base "." ext "...")
    try
        tabs := ReadTabs(path)
    catch as e {
        SetBusy(false)
        Status("Ready.")
        Alert(OpenError(e))
        return
    }
    SetBusy(false)

    S.src := path
    S.tabs := tabs
    S.edSrc.Value := path
    if (S.outAuto || S.edOut.Value = "") {
        S.edOut.Value := dir "\" base "_split"
        S.outAuto := true
    }

    ; Sheet picker for row mode: worksheets only (chart sheets have no rows),
    ; preselecting the first visible one.
    S.wsNames := [], labels := [], pick := 0
    for info in tabs {
        if !info.isWs
            continue
        S.wsNames.Push(info.name)
        labels.Push(info.name (info.visible = XL_SHEET_VISIBLE ? "" : "  (hidden)"))
        if (!pick && info.visible = XL_SHEET_VISIBLE)
            pick := labels.Length
    }
    S.ddSheet.Delete()
    S.ddSheet.Add(labels)
    if (labels.Length)
        S.ddSheet.Choose(pick || 1)

    Status("Ready - " base "." ext "  (" tabs.Length " tab" (tabs.Length = 1 ? "" : "s") ")")
    UpdatePreview()
}


; ============================================================================
;  Preview - what a run would do with the current settings
; ============================================================================
UpdatePreview(*) {
    global S
    if (S.busy)
        return
    ApplyMode()
    if (S.src = "") {
        S.log.Value := "Choose a workbook to see what will be written."
        return
    }
    S.log.Value := S.rbRows.Value ? PreviewRows() : PreviewTabs()
}

PreviewTabs() {
    global S
    srcBase := SourceBase()
    out := [], count := 0
    for info in S.tabs {
        why := TabSkipReason(info)
        if (why != "") {
            out.Push(Line("skip (" why ")", info.name))
        } else {
            out.Push(Line("will save", TabFileBase(info, srcBase) ".xlsx"))
            count++
        }
    }
    text := S.tabs.Length " tab" (S.tabs.Length = 1 ? "" : "s") " in this workbook - "
        . count " file" (count = 1 ? "" : "s") " will be written.`r`n`r`n"
    for l in out
        text .= l "`r`n"
    return RTrim(text, "`r`n")
}

PreviewRows() {
    global S
    plan := RowPlan(SelectedTab())
    if (plan.err != "")
        return plan.err

    count := plan.blocks.Length
    text := "Sheet  : " plan.name "`r`n"
        . "Rows   : " Num(plan.total) " data row" (plan.total = 1 ? "" : "s")
        . " (" plan.first "-" plan.last ")"
        . (plan.hdr ? " below " plan.hdr " header row" (plan.hdr = 1 ? "" : "s") : ", no header") "`r`n"
        . "Output : " (S.cbSeparate.Value
            ? count " file" (count = 1 ? "" : "s")
            : "1 workbook with " count " tab" (count = 1 ? "" : "s")) "`r`n"
        . (plan.note != "" ? "Note   : " plan.note "`r`n" : "")
        . (S.cbValues.Value ? "" : "Note   : formulas are kept - any that use rows outside their block will show #REF!`r`n")
        . "`r`n"

    srcBase := SafeName(SourceBase())
    shown := Min(count, 200)
    Loop shown {
        b := plan.blocks[A_Index]
        label := BlockLabel(A_Index, count)
        text .= Line(S.cbSeparate.Value ? srcBase "_" label ".xlsx" : label, BlockRows(b)) "`r`n"
    }
    if (count > shown)
        text .= "... and " (count - shown) " more"
    return RTrim(text, "`r`n")
}


; ============================================================================
;  Plans - shared by the preview and the run, so they always agree
; ============================================================================
TabSkipReason(info) {
    global S, XL_SHEET_VISIBLE
    if (!S.cbHidden.Value && info.visible != XL_SHEET_VISIBLE)
        return "hidden"
    if (S.cbSkipEmpty.Value && info.empty)
        return "empty"
    return ""
}

TabFileBase(info, srcBase) {
    global S
    base := SafeName(info.name)
    if (S.cbPrefix.Value)
        base := SafeName(srcBase) " - " base
    return base
}

; Works out the row range of every block. Rows are given as source row
; numbers; blocks from "Split into" differ in size by at most one row.
RowPlan(info) {
    global S, MAX_BLOCKS
    if !IsObject(info)
        return { err: "This workbook has no worksheets to split by rows." }
    if (info.lastRow = 0)
        return { err: "'" info.name "' is empty." }

    hdr := ToInt(S.edHeader.Value)
    first := hdr + 1, last := info.lastRow, total := last - hdr
    if (total < 1)
        return { err: "'" info.name "' has no rows below the " hdr " header row" (hdr = 1 ? "" : "s") "." }

    blocks := [], note := ""
    if (S.rbCount.Value) {
        asked := ToInt(S.edCount.Value)
        if (asked < 1)
            return { err: "Enter how many blocks to split into." }
        n := Min(asked, total, MAX_BLOCKS)
        if (n < asked)
            note := "Only " Num(n) " block" (n = 1 ? "" : "s") " - "
                . (n = total ? "there are just " Num(total) " data rows." : "the limit is " MAX_BLOCKS ".")
        size := total // n, extra := Mod(total, n), r := first
        Loop n {
            len := size + (A_Index <= extra)
            blocks.Push({ s: r, e: r + len - 1 })
            r += len
        }
    } else {
        size := ToInt(S.edSize.Value)
        if (size < 1)
            return { err: "Enter how many rows each block should have." }
        n := Ceil(total / size)
        if (n > MAX_BLOCKS)
            return { err: "That makes " Num(n) " blocks - the limit is " MAX_BLOCKS ". Use bigger blocks." }
        r := first
        Loop n {
            blocks.Push({ s: r, e: Min(r + size - 1, last) })
            r += size
        }
    }
    return { err: "", note: note, name: info.name, hdr: hdr, first: first, last: last, total: total, blocks: blocks }
}

SelectedTab() {
    global S
    i := S.ddSheet.Value
    if (!i || i > S.wsNames.Length)
        return ""
    return FindTab(S.tabs, S.wsNames[i])
}

FindTab(tabs, name) {
    for info in tabs
        if (info.name = name)
            return info
    return ""
}


; ============================================================================
;  The actual split
; ============================================================================
RunSplit(*) {
    global S

    if (S.busy)
        return
    if (S.src = "" || !FileExist(S.src)) {
        Alert("Pick a source workbook first.")
        return
    }
    if (S.rbRows.Value) {
        plan := RowPlan(SelectedTab())
        if (plan.err != "") {
            Alert(plan.err)
            return
        }
    }

    outDir := Trim(S.edOut.Value)
    if (outDir = "") {
        Alert("Pick an output folder.")
        return
    }
    if !DirExist(outDir) {
        try
            DirCreate(outDir)
        catch as e {
            Alert("Could not create the output folder:`n" outDir "`n`n" ErrText(e))
            return
        }
    }

    SetBusy(true)
    S.log.Value := ""
    Log("=== " FormatTime(, "yyyy-MM-dd HH:mm:ss") " ===")
    Log("Source: " S.src)
    Log("Output: " outDir)
    Log("")

    try
        S.xl := StartExcel()
    catch as e {
        Log("Could not start Excel: " ErrText(e))
        Alert("Could not start Excel. Is Microsoft Excel installed?")
        SetBusy(false)
        return
    }

    Status("Opening workbook...")
    try
        wb := OpenReadOnly(S.xl, S.src)
    catch as e {
        Log("Could not open the workbook: " ErrText(e))
        Alert(OpenError(e))
        QuitExcel()
        SetBusy(false)
        return
    }

    ; Snapshot the tabs again - the file may have changed since it was picked,
    ; and new workbooks get created as we go.
    tabs := SnapshotTabs(wb)
    res := S.rbRows.Value
        ? SplitRows(S.xl, wb, tabs, outDir)
        : SplitTabs(S.xl, wb, tabs, outDir)

    try wb.Close(false)
    wb := ""
    QuitExcel()

    S.pb.Value := 100
    Log("")
    Log("Done. " res.made " file" (res.made = 1 ? "" : "s") " written"
        . (res.skipped ? ", " res.skipped " skipped" : "")
        . (res.failed ? ", " res.failed " failed" : "") ".")
    Status(res.made " file" (res.made = 1 ? "" : "s") " written to " outDir)
    SetBusy(false)

    if (res.made && S.cbOpen.Value)
        try Run('explorer.exe "' outDir '"')
}

; One file per tab.
SplitTabs(xl, wb, tabs, outDir) {
    global S

    Log("Found " tabs.Length " tab" (tabs.Length = 1 ? "" : "s") " in the workbook.")
    Log("")

    srcBase := SourceBase()
    res := { made: 0, skipped: 0, failed: 0 }

    for i, info in tabs {
        Sleep(-1)                                   ; keep the GUI responsive / allow Stop
        if (S.cancel) {
            Log("")
            Log("Stopped by user.")
            break
        }
        Status("Tab " i " of " tabs.Length ": " info.name)

        why := TabSkipReason(info)
        if (why != "") {
            Log(Line("skip (" why ")", info.name))
            res.skipped++
            Advance(i, tabs.Length)
            continue
        }

        base := TabFileBase(info, srcBase)
        target := PickPath(outDir "\" base ".xlsx", S.cbOverwrite.Value)
        if (target = "") {
            Log(Line("skip (exists)", base ".xlsx"))
            res.skipped++
            Advance(i, tabs.Length)
            continue
        }

        nwb := ""
        try {
            Unhide(wb.Sheets(info.idx))
            wb.Sheets(info.idx).Copy()              ; no args -> new one-sheet workbook
            nwb := xl.ActiveWorkbook
            kept := S.cbValues.Value && info.isWs && !Flatten(xl, nwb.Sheets(1))
            SaveNew(nwb, target)
            nwb := ""
            res.made++
            SplitPath(target, &tname)
            Log(Line("saved", tname (kept ? "  (formulas kept - could not convert)" : "")))
        } catch as e {
            res.failed++
            Log(Line("FAILED", info.name "  ->  " ErrText(e)))
            if (nwb != "")
                try nwb.Close(false)
        }
        Advance(i, tabs.Length)
    }
    return res
}

; One sheet cut into blocks of rows, each block with the header rows on top.
; Every block starts as a full copy of the sheet - so column widths, page
; setup, frozen panes and formatting all come along - and the rows outside
; the block are then deleted.
SplitRows(xl, wb, tabs, outDir) {
    global S

    res := { made: 0, skipped: 0, failed: 0 }
    name := S.wsNames[S.ddSheet.Value]
    info := FindTab(tabs, name)
    plan := RowPlan(info)
    if (plan.err != "") {
        Log("Cannot split: " (IsObject(info) ? plan.err : "'" name "' is no longer in the workbook."))
        res.failed++
        return res
    }

    src := wb.Sheets(info.idx)
    Unhide(src)
    if (src.ProtectContents) {
        ; An empty password unprotects a sheet that has none, and fails
        ; (instead of prompting) on one that does.
        try
            src.Unprotect("")
        catch {
            Log("Cannot split: '" name "' is protected with a password.")
            res.failed++
            return res
        }
        Log("Note: '" name "' is protected - the blocks are saved unprotected.")
    }
    ; Flatten the source once, in memory only - it was opened read-only and is
    ; never saved. Doing it before the cut means formulas that look at rows
    ; outside a block still give the values they had in the original.
    if (S.cbValues.Value && !Flatten(xl, src))
        Log("Note: could not convert formulas to values - they were kept.")

    count := plan.blocks.Length
    Log("Sheet '" name "': " Num(plan.total) " data rows (" plan.first "-" plan.last "), "
        . plan.hdr " header row" (plan.hdr = 1 ? "" : "s") ", " count " block" (count = 1 ? "" : "s") ".")
    Log("")

    srcBase := SafeName(SourceBase())

    if (S.cbSeparate.Value) {
        for i, b in plan.blocks {
            Sleep(-1)
            if (S.cancel) {
                Log("")
                Log("Stopped by user.")
                break
            }
            label := BlockLabel(i, count)
            Status("Block " i " of " count ": " BlockRows(b))

            target := PickPath(outDir "\" srcBase "_" label ".xlsx", S.cbOverwrite.Value)
            if (target = "") {
                Log(Line("skip (exists)", srcBase "_" label ".xlsx"))
                res.skipped++
                Advance(i, count)
                continue
            }

            nwb := ""
            try {
                src.Copy()                          ; no args -> new one-sheet workbook
                nwb := xl.ActiveWorkbook
                CutToBlock(nwb.Sheets(1), plan, b, label)
                SaveNew(nwb, target)
                nwb := ""
                res.made++
                SplitPath(target, &tname)
                Log(Line("saved", tname "  " BlockRows(b)))
            } catch as e {
                res.failed++
                Log(Line("FAILED", label "  ->  " ErrText(e)))
                if (nwb != "")
                    try nwb.Close(false)
            }
            Advance(i, count)
        }
        return res
    }

    ; All blocks as tabs of one new workbook.
    target := PickPath(outDir "\" srcBase "_blocks.xlsx", S.cbOverwrite.Value)
    if (target = "") {
        Log(Line("skip (exists)", srcBase "_blocks.xlsx"))
        res.skipped++
        return res
    }
    nwb := "", label := ""
    try {
        for i, b in plan.blocks {
            Sleep(-1)
            if (S.cancel)
                break
            label := BlockLabel(i, count)
            Status("Block " i " of " count ": " BlockRows(b))
            if (nwb = "") {
                src.Copy()                          ; first block makes the workbook
                nwb := xl.ActiveWorkbook
            } else {
                src.Copy(, nwb.Sheets(nwb.Sheets.Count))
            }
            CutToBlock(nwb.Sheets(nwb.Sheets.Count), plan, b, label)
            Log(Line("tab", label "  " BlockRows(b)))
            Advance(i, count)
        }
        if (S.cancel) {
            if (nwb != "")
                nwb.Close(false)
            Log("")
            Log("Stopped by user - " srcBase "_blocks.xlsx was not saved.")
            return res
        }
        nwb.Sheets(1).Activate()
        SaveNew(nwb, target)
        res.made++
        SplitPath(target, &tname)
        Log("")
        Log(Line("saved", tname "  (" count " tab" (count = 1 ? "" : "s") ")"))
    } catch as e {
        res.failed++
        Log(Line("FAILED", (label = "" ? srcBase "_blocks.xlsx" : label) "  ->  " ErrText(e)))
        if (nwb != "")
            try nwb.Close(false)
    }
    return res
}

; Trim a full copy of the sheet down to the header rows plus one block.
CutToBlock(sh, plan, b, label) {
    if (b.e < plan.last)
        sh.Rows((b.e + 1) ":" plan.last).Delete()
    if (b.s > plan.first)
        sh.Rows(plan.first ":" (b.s - 1)).Delete()
    sh.Name := label
    try sh.Range("A1").Select()
}

; Excel refuses to copy a hidden sheet, and a one-sheet workbook can't be
; saved with its only sheet hidden - so show it first. This only touches the
; read-only source in memory; it is never saved.
Unhide(sh) {
    global XL_SHEET_VISIBLE
    if (sh.Visible != XL_SHEET_VISIBLE)
        sh.Visible := XL_SHEET_VISIBLE
}

SaveNew(nwb, target) {
    global XLSX
    if FileExist(target)
        FileDelete(target)
    nwb.SaveAs(target, XLSX)
    nwb.Close(false)
}

Advance(i, total) {
    global S
    S.pb.Value := Round(i / total * 100)
}


; ============================================================================
;  Excel
; ============================================================================
; A COM-started Excel is already hidden. Don't set Visible := false anyway:
; with a workbook opened afterwards, that keeps the process alive after Quit().
StartExcel() {
    xl := ComObject("Excel.Application")
    xl.DisplayAlerts := false
    xl.ScreenUpdating := false
    xl.EnableEvents := false
    try xl.AskToUpdateLinks := false
    try xl.AlertBeforeOverwriting := false
    return xl
}

; UpdateLinks = 0, ReadOnly = true, and an empty password so a protected
; workbook fails to open instead of prompting in the hidden instance.
OpenReadOnly(xl, path) => xl.Workbooks.Open(path, 0, true, , "")

ReadTabs(path) {
    xl := StartExcel()
    try {
        wb := OpenReadOnly(xl, path)
        tabs := SnapshotTabs(wb)
        wb.Close(false)
    } finally {
        Quit(xl)
    }
    return tabs
}

SnapshotTabs(wb) {
    tabs := []
    Loop wb.Sheets.Count {
        sh := wb.Sheets(A_Index)
        isWs := false, lastRow := 0
        try {
            ur := sh.UsedRange                      ; chart sheets have no cells
            isWs := true
            lastRow := FindLastRow(sh)
        }
        tabs.Push({ idx: A_Index, name: sh.Name, visible: sh.Visible, isWs: isWs,
                    empty: IsEmptySheet(sh), lastRow: lastRow })
    }
    return tabs
}

; Last row holding anything - a value or a formula, even one showing "".
; Unlike UsedRange this ignores rows that only carry formatting.
FindLastRow(sh) {
    global XL_FORMULAS, XL_PART, XL_BY_ROWS, XL_PREVIOUS
    try return sh.Cells.Find("*", sh.Cells(1, 1), XL_FORMULAS, XL_PART, XL_BY_ROWS, XL_PREVIOUS).Row
    return 0
}

; Replace formulas with their results, leaving formatting alone.
Flatten(xl, sh) {
    global XL_PASTE_VALUES
    try {
        ur := sh.UsedRange
        ur.Copy()
        ur.PasteSpecial(XL_PASTE_VALUES)
        xl.CutCopyMode := 0
        return true
    } catch {
        try {
            ur := sh.UsedRange
            ur.Value2 := ur.Value2                  ; fallback
            return true
        }
    }
    return false
}

IsEmptySheet(sh) {
    try {
        if (sh.UsedRange.Cells.Count > 1)
            return false
        return (sh.Range("A1").Text = "")
    }
    return false                                    ; chart sheets etc. are never "empty"
}

Quit(xl) {
    try xl.ScreenUpdating := true
    try xl.EnableEvents := true
    try xl.DisplayAlerts := true
    try xl.Quit()
}

QuitExcel() {
    global S
    if (S.xl != "")
        Quit(S.xl)
    S.xl := ""
}

; Closing the window mid-run must not leave a hidden Excel behind.
QuitOnExit(*) {
    QuitExcel()
}


; ============================================================================
;  Helpers
; ============================================================================
SourceBase() {
    global S
    SplitPath(S.src, , , , &base)
    return base
}

BlockLabel(i, count) => "Block_" Format("{:0" StrLen(count) "}", i)

BlockRows(b) => b.s = b.e ? "row " b.s : "rows " b.s "-" b.e "  (" Num(b.e - b.s + 1) ")"

Line(tag, text) => Format("{:-14}: {}", tag, text)

Num(n) => RegExReplace(n, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0,")

ToInt(v) => IsInteger(v) ? Integer(v) : 0

SafeName(name) {
    name := RegExReplace(name, "[\r\n\t]+", " ")
    name := RegExReplace(name, '[\\/:*?"<>|]', "_")
    name := Trim(name, " .")
    if (name = "")
        name := "Sheet"
    if RegExMatch(name, "i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$")
        name := "_" name
    if (StrLen(name) > 100)
        name := Trim(SubStr(name, 1, 100))
    return name
}

PickPath(path, overwrite) {
    if (overwrite || !FileExist(path))
        return path
    SplitPath(path, , &dir, &ext, &base)
    Loop 199 {
        cand := dir "\" base " (" (A_Index + 1) ")." ext
        if !FileExist(cand)
            return cand
    }
    return ""
}

; Excel words a password-protected workbook as "the password you supplied is
; not correct" - confusing when nobody typed one.
OpenError(e) {
    if InStr(ErrText(e), "password")
        return "This workbook is password-protected, so SheetSplit can't open it.`n`n"
            . "Remove the password in Excel first (File > Info > Protect Workbook)."
    return "Could not open that workbook.`n`n" ErrText(e)
}

ErrText(e) {
    if !(e is Error)
        return String(e)
    t := e.Message
    if (e.Extra != "")
        t .= " (" e.Extra ")"
    return RegExReplace(t, "\s+", " ")
}

SetBusy(on) {
    global S
    S.busy := on
    if (on) {
        S.cancel := false
        S.pb.Value := 0
    }
    for c in S.inputs
        c.Enabled := !on
    S.btnGo.Enabled := !on
    S.btnStop.Enabled := on
    if (!on)
        ApplyMode()
    Sleep(-1)
}

Status(text) {
    global S
    S.status.Value := text
    Sleep(-1)
}

Log(text) {
    global S
    S.log.Value .= (S.log.Value = "" ? "" : "`r`n") text
    try SendMessage(0x0115, 7, 0, S.log)            ; WM_VSCROLL / SB_BOTTOM
    Sleep(-1)
}

Alert(msg) {
    global S
    MsgBox(msg, "SheetSplit", "Icon! Owner" S.gui.Hwnd)
}
