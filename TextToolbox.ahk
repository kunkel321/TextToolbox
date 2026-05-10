#Requires AutoHotkey v2.0
#SingleInstance Force
TraySetIcon("shell32.dll", 75)

; ============================================================
;  App:             TextToolbox.ahk
;  Version Date:    5-10-2026 
;  By:              Kunkel321 with Claude AI
;  GitHub:          https://github.com/kunkel321/TextToolbox
;  AHK Forum:       https://www.autohotkey.com/boards/viewtopic.php?f=83&t=140654
;  Description:     Dual-pane text transformation utility based on Tidbit's 
;  same-name tool.  (Also see onboard help via F1 key.)
;  Settings persisted to ttSettings.ini (same folder as script).
; ============================================================

; ---------- Config / Globals --------------------------------
IniFile      := A_ScriptDir "\ttSettings.ini"
UndoStack    := []      ; Array of Maps {in, out}, index 1 = most recent
RedoStack    := []      ; Array of Maps {in, out}, index 1 = most recent
UNDO_MAX     := 20     ; max undo/redo levels
HISTORY_MAX  := 10   ; max items remembered per combo history
FindHistory  := []   ; Array of find strings, index 1 = most recent
ReplHistory  := []   ; Array of replace strings, index 1 = most recent

; ---------- User option -------------------------------------
; Set to true to automatically paste clipboard contents into
; the Input pane when the script starts.
AUTO_PASTE_CLIPBOARD := true

; Set to true to enable debug logging to ttDebug.log in the script folder.
CODE_DEBUG_LOG := false

Debug(msg) {
    global CODE_DEBUG_LOG
    if CODE_DEBUG_LOG
        FileAppend("Debug: " FormatTime(A_Now, "MMM-dd hh:mm:ss") ": " msg "`n",
                   A_ScriptDir "\ttDebug.log")
}

; Snapshot both panes onto the undo stack; clears the redo stack.
; Call this BEFORE any operation that modifies EditIn or EditOut.
PushUndo(inTxt, outTxt) {
    global UndoStack, RedoStack, UNDO_MAX
    UndoStack.InsertAt(1, Map("in", inTxt, "out", outTxt))
    while (UndoStack.Length > UNDO_MAX)
        UndoStack.Pop()
    RedoStack := []   ; a new action invalidates the redo chain
}

PAD      := 10
BTN_H    := 28
BTN_W    := 100
TAB_H    := 220   ; fixed height of the tab/controls strip (top section)
LB_W     := 118   ; width of the left-side navigation listbox
PANE_H   := 280   ; initial height of the dual edit-pane area (resizable)
WIN_W    := 780
TOOL_H   := 26    ; toolbar row height (fixed)
; Layout (top to bottom):
;   PAD  →  tab/listbox area (TAB_H)  →  PAD  →  toolbar (TOOL_H)  →  PAD  →  panes (PANE_H)  →  PAD
WIN_H    := PAD + TAB_H + PAD + TOOL_H + PAD + PANE_H + PAD
SplitH   := true  ; true = side-by-side, false = stacked
STRIP_W  := 24    ; width of the narrow vertical strip in vert mode

; ============================================================
;  LOAD SETTINGS
; ============================================================
LoadSettings() {
    global WIN_W, WIN_H, SplitH, PANE_H, PAD, TAB_H, TOOL_H
    WIN_W  := Integer(IniRead(IniFile, "Window", "W",      WIN_W))
    WIN_H  := Integer(IniRead(IniFile, "Window", "H",      WIN_H))
    PANE_H := Integer(IniRead(IniFile, "Window", "PaneH",  PANE_H))
    SplitH := IniRead(IniFile, "Window", "SplitH", "1") = "1"
}

; ============================================================
;  BUILD GUI
; ============================================================
LoadSettings()

g := Gui("+Resize +MinSize520x400", "Text Toolbox")
g.BackColor := "F0F0F0"
g.SetFont("s9", "Segoe UI")

; ---------- TAB AREA (top, fixed position) -------------------
; LbNav and Tabs sit at y=PAD and never move.
; Tab child controls are parented here and show/hide automatically — no repositioning needed.
tabList := ["Case", "Sort", "Find/Replace", "Remove", "Extract", "Wrap/Indent",
            "Counter", "Padding", "CSV View", "Compare", "N-Grams", "Stats"]

LbNav := g.Add("ListBox",
    "x" PAD " y" PAD " w" LB_W " h" TAB_H " AltSubmit", tabList)

; Tab control: Buttons style so the tab strip can be hidden.
; GetPos captures TX,TY (the content-area origin) before we collapse the strip.
Tabs := g.Add("Tab", "x+m yp w" (WIN_W - PAD*2 - LB_W) " h" TAB_H " Buttons", tabList)
Tabs.GetPos(&TX, &TY)
; Collapse the Tab control to 0×0 — this hides the button strip while leaving
; child controls visible (Tab1 children are GUI-parented, not Tab-parented).
; The LbNav sits in front of where the strip was, covering any remnant rendering.
Tabs.Move(, , 0, 0)

BuildTab_Case()
BuildTab_Sort()
BuildTab_FindReplace()
BuildTab_Remove()
BuildTab_Extract()
BuildTab_WrapIndent()
BuildTab_Counter()
BuildTab_Padding()
BuildTab_CsvView()
BuildTab_Compare()
BuildTab_Stats()
BuildTab_NGrams()

Tabs.UseTab(0)

; ---------- Toolbar row (below tab area, fixed position) -----
toolY := PAD + TAB_H + PAD
g.Add("Text",    "x10          y" (toolY+2) " w55",      "Input:")
BtnPaste  := g.Add("Button",   "x68          y" toolY  " w52 h22", "Paste")
BtnClrIn  := g.Add("Button",   "x122         y" toolY  " w52 h22", "Clear")
ChkWrap   := g.Add("Checkbox", "x180         y" (toolY+2) " Checked", "Wrap")
BtnSplit  := g.Add("Button",   "x236         y" toolY  " w72 h22", "⇕ Vert")
g.Add("Text",    "x318         y" (toolY+2) " w55",      "Output:")
BtnCopy   := g.Add("Button",   "x370         y" toolY  " w72 h22", "Copy Out")
BtnClrOut := g.Add("Button",   "x444         y" toolY  " w52 h22", "Clear")
g.Add("Text",    "x502         y" (toolY+2) " w4",       "|")
BtnUndo   := g.Add("Button",   "x510         y" toolY  " w56 h22", "Undo")
BtnRedo   := g.Add("Button",   "x568         y" toolY  " w56 h22", "Redo")
BtnHelp   := g.Add("Button",   "x630         y" toolY  " w26 h22", "?")

; ---------- Dual panes + mid-strip (positions set by LayoutPanes) ---
EditIn   := g.Add("Edit",   "x0 y0 w10 h10 Multi VScroll Wrap", "")
EditOut  := g.Add("Edit",   "x0 y0 w10 h10 Multi VScroll Wrap", "")
BtnApply := g.Add("Button", "x0 y0 w10 h10 Default", "Apply >")
BtnSwap  := g.Add("Button", "x0 y0 w10 h10",         "Swap ↔")

; ---------- EVENTS ------------------------------------------
BtnPaste.OnEvent("Click",  OnPaste)
BtnClrIn.OnEvent("Click",  OnClearInput)
BtnClrOut.OnEvent("Click", OnClearOutput)
BtnSplit.OnEvent("Click",  OnToggleSplit)
ChkWrap.OnEvent("Click",   OnToggleWrap)
BtnApply.OnEvent("Click",  OnApply)
BtnUndo.OnEvent("Click",   OnUndo)
BtnRedo.OnEvent("Click",   OnRedo)
BtnSwap.OnEvent("Click",   OnSwap)
BtnCopy.OnEvent("Click",   OnCopyOutput)
BtnHelp.OnEvent("Click",   (*) => ShowHelp(0))   ; 0 = general overview
LbNav.OnEvent("Change",    OnNavChange)
g.OnEvent("Size",          OnGuiSize)
g.OnEvent("Close",         OnClose)

; Hotkeys — active only while Text Toolbox window is in the foreground
; ^z / ^y are skipped when an Edit control has focus (let Windows handle native undo there)
#HotIf WinActive("Text Toolbox ahk_class AutoHotkeyGUI") && !EditHasFocus()
^z:: OnUndo()
^y:: OnRedo()
#HotIf WinActive("Text Toolbox ahk_class AutoHotkeyGUI")
Esc:: OnClose(g)
F1::  ShowHelp(Tabs.Value)   ; F1 = help for the current tab
#HotIf

; Returns true when the focused control is an Edit (so ^z/^y stay native inside edit boxes)
EditHasFocus() {
    focHwnd := ControlGetFocus("A")
    try {
        cls := WinGetClass("ahk_id " focHwnd)
        return (cls = "Edit" || cls = "RichEdit20W" || cls = "RICHEDIT50W")
    }
    return false
}

; ---------- RESTORE last tab & do initial layout ------------
lastTab := Integer(IniRead(IniFile, "Window", "LastTab", 1))
Tabs.Choose(lastTab)
LbNav.Value := lastTab

UpdateSplitBtn()
g.Show("w" WIN_W " h" WIN_H)
; Trigger layout once the message loop is running and client area is real
SetTimer(InitialLayout, -50)

InitialLayout() {
    global AUTO_PASTE_CLIPBOARD, EditIn
    g.GetClientPos(, , &cW, &cH)
    LayoutPanes(cW, cH)
    LoadHistory()
    if (AUTO_PASTE_CLIPBOARD && A_Clipboard != "")
        EditIn.Value := A_Clipboard
}

; ============================================================
;  PANE LAYOUT  — called on init and every resize
; ============================================================
; ============================================================
;  PANE LAYOUT  — called on init and every resize
;  Tab area is fixed at top; only the edit panes stretch.
; ============================================================
LayoutPanes(W, H) {
    global PAD, TAB_H, TOOL_H, PANE_H, SplitH, STRIP_W, LB_W
    global EditIn, EditOut, BtnApply, BtnSwap

    innerW := W - PAD*2
    paneY  := PAD + TAB_H + PAD + TOOL_H + PAD

    ; Panes fill all remaining vertical space
    PANE_H := Max(H - paneY - PAD, 80)

    if SplitH {
        ; Side-by-side: narrow strip column between panes
        stripX := PAD + (innerW - STRIP_W) // 2
        leftW  := stripX - PAD
        rightX := stripX + STRIP_W
        rightW := PAD + innerW - rightX

        EditIn.Move( PAD,    paneY, leftW,  PANE_H)
        EditOut.Move(rightX, paneY, rightW, PANE_H)

        halfBtnH := (PANE_H - PAD) // 2
        BtnApply.Move(stripX, paneY,                  STRIP_W, halfBtnH)
        BtnSwap.Move( stripX, paneY + halfBtnH + PAD, STRIP_W, PANE_H - halfBtnH - PAD)
    } else {
        ; Stacked: Apply + Swap span the full width between the panes, growing with window
        stripH   := 20
        gap      := PAD                        ; small gap between the two buttons
        btnWide  := (innerW - gap) // 2        ; each button gets half of inner width
        btnApplyX := PAD
        btnSwapX  := PAD + btnWide + gap
        halfH    := (PANE_H - stripH) // 2

        EditIn.Move( PAD, paneY,                  innerW, halfH)
        EditOut.Move(PAD, paneY + halfH + stripH, innerW, PANE_H - halfH - stripH)

        BtnApply.Move(btnApplyX, paneY + halfH, btnWide, stripH)
        BtnSwap.Move( btnSwapX,  paneY + halfH, btnWide, stripH)
    }

    ; LbNav height tracks TAB_H (fixed, but resize it in case TAB_H ever changes)
    LbNav.Move(PAD, PAD, LB_W, TAB_H)
}

UpdateSplitBtn() {
    global BtnSplit, BtnApply, BtnSwap, SplitH
    if SplitH {
        ; Side-by-side panes, vertical divider between them = Vertical split
        BtnSplit.Text := "⇕ Vert"
        BtnApply.Text := "▶`nA`np`np`nl`ny`n▶"
        BtnSwap.Text  := "▶`nS`nw`na`np`n◀"
    } else {
        ; Stacked panes, horizontal divider between them = Horizontal split
        BtnSplit.Text := "⇔ Horiz"
        BtnApply.Text := "▼ Apply ▼"
        BtnSwap.Text  := "▼ Swap ▲"
    }
}

; ============================================================
;  TAB BUILDERS
; ============================================================

BuildTab_Case() {
    global Tabs, g, TX, TY
    Tabs.UseTab(1)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Convert case of each line:")
    global CaseR1  := g.Add("Radio", "x" (TX+8) "  y+8 Group", "UPPERCASE")
    global CaseR2  := g.Add("Radio", "x" (TX+8) "  y+4",       "lowercase")
    global CaseR3  := g.Add("Radio", "x" (TX+8) "  y+4",       "Title Case")
    global CaseR4  := g.Add("Radio", "x" (TX+8) "  y+4",       "Sentence case")
    global CaseR5  := g.Add("Radio", "x" (TX+8) "  y+4",       "iNVERT cAsE")
    global CaseR6  := g.Add("Radio", "x" (TX+8) "  y+4",       "sArCaStIc CaSe")
    CaseR1.GetPos(, &cRgY)
    global CaseR7  := g.Add("Radio", "x" (TX+190) " y" cRgY,    "kebab-case")
    global CaseR8  := g.Add("Radio", "x" (TX+190) " y+4",       "snake_case")
    global CaseR9  := g.Add("Radio", "x" (TX+190) " y+4",       "PascalCase")
    global CaseR10 := g.Add("Radio", "x" (TX+190) " y+4",       "camelCase")
    global CaseR11 := g.Add("Radio", "x" (TX+190) " y+4",       "SCREAMING_SNAKE_CASE")
    global CaseR12 := g.Add("Radio", "x" (TX+190) " y+4",       "Train-Case")
}

BuildTab_Sort() {
    global Tabs, g, TX, TY
    Tabs.UseTab(2)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Sort lines:")
    global SortR1  := g.Add("Radio", "x" (TX+8) " y+8 Group", "A → Z")
    global SortR2  := g.Add("Radio", "x" (TX+8) " y+4",       "Z → A")
    global SortR3  := g.Add("Radio", "x" (TX+8) " y+4",       "Random")
    global SortR4  := g.Add("Radio", "x" (TX+8) " y+4",       "Length (short→long)")
    global SortR5  := g.Add("Radio", "x" (TX+8) " y+4",       "Length (long→short)")
    global SortR6  := g.Add("Radio", "x" (TX+8) " y+4",       "Numeric (natural sort)")
    global ChkSortDupe := g.Add("Checkbox", "x" (TX+8) " y+10", "Remove duplicate lines")
    global ChkSortTrim := g.Add("Checkbox", "x" (TX+8) " y+4",  "Trim whitespace before sorting")
    g.Add("Text", "x" (TX+8) " y+10", "Sort key:")
    global DdlSortKey := g.Add("DropDownList",
        "x" (TX+62) " yp-3 w148 Choose1",
        ["Whole line", "Word 1", "Word 2", "Last word", "After delimiter"])
    global EditSortDelim := g.Add("Edit", "x+8 yp w28 h22", "-")
    global TxtSortDelim  := g.Add("Text", "x+4 yp+3", "delim")
    ; Enable/disable delimiter box based on DDL selection
    DdlSortKey.OnEvent("Change", (*) => UpdateSortKeyUI())
    UpdateSortKeyUI()
}

UpdateSortKeyUI() {
    global DdlSortKey, EditSortDelim, TxtSortDelim
    isDelim := (DdlSortKey.Value = 5)   ; "After delimiter" is item 5
    EditSortDelim.Enabled := isDelim
    TxtSortDelim.Enabled  := isDelim
}

BuildTab_FindReplace() {
    global Tabs, g, TX, TY
    Tabs.UseTab(3)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Find:")
    global CboFind := g.Add("ComboBox", "x" (TX+8) " y+4 w380", [])
    g.Add("Text", "x" (TX+8) " y+8", "Replace with:")
    global CboRepl := g.Add("ComboBox", "x" (TX+8) " y+4 w380", [])
    global ChkFRCase  := g.Add("Checkbox", "x" (TX+8) " y+10", "Case sensitive")
    global ChkFRRegex := g.Add("Checkbox", "x" (TX+8) " y+4",  "Use regex")
    global ChkFRAll   := g.Add("Checkbox", "x" (TX+8) " y+4 Checked", "Replace all occurrences")
}

BuildTab_Remove() {
    global Tabs, g, TX, TY
    Tabs.UseTab(4)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Remove from text:")
    global ChkRmBlank    := g.Add("Checkbox", "x" (TX+8) " y+8",  "Blank / empty lines")
    global ChkRmDupe     := g.Add("Checkbox", "x" (TX+8) " y+4",  "Duplicate lines")
    global ChkRmLead     := g.Add("Checkbox", "x" (TX+8) " y+4",  "Leading whitespace (per line)")
    global ChkRmTrail    := g.Add("Checkbox", "x" (TX+8) " y+4",  "Trailing whitespace (per line)")
    global ChkRmHTML     := g.Add("Checkbox", "x" (TX+8) " y+4",  "HTML tags  <…>")
    global ChkRmBBCode   := g.Add("Checkbox", "x" (TX+8) " y+4",  "BBCode tags  […]")
    global ChkRmNonASCII := g.Add("Checkbox", "x" (TX+8) " y+4",  "Non-ASCII characters")
    g.Add("Text", "x" (TX+8) " y+10", "Trim characters from LEFT:")
    global NumTrimL := g.Add("Edit", "x+6 yp-2 w40 h22 Number", "0")
    g.Add("Text", "x+8 yp+2", "RIGHT:")
    global NumTrimR := g.Add("Edit", "x+6 yp-2 w40 h22 Number", "0")
}

BuildTab_Extract() {
    global Tabs, g, TX, TY
    Tabs.UseTab(5)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Extract from text:")

    ; Row 1 — what to extract (radio group)
    global ExtR1 := g.Add("Radio", "x" (TX+8)   " y+10 Group", "Numbers")
    global ExtR2 := g.Add("Radio", "x" (TX+100)  " yp",         "Email addresses")
    global ExtR3 := g.Add("Radio", "x" (TX+240)  " yp",         "URLs")
    global ExtR4 := g.Add("Radio", "x" (TX+8)   " y+6",         "Lines matching pattern:")
    global ExtR5 := g.Add("Radio", "x" (TX+8)   " y+4",         "Custom pattern (regex):")

    ; Pattern edit boxes — shown/enabled based on radio selection
    global EditExtLinePattern := g.Add("Edit", "x" (TX+175) " yp-26 w250 h22", "")
    global EditExtPattern     := g.Add("Edit", "x" (TX+175) " yp+26 w250 h22", "")

    ; Row 2 — output options
    g.Add("Text", "x" (TX+8) " y+12", "Output options:")
    global ChkExtUnique   := g.Add("Checkbox", "x" (TX+110) " yp",    "Unique only")
    global ChkExtPerLine  := g.Add("Checkbox", "x" (TX+210) " yp",    "One per line")
    global ChkExtSemicolon := g.Add("Checkbox", "x" (TX+310) " yp",   "Semicolon-separated")

    ; Wire radio changes to enable/disable the right edit boxes
    for r in [ExtR1, ExtR2, ExtR3, ExtR4, ExtR5]
        r.OnEvent("Click", UpdateExtractUI)

    ; Set defaults
    ExtR1.Value := 1
    ChkExtPerLine.Value := 1
    ChkExtPerLine.OnEvent("Click", UpdateExtractUI)
    UpdateExtractUI()
}

UpdateExtractUI(*) {
    global ExtR1, ExtR2, ExtR3, ExtR4, ExtR5
    global EditExtLinePattern, EditExtPattern
    global ChkExtSemicolon, ChkExtPerLine

    lineMode    := ExtR4.Value
    patternMode := ExtR5.Value
    EditExtLinePattern.Enabled := lineMode
    EditExtPattern.Enabled     := patternMode

    ; Semicolon option only makes sense for email (or any single-line output)
    ; Grey it out when One-per-line is checked
    ChkExtSemicolon.Enabled := !ChkExtPerLine.Value
}

BuildTab_WrapIndent() {
    global Tabs, g, TX, TY
    Tabs.UseTab(6)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Word wrap at column:")
    global NumWrapCol := g.Add("Edit", "x+6 yp-2 w50 h22 Number", "80")
    g.Add("Text",  "x" (TX+8) " y+10", "Indent / Dedent:")
    global IndentRG := g.Add("Radio", "x" (TX+8) " y+4 Group", "Indent (add spaces)")
    global IndentR2 := g.Add("Radio", "x" (TX+8) " y+4",       "Dedent (remove spaces)")
    g.Add("Text",  "x" (TX+8) " y+8", "Spaces per level:")
    global NumIndent := g.Add("Edit", "x+6 yp-2 w40 h22 Number", "4")
}

BuildTab_Counter() {
    global Tabs, g, TX, TY
    Tabs.UseTab(7)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Prepend a counter to each line.")
    g.Add("Text", "x" (TX+8) " y+12", "Start at:")
    global NumCtrStart := g.Add("Edit", "x+6 yp-2 w50 h22 Number", "1")
    g.Add("Text", "x+10 yp+2", "Step:")
    global NumCtrStep  := g.Add("Edit", "x+6 yp-2 w40 h22 Number", "1")
    g.Add("Text", "x" (TX+8) " y+8", "Separator:")
    global EditCtrSep  := g.Add("Edit", "x+6 yp-2 w60 h22", ". ")
    global ChkCtrAlign := g.Add("Checkbox", "x" (TX+8) " y+10", "Right-align numbers")
}

BuildTab_Padding() {
    global Tabs, g, TX, TY
    Tabs.UseTab(8)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Pad all lines to equal length:")
    global PadR1 := g.Add("Radio", "x" (TX+8) " y+8 Group", "Left-align  (pad right)")
    global PadR2 := g.Add("Radio", "x" (TX+8) " y+4",       "Right-align  (pad left)")
    global PadR3 := g.Add("Radio", "x" (TX+8) " y+4",       "Center")
    global PadRG := PadR1   ; kept for backward compat reference
    g.Add("Text", "x" (TX+8) " y+10", "Pad character:")
    global EditPadChar := g.Add("Edit", "x+6 yp-2 w30 h22", " ")
    g.Add("Text", "x" (TX+8) " y+8 w380", "(Leave blank to use space.  Fixed-width fonts recommended.)")
}

BuildTab_CsvView() {
    global Tabs, g, TX, TY
    Tabs.UseTab(9)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "CSV Viewer  (press F1 for tips)    Remove col:")
    global DdlCsvRemove := g.Add("DropDownList", "x+6 yp-3 w160 Choose1", ["(load CSV first)"])
    global BtnCsvRemove := g.Add("Button", "x+6 yp w70 h22", "Remove")
    global LvCsv := g.Add("ListView",
        "x" (TX+8) " y+6 w540 h140 Grid", ["(paste CSV below)"])
    ; Enable column header drag-to-reorder (LVS_EX_HEADERDRAGDROP = 0x10)
    exStyle := SendMessage(0x1037, 0, 0, , "ahk_id " LvCsv.Hwnd)  ; LVM_GETEXTENDEDLISTVIEWSTYLE
    SendMessage(0x1036, 0, exStyle | 0x10, , "ahk_id " LvCsv.Hwnd) ; LVM_SETEXTENDEDLISTVIEWSTYLE
    global BtnCsvExport  := g.Add("Button",   "x" (TX+8) " y+6 w160 h22", "▼ Send to Output pane")
    global ChkCsvQuote   := g.Add("Checkbox", "x+8 yp+3",                  'Use double-quotes')

    LvCsv.OnEvent("ColClick", OnCsvColClick)
    BtnCsvRemove.OnEvent("Click", OnCsvRemoveCol)
    BtnCsvExport.OnEvent("Click", OnCsvExport)
}

BuildTab_Compare() {
    global Tabs, g, TX, TY
    Tabs.UseTab(10)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Original text in Input pane.  Comparison text:")

    ; Controls column (left of the edit box)
    ctrlX := TX+8
    editX := TX+170   ; edit box starts here, leaving ~160px for controls

    g.Add("Text", "x" ctrlX " y+10", "Return lines:")
    global CmpR1 := g.Add("Radio", "x" ctrlX " y+6 Group", "Original only")
    global CmpR2 := g.Add("Radio", "x" ctrlX " y+4",       "Comparison only")
    global CmpR3 := g.Add("Radio", "x" ctrlX " y+4 Checked", "In both")
    g.Add("Text", "x" ctrlX " y+10", "Post-process:")
    global ChkCmpDupe := g.Add("Checkbox", "x" ctrlX " y+6",  "Remove duplicates")
    global ChkCmpSort := g.Add("Checkbox", "x" ctrlX " y+4",  "Alphabetize")

    ; Comparison text edit box, aligned to right of controls
    g.Add("Text", "x" editX " y" (TY+8), "Comparison text:")
    global EditCmpRight := g.Add("Edit",
        "x" editX " y+4 w390 h120 Multi VScroll -Wrap", "")
}

BuildTab_Stats() {
    global Tabs, g, TX, TY
    Tabs.UseTab(12)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Apply to count statistics for the Input text.")
    global LvStats := g.Add("ListView",
        "x" (TX+8) " y+8 w400 h175 Grid NoSort -Hdr", ["Statistic", "Value"])
    LvStats.Add("", "Words",                          "—")
    LvStats.Add("", "Unique words",                   "—")
    LvStats.Add("", "Type-Token Ratio (vocab %)",     "—")
    LvStats.Add("", "Lines",                          "—")
    LvStats.Add("", "Paragraphs (est.)",              "—")
    LvStats.Add("", "Sentences (est.)",               "—")
    LvStats.Add("", "Characters (with spaces)",       "—")
    LvStats.Add("", "Characters (no spaces)",         "—")
    LvStats.Add("", "Avg word length (chars)",        "—")
    LvStats.Add("", "Avg sentence length (words)",    "—")
    LvStats.Add("", "Longest word",                   "—")
    LvStats.Add("", "Total syllables (est.)",         "—")
    LvStats.Add("", "Hapax legomena (appear once)",   "—")
    LvStats.Add("", "Flesch-Kincaid Reading Ease",    "—")
    LvStats.ModifyCol(1, 210)
    LvStats.ModifyCol(2, 160)
}

BuildTab_NGrams() {
    global Tabs, g, TX, TY
    Tabs.UseTab(11)
    g.Add("Text", "x" (TX+8) " y" (TY+8), "Group size:")
    global DdlNgramSize := g.Add("DropDownList",
        "x" (TX+78) " yp-3 w80 Choose2",
        ["1 word", "2 words", "3 words", "4 words", "5 words", "6 words"])
    global BtnNgramRefresh := g.Add("Button", "x+8 yp w70 h22", "Refresh")
    global BtnNgramExport  := g.Add("Button", "x+6 yp w70 h22", "Export")
    global LvNgrams := g.Add("ListView",
        "x" (TX+8) " y+8 w540 h170 Grid", ["Phrase / Word", "Count", "Percent"])
    LvNgrams.ModifyCol(1, 330)
    LvNgrams.ModifyCol(2,  80)
    LvNgrams.ModifyCol(3,  90)
    BtnNgramRefresh.OnEvent("Click", OnNgramRefresh)
    BtnNgramExport.OnEvent("Click",  OnNgramExport)
    LvNgrams.OnEvent("ColClick",     OnNgramColClick)
}

; ============================================================
;  RESIZE
; ============================================================
OnGuiSize(GuiObj, MinMax, W, H) {
    if (MinMax = -1)
        return
    static WM_SETREDRAW    := 0x000B
    static RDW_INVALIDATE  := 0x0001
    static RDW_ERASE       := 0x0004
    static RDW_ALLCHILDREN := 0x0080
    DllCall("user32\SendMessageW", "Ptr", GuiObj.Hwnd, "UInt", WM_SETREDRAW, "Ptr", 0, "Ptr", 0)
    LayoutPanes(W, H)
    DllCall("user32\SendMessageW", "Ptr", GuiObj.Hwnd, "UInt", WM_SETREDRAW, "Ptr", 1, "Ptr", 0)
    DllCall("RedrawWindow", "Ptr", GuiObj.Hwnd, "Ptr", 0, "Ptr", 0,
            "UInt", RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN)
}

; ============================================================
;  BUTTON HANDLERS
; ============================================================

OnNavChange(ctrl, *) {
    global Tabs, LbNav
    if LbNav.Value
        Tabs.Choose(LbNav.Value)
}

OnToggleSplit(*) {
    global SplitH
    SplitH := !SplitH
    UpdateSplitBtn()
    g.GetClientPos(, , &W, &H)
    LayoutPanes(W, H)
}

OnToggleWrap(*) {
    global EditIn, EditOut, ChkWrap, g
    wrapIn  := ChkWrap.Value ? "Multi VScroll Wrap"          : "Multi VScroll HScroll -Wrap"
    wrapOut := ChkWrap.Value ? "Multi VScroll Wrap"          : "Multi VScroll HScroll -Wrap"

    savedIn := EditIn.Value
    EditIn.GetPos(&inX, &inY, &inW, &inH)
    DllCall("DestroyWindow", "Ptr", EditIn.Hwnd)
    EditIn := g.Add("Edit", "x" inX " y" inY " w" inW " h" inH " " wrapIn, savedIn)

    savedOut := EditOut.Value
    EditOut.GetPos(&outX, &outY, &outW, &outH)
    DllCall("DestroyWindow", "Ptr", EditOut.Hwnd)
    EditOut := g.Add("Edit", "x" outX " y" outY " w" outW " h" outH " " wrapOut, savedOut)
}

OnPaste(*) {
    global EditIn
    EditIn.Value := A_Clipboard
}

OnClearInput(*) {
    global EditIn, EditOut
    PushUndo(EditIn.Value, EditOut.Value)
    EditIn.Value := ""
}

OnClearOutput(*) {
    global EditOut
    EditOut.Value := ""
}

OnSwap(*) {
    global EditIn, EditOut
    PushUndo(EditIn.Value, EditOut.Value)
    tmp           := EditOut.Value
    EditOut.Value := EditIn.Value
    EditIn.Value  := tmp
}

OnCopyOutput(*) {
    global EditOut
    A_Clipboard := EditOut.Value
    ToolTip("Copied!", , , 1)
    SetTimer(() => ToolTip("", , , 1), -1200)
}

OnUndo(*) {
    global EditIn, EditOut, UndoStack, RedoStack, UNDO_MAX
    if (UndoStack.Length = 0) {
        ToolTip("Nothing to undo.", , , 2)
        SetTimer(() => ToolTip("", , , 2), -1200)
        return
    }
    ; Push current state onto redo stack before restoring
    RedoStack.InsertAt(1, Map("in", EditIn.Value, "out", EditOut.Value))
    while (RedoStack.Length > UNDO_MAX)
        RedoStack.Pop()

    snap := UndoStack.RemoveAt(1)
    EditIn.Value  := snap["in"]
    EditOut.Value := snap["out"]

    cnt := UndoStack.Length
    msg := (cnt > 0) ? cnt " more undo level" (cnt = 1 ? "" : "s") "." : "No more undo levels."
    ToolTip(msg, , , 2)
    SetTimer(() => ToolTip("", , , 2), -1500)
}

OnRedo(*) {
    global EditIn, EditOut, UndoStack, RedoStack, UNDO_MAX
    if (RedoStack.Length = 0) {
        ToolTip("Nothing to redo.", , , 2)
        SetTimer(() => ToolTip("", , , 2), -1200)
        return
    }
    ; Push current state onto undo stack before restoring (without clearing redo)
    UndoStack.InsertAt(1, Map("in", EditIn.Value, "out", EditOut.Value))
    while (UndoStack.Length > UNDO_MAX)
        UndoStack.Pop()

    snap := RedoStack.RemoveAt(1)
    EditIn.Value  := snap["in"]
    EditOut.Value := snap["out"]

    cnt := RedoStack.Length
    msg := (cnt > 0) ? cnt " more redo level" (cnt = 1 ? "" : "s") "." : "No more redo levels."
    ToolTip(msg, , , 2)
    SetTimer(() => ToolTip("", , , 2), -1500)
}

LoadHistory() {
    global CboFind, CboRepl, IniFile, HISTORY_MAX, FindHistory, ReplHistory
    Loop HISTORY_MAX {
        val := IniRead(IniFile, "FindHistory", "Item" A_Index, "")
        if (val != "") {
            FindHistory.Push(val)
            CboFind.Add([val])
        }
        val := IniRead(IniFile, "ReplaceHistory", "Item" A_Index, "")
        if (val != "") {
            ReplHistory.Push(val)
            CboRepl.Add([val])
        }
    }
}

SaveHistory() {
    global IniFile, HISTORY_MAX, FindHistory, ReplHistory
    try IniDelete(IniFile, "FindHistory")
    try IniDelete(IniFile, "ReplaceHistory")
    Loop Min(FindHistory.Length, HISTORY_MAX)
        IniWrite(FindHistory[A_Index], IniFile, "FindHistory", "Item" A_Index)
    Loop Min(ReplHistory.Length, HISTORY_MAX)
        IniWrite(ReplHistory[A_Index], IniFile, "ReplaceHistory", "Item" A_Index)
}

; Add a value to a ComboBox history, deduplicating and capping at HISTORY_MAX.
; Uses a backing Array as source of truth; rebuilds the combo from it.
AddToHistory(cbo, arr, val) {
    global HISTORY_MAX
    if (val = "")
        return
    ; Remove duplicate (case-insensitive)
    i := 1
    while (i <= arr.Length) {
        if (StrLower(arr[i]) == StrLower(val))
            arr.RemoveAt(i)
        else
            i++
    }
    ; Prepend to front
    arr.InsertAt(1, val)
    ; Trim to max
    while (arr.Length > HISTORY_MAX)
        arr.Pop()
    ; Rebuild combo from array
    cbo.Delete()
    for item in arr
        cbo.Add([item])
    ; Select first (most recent) so .Text reflects it
    cbo.Value := 1
}

OnClose(GuiObj, *) {
    GuiObj.GetPos(, , &ww, &wh)
    IniWrite(ww,          IniFile, "Window", "W")
    IniWrite(wh,          IniFile, "Window", "H")
    IniWrite(PANE_H,      IniFile, "Window", "PaneH")
    IniWrite(SplitH?1:0,  IniFile, "Window", "SplitH")
    IniWrite(Tabs.Value,  IniFile, "Window", "LastTab")
    SaveHistory()
    ExitApp()
}

; ============================================================
;  APPLY — dispatcher
; ============================================================
OnApply(*) {
    global EditIn, EditOut, Tabs

    inputText := EditIn.Value
    if (inputText = "") {
        ToolTip("Input is empty.", , , 3)
        SetTimer(() => ToolTip("", , , 3), -1500)
        return
    }

    PushUndo(inputText, EditOut.Value)
    result    := ""

    switch Tabs.Value {
        case 1:  result := Apply_Case(inputText)
        case 2:  result := Apply_Sort(inputText)
        case 3:  result := Apply_FindReplace(inputText)
        case 4:  result := Apply_Remove(inputText)
        case 5:  result := Apply_Extract(inputText)
        case 6:  result := Apply_WrapIndent(inputText)
        case 7:  result := Apply_Counter(inputText)
        case 8:  result := Apply_Padding(inputText)
        case 9:  Apply_CsvView(inputText)  ; populates LV
                 return
        case 10: result := Apply_Compare(inputText)
        case 11: Apply_NGrams(inputText)   ; populates LV
                 return
        case 12: Apply_Stats(inputText)    ; populates LV
                 return
        default: result := inputText
    }

    EditOut.Value := result
}

; ============================================================
;  TRANSFORM FUNCTIONS
; ============================================================

Apply_Case(txt) {
    lines := StrSplit(txt, "`n", "`r")
    out   := []
    mode  := GetCaseMode()
    for line in lines {
        switch mode {
            case 1:  out.Push(StrUpper(line))
            case 2:  out.Push(StrLower(line))
            case 3:  out.Push(TitleCase(line))
            case 4:  out.Push(SentenceCase(line))
            case 5:  out.Push(InvertCase(line))
            case 6:  out.Push(SarcasticCase(line))
            case 7:  out.Push(KebabCase(line))
            case 8:  out.Push(SnakeCase(line))
            case 9:  out.Push(PascalCase(line))
            case 10: out.Push(CamelCase(line))
            case 11: out.Push(ScreamingSnakeCase(line))
            case 12: out.Push(TrainCase(line))
            default: out.Push(line)
        }
    }
    return JoinLines(out)
}

GetCaseMode() {
    global CaseR1, CaseR2, CaseR3, CaseR4, CaseR5, CaseR6
    global CaseR7, CaseR8, CaseR9, CaseR10, CaseR11, CaseR12
    radios := [CaseR1, CaseR2, CaseR3, CaseR4, CaseR5, CaseR6,
               CaseR7, CaseR8, CaseR9, CaseR10, CaseR11, CaseR12]
    for i, r in radios
        if r.Value
            return i
    return 1
}

Apply_Sort(txt) {
    global SortR1, SortR2, SortR3, SortR4, SortR5, SortR6
    global ChkSortDupe, ChkSortTrim, DdlSortKey, EditSortDelim

    doTrim  := ChkSortTrim.Value
    doDupe  := ChkSortDupe.Value
    keyMode := DdlSortKey.Value      ; 1=whole line, 2=word1, 3=word2, 4=last word, 5=after delim
    delim   := EditSortDelim.Value
    if (delim = "")
        delim := "-"

    ; Determine sort mode
    mode := 1
    for i, r in [SortR1, SortR2, SortR3, SortR4, SortR5, SortR6]
        if r.Value {
            mode := i
            break
        }

    ; A→Z and Z→A with whole-line key: use AHK's built-in Sort (fast path)
    if (mode = 1 || mode = 2) && (keyMode = 1) {
        sorted := doTrim ? TrimLines(txt) : txt
        opts   := (mode = 2) ? "R" : ""
        if doDupe
            opts .= " U"
        sorted := Sort(sorted, opts)
        return sorted
    }

    ; All other cases: split into array, sort with key extractor
    lines := StrSplit(txt, "`n", "`r")
    work  := []
    for line in lines
        work.Push(doTrim ? Trim(line) : line)

    if (mode = 1 || mode = 2) {
        ; A→Z or Z→A with a custom sort key — insertion sort with SortKeyOf
        reverse := (mode = 2)
        n := work.Length
        i := 2
        while i <= n {
            cur    := work[i]
            curKey := StrLower(SortKeyOf(cur, keyMode, delim))
            j := i - 1
            while j >= 1 {
                cmpKey := StrLower(SortKeyOf(work[j], keyMode, delim))
                cmp    := StrCompare(cmpKey, curKey)
                swapNeeded := reverse ? (cmp < 0) : (cmp > 0)
                if swapNeeded {
                    work[j+1] := work[j]
                    j--
                } else
                    break
            }
            work[j+1] := cur
            i++
        }
    } else if (mode = 3) {
        ; Random — Fisher-Yates shuffle
        n := work.Length
        Loop n - 1 {
            i := n - A_Index + 1
            j := Random(1, i)
            tmp := work[i], work[i] := work[j], work[j] := tmp
        }
    } else if (mode = 4) {
        ; Length short→long
        SortByLength(work, false)
    } else if (mode = 5) {
        ; Length long→short
        SortByLength(work, true)
    } else if (mode = 6) {
        ; Numeric / natural sort — key-aware
        reverse := false
        n := work.Length
        i := 2
        while i <= n {
            cur    := work[i]
            curKey := SortKeyOf(cur, keyMode, delim)
            j := i - 1
            while j >= 1 {
                cmpKey := SortKeyOf(work[j], keyMode, delim)
                cmp    := NaturalCompare(cmpKey, curKey)
                if cmp > 0 {
                    work[j+1] := work[j]
                    j--
                } else
                    break
            }
            work[j+1] := cur
            i++
        }
    }

    ; Remove duplicates (case-insensitive, preserving order)
    if doDupe {
        seen := Map()
        deduped := []
        for line in work {
            key := StrLower(line)
            if !seen.Has(key) {
                seen[key] := true
                deduped.Push(line)
            }
        }
        work := deduped
    }

    return JoinLines(work)
}

; Return the portion of a line used as the sort key.
;   keyMode 1 = whole line
;   keyMode 2 = first word  (split on whitespace)
;   keyMode 3 = second word
;   keyMode 4 = last word
;   keyMode 5 = text after the first occurrence of delim
SortKeyOf(line, keyMode, delim) {
    if keyMode = 1
        return line
    if keyMode = 5 {
        pos := InStr(line, delim)
        return pos ? LTrim(SubStr(line, pos + StrLen(delim))) : line
    }
    ; Word-based modes — split on runs of whitespace
    words := StrSplit(Trim(line), " ", "`t")
    ; Collapse empty tokens that StrSplit may produce from multiple spaces
    cleanWords := []
    for w in words
        if w != ""
            cleanWords.Push(w)
    if cleanWords.Length = 0
        return line
    if keyMode = 2
        return cleanWords[1]
    if keyMode = 3
        return cleanWords.Length >= 2 ? cleanWords[2] : cleanWords[1]
    if keyMode = 4
        return cleanWords[cleanWords.Length]
    return line
}

SortByLength(arr, descending) {
    ; Insertion sort (stable, fine for typical text sizes)
    n := arr.Length
    i := 2
    while i <= n {
        key := arr[i]
        kLen := StrLen(key)
        j := i - 1
        while j >= 1 {
            cLen := StrLen(arr[j])
            swapNeeded := descending ? (cLen < kLen) : (cLen > kLen)
            if swapNeeded {
                arr[j+1] := arr[j]
                j--
            } else
                break
        }
        arr[j+1] := key
        i++
    }
}

NaturalCompare(a, b) {
    ; Returns -1, 0, or 1 for natural (human) ordering
    pa := 1, pb := 1
    la := StrLen(a), lb := StrLen(b)
    while (pa <= la && pb <= lb) {
        ca := SubStr(a, pa, 1)
        cb := SubStr(b, pb, 1)
        isDigA := RegExMatch(ca, "\d")
        isDigB := RegExMatch(cb, "\d")
        if isDigA && isDigB {
            ; Collect full numeric run from each
            numA := "", numB := ""
            while (pa <= la && RegExMatch(SubStr(a, pa, 1), "\d"))
                numA .= SubStr(a, pa++, 1)
            while (pb <= lb && RegExMatch(SubStr(b, pb, 1), "\d"))
                numB .= SubStr(b, pb++, 1)
            ia := Integer(numA), ib := Integer(numB)
            if ia < ib
                return -1
            else if ia > ib
                return 1
        } else {
            uca := StrUpper(ca), ucb := StrUpper(cb)
            if uca < ucb
                return -1
            else if uca > ucb
                return 1
            pa++, pb++
        }
    }
    if la < lb
        return -1
    else if la > lb
        return 1
    return 0
}

Apply_FindReplace(txt) {
    findStr := CboFind.Text
    replStr := CboRepl.Text

    if (findStr = "") {
        ToolTip("Find field is empty.", , , 3)
        SetTimer(() => ToolTip("", , , 3), -1500)
        return txt
    }
    ; Add both to history (replace only if non-empty)
    AddToHistory(CboFind, FindHistory, findStr)
    if (replStr != "")
        AddToHistory(CboRepl, ReplHistory, replStr)

    caseSens := ChkFRCase.Value
    useRegex := ChkFRRegex.Value
    replAll  := ChkFRAll.Value

    if useRegex {
        flags := caseSens ? ")" : "i)"
        pat   := flags . findStr
        if replAll
            return RegExReplace(txt, pat, replStr)
        else {
            if RegExMatch(txt, pat, &m)
                return SubStr(txt, 1, m.Pos - 1) . replStr . SubStr(txt, m.Pos + m.Len)
            return txt
        }
    } else {
        if caseSens {
            pat := ")" . RegExEscape(findStr)
            if replAll
                return RegExReplace(txt, pat, replStr)
            else {
                if RegExMatch(txt, pat, &m)
                    return SubStr(txt, 1, m.Pos - 1) . replStr . SubStr(txt, m.Pos + m.Len)
                return txt
            }
        } else {
            limit := replAll ? -1 : 1
            return StrReplace(txt, findStr, replStr, false, , limit)
        }
    }
}

RegExEscape(s) {
    static meta := "\.+*?[^]$(){}=!<>|:-#"
    out := ""
    Loop Parse, s {
        ch  := A_LoopField
        out .= (InStr(meta, ch) ? "\" : "") . ch
    }
    return out
}

Apply_Remove(txt) {
    lines := StrSplit(txt, "`n", "`r")
    out   := []
    seen  := Map()

    trimL := Integer(NumTrimL.Value)
    trimR := Integer(NumTrimR.Value)

    for line in lines {
        if (trimL > 0)
            line := SubStr(line, trimL + 1)
        if (trimR > 0 && StrLen(line) > 0)
            line := SubStr(line, 1, Max(0, StrLen(line) - trimR))
        if ChkRmLead.Value
            line := LTrim(line)
        if ChkRmTrail.Value
            line := RTrim(line)
        if ChkRmHTML.Value
            line := RegExReplace(line, "<[^>]*>", "")
        if ChkRmBBCode.Value
            line := RegExReplace(line, "\[[^\]]*\]", "")
        if ChkRmNonASCII.Value
            line := RegExReplace(line, "[^\x00-\x7F]", "")
        if ChkRmBlank.Value && (Trim(line) = "")
            continue
        if ChkRmDupe.Value {
            key := StrLower(line)
            if seen.Has(key)
                continue
            seen[key] := true
        }
        out.Push(line)
    }
    return JoinLines(out)
}

Apply_Extract(txt) {
    global ExtR1, ExtR2, ExtR3, ExtR4, ExtR5
    global EditExtLinePattern, EditExtPattern
    global ChkExtUnique, ChkExtPerLine, ChkExtSemicolon

    doUnique    := ChkExtUnique.Value
    doPerLine   := ChkExtPerLine.Value
    doSemicolon := ChkExtSemicolon.Value

    matches := []

    if ExtR1.Value {
        ; Numbers — integers, decimals, negatives
        pos := 1
        while RegExMatch(txt, "-?\d+(?:\.\d+)?", &m, pos) {
            matches.Push(m[0])
            pos := m.Pos + m.Len
        }

    } else if ExtR2.Value {
        ; Email addresses
        pos := 1
        while RegExMatch(txt, "[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", &m, pos) {
            matches.Push(m[0])
            pos := m.Pos + m.Len
        }

    } else if ExtR3.Value {
        ; URLs
        pos := 1
        while RegExMatch(txt, 'i)\b(?:https?://|www\.)[^\s"<>]+(?<![.,;:!?])', &m, pos) {
            matches.Push(m[0])
            pos := m.Pos + m.Len
        }

    } else if ExtR4.Value {
        ; Lines matching pattern
        pat := Trim(EditExtLinePattern.Value)
        if pat = "" {
            ToolTip("Enter a pattern to match.", , , 3)
            SetTimer(() => ToolTip("", , , 3), -1800)
            return ""
        }
        Loop Parse, txt, "`n", "`r" {
            try {
                if RegExMatch(A_LoopField, pat)
                    matches.Push(A_LoopField)
            } catch {
                ToolTip("Invalid regex pattern.", , , 3)
                SetTimer(() => ToolTip("", , , 3), -1800)
                return ""
            }
        }

    } else if ExtR5.Value {
        ; Custom regex — extract captured group 1 (or whole match if no group)
        pat := Trim(EditExtPattern.Value)
        if pat = "" {
            ToolTip("Enter a regex pattern.", , , 3)
            SetTimer(() => ToolTip("", , , 3), -1800)
            return ""
        }
        pos := 1
        while RegExMatch(txt, pat, &m, pos) {
            try {
                val := (m.Count >= 1) ? m[1] : m[0]
                matches.Push(val)
            } catch {
                matches.Push(m[0])
            }
            newPos := m.Pos + m.Len
            if newPos <= pos  ; safety — avoid infinite loop on zero-length match
                break
            pos := newPos
        }
    }

    ; Deduplicate (preserving order)
    if doUnique {
        seen := Map()
        deduped := []
        for v in matches {
            key := StrLower(v)
            if !seen.Has(key) {
                seen[key] := true
                deduped.Push(v)
            }
        }
        matches := deduped
    }

    if matches.Length = 0
        return ""

    ; Format output
    if doSemicolon && !doPerLine
        return JoinArray(matches, "; ")

    return JoinLines(matches)
}

; Join an array with an arbitrary separator
JoinArray(arr, sep) {
    out := ""
    for i, v in arr
        out .= (i = 1 ? "" : sep) . v
    return out
}

Apply_WrapIndent(txt) {
    global NumWrapCol, IndentRG, NumIndent

    wrapCol  := Integer(NumWrapCol.Value)
    spaces   := Integer(NumIndent.Value)
    isIndent := IndentRG.Value ? true : false

    lines := StrSplit(txt, "`n", "`r")
    out   := []

    for line in lines {
        ; --- Hard wrap ---
        if (wrapCol > 0 && StrLen(line) > wrapCol) {
            ; Wrap at word boundaries
            wrapped := WordWrap(line, wrapCol)
            for wl in wrapped
                out.Push(wl)
        } else {
            out.Push(line)
        }
    }

    ; --- Indent / Dedent ---
    if (spaces > 0) {
        pad := ""
        Loop spaces
            pad .= " "
        final := []
        for line in out {
            if isIndent {
                final.Push(pad . line)
            } else {
                ; Remove up to `spaces` leading spaces
                removed := 0
                while (removed < spaces && SubStr(line, 1, 1) = " ") {
                    line := SubStr(line, 2)
                    removed++
                }
                final.Push(line)
            }
        }
        out := final
    }

    return JoinLines(out)
}

; Word-wrap a single line at maxCol characters, breaking at spaces
WordWrap(line, maxCol) {
    result := []
    while StrLen(line) > maxCol {
        ; Find last space at or before maxCol
        breakPos := 0
        Loop maxCol {
            if SubStr(line, A_Index, 1) = " "
                breakPos := A_Index
        }
        if breakPos = 0 {
            ; No space found — hard break at maxCol
            result.Push(SubStr(line, 1, maxCol))
            line := SubStr(line, maxCol + 1)
        } else {
            result.Push(SubStr(line, 1, breakPos - 1))
            line := SubStr(line, breakPos + 1)
        }
    }
    result.Push(line)
    return result
}

Apply_Counter(txt) {
    global NumCtrStart, NumCtrStep, EditCtrSep, ChkCtrAlign

    start := Integer(NumCtrStart.Value)
    step  := Integer(NumCtrStep.Value)
    sep   := EditCtrSep.Value   ; e.g. ". " or ": " or "\t"
    doAlign := ChkCtrAlign.Value

    lines := StrSplit(txt, "`n", "`r")
    out   := []

    ; Calculate max counter value to determine padding width
    maxNum := start + step * (lines.Length - 1)
    padW   := StrLen(String(maxNum))

    n := start
    for line in lines {
        numStr := doAlign ? Format("{:0" padW "}", n) : String(n)
        out.Push(numStr . sep . line)
        n += step
    }
    return JoinLines(out)
}

Apply_Padding(txt) {
    global EditPadChar

    mode  := GetPadMode()
    padCh := EditPadChar.Value
    if (StrLen(padCh) = 0)
        padCh := " "
    padCh := SubStr(padCh, 1, 1)   ; only first character

    lines := StrSplit(txt, "`n", "`r")

    ; Find maximum line length
    maxLen := 0
    for line in lines
        maxLen := Max(maxLen, StrLen(line))

    out := []
    for line in lines {
        diff := maxLen - StrLen(line)
        if diff = 0 {
            out.Push(line)
            continue
        }
        if (mode = 1) {
            ; Left-align: pad right
            out.Push(line . RepeatChar(padCh, diff))
        } else if (mode = 2) {
            ; Right-align: pad left
            out.Push(RepeatChar(padCh, diff) . line)
        } else {
            ; Center: split padding
            leftPad  := diff // 2
            rightPad := diff - leftPad
            out.Push(RepeatChar(padCh, leftPad) . line . RepeatChar(padCh, rightPad))
        }
    }
    return JoinLines(out)
}

GetPadMode() {
    global PadRG
    ; PadRG is first radio.  Use Gui control tab to read sibling radios.
    ; Simplest reliable way: test Value of the first; if 0, next radio in
    ; the same group is at PadRG.Hwnd+1... but that's fragile.
    ; Instead declare all three as globals in BuildTab_Padding (done below).
    global PadR1, PadR2, PadR3
    if PadR1.Value
        return 1
    if PadR2.Value
        return 2
    return 3
}

Apply_CsvView(txt) {
    global LvCsv, CsvSortCol, CsvSortAsc, CsvHeaders
    CsvSortCol := 0
    CsvSortAsc := true
    CsvHeaders := []

    ; Clear existing columns and rows
    LvCsv.Delete()
    Loop 200 {
        try
            LvCsv.DeleteCol(1)
        catch
            break
    }

    lines := StrSplit(txt, "`n", "`r")
    if (lines.Length = 0 || Trim(lines[1]) = "") {
        LvCsv.InsertCol(1, 140, "(no data)")
        return
    }

    ; Parse first line as header
    headers := ParseCSVLine(lines[1])
    if (headers.Length = 0) {
        LvCsv.InsertCol(1, 140, "(empty)")
        return
    }

    ; Add columns and store in CsvHeaders
    CsvHeaders := headers.Clone()
    for i, h in headers
        LvCsv.InsertCol(i, 100, h)

    ; Add data rows
    rowStart := (lines.Length > 1) ? 2 : lines.Length + 1
    Loop lines.Length - 1 {
        rowLine := lines[A_Index + 1]
        if (Trim(rowLine) = "")
            continue
        fields := ParseCSVLine(rowLine)
        ; Pad to header count if needed
        while fields.Length < headers.Length
            fields.Push("")
        LvCsv.Add("", fields*)
    }

    ; Auto-size columns (cap at 200px)
    Loop headers.Length
        LvCsv.ModifyCol(A_Index, "AutoHdr")

    PopulateCsvRemoveDdl()
}

; Parse one RFC-4180 CSV line into an Array of field strings
ParseCSVLine(line) {
    fields := []
    pos    := 1
    len    := StrLen(line)

    while pos <= len {
        ch := SubStr(line, pos, 1)
        if (ch = '"') {
            ; Quoted field
            field := ""
            pos++
            while pos <= len {
                c := SubStr(line, pos, 1)
                if (c = '"') {
                    ; Peek ahead for escaped quote
                    if (pos + 1 <= len && SubStr(line, pos+1, 1) = '"') {
                        field .= '"'
                        pos += 2
                    } else {
                        pos++   ; closing quote
                        break
                    }
                } else {
                    field .= c
                    pos++
                }
            }
            fields.Push(field)
            ; Consume trailing comma
            if (pos <= len && SubStr(line, pos, 1) = ",")
                pos++
        } else {
            ; Unquoted field — read to next comma
            start := pos
            while (pos <= len && SubStr(line, pos, 1) != ",")
                pos++
            fields.Push(SubStr(line, start, pos - start))
            if (pos <= len)
                pos++   ; skip comma
        }
    }
    return fields
}

; Returns an array (1-based) of logical column indices in current visual order.
; e.g. [3,1,2] means visual col 1 = logical col 3, etc.
LvGetColOrder(lv, colCount) {
    buf := Buffer(4 * colCount, 0)
    SendMessage(0x103B, colCount, buf.Ptr, , "ahk_id " lv.Hwnd)  ; LVM_GETCOLUMNORDERARRAY
    order := []
    Loop colCount
        order.Push(NumGet(buf, (A_Index - 1) * 4, "int") + 1)  ; +1 → 1-based
    return order
}

; Count columns in a ListView via its header control
LvColCount(lv) {
    hHeader := SendMessage(0x101F, 0, 0, , "ahk_id " lv.Hwnd)  ; LVM_GETHEADER
    return SendMessage(0x1200, 0, 0, , "ahk_id " hHeader)       ; HDM_GETITEMCOUNT
}
CsvSortCol := 0      ; last sorted column index (0 = unsorted)
CsvSortAsc := true   ; true = ascending
CsvHeaders := []     ; column header names — maintained as source of truth

; Sort ListView by column when header is clicked
OnCsvColClick(ctrl, colIndex) {
    global CsvSortCol, CsvSortAsc, CsvHeaders
    rowCount := ctrl.GetCount()
    colCount  := LvColCount(ctrl)
    order     := LvGetColOrder(ctrl, colCount)

    ; Find visual position of the clicked logical column
    visualSortPos := 1
    for vi, li in order {
        if (li = colIndex) {
            visualSortPos := vi
            break
        }
    }

    ; Toggle or set sort direction
    if (CsvSortCol = visualSortPos)
        CsvSortAsc := !CsvSortAsc
    else {
        CsvSortCol := visualSortPos
        CsvSortAsc := true
    }

    ; Snapshot rows in visual column order BEFORE touching the LV
    rows := []
    Loop rowCount {
        ri := A_Index
        row := []
        for logicalIdx in order
            row.Push(ctrl.GetText(ri, logicalIdx))
        rows.Push(row)
    }

    BubbleSort(rows, visualSortPos, CsvSortAsc)

    ; Bake visual order into CsvHeaders
    newHeaders := []
    for logicalIdx in order
        newHeaders.Push(CsvHeaders[logicalIdx])
    CsvHeaders := newHeaders

    ; Repopulate rows — now in visual order = new logical order
    ctrl.Delete()
    for row in rows
        ctrl.Add("", row*)

    ; Reset column order to identity — data is now logically ordered correctly
    identBuf := Buffer(4 * colCount, 0)
    Loop colCount
        NumPut("int", A_Index - 1, identBuf, (A_Index - 1) * 4)
    SendMessage(0x103C, colCount, identBuf.Ptr, , "ahk_id " ctrl.Hwnd)  ; LVM_SETCOLUMNORDERARRAY

    ; CsvSortCol now refers to the visual/logical position (they're the same after identity reset)
    CsvSortCol := visualSortPos

    ; Re-autosize all columns
    Loop colCount
        ctrl.ModifyCol(A_Index, "AutoHdr")
}

; Simple stable sort for small-to-medium CSV data
BubbleSort(rows, col, asc) {
    n := rows.Length
    Loop n - 1 {
        i := A_Index
        Loop n - i {
            j := A_Index
            a := rows[j][col]
            b := rows[j+1][col]
            ; Try numeric comparison first
            swap := IsNumber(a) && IsNumber(b)
                  ? (asc ? (Number(a) > Number(b))      : (Number(a) < Number(b)))
                  : (asc ? (StrCompare(a, b) > 0)       : (StrCompare(a, b) < 0))
            if swap {
                tmp       := rows[j]
                rows[j]   := rows[j+1]
                rows[j+1] := tmp
            }
        }
    }
}

; Populate the Remove Column DDL with current header names
PopulateCsvRemoveDdl() {
    global DdlCsvRemove, CsvHeaders
    DdlCsvRemove.Delete()
    if CsvHeaders.Length = 0 {
        DdlCsvRemove.Add(["(no columns)"])
        DdlCsvRemove.Choose(1)
        return
    }
    for h in CsvHeaders
        DdlCsvRemove.Add([h])
    DdlCsvRemove.Choose(1)
}

; Remove the column selected in DdlCsvRemove
OnCsvRemoveCol(*) {
    global LvCsv, DdlCsvRemove, CsvHeaders
    colIdx := DdlCsvRemove.Value
    if (colIdx < 1 || colIdx > CsvHeaders.Length)
        return
    CsvRemoveColumn(LvCsv, colIdx)
    PopulateCsvRemoveDdl()
}

; Remove a column from the ListView by index (1-based)
CsvRemoveColumn(ctrl, colIdx) {
    global CsvHeaders
    colCount := LvColCount(ctrl)
    if (colIdx < 1 || colIdx > colCount)
        return

    ; Snapshot all rows minus the target column
    rowCount := ctrl.GetCount()
    rows := []
    Loop rowCount {
        ri := A_Index
        row := []
        Loop colCount {
            if (A_Index != colIdx)
                row.Push(ctrl.GetText(ri, A_Index))
        }
        rows.Push(row)
    }

    ; Update CsvHeaders
    CsvHeaders.RemoveAt(colIdx)

    ; Rebuild the ListView
    ctrl.Delete()
    Loop 200 {
        try
            ctrl.DeleteCol(1)
        catch
            break
    }
    for i, h in CsvHeaders
        ctrl.InsertCol(i, 100, h)
    for row in rows
        ctrl.Add("", row*)
    Loop CsvHeaders.Length
        ctrl.ModifyCol(A_Index, "AutoHdr")
}

OnCsvExport(*) {
    global LvCsv, EditOut, CsvHeaders, ChkCsvQuote

    colCount   := LvColCount(LvCsv)
    rowCount   := LvCsv.GetCount()
    forceQuote := ChkCsvQuote.Value
    if (colCount = 0)
        return

    order := LvGetColOrder(LvCsv, colCount)
    csvLines := []

    headerFields := []
    for logicalIdx in order
        headerFields.Push(CsvQuote(logicalIdx <= CsvHeaders.Length ? CsvHeaders[logicalIdx] : "", forceQuote))
    csvLines.Push(JoinArr(headerFields, ","))

    Loop rowCount {
        ri := A_Index
        rowFields := []
        for logicalIdx in order
            rowFields.Push(CsvQuote(LvCsv.GetText(ri, logicalIdx), forceQuote))
        csvLines.Push(JoinArr(rowFields, ","))
    }

    EditOut.Value := JoinLines(csvLines)
}

; Quote a CSV field if it contains comma, quote, or newline — or if forceQuote is true
CsvQuote(val, forceQuote := false) {
    if (forceQuote || RegExMatch(val, '[,"`n`r]'))
        return '"' StrReplace(val, '"', '""') '"'
    return val
}

; Join an array of strings with a delimiter
JoinArr(arr, delim) {
    out := ""
    for i, v in arr
        out .= (i > 1 ? delim : "") v
    return out
}

Apply_Compare(txt) {
    global EditCmpRight, CmpR1, CmpR2, CmpR3, ChkCmpDupe, ChkCmpSort

    cmpTxt := EditCmpRight.Value

    linesL := StrSplit(txt,    "`n", "`r")
    linesR := StrSplit(cmpTxt, "`n", "`r")

    ; Build lookup sets (case-sensitive)
    setL := Map(), setR := Map()
    for line in linesL
        setL[line] := true
    for line in linesR
        setR[line] := true

    ; Determine which mode is selected
    mode := CmpR1.Value ? "orig" : CmpR2.Value ? "cmp" : "both"

    ; Collect result lines (preserving order from the relevant source)
    result := []
    srcLines := (mode = "cmp") ? linesR : linesL
    for line in srcLines {
        if (mode = "orig") && setR.Has(line)   ; original only = in L, not in R
            continue
        if (mode = "cmp")  && setL.Has(line)   ; comparison only = in R, not in L
            continue
        if (mode = "both") && !setR.Has(line)  ; in both = in L and in R
            continue
        result.Push(line)
    }

    ; Remove duplicates if requested (preserves first occurrence order)
    if ChkCmpDupe.Value {
        seen := Map()
        deduped := []
        for line in result {
            if !seen.Has(line) {
                seen[line] := true
                deduped.Push(line)
            }
        }
        result := deduped
    }

    ; Alphabetize if requested
    if ChkCmpSort.Value {
        joined := ""
        for line in result
            joined .= line "`n"
        joined := RTrim(joined, "`n")
        joined := Sort(joined, "D`n")
        result := StrSplit(joined, "`n")
    }

    return JoinLines(result)
}

HasValue(arr, val) {
    for item in arr
        if (item == val)
            return true
    return false
}

Apply_Stats(txt) {
    global LvStats

    ; --- Tokenize words (letters/apostrophes only, strip edge apostrophes) ---
    wordList := []
    pos := 1
    while RegExMatch(txt, "[a-zA-Z']+", &m, pos) {
        w := RegExReplace(m[0], "^'+|'+$", "")
        if w != ""
            wordList.Push(StrLower(w))
        pos := m.Pos + m.Len
    }
    totalWords := wordList.Length

    ; --- Unique words & TTR ---
    uniqMap := Map()
    for w in wordList
        uniqMap[w] := (uniqMap.Has(w) ? uniqMap[w] + 1 : 1)
    uniqueWords := uniqMap.Count
    ttr := totalWords > 0 ? Round(uniqueWords / totalWords * 100, 1) "%" : "—"

    ; --- Hapax legomena ---
    hapax := 0
    for w, c in uniqMap
        if c = 1
            hapax++

    ; --- Line count ---
    lineCount := StrSplit(txt, "`n", "`r").Length

    ; --- Paragraph count ---
    paraCount := 0
    if Trim(txt) != "" {
        Loop Parse, txt, "`n" {
            if Trim(A_LoopField) = "" && A_Index > 1
                paraCount++
        }
        paraCount++   ; at least 1
    }

    ; --- Sentence count (punctuation-based, fallback to lines) ---
    sentList := []
    raw := RegExReplace(txt, "\s+", " ")
    pos := 1
    while RegExMatch(raw, "[A-Za-z][^.!?]*[.!?]", &m, pos) {
        s := Trim(m[0])
        if StrLen(s) > 3
            sentList.Push(s)
        pos := m.Pos + m.Len
    }
    if sentList.Length = 0 {
        Loop Parse, txt, "`n", "`r" {
            s := Trim(A_LoopField)
            if StrLen(s) > 1
                sentList.Push(s)
        }
    }
    sentCount := sentList.Length

    ; --- Character counts ---
    charCount    := StrLen(txt)
    noSpaceCount := StrLen(RegExReplace(txt, "\s", ""))

    ; --- Avg word length (letter chars only) ---
    totalLetters := 0
    for w in wordList
        totalLetters += StrLen(w)
    avgWordLen := totalWords > 0 ? Round(totalLetters / totalWords, 2) : 0

    ; --- Avg sentence length ---
    totalSentWords := 0
    for s in sentList {
        pos2 := 1
        while RegExMatch(s, "[a-zA-Z']+", &m2, pos2) {
            totalSentWords++
            pos2 := m2.Pos + m2.Len
        }
    }
    avgSentLen := sentCount > 0 ? Round(totalSentWords / sentCount, 1) : 0

    ; --- Longest word ---
    longestWord := ""
    for w in wordList
        if StrLen(w) > StrLen(longestWord)
            longestWord := w

    ; --- Syllables & Flesch-Kincaid ---
    totalSyllables := 0
    for w in wordList
        totalSyllables += CountSyllables_TTB(w)
    fk := 0
    if totalWords > 0 && sentCount > 0 {
        fk := Round(206.835
            - 1.015 * (totalWords / sentCount)
            - 84.6  * (totalSyllables / totalWords), 1)
        fk := Round(Max(0, Min(100, fk)), 1)
    }
    fkLabel := fk >= 90 ? "Very Easy"
             : fk >= 80 ? "Easy"
             : fk >= 70 ? "Fairly Easy"
             : fk >= 60 ? "Standard"
             : fk >= 50 ? "Fairly Difficult"
             : fk >= 30 ? "Difficult"
             : "Very Difficult"

    ; --- Populate ListView rows ---
    LvStats.Modify(1,  , , totalWords)
    LvStats.Modify(2,  , , uniqueWords)
    LvStats.Modify(3,  , , ttr)
    LvStats.Modify(4,  , , lineCount)
    LvStats.Modify(5,  , , paraCount)
    LvStats.Modify(6,  , , sentCount)
    LvStats.Modify(7,  , , charCount)
    LvStats.Modify(8,  , , noSpaceCount)
    LvStats.Modify(9,  , , avgWordLen)
    LvStats.Modify(10, , , avgSentLen)
    LvStats.Modify(11, , , longestWord)
    LvStats.Modify(12, , , totalSyllables)
    LvStats.Modify(13, , , hapax)
    LvStats.Modify(14, , , fk " / 100  (" fkLabel ")")
}

; Simple syllable heuristic: count vowel groups, strip trailing-e first.
CountSyllables_TTB(word) {
    word := StrLower(RegExReplace(word, "e$", ""))
    cnt := 0
    pos := 1
    while RegExMatch(word, "[aeiouy]+", &m, pos) {
        cnt++
        pos := m.Pos + m.Len
    }
    return Max(1, cnt)
}

; ============================================================
;  N-GRAMS
; ============================================================

Apply_NGrams(txt) {
    global LvNgrams, DdlNgramSize
    ; Parse N from DDL text: "1 word" or "N words"
    nVal := Integer(SubStr(DdlNgramSize.Text, 1, 1))
    PopulateNgrams_TTB(txt, nVal)
}

OnNgramRefresh(*) {
    global EditIn, DdlNgramSize, LvNgrams
    txt := EditIn.Value
    if txt = "" {
        ToolTip("Input is empty.", , , 3)
        SetTimer(() => ToolTip("", , , 3), -1500)
        return
    }
    nVal := Integer(SubStr(DdlNgramSize.Text, 1, 1))
    PopulateNgrams_TTB(txt, nVal)
}

PopulateNgrams_TTB(txt, n) {
    global LvNgrams, BtnNgramRefresh

    BtnNgramRefresh.Enabled := false
    BtnNgramRefresh.Text    := "Working…"

    try {
        ; Tokenize words
        wordList := []
        pos := 1
        while RegExMatch(txt, "[a-zA-Z']+", &m, pos) {
            w := RegExReplace(m[0], "^'+|'+$", "")
            if w != ""
                wordList.Push(StrLower(w))
            pos := m.Pos + m.Len
        }

        LvNgrams.Delete()

        if wordList.Length = 0 {
            ToolTip("No words found in Input.", , , 3)
            SetTimer(() => ToolTip("", , , 3), -1800)
            return
        }

        freqMap := Map()
        if n = 1 {
            ; Word frequency mode
            for w in wordList
                freqMap[w] := (freqMap.Has(w) ? freqMap[w] + 1 : 1)
        } else {
            ; N-gram mode
            maxI := wordList.Length - n + 1
            Loop maxI {
                i := A_Index
                phrase := ""
                Loop n
                    phrase .= (A_Index = 1 ? "" : " ") . wordList[i + A_Index - 1]
                freqMap[phrase] := (freqMap.Has(phrase) ? freqMap[phrase] + 1 : 1)
            }
        }

        total := 0
        for p, c in freqMap
            total += c

        ; Collect, sort by count descending
        rows := []
        for p, c in freqMap
            rows.Push([p, c])
        rows := SortRowsByCol(rows, 2, "desc")

        for row in rows {
            pct := total > 0 ? Round(row[2] / total * 100, 2) "%" : "0%"
            LvNgrams.Add("", row[1], row[2], pct)
        }

        label := n = 1 ? "word" : n "-word phrase"
        ToolTip(rows.Length " unique " label (rows.Length = 1 ? "" : "s") " found.", , , 3)
        SetTimer(() => ToolTip("", , , 3), -2000)
    }
    finally {
        BtnNgramRefresh.Text    := "Refresh"
        BtnNgramRefresh.Enabled := true
    }
}

OnNgramExport(*) {
    global LvNgrams, EditOut
    rowCount := LvNgrams.GetCount()
    if rowCount = 0 {
        ToolTip("N-Grams list is empty — apply first.", , , 3)
        SetTimer(() => ToolTip("", , , 3), -1800)
        return
    }

    ; Measure column widths for alignment
    maxPhraseLen := StrLen("Phrase / Word")
    maxCountLen  := StrLen("Count")
    Loop rowCount {
        w1 := StrLen(LvNgrams.GetText(A_Index, 1))
        w2 := StrLen(LvNgrams.GetText(A_Index, 2))
        if w1 > maxPhraseLen
            maxPhraseLen := w1
        if w2 > maxCountLen
            maxCountLen := w2
    }

    ; Build header + separator
    sep := RepeatChar("-", maxPhraseLen + 2 + maxCountLen + 2 + 10)
    out := PadRightTTB("Phrase / Word", maxPhraseLen)
          . "  " PadRightTTB("Count", maxCountLen)
          . "  Percent`n"
          . sep . "`n"

    Loop rowCount {
        p   := LvNgrams.GetText(A_Index, 1)
        cnt := LvNgrams.GetText(A_Index, 2)
        pct := LvNgrams.GetText(A_Index, 3)
        out .= PadRightTTB(p, maxPhraseLen)
            . "  " PadRightTTB(cnt, maxCountLen)
            . "  " pct . "`n"
    }

    PushUndo(EditIn.Value, EditOut.Value)
    EditOut.Value := RTrim(out, "`n")
    ToolTip("Exported " rowCount " rows to Output pane.", , , 3)
    SetTimer(() => ToolTip("", , , 3), -2000)
}

; Sort a nested Array of Arrays by column index (1-based), direction "asc"/"desc".
; Numeric columns are compared numerically; string columns use StrCompare.
SortRowsByCol(rows, col, dir := "asc") {
    lines := ""
    isNumeric := true
    for row in rows {
        val := RegExReplace(row[col], "%", "")
        if !IsNumber(val)
            isNumeric := false
        payload := ""
        for v in row
            payload .= (payload = "" ? "" : "`t") . v
        lines .= payload . "`n"
    }
    lines := RTrim(lines, "`n")

    colIdx  := col
    isNum   := isNumeric
    reverse := (dir = "desc") ? -1 : 1

    SortCb(a, b, *) {
        vA := StrSplit(a, "`t")[colIdx]
        vB := StrSplit(b, "`t")[colIdx]
        vA := RegExReplace(vA, "%", "")
        vB := RegExReplace(vB, "%", "")
        if isNum {
            nA := Number(vA), nB := Number(vB)
            return reverse * (nA > nB ? 1 : nA < nB ? -1 : 0)
        }
        return reverse * StrCompare(vA, vB)
    }

    lines := Sort(lines, "", SortCb)
    result := []
    Loop Parse, lines, "`n" {
        if A_LoopField = ""
            continue
        result.Push(StrSplit(A_LoopField, "`t"))
    }
    return result
}

OnNgramColClick(lvCtrl, colNum) {
    static sortState := Map()
    key := colNum
    dir := (sortState.Has(key) && sortState[key] = "desc") ? "asc" : "desc"
    sortState[key] := dir

    rowCount := lvCtrl.GetCount()
    colCount := lvCtrl.GetCount("Col")
    rows := []
    Loop rowCount {
        r := A_Index
        row := []
        Loop colCount
            row.Push(lvCtrl.GetText(r, A_Index))
        rows.Push(row)
    }
    if rows.Length = 0
        return
    rows := SortRowsByCol(rows, colNum, dir)
    lvCtrl.Delete()
    for row in rows
        lvCtrl.Add("", row*)
}

PadRightTTB(str, width) {
    str := String(str)
    Loop (width - StrLen(str))
        str .= " "
    return str
}

; ============================================================
;  HELP SYSTEM
; ============================================================
; tabIndex 0 = general overview; 1-10 = per-tab help.
; Called by the [?] toolbar button (general) and F1 hotkey (current tab).
ShowHelp(tabIndex) {
    static helpGui := 0

    ; Destroy any existing help window first
    if IsObject(helpGui)
        try helpGui.Destroy()

    helpTexts := Map(
        0,
"Main Gui`n`nOVERVIEW`n" .
"Text Toolbox is a dual-pane text transformation utility.  Paste or type text into the " .
"Input pane on the left (or top), choose a tab and options, then click Apply.  " .
"The transformed result appears in the Output pane.`n`n" .
"DEDICATION`n" .
"TextToolbox is dedicated to the AHK forum member, Tidbit, and to his memory.  " .
"The app is inspired by, and based on, his Text Toolbox v1.1,`n" .
"https://www.autohotkey.com/boards/viewtopic.php?f=6&t=14515 `n" .
"which I've used regularly since it was shared, in 2016.  The it looks similar and " .
"functions similarly.  First versions of this tool were created by feeding screenshots " .
"of Tidbit's tool to Claude AI.  Thanks go to forum user Just me, for helping to hide the " .
"TabBar in AHKv2.`n`n" .
"TOOLBAR (top row)`n" .
"  Paste        — pastes clipboard into Input pane`n" .
"  Clear (in)   — clears Input pane (undo-able)`n" .
"  Wrap         — toggles word-wrap on both panes`n" .
"  ⇕ Vert / ⇔ Horiz — toggles side-by-side vs. stacked pane layout`n" .
"  Copy Out     — copies Output pane to clipboard`n" .
"  Clear (out)  — clears Output pane`n" .
"  Undo / Redo  — multi-level undo/redo (up to 20 levels)`n" .
"  [?]          — show this help window`n`n" .
"NAVIGATION`n" .
"  Click a tab name in the left list to switch tabs.`n`n" .
"APPLY / SWAP STRIP`n" .
"  The narrow strip between the two panes holds Apply and Swap.`n" .
"  Apply runs the active tab's transform on the Input text.`n" .
"  Swap exchanges the contents of the two panes (undo-able).`n`n" .
"KEYBOARD SHORTCUTS`n" .
"  Ctrl+Z / Ctrl+Y  — Undo / Redo (when no Edit box has focus)`n" .
"  F1               — Help for the currently active tab`n" .
"  Esc              — Close`n`n" .
"SETTINGS`n" .
"  Window size, split orientation, last active tab, and Find/Replace " .
"history are saved automatically to ttSettings.ini in the script folder.",

        1,
"Case Tab`n`n" .
"Converts the capitalisation of every line in the Input pane.`n`n" .
"MODES`n" .
"  UPPERCASE         — all letters uppercased`n" .
"  lowercase         — all letters lowercased`n" .
"  Title Case        — first letter of every word uppercased`n" .
"  Sentence case     — first letter of each line uppercased, rest lower`n" .
"  iNVERT cAsE       — swaps upper↔lower for every letter`n" .
"  sArCaStIc CaSe    — alternates lower/upper on each alphabetic character`n" .
"  kebab-case        — words joined with hyphens, all lower`n" .
"  snake_case        — words joined with underscores, all lower`n" .
"  PascalCase        — each word capitalised, no separator`n" .
"  camelCase         — like Pascal but first word stays lowercase`n" .
"  SCREAMING_SNAKE   — snake_case in all caps`n" .
"  Train-Case        — like kebab but each word is Title-cased`n`n" .
"TIP: For the compound-word modes (kebab, snake, Pascal, etc.) the input " .
"is split on spaces, hyphens, underscores, and dots before rejoining.",

        2,
"Sort Tab`n`n" .
"Sorts the lines of the Input pane and writes the result to Output.`n`n" .
"MODES`n" .
"  A → Z               — alphabetical ascending (case-insensitive)`n" .
"  Z → A               — alphabetical descending`n" .
"  Random              — random shuffle`n" .
"  Length (short→long) — shorter lines first`n" .
"  Length (long→short) — longer lines first`n" .
"  Numeric             — natural/numeric sort (e.g. 2 before 10)`n`n" .
"OPTIONS`n" .
"  Remove duplicate lines  — keeps only the first occurrence of each line`n" .
"  Trim whitespace         — strips leading/trailing spaces before comparing`n`n" .
"SORT KEY`n" .
"  Controls which part of each line is used for the comparison.  " .
"The full original line is always kept in the output — only the comparison key changes.`n`n" .
"  Whole line       — default; sorts on the entire line`n" .
"  Word 1           — sorts on the first word only`n" .
"  Word 2           — sorts on the second word (e.g. last name in 'First Last' lists)`n" .
"  Last word        — sorts on the final word`n" .
"  After delimiter  — sorts on the text following the first occurrence of the " .
"delimiter character (set in the small box to the right of the dropdown)`n`n" .
"  The Sort key applies to A→Z, Z→A, and Numeric modes.  " .
"Length and Random always use the whole line.`n`n" .
"EXAMPLES`n" .
"  'First Last' name list → mode A→Z, key Word 2`n" .
"  'red-dog / blue-fish'  → mode A→Z, key After delimiter, delim '-'`n" .
"  '10. item / 2. item'   → mode Numeric, key Word 2 (skips the number prefix)",

        3,
"Find / Replace Tab`n`n" .
"Searches the Input text and writes the substituted result to Output.`n`n" .
"CONTROLS`n" .
"  Find            — the text (or pattern) to search for`n" .
"  Replace with    — the replacement text`n" .
"  Case sensitive  — when unchecked, search ignores capitalisation`n" .
"  Use regex       — treats the Find field as a PCRE regular expression`n" .
"  Replace all     — replaces every match; uncheck to replace only the first`n`n" .
"HISTORY`n" .
"  Both ComboBoxes remember up to 10 recent entries.  " .
"Click the dropdown arrow to revisit a previous search or replacement.  " .
"History is saved to ttSettings.ini and restored on next launch.`n`n" .
"REGEX EXAMPLES`n" .
"  Swap first two words on each line`n" .
"    Find:    ^(\w+)\s+(\w+)`n" .
"    Replace: $2 $1`n`n" .
"  Remove leading numbers/bullets (e.g. '1. ' or '- ')`n" .
"    Find:    ^\s*[\d]+[.)]\s*|^\s*[-*•]\s*`n" .
"    Replace: (leave blank)`n`n" .
"  Collapse multiple blank lines into one`n" .
"    Find:    \n{3,}`n" .
"    Replace: \n\n`n`n" .
"  Trim trailing whitespace from every line`n" .
"    Find:    [ \t]+$`n" .
"    Replace: (leave blank)`n`n" .
"  Wrap every line in HTML <li> tags`n" .
"    Find:    ^(.+)$`n" .
"    Replace: <li>$1</li>`n`n" .
"  Extract email addresses (keep only the match)`n" .
"    Find:    ^.*?([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}).*$`n" .
"    Replace: $1`n`n" .
"  Convert Windows path separators to forward slashes`n" .
"    Find:    \\`n" .
"    Replace: /`n`n" .
"  Add a comma to the end of every non-empty line`n" .
"    Find:    (.+)$`n" .
"    Replace: $1,`n`n" .
"  Double-space (insert blank line after each line)`n" .
"    Find:    \n`n" .
"    Replace: \n\n`n`n" .
"QUICK REFERENCE`n" .
"  .   any char      \d  digit       \w  word char   \s  whitespace`n" .
"  *   0 or more     +   1 or more   ?   0 or 1      {n} exactly n`n" .
"  ^   line start    $   line end    |   or          ()  capture group`n" .
"  $1 $2 …  insert capture groups in the Replace field",

        4,
"Remove Tab`n`n" .
"Strips unwanted content from the Input text.  Multiple options can be " .
"checked at once; they are applied in the order shown.`n`n" .
"OPTIONS`n" .
"  Blank / empty lines          — removes lines that are empty or whitespace-only`n" .
"  Duplicate lines              — keeps only the first occurrence of each line`n" .
"  Leading whitespace           — strips spaces/tabs from the start of each line`n" .
"  Trailing whitespace          — strips spaces/tabs from the end of each line`n" .
"  HTML tags <…>                — removes anything matching <…> angle-bracket tags`n" .
"  BBCode tags […]              — removes anything matching […] square-bracket tags`n" .
"  Non-ASCII characters         — removes any character outside the 0-127 ASCII range`n`n" .
"TRIM N CHARS`n" .
"  LEFT  — removes the first N characters from every line`n" .
"  RIGHT — removes the last N characters from every line`n`n" .
"TIP: Set LEFT=0 and RIGHT=0 to skip char-trimming while still using the checkboxes above.",

        5,
"Extract Tab`n`n" .
"Pulls specific content out of the Input text and writes only the matches to Output.`n`n" .
"MODES`n" .
"  Numbers              — extracts integers, decimals, and negative numbers`n" .
"                         Example: 'Score = 85- Average' → '85'`n" .
"  Email addresses      — extracts anything matching an email pattern`n" .
"  URLs                 — extracts http/https URLs`n" .
"  Lines matching       — keeps only lines that contain the given pattern (like grep)`n" .
"                         Plain text or regex; e.g. 'Score' or '\d{2,3}'`n" .
"  Custom pattern       — extracts regex matches; if your pattern has a capture`n" .
"                         group (…) only the captured portion is returned`n`n" .
"OUTPUT OPTIONS`n" .
"  Unique only          — removes duplicate matches (case-insensitive)`n" .
"  One per line         — each match on its own line (default)`n" .
"  Semicolon-separated  — joins all matches with '; ' on a single line`n" .
"                         (useful for pasting email lists into Outlook)`n" .
"                         Only available when 'One per line' is unchecked`n`n" .
"TIPS`n" .
"  For test scores like 'SS = 85- Average', Numbers mode extracts all numeric`n" .
"  values.  If a line has multiple numbers (e.g. 'Item 3: score 85'), use`n" .
"  'Lines matching' with a pattern like '\d{2,3}' to filter first, then`n" .
"  run Numbers mode in a second pass.`n" .
"  Custom pattern with a capture group: pattern '\bScore\s*=\s*(\d+)' would`n" .
"  extract only the digits after 'Score ='.",

        6,
"Wrap / Indent Tab`n`n" .
"Reformats line lengths and indentation.`n`n" .
"WORD WRAP`n" .
"  Hard-wraps the text at the specified column width.  " .
"Long lines are broken at word boundaries so no line exceeds the column count.  " .
"Set the column number in the 'Word wrap at column' field (default 80).`n`n" .
"INDENT / DEDENT`n" .
"  Indent  — adds N spaces to the beginning of every non-empty line`n" .
"  Dedent  — removes up to N spaces from the beginning of every line`n" .
"  'Spaces per level' sets how many spaces N is (default 4)`n`n" .
"TIP: Apply Wrap first, then Indent in a second Apply step if you need both.  " .
"Or use Swap to move the output back to Input and apply the second transform.",

        7,
"Counter Tab`n`n" .
"Prepends an incrementing number to each line.`n`n" .
"OPTIONS`n" .
"  Start at   — the number assigned to the first line (default 1)`n" .
"  Step       — how much each counter increments (default 1; use 2 for even/odd lists)`n" .
"  Separator  — text inserted between the number and the line content (default '. ')`n" .
"  Right-align numbers — pads shorter numbers with spaces so all numbers " .
"                        align to the width of the largest number`n`n" .
"EXAMPLES`n" .
"  Start=1, Step=1, Sep='. '   →  1. Alpha  /  2. Beta  /  3. Gamma`n" .
"  Start=0, Step=10, Sep=': '  →  0: Alpha  /  10: Beta  /  20: Gamma`n" .
"  With right-align on a 12-line list: ' 1. Alpha' … '12. Omega'",

        8,
"Padding Tab`n`n" .
"Pads all lines to the same length so they form a neat rectangular block.  " .
"Useful for monospace output, columns, or ASCII art.`n`n" .
"ALIGNMENT`n" .
"  Left-align   — content at left, padding added to the right`n" .
"  Right-align  — content at right, padding added to the left`n" .
"  Center       — content centred, padding split evenly on both sides`n`n" .
"PAD CHARACTER`n" .
"  The character used to fill the extra space (default: space).  " .
"You can enter any single character — e.g. '-', '.', '0'.`n`n" .
"TIP: For best results use a fixed-width (monospace) font when viewing " .
"the output, such as Courier New or Consolas.",

        9,
"CSV View Tab`n`n" .
"Parses CSV (comma-separated values) text from the Input pane and " .
"displays it as a sortable, interactive table.`n`n" .
"USAGE`n" .
"  1. Paste CSV into the Input pane.`n" .
"  2. Switch to CSV View and click Apply (or press the Apply button).`n" .
"  3. The first row is treated as column headers.`n`n" .
"INTERACTIVE FEATURES`n" .
"  Click a column header  — sort ascending; click again for descending`n" .
"  Drag a column header   — reorder columns`n" .
"  Remove col dropdown    — select a column name then click Remove to delete it`n`n" .
"EXPORT`n" .
"  '▼ Send to Output pane' exports the current table (respecting column " .
"order and removed columns) back to CSV text in the Output pane.`n" .
"  'Use double-quotes' — forces quoting on every field (not just fields " .
"that contain commas or quotes).`n`n" .
"LIMITATION`n" .
"  Do not both reorder columns AND sort in the same session — the " .
"combination can corrupt column display.  Workaround: reorder → export " .
"→ Swap → Apply → sort → export again.",

        10,
"Compare Tab`n`n" .
"Compares two sets of lines and returns only the lines matching a chosen " .
"relationship.  The Input pane holds the 'Original' text; type or paste " .
"the second set into the 'Comparison text' box on the right.`n`n" .
"RETURN MODES`n" .
"  Original only   — lines present in Original but NOT in Comparison`n" .
"  Comparison only — lines present in Comparison but NOT in Original`n" .
"  In both         — lines that appear in both sets (intersection)`n`n" .
"POST-PROCESS OPTIONS`n" .
"  Remove duplicates — de-duplicates the result list`n" .
"  Alphabetize       — sorts the result list A→Z`n`n" .
"TIP: Comparison is line-by-line and case-sensitive.  Use the Remove tab " .
"with 'Trim whitespace' first if your lines may have stray spaces.",

        11,
"N-Grams Tab`n`n" .
"Counts how often words or multi-word phrases appear in the Input text " .
"and displays the results in a sortable table.`n`n" .
"CONTROLS`n" .
"  Group size  — number of words per phrase (1–6)`n" .
"               1 word  = simple word-frequency count`n" .
"               2 words = bigrams (pairs), 3 = trigrams, etc.`n" .
"  Refresh     — (re)build the table from the current Input text`n" .
"  Export      — write the table as aligned plain text to the Output pane`n`n" .
"COLUMNS`n" .
"  Phrase / Word — the word or phrase`n" .
"  Count         — how many times it appears`n" .
"  Percent       — count as a share of all token occurrences`n`n" .
"  Click any column header to sort ascending/descending.`n`n" .
"EXPORT FORMAT`n" .
"  The Export button writes a fixed-width, human-readable table to the " .
"Output pane.  You can then copy, save, or further process it.  " .
"The export respects the current sort order of the table.`n`n" .
"TIPS`n" .
"  N=1 is a quick word-cloud alternative — it shows the most repeated words.`n" .
"  N=2 (bigrams) often reveals characteristic phrases and collocations.`n" .
"  Very large texts with high N values may take a moment to compute.",

        12,
"Stats Tab`n`n" .
"Counts various statistics about the text in the Input pane.  " .
"Click Apply to analyse and populate the table.`n`n" .
"STATISTICS`n" .
"  Words                      — total word count`n" .
"  Unique words               — count of distinct words (case-insensitive)`n" .
"  Type-Token Ratio           — unique words / total words as a percentage; " .
"higher = more varied vocabulary`n" .
"  Lines                      — total line count (including blank lines)`n" .
"  Paragraphs (est.)          — blank-line-separated blocks`n" .
"  Sentences (est.)           — punctuation-delimited sentences; falls back to " .
"lines for lists/lyrics`n" .
"  Characters (with spaces)   — total character count`n" .
"  Characters (no spaces)     — character count excluding all whitespace`n" .
"  Avg word length            — mean letter-count per word`n" .
"  Avg sentence length        — mean words per sentence`n" .
"  Longest word               — the word with the most letters`n" .
"  Total syllables (est.)     — vowel-group heuristic; used for Flesch-Kincaid`n" .
"  Hapax legomena             — words appearing exactly once; high count = rich vocab`n" .
"  Flesch-Kincaid Reading Ease — 0-100 score: 90+ = very easy, 60 = standard, " .
"30- = difficult (academic/legal)`n`n" .
"TIP: Stats does not write to the Output pane — it only populates the " .
"table.  Use the N-Grams tab if you want exportable word/phrase frequency data."
    )

    tabNames := ["Case", "Sort", "Find/Replace", "Remove", "Extract", "Wrap/Indent",
                 "Counter", "Padding", "CSV View", "Compare", "N-Grams", "Stats"]

    if (tabIndex = 0)
        title := "Text Toolbox — Help"
    else
        title := "Help — " tabNames[tabIndex] " tab"

    helpText := helpTexts.Has(tabIndex) ? helpTexts[tabIndex] : helpTexts[0]

    helpGui := Gui("+AlwaysOnTop", title)
    helpGui.SetFont("s9", "Segoe UI")
    helpGui.Add("Edit",
        "x10 y10 w560 h380 ReadOnly -E0x200 -WantReturn -TabStop Multi VScroll",
        helpText)
    closeBtn := helpGui.Add("Button", "x240 y+8 w100 Default", "Close")
    closeBtn.OnEvent("Click", (*) => helpGui.Destroy())
    helpGui.OnEvent("Escape", (*) => helpGui.Destroy())
    helpGui.Show("AutoSize")
}

; ============================================================
;  CASE HELPERS
; ============================================================

SplitWords(s) {
    s := RegExReplace(s, "[-_.]", " ")
    s := Trim(s)
    return StrSplit(RegExReplace(s, "\s+", " "), " ")
}

TitleCase(s) {
    out     := ""
    capNext := true
    Loop Parse, s {
        ch := A_LoopField
        if RegExMatch(ch, "\s") {
            out .= ch
            capNext := true
        } else if capNext {
            out .= StrUpper(ch)
            capNext := false
        } else {
            out .= StrLower(ch)
        }
    }
    return out
}

SentenceCase(s) {
    s := StrLower(s)
    if RegExMatch(s, "^\s*\K\S", &m)
        s := SubStr(s, 1, m.Pos-1) . StrUpper(m[0]) . SubStr(s, m.Pos+1)
    return s
}

InvertCase(s) {
    out := ""
    Loop Parse, s {
        ch  := A_LoopField
        out .= (ch == StrUpper(ch)) ? StrLower(ch) : StrUpper(ch)
    }
    return out
}

SarcasticCase(s) {
    out  := ""
    flip := true
    Loop Parse, s {
        ch := A_LoopField
        if RegExMatch(ch, "[a-zA-Z]") {
            out  .= flip ? StrLower(ch) : StrUpper(ch)
            flip := !flip
        } else {
            out .= ch
        }
    }
    return out
}

KebabCase(s) {
    words := SplitWords(StrLower(s))
    result := ""
    for i, w in words
        result .= (i > 1 ? "-" : "") . w
    return result
}

SnakeCase(s) {
    words := SplitWords(StrLower(s))
    result := ""
    for i, w in words
        result .= (i > 1 ? "_" : "") . w
    return result
}

PascalCase(s) {
    words  := SplitWords(s)
    result := ""
    for w in words
        result .= StrUpper(SubStr(w, 1, 1)) . StrLower(SubStr(w, 2))
    return result
}

CamelCase(s) {
    words  := SplitWords(s)
    result := ""
    for i, w in words {
        if (i = 1)
            result .= StrLower(w)
        else
            result .= StrUpper(SubStr(w, 1, 1)) . StrLower(SubStr(w, 2))
    }
    return result
}

ScreamingSnakeCase(s) {
    words := SplitWords(StrUpper(s))
    result := ""
    for i, w in words
        result .= (i > 1 ? "_" : "") . w
    return result
}

TrainCase(s) {
    words  := SplitWords(s)
    result := ""
    for i, w in words
        result .= (i > 1 ? "-" : "") . StrUpper(SubStr(w, 1, 1)) . StrLower(SubStr(w, 2))
    return result
}

; ============================================================
;  UTILITY
; ============================================================

RepeatChar(ch, n) {
    out := ""
    Loop n
        out .= ch
    return out
}

; Trim leading and trailing whitespace from each line without splitting into an array
TrimLines(txt) {
    txt := RegExReplace(txt, "m)^[ \t]+", "")   ; strip leading whitespace per line
    txt := RegExReplace(txt, "m)[ \t]+$", "")   ; strip trailing whitespace per line
    return txt
}

JoinLines(arr) {
    result := ""
    for i, line in arr
        result .= (i > 1 ? "`n" : "") . line
    return result
}
