Attribute VB_Name = "StickShiftDashboard"
' =====================================================================
'  StickShift Dashboard — one-time setup for a macro-button control panel
'
'  Run CreateStickShiftDashboard once to build the "StickShift" sheet.
'  Re-run at any time to reset the sheet (e.g. after importing into a
'  new workbook). The sheet itself has no persistent state — it is a
'  pure UI layer over the macros.
'
'  Public surface used by other modules:
'    CONTEXT_CELL        — cell address where the current context path lives
'    RefreshContextDisplay — rewrites CONTEXT_CELL from BundleRootRaw()
'    ShowStickShiftReadme  — renders README into the "Read Me" sheet
'
'  Requires the other modules in the same workbook:
'    StickShiftConfig         -> BundleRootRaw, SetBundleRoot
'    StickShiftWriteApply     -> ApplyStickShiftWrite, ApplyWriteEnvelopeText
'    StickShiftIndexGenerator -> GenerateStickShiftIndexes
'    StickShiftLint           -> RunStickShiftLint
'    StickShiftContextBundle  -> BuildContextBundle
'    StickShiftBootstrap      -> BootstrapBundle
' =====================================================================

Option Explicit

Public Const CONTEXT_CELL As String = "B6"

Private Const SHEET_NAME As String = "StickShift"


Sub CreateStickShiftDashboard()
    ' -- Get or create the sheet --------------------------------------------------
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SHEET_NAME
    Else
        ws.Cells.Clear
        Dim shp As Shape
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' -- Column widths (in character units) ---------------------------------------
    ws.Columns("A").ColumnWidth = 2      ' left margin
    ws.Columns("B").ColumnWidth = 24     ' button + value left
    ws.Columns("C").ColumnWidth = 16     ' value right / inner gap
    ws.Columns("D").ColumnWidth = 22     ' context button + description right
    ws.Columns("E").ColumnWidth = 2      ' right margin

    ' -- Row heights (in points) --------------------------------------------------
    ws.Rows("1").RowHeight = 8      ' top padding
    ws.Rows("2").RowHeight = 36     ' title + Read Me button
    ws.Rows("3").RowHeight = 20     ' tagline
    ws.Rows("4").RowHeight = 10     ' gap before context bar
    ws.Rows("5").RowHeight = 16     ' context bar label
    ws.Rows("6").RowHeight = 22     ' context value (CONTEXT_CELL)
    ws.Rows("7").RowHeight = 46     ' Switch Context button
    ws.Rows("8").RowHeight = 34     ' Switch Context description
    ws.Rows("9").RowHeight = 14     ' divider
    ws.Rows("10").RowHeight = 46    ' button 1 Initialize Context
    ws.Rows("11").RowHeight = 34    ' description 1
    ws.Rows("12").RowHeight = 20    ' inline section divider
    ws.Rows("13").RowHeight = 46    ' button 2 Build Context Bundle
    ws.Rows("14").RowHeight = 34    ' description 2
    ws.Rows("15").RowHeight = 46    ' button 3 Apply Write Envelope
    ws.Rows("16").RowHeight = 34    ' description 3
    ws.Rows("17").RowHeight = 10    ' light divider
    ws.Rows("18").RowHeight = 46    ' button 4 Regenerate Index
    ws.Rows("19").RowHeight = 34    ' description 4
    ws.Rows("20").RowHeight = 46    ' button 5 Run Linter
    ws.Rows("21").RowHeight = 34    ' description 5
    ws.Rows("22").RowHeight = 10    ' gap
    ws.Rows("23").RowHeight = 20    ' footer hint
    ws.Rows("24").RowHeight = 14    ' footer brand
    ws.Rows("25").RowHeight = 8     ' bottom padding

    ' -- Background ---------------------------------------------------------------
    ws.Cells.Interior.Color = RGB(248, 250, 252)

    ' -- Title --------------------------------------------------------------------
    Dim r As Range
    Set r = ws.Range("B2:C2"): r.Merge
    r.Value = "StickShift"
    r.Font.Size = 17: r.Font.Bold = True
    r.Font.Color = RGB(15, 23, 42)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    ' -- Read Me button (top-right) -----------------------------------------------
    Dim readMeCell As Range
    Set readMeCell = ws.Range("D2")
    Const RM_INSET As Double = 4
    Dim rmShape As Shape
    Set rmShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                      readMeCell.Left + RM_INSET, _
                                      readMeCell.Top + RM_INSET, _
                                      readMeCell.Width - RM_INSET * 2, _
                                      readMeCell.Height - RM_INSET * 2)
    With rmShape
        .Name = "btn_ShowStickShiftReadme"
        .OnAction = "ShowStickShiftReadme"
        .Fill.ForeColor.RGB = RGB(100, 116, 139)
        .Fill.Solid
        .Line.Visible = msoFalse
        On Error Resume Next
        .Adjustments(1) = 0.18
        On Error GoTo 0
        With .TextFrame2
            .TextRange.Text = ChrW(128218) & " Read Me"
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange.Font
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Size = 9
                .Bold = msoTrue
            End With
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    ' -- Tagline ------------------------------------------------------------------
    Set r = ws.Range("B3:D3"): r.Merge
    r.Value = "Everything automatic AI promises " & ChrW(8212) & " except your hand's on the StickShift."
    r.Font.Size = 9: r.Font.Color = RGB(100, 116, 139)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    r.Borders(xlEdgeBottom).Weight = xlThin

    ' -- Context bar label --------------------------------------------------------
    Set r = ws.Range("B5:D5"): r.Merge
    r.Value = "Current context:"
    r.Font.Size = 8: r.Font.Bold = True
    r.Font.Color = RGB(71, 85, 105)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(241, 245, 249)
    r.Borders(xlEdgeTop).LineStyle = xlContinuous
    r.Borders(xlEdgeTop).Color = RGB(203, 213, 225)
    r.Borders(xlEdgeTop).Weight = xlThin

    ' -- Context value cell (CONTEXT_CELL) ----------------------------------------
    Dim rootVal As String
    rootVal = StickShiftConfig.BundleRootRaw()
    If rootVal = "" Then rootVal = "(none set " & ChrW(8212) & " click Switch Context)"

    Set r = ws.Range("B6:D6"): r.Merge
    r.Value = rootVal
    r.Font.Size = 9
    r.Font.Color = IIf(StickShiftConfig.BundleRootRaw() = "", RGB(148, 163, 184), RGB(15, 23, 42))
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(241, 245, 249)
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    r.Borders(xlEdgeBottom).Weight = xlThin
    r.Borders(xlEdgeLeft).LineStyle = xlContinuous
    r.Borders(xlEdgeLeft).Color = RGB(109, 40, 217)
    r.Borders(xlEdgeLeft).Weight = xlMedium

    ' -- Switch Context button + description (rows 7-8) ---------------------------
    MakeButton ws, 7, _
        "Switch Context", "SetBundleRoot", RGB(71, 85, 105), _
        "Pick or change the folder StickShift is pointed at. Switch anytime " & Chr(10) & _
        Chr(8212) & " personal, shared, or per-project. Remembered per machine."

    ' -- Divider ------------------------------------------------------------------
    Set r = ws.Range("B9:D9"): r.Merge
    r.Interior.Color = RGB(248, 250, 252)
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).Color = RGB(203, 213, 225)
    r.Borders(xlEdgeBottom).Weight = xlThin

    ' -- Button 1: Initialize Context ---------------------------------------------
    MakeButton ws, 10, _
        "1  Initialize Context", "BootstrapBundle", RGB(5, 150, 105), _
        "Seeds a new/empty context with starter files (incl. the skill-authoring skill)" & Chr(10) & _
        "and builds the listings. Safe to re-run " & Chr(8212) & " skips files that already exist."

    ' -- Inline section divider ---------------------------------------------------
    Set r = ws.Range("B12:D12"): r.Merge
    r.Value = ChrW(8593) & " one-time setup   " & ChrW(183) & "   " & ChrW(8595) & " the loop"
    r.Font.Size = 8: r.Font.Color = RGB(148, 163, 184): r.Font.Italic = True
    r.HorizontalAlignment = xlCenter
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    ' -- Button 2: Build Context Bundle -------------------------------------------
    MakeButton ws, 13, _
        "2  Build Context Bundle", "BuildContextBundle", RGB(109, 40, 217), _
        "Copy a <CONTEXT_REQUEST> from the chat, then click. Assembles" & Chr(10) & _
        "StickShift-context.md in the -dist folder to paste back into the chat."

    ' -- Button 3: Apply Write Envelope -------------------------------------------
    MakeButton ws, 15, _
        "3  Apply Write Envelope", "ApplyStickShiftWrite", RGB(37, 99, 235), _
        "Copy a <VBA_WRITE> block from the chat, then click. Writes the files" & Chr(10) & _
        "into your context. Every write is recorded in log.md."

    ' -- Light divider between everyday and maintenance ----------------------------
    Set r = ws.Range("B17:D17"): r.Merge
    r.Interior.Color = RGB(248, 250, 252)
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    r.Borders(xlEdgeBottom).Weight = xlThin

    ' -- Button 4: Regenerate Index -----------------------------------------------
    MakeButton ws, 18, _
        "4  Regenerate Index", "GenerateStickShiftIndexes", RGB(22, 163, 74), _
        "Rebuilds the listings. Runs automatically after every write; this is the" & Chr(10) & _
        "manual version."

    ' -- Button 5: Run Linter -----------------------------------------------------
    MakeButton ws, 20, _
        "5  Run Linter", "RunStickShiftLint", RGB(220, 85, 10), _
        "Scans the context: broken links, WIP violations, stalls, active-to-archived." & Chr(10) & _
        "Findings appear colour-coded in the ""StickShift Lint Report"" sheet."

    ' -- Footer -------------------------------------------------------------------
    Set r = ws.Range("B23:D23"): r.Merge
    r.Value = "First time? Set your context in the bar above, then click Initialize Context."
    r.Font.Size = 8: r.Font.Color = RGB(100, 116, 139)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    Set r = ws.Range("B24:D24"): r.Merge
    r.Value = ChrW(9881) & " Part of StickShift"
    r.Font.Size = 8: r.Font.Color = RGB(148, 163, 184)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    ws.Range("A1").Select
    MsgBox "StickShift dashboard ready. Set your context in the top bar, then click Initialize Context.", _
           vbInformation, "StickShift"
End Sub


Public Sub RefreshContextDisplay()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim rootVal As String
    rootVal = StickShiftConfig.BundleRootRaw()

    Dim r As Range
    Set r = ws.Range(CONTEXT_CELL)
    If rootVal = "" Then
        r.Value = "(none set " & ChrW(8212) & " click Switch Context)"
        r.Font.Color = RGB(148, 163, 184)
    Else
        r.Value = rootVal
        r.Font.Color = RGB(15, 23, 42)
    End If
End Sub


Public Sub ShowStickShiftReadme()
    Const RM_SHEET As String = "Read Me"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(RM_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = RM_SHEET
    Else
        ws.Cells.Clear
        Dim shp As Shape
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Interior.Color = RGB(248, 250, 252)

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 80
    ws.Columns("C").ColumnWidth = 3

    Dim readmeText As String
    readmeText = _
        "StickShift " & ChrW(8212) & " READ ME FIRST" & vbLf & _
        "Everything automatic AI promises " & ChrW(8212) & " except your hand" & Chr(39) & "s on the StickShift." & vbLf & vbLf & _
        "A manual-activation knowledge agent. It runs entirely inside this workbook." & vbLf & vbLf & _
        String(60, "-") & vbLf & _
        "WHAT THIS IS" & vbLf & _
        String(60, "-") & vbLf & _
        "This connects your AI chat to a folder of markdown notes (a " & Chr(34) & "context" & Chr(34) & ")." & vbLf & _
        "The AI asks for context, or proposes file edits. You copy what it gives you," & vbLf & _
        "click a button here, and the workbook does the work." & vbLf & vbLf & _
        "Nothing runs on its own. Nothing happens until you click. You stay in the loop." & vbLf & vbLf & _
        String(60, "-") & vbLf & _
        "YOUR CONTEXT  (the bar at the top)" & vbLf & _
        String(60, "-") & vbLf & _
        "The top of the dashboard shows your current context - the folder StickShift is" & vbLf & _
        "pointed at right now. The " & Chr(34) & "Switch Context" & Chr(34) & " button changes it." & vbLf & vbLf & _
        "You can keep more than one: a personal context, a shared team context, a" & vbLf & _
        "per-project context. Switch between them anytime; StickShift remembers the last" & vbLf & _
        "one per machine." & vbLf & vbLf & _
        String(60, "-") & vbLf & _
        "SET UP - TWO CLICKS" & vbLf & _
        String(60, "-") & vbLf & _
        "1. Enable macros: when the file opens, click the yellow bar at the top" & vbLf & _
        "   (" & Chr(34) & "Enable Content" & Chr(34) & ")." & vbLf & _
        "   (If you see a red " & Chr(34) & "macros have been blocked" & Chr(34) & " bar instead: close Excel," & vbLf & _
        "   right-click the file -> Properties -> tick " & Chr(34) & "Unblock" & Chr(34) & " -> OK, then reopen.)" & vbLf & vbLf & _
        "2. Set your context: click " & Chr(34) & "Switch Context" & Chr(34) & " in the top bar and pick (or make) a" & vbLf & _
        "   folder - for example C:\StickShift. The name doesn" & Chr(39) & "t matter; the tool" & vbLf & _
        "   remembers the path." & vbLf & vbLf & _
        "3. Click button 1, " & Chr(34) & "Initialize Context" & Chr(34) & ". This seeds the starter files -" & vbLf & _
        "   including a skill that teaches the AI how to write more skills - and builds" & vbLf & _
        "   the listings the AI reads. That" & Chr(39) & "s setup done." & vbLf & vbLf & _
        "   (Switching to a context someone already set up? Skip step 3 - just start" & vbLf & _
        "   using it.)" & vbLf & vbLf & _
        String(60, "-") & vbLf & _
        "THE LOOP - THIS IS THE WHOLE THING" & vbLf & _
        String(60, "-") & vbLf & _
        "   copy from chat  ->  click a button here  ->  get the result  ->  back to chat" & vbLf & vbLf & _
        "Two buttons do the real work:" & vbLf & vbLf & _
        "   Button 2 - BUILD CONTEXT BUNDLE" & vbLf & _
        "     Use when the AI replies with a <CONTEXT_REQUEST> block." & vbLf & _
        "     Copy the whole block, click button 2, and the assembled context opens as" & vbLf & _
        "     StickShift-context.md. Paste that back into the chat." & vbLf & vbLf & _
        "   Button 3 - APPLY WRITE ENVELOPE" & vbLf & _
        "     Use when the AI replies with a <VBA_WRITE> block (file edits)." & vbLf & _
        "     Copy it, click button 3, and the files are written into your context." & vbLf & _
        "     Every write is recorded in log.md, so you can always see what changed." & vbLf & vbLf & _
        String(60, "-") & vbLf & _
        "TRY THIS FIRST - your first round-trips" & vbLf & _
        String(60, "-") & vbLf & _
        "Your context comes seeded with a " & Chr(34) & "skill" & Chr(34) & " - a reusable procedure the AI can pick" & vbLf & _
        "up and follow. This one teaches the AI how to write more skills. So your first" & vbLf & _
        "session is: watch it find the skill, then use it to make your own." & vbLf & vbLf & _
        "ROUND-TRIP 1 - see retrieval work:" & vbLf & _
        "1. In your AI chat, ask: " & Chr(34) & "What skills are in my context?" & Chr(34) & vbLf & _
        "2. The AI replies with a <CONTEXT_REQUEST> block - that" & Chr(39) & "s it asking to look in" & vbLf & _
        "   the folder. Copy the whole block." & vbLf & _
        "3. Click button 2, " & Chr(34) & "Build Context Bundle" & Chr(34) & ". StickShift-context.md opens. Paste" & vbLf & _
        "   its contents back into the chat." & vbLf & _
        "4. The AI now lists your skills - including " & Chr(34) & "Skill MD Authoring" & Chr(34) & "." & vbLf & vbLf & _
        "ROUND-TRIP 2 - make something of your own:" & vbLf & _
        "5. Ask: " & Chr(34) & "Use the skill-authoring skill to help me write a new skill for" & vbLf & _
        "   <a task you actually do>." & Chr(34) & vbLf & _
        "6. The AI may ask for the full procedure first (another <CONTEXT_REQUEST> - same" & vbLf & _
        "   button 2). Then it walks you through it and hands you a <VBA_WRITE> block." & vbLf & _
        "7. Copy that block, click button 3, " & Chr(34) & "Apply Write Envelope" & Chr(34) & "." & vbLf & _
        "8. Open your skills folder - your new skill is there. Open log.md - the write is" & vbLf & _
        "   recorded, with a timestamp." & vbLf & vbLf & _
        "That is the whole agent: it found context for you, helped you build something," & vbLf & _
        "wrote it where it belongs, and kept a record - and nothing happened until you" & vbLf & _
        "clicked. You were in the loop the entire time." & vbLf & vbLf & _
        String(60, "-") & vbLf & _
        "THE OTHER BUTTONS - not needed on day one" & vbLf & _
        String(60, "-") & vbLf & _
        "   Button 4 - Regenerate Index : rebuilds the listings. Runs automatically after" & vbLf & _
        "              every write; button 4 is the manual version." & vbLf & _
        "   Button 5 - Run Linter       : checks the context for problems (broken links," & vbLf & _
        "              stalls, etc.). Findings appear in the " & Chr(34) & "StickShift Lint Report" & Chr(34) & " sheet." & vbLf & vbLf & _
        String(60, "-") & vbLf & _
        "IF YOU GET STUCK" & vbLf & _
        String(60, "-") & vbLf & _
        "Note exactly where you hesitated - that spot is genuinely useful; it tells us" & vbLf & _
        "what to make smoother for the next person."

    Dim r As Range
    Set r = ws.Range("B2")
    r.Value = readmeText
    r.Font.Name = "Consolas"
    r.Font.Size = 10
    r.Font.Color = RGB(30, 41, 59)
    r.WrapText = True
    r.VerticalAlignment = xlVAlignTop
    r.Interior.Color = RGB(248, 250, 252)

    ws.Rows("2").RowHeight = 1200

    ' Back to dashboard affordance
    Dim backCell As Range
    Set backCell = ws.Range("B1")
    backCell.Value = ChrW(8592) & " Back to dashboard"
    backCell.Font.Size = 9
    backCell.Font.Color = RGB(109, 40, 217)
    backCell.Font.Underline = xlUnderlineStyleSingle
    backCell.Interior.Color = RGB(248, 250, 252)
    ws.Rows("1").RowHeight = 20

    ws.Range("A1").Select
End Sub


Private Sub MakeButton(ByVal ws As Worksheet, ByVal btnRow As Long, _
                        ByVal caption As String, ByVal macroName As String, _
                        ByVal btnColor As Long, ByVal descText As String)

    Const INSET As Double = 4

    Dim cell As Range: Set cell = ws.Cells(btnRow, 2)

    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                cell.Left + INSET, _
                                cell.Top + INSET, _
                                cell.Width - INSET * 2, _
                                cell.Height - INSET * 2)
    With s
        .Name = "btn_" & macroName
        .OnAction = macroName

        .Fill.ForeColor.RGB = btnColor
        .Fill.Solid
        .Line.Visible = msoFalse

        On Error Resume Next
        .Adjustments(1) = 0.18
        On Error GoTo 0

        With .Shadow
            .Visible = msoTrue
            .OffsetX = 0
            .OffsetY = 2
            .Transparency = 0.75
            .Size = 102
            .ForeColor.RGB = RGB(0, 0, 0)
        End With

        With .TextFrame2
            .TextRange.Text = caption
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange.Font
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Size = 11
                .Bold = msoTrue
            End With
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim descRange As Range
    Set descRange = ws.Range(ws.Cells(btnRow + 1, 2), ws.Cells(btnRow + 1, 4))
    descRange.Merge
    descRange.Value = descText
    descRange.Font.Size = 8.5
    descRange.Font.Color = RGB(51, 65, 85)
    descRange.VerticalAlignment = xlVAlignCenter
    descRange.HorizontalAlignment = xlLeft
    descRange.WrapText = True
    descRange.Interior.Color = RGB(241, 245, 249)

    With descRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(226, 232, 240)
        .Weight = xlThin
    End With
    With descRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = btnColor
        .Weight = xlMedium
    End With
End Sub
