Attribute VB_Name = "modWhatIsHidden"
'===============================================================================
' What Is Hidden Here? - CA edition
'
' Run WhatIsHiddenHere on any workbook a client sends you (it audits the ACTIVE
' workbook). It never changes the data: it only adds a "WHAT WAS HIDDEN" report:
'   1  hidden sheets                      8  named cells parked far away
'   2  VERY hidden sheets (not in Unhide) 9  hidden names (Name Manager can't see)
'   3  hidden rows  (with their totals)  10  notes / comments on cells
'   4  hidden columns (with headers)     11  hidden shapes / text boxes
'   5  rows folded into a collapsed group12  tiny fonts (2 pt and below)
'   6  white-on-white text               13  links to other workbooks
'   7  values blanked by ;;; format      14  personal data (PAN / Aadhaar / mobile)
'                                            - including inside hidden areas
' Reveal_All un-hides everything (asks first). Copy this module into
' Personal.xlsb to have it in every workbook: Alt+F11 > File > Import File.
'===============================================================================
Option Explicit

Private Const REPORT As String = "WHAT WAS HIDDEN"
Private mRep As Worksheet
Private mRow As Long
Private mN As Long
Private mRx As Object

Public Sub WhatIsHiddenHere()
    Dim wb As Workbook, t0 As Double
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    t0 = Timer
    Application.ScreenUpdating = False
    WhatIsHidden_Run wb
    mRep.Range("A2").Value = mN & " hidden thing(s) found in " & Format$(Timer - t0, "0.0") & " seconds.   Hidden does not mean absent."
    Application.ScreenUpdating = True
    mRep.Activate
    mRep.Range("A1").Select
End Sub

'Builds the report; returns the number of findings (used by the self-test)
Public Function WhatIsHidden_Run(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Set mRx = CreateObject("VBScript.RegExp")
    mN = 0
    NewReport wb
    For Each ws In wb.Worksheets
        If ws.Name <> REPORT Then
            CheckSheetVisibility ws
            CheckRows ws
            CheckColumns ws
            CheckCells ws
            CheckFarAway ws
            CheckComments ws
            CheckShapes ws
        End If
    Next ws
    CheckNames wb
    CheckLinks wb
    For Each ws In wb.Worksheets
        If ws.Name <> REPORT Then CheckPersonalData ws
    Next ws
    FinishReport
    WhatIsHidden_Run = mN
End Function

'-------------------------------------------------------------------------------
' Report
'-------------------------------------------------------------------------------
Private Sub NewReport(ByVal wb As Workbook)
    Dim prev As Boolean
    prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(REPORT).Delete
    On Error GoTo 0
    Application.DisplayAlerts = prev
    Set mRep = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    mRep.Name = REPORT
    mRep.Range("A1").Value = "WHAT THIS FILE WAS HIDING  -  " & wb.Name
    mRep.Range("A1").Font.Size = 16: mRep.Range("A1").Font.Bold = True: mRep.Range("A1").Font.Color = RGB(192, 0, 0)
    mRep.Range("A2").Font.Italic = True
    mRep.Range("A4:F4").Value = Array("#", "Sheet", "Where", "What was hidden", "Peek (first values)", "How to reveal it")
    mRep.Range("A4:F4").Font.Bold = True
    mRep.Range("A4:F4").Font.Color = RGB(255, 255, 255)
    mRep.Range("A4:F4").Interior.Color = RGB(31, 56, 100)
    mRow = 5
End Sub

Private Sub Add(ByVal sheetName As String, ByVal where As String, ByVal what As String, ByVal peek As String, ByVal how As String, Optional ByVal severe As Boolean = False)
    mN = mN + 1
    mRep.Cells(mRow, 1).Value = mN
    mRep.Cells(mRow, 2).Value = sheetName
    mRep.Cells(mRow, 3).Value = "'" & where
    mRep.Cells(mRow, 4).Value = what
    mRep.Cells(mRow, 5).Value = "'" & Left$(peek, 200)
    mRep.Cells(mRow, 6).Value = how
    If severe Then mRep.Range(mRep.Cells(mRow, 1), mRep.Cells(mRow, 6)).Interior.Color = RGB(255, 228, 228)
    mRow = mRow + 1
End Sub

Private Sub FinishReport()
    With mRep
        .Columns(1).ColumnWidth = 4
        .Columns(2).ColumnWidth = 16
        .Columns(3).ColumnWidth = 16
        .Columns(4).ColumnWidth = 46
        .Columns(5).ColumnWidth = 56
        .Columns(6).ColumnWidth = 52
        .Range("A5:F" & mRow).WrapText = True
        .Range("A5:F" & mRow).VerticalAlignment = xlTop
        .Cells(mRow + 1, 1).Value = "Nobody in this room wrote a line of this macro. Before you paste a client's file into any AI tool, run it."
        .Cells(mRow + 1, 1).Font.Italic = True
        .Cells(mRow + 1, 1).Font.Color = RGB(89, 89, 89)
        .Activate
        ActiveWindow.DisplayGridlines = False
    End With
End Sub

'-------------------------------------------------------------------------------
' Checks
'-------------------------------------------------------------------------------
Private Sub CheckSheetVisibility(ByVal ws As Worksheet)
    Select Case ws.Visible
        Case xlSheetHidden
            Add ws.Name, "whole sheet", "Hidden sheet", SheetPeek(ws), "Right-click any sheet tab > Unhide"
        Case xlSheetVeryHidden
            Add ws.Name, "whole sheet", "VERY hidden sheet - it does not even appear in the Unhide list", SheetPeek(ws), _
                "Alt+F11 > select the sheet > Properties (F4) > Visible = -1 xlSheetVisible", True
    End Select
End Sub

Private Function SheetPeek(ByVal ws As Worksheet) As String
    Dim c As Range, n As Long, s As String, scan As Range
    On Error Resume Next
    Set scan = ws.UsedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If scan Is Nothing Then SheetPeek = "(empty)": Exit Function
    For Each c In scan
        s = s & IIf(s = "", "", " | ") & CellText(c)
        n = n + 1
        If n >= 4 Then Exit For
    Next c
    SheetPeek = s
End Function

'Indian digit grouping: 10435000 -> 1,04,35,000
Private Function Num(ByVal x As Double) As String
    Dim neg As Boolean, ip As String, dp As String, out As String
    neg = (x < 0): x = Abs(x)
    ip = Format$(Int(x), "0")
    If x <> Int(x) Then dp = Mid$(Format$(x - Int(x), "0.00"), 2)
    If Len(ip) > 3 Then
        out = "," & Right$(ip, 3): ip = Left$(ip, Len(ip) - 3)
        Do While Len(ip) > 2
            out = "," & Right$(ip, 2) & out: ip = Left$(ip, Len(ip) - 2)
        Loop
        out = ip & out
    Else
        out = ip
    End If
    Num = IIf(neg, "-", "") & out & dp
End Function

Private Function CellText(ByVal c As Range) As String
    If IsError(c.Value) Then CellText = "#ERR": Exit Function
    If IsNumeric(c.Value) And Not IsEmpty(c.Value) And VarType(c.Value) <> vbString Then
        CellText = Num(c.Value)
    Else
        CellText = Left$(CStr(c.Value), 90)
    End If
End Function

'Hidden rows in blocks; collapsed outline groups are reported separately
Private Sub CheckRows(ByVal ws As Worksheet)
    Dim ur As Range, i As Long, first As Long, lastR As Long, grouped As Boolean, inBlock As Boolean
    Set ur = ws.UsedRange
    lastR = ur.Row + ur.Rows.Count - 1
    For i = ur.Row To lastR + 1
        If i <= lastR And ws.Rows(i).Hidden Then
            If Not inBlock Then first = i: inBlock = True: grouped = (ws.Rows(i).OutlineLevel > 1)
        ElseIf inBlock Then
            ReportRowBlock ws, first, i - 1, grouped
            inBlock = False
        End If
    Next i
End Sub

Private Sub ReportRowBlock(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal grouped As Boolean)
    Dim rng As Range, c As Range, tot As Double, nNum As Long, txt As String, numCol As Long, hdr As String, r As Long
    Set rng = ws.Range(ws.Rows(r1), ws.Rows(r2))
    For Each c In Intersect(rng, ws.UsedRange).Cells
        If Not IsEmpty(c.Value) And Not IsError(c.Value) Then
            If IsNumeric(c.Value) And VarType(c.Value) <> vbString Then
                If numCol = 0 Then numCol = c.Column
                If c.Column = numCol Then tot = tot + c.Value: nNum = nNum + 1
            ElseIf txt = "" Then
                txt = Trim$(Left$(CStr(c.Value), 50))
            End If
        End If
    Next c
    If numCol > 0 Then
        For r = 1 To r1 - 1
            If Not IsNumeric(ws.Cells(r, numCol).Value) And Trim$(CStr(ws.Cells(r, numCol).Value)) <> "" Then hdr = CStr(ws.Cells(r, numCol).Value)
            If hdr <> "" Then Exit For
        Next r
    End If
    If grouped Then
        Add ws.Name, "rows " & r1 & "-" & r2, (r2 - r1 + 1) & " row(s) folded into a collapsed group", _
            IIf(txt <> "", txt & "; ", "") & IIf(hdr = "", "values", hdr) & " in these rows: " & Num(tot), "Click the + in the left margin (or Data > Ungroup)"
    Else
        Add ws.Name, "rows " & r1 & "-" & r2, (r2 - r1 + 1) & " hidden row(s) - and the visible total does not include them", _
            IIf(txt <> "", txt & "; ", "") & IIf(hdr = "", "values", hdr) & " in these rows: " & Num(tot), "Select rows " & (r1 - 1) & " to " & (r2 + 1) & " > right-click > Unhide", True
    End If
End Sub

Private Sub CheckColumns(ByVal ws As Worksheet)
    Dim ur As Range, i As Long, first As Long, lastC As Long, inBlock As Boolean
    Set ur = ws.UsedRange
    lastC = ur.Column + ur.Columns.Count - 1
    For i = ur.Column To lastC + 1
        If i <= lastC And ws.Columns(i).Hidden Then
            If Not inBlock Then first = i: inBlock = True
        ElseIf inBlock Then
            ReportColBlock ws, first, i - 1
            inBlock = False
        End If
    Next i
End Sub

Private Sub ReportColBlock(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long)
    Dim i As Long, r As Long, hdrs As String, h As String, addr As String
    For i = c1 To c2
        h = ""
        For r = 1 To 10
            If Trim$(CStr(ws.Cells(r, i).Value)) <> "" Then h = CStr(ws.Cells(r, i).Value): Exit For
        Next r
        hdrs = hdrs & IIf(hdrs = "", "", ", ") & IIf(h = "", "(no header)", h)
    Next i
    addr = Split(ws.Cells(1, c1).Address(True, False), "$")(0) & ":" & Split(ws.Cells(1, c2).Address(True, False), "$")(0)
    Add ws.Name, "columns " & addr, (c2 - c1 + 1) & " hidden column(s)", hdrs, "Select the columns either side > right-click > Unhide", True
End Sub

'White-on-white text, ;;; formats and tiny fonts
Private Sub CheckCells(ByVal ws As Worksheet)
    Dim scan As Range, c As Range, wAddr As String, wPeek As String, nW As Long, fAddr As String, fPeek As String, nF As Long
    Dim tAddr As String, tPeek As String, nT As Long
    On Error Resume Next
    Set scan = ws.UsedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If scan Is Nothing Then Exit Sub
    For Each c In scan
        If FontMatchesFill(c) Then
            nW = nW + 1
            If nW <= 3 Then wAddr = wAddr & IIf(wAddr = "", "", ", ") & c.Address(False, False): wPeek = wPeek & IIf(wPeek = "", "", " | ") & CellText(c)
        End If
        If Left$(c.NumberFormat, 3) = ";;;" Then
            nF = nF + 1
            If nF <= 3 Then fAddr = fAddr & IIf(fAddr = "", "", ", ") & c.Address(False, False): fPeek = fPeek & IIf(fPeek = "", "", " | ") & CellText(c)
        End If
        If c.Font.Size <= 2 Then
            nT = nT + 1
            If nT <= 3 Then tAddr = tAddr & IIf(tAddr = "", "", ", ") & c.Address(False, False): tPeek = tPeek & IIf(tPeek = "", "", " | ") & CellText(c)
        End If
    Next c
    If nW > 0 Then Add ws.Name, wAddr & IIf(nW > 3, " ...", ""), nW & " cell(s) written in the same colour as the background", wPeek, "Select all (Ctrl+A) > font colour Automatic", True
    If nF > 0 Then Add ws.Name, fAddr & IIf(nF > 3, " ...", ""), nF & " value(s) blanked by the ;;; number format - visible only in the formula bar", fPeek, "Ctrl+1 > Number > General"
    If nT > 0 Then Add ws.Name, tAddr & IIf(nT > 3, " ...", ""), nT & " cell(s) in a microscopic font (2 pt or less)", tPeek, "Set the font size back to 10"
End Sub

'True when the text colour is (almost) the colour of the fill behind it
Private Function FontMatchesFill(ByVal c As Range) As Boolean
    Dim fc As Long, bc As Long
    If IsNull(c.Font.Color) Then Exit Function
    fc = c.Font.Color
    If c.Interior.ColorIndex = xlNone Then bc = RGB(255, 255, 255) Else bc = c.Interior.Color
    FontMatchesFill = (Abs(Lum(fc) - Lum(bc)) < 12)
End Function

Private Function Lum(ByVal clr As Long) As Double
    Lum = ((clr Mod 256) * 299 + ((clr \ 256) Mod 256) * 587 + (clr \ 65536) * 114) / 1000
End Function

'Values parked far away from the main table (e.g. a "real total" in BZ2000)
Private Sub CheckFarAway(ByVal ws As Worksheet)
    Dim scan As Range, c As Range, lastR As Long, lastC As Long, n As Long
    MainBlock ws, lastR, lastC
    On Error Resume Next
    Set scan = ws.UsedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If scan Is Nothing Then Exit Sub
    For Each c In scan
        If (c.Row > lastR + 40 Or c.Column > lastC + 15) And Not (c.EntireRow.Hidden Or c.EntireColumn.Hidden) And NameOf(c) = "" Then
            n = n + 1
            If n <= 3 Then Add ws.Name, c.Address(False, False), "Value parked far outside the visible table (" & (c.Row - lastR) & " rows down, " & (c.Column - lastC) & " columns right)", _
                               CellText(c), "Ctrl+End jumps to the last used cell; Ctrl+G > type the cell or name", True
        End If
    Next c
End Sub

'Extent of the normal table: last row of the first 20 columns, last column of the first 50 rows
Private Sub MainBlock(ByVal ws As Worksheet, ByRef lastR As Long, ByRef lastC As Long)
    Dim i As Long, x As Long
    lastR = 1: lastC = 1
    For i = 1 To 20
        x = ws.Cells(ws.Rows.Count, i).End(xlUp).Row
        If x > lastR Then lastR = x
    Next i
    For i = 1 To 50
        x = ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
        If x > lastC Then lastC = x
    Next i
End Sub

Private Function NameOf(ByVal c As Range) As String
    Dim nm As Name, rg As Range
    For Each nm In c.Worksheet.Parent.Names
        Set rg = Nothing
        On Error Resume Next
        Set rg = nm.RefersToRange
        On Error GoTo 0
        If Not rg Is Nothing Then
            If rg.Address(External:=True) = c.Address(External:=True) Then NameOf = nm.Name: Exit Function
        End If
    Next nm
End Function

Private Sub CheckComments(ByVal ws As Worksheet)
    Dim cm As Comment, n As Long, addr As String, peek As String
    For Each cm In ws.Comments
        n = n + 1
        If n <= 3 Then addr = addr & IIf(addr = "", "", ", ") & cm.Parent.Address(False, False): peek = peek & IIf(peek = "", "", " | ") & Replace(cm.Text, vbLf, " ")
    Next cm
    If n > 0 Then Add ws.Name, addr & IIf(n > 3, " ...", ""), n & " note(s) on cells - travel with the file, easy to miss", peek, "Review > Show All Notes"
End Sub

Private Sub CheckShapes(ByVal ws As Worksheet)
    Dim s As Shape, t As String
    For Each s In ws.Shapes
        If s.Type = msoComment Then GoTo NextShape
        If Not s.Visible Then
            t = ""
            On Error Resume Next
            t = s.TextFrame2.TextRange.Text
            On Error GoTo 0
            Add ws.Name, s.Name, "Hidden shape / text box", IIf(t = "", "(no text)", t), "Home > Find & Select > Selection Pane > click the eye", True
        End If
NextShape:
    Next s
End Sub

Private Sub CheckNames(ByVal wb As Workbook)
    Dim nm As Name, v As String, lastR As Long, lastC As Long, rg As Range
    For Each nm In wb.Names
        If Left$(nm.Name, 6) <> "_xlnm." And InStr(nm.Name, "Print_") = 0 Then
            v = ""
            Set rg = Nothing
            On Error Resume Next
            Set rg = nm.RefersToRange
            v = CStr(Application.Evaluate(nm.RefersTo))
            On Error GoTo 0
            If Not nm.Visible Then
                Add "(workbook)", nm.Name, "Hidden name - not shown in Name Manager", nm.RefersTo & "  =  " & v, "VBA: ThisWorkbook.Names(""" & nm.Name & """).Visible = True", True
            ElseIf Not rg Is Nothing Then
                MainBlock rg.Worksheet, lastR, lastC
                If rg.Row > lastR + 40 Or rg.Column > lastC + 15 Then
                    Add rg.Worksheet.Name, nm.Name, "Named cell pointing far away from the table", nm.RefersTo & "  =  " & CellText(rg.Cells(1, 1)), "Ctrl+G > " & nm.Name, True
                End If
            End If
        End If
    Next nm
End Sub

Private Sub CheckLinks(ByVal wb As Workbook)
    Dim l As Variant, i As Long
    l = wb.LinkSources(xlExcelLinks)
    If IsEmpty(l) Then Exit Sub
    For i = LBound(l) To UBound(l)
        Add "(workbook)", "links", "Formulas pull numbers from another workbook", CStr(l(i)), "Data > Edit Links"
    Next i
End Sub

'PAN / Aadhaar / mobile anywhere - including inside hidden rows, columns and sheets
Private Sub CheckPersonalData(ByVal ws As Worksheet)
    Dim scan As Range, c As Range, s As String, nP As Long, nA As Long, nM As Long, hid As Long, ex As String
    On Error Resume Next
    Set scan = ws.UsedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If scan Is Nothing Then Exit Sub
    For Each c In scan
        s = Replace(CStr(c.Value), " ", "")
        If Rx(s, "^[A-Z]{5}[0-9]{4}[A-Z]$") Then nP = nP + 1: Hit c, hid, ex
        If Rx(s, "^[2-9][0-9]{11}$") Then nA = nA + 1: Hit c, hid, ex
        If Rx(s, "^(\+?91)?[6-9][0-9]{9}$") Then nM = nM + 1: Hit c, hid, ex
    Next c
    If nP + nA + nM > 0 Then
        Add ws.Name, "personal data", "Personal data: " & nP & " PAN, " & nA & " Aadhaar, " & nM & " mobile number(s)" & IIf(hid > 0, " - " & hid & " of them in hidden rows / columns / sheets", ""), _
            ex, "Delete or mask before sharing, and before pasting into any AI tool (DPDP Act)", True
    End If
End Sub

Private Sub Hit(ByVal c As Range, ByRef hid As Long, ByRef ex As String)
    If c.EntireRow.Hidden Or c.EntireColumn.Hidden Or c.Worksheet.Visible <> xlSheetVisible Then hid = hid + 1
    If Len(ex) < 60 Then ex = ex & IIf(ex = "", "", " | ") & Mask(CStr(c.Value))
End Sub

Private Function Mask(ByVal s As String) As String
    s = Replace(s, " ", "")
    If Len(s) > 4 Then Mask = String$(Len(s) - 4, "x") & Right$(s, 4) Else Mask = s
End Function

Private Function Rx(ByVal s As String, ByVal pattern As String) As Boolean
    mRx.Pattern = pattern
    mRx.IgnoreCase = False
    Rx = mRx.Test(s)
End Function

'-------------------------------------------------------------------------------
' Un-hide everything (asks first)
'-------------------------------------------------------------------------------
Public Sub Reveal_All()
    Dim wb As Workbook, ws As Worksheet, s As Shape, nm As Name, c As Range, scan As Range
    Set wb = ActiveWorkbook
    If MsgBox("Un-hide every sheet, row, column, group, shape and name, and make white / ;;; / tiny text visible in " & wb.Name & "?", vbQuestion + vbYesNo, "Reveal all") <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        ws.Visible = xlSheetVisible
        ws.Cells.EntireRow.Hidden = False
        ws.Cells.EntireColumn.Hidden = False
        On Error Resume Next
        ws.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
        Set scan = Nothing
        Set scan = ws.UsedRange.SpecialCells(xlCellTypeConstants)
        On Error GoTo 0
        If Not scan Is Nothing Then
            For Each c In scan
                If FontMatchesFill(c) Then c.Font.Color = RGB(192, 0, 0)
                If Left$(c.NumberFormat, 3) = ";;;" Then c.NumberFormat = "General": c.Font.Color = RGB(192, 0, 0)
                If c.Font.Size <= 2 Then c.Font.Size = 10: c.Font.Color = RGB(192, 0, 0)
            Next c
        End If
        For Each s In ws.Shapes
            s.Visible = msoTrue
        Next s
    Next ws
    For Each nm In wb.Names
        nm.Visible = True
    Next nm
    Application.ScreenUpdating = True
    MsgBox "Everything is visible. Revealed values are in red.", vbInformation, "Reveal all"
End Sub
