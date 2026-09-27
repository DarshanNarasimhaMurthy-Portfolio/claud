Attribute VB_Name = "GloveTracker"
'==============================================================
' Glove Tracker macros
'   IssueGloves      - Dashboard "ISSUE GLOVES" button
'   UndoLastIssue    - Dashboard "UNDO LAST ISSUE" button
'   AddColleague     - Dashboard "ADD COLLEAGUE" button
'   SetupButtons     - run once (Alt+F8) to draw the buttons
'==============================================================
Option Explicit

Private Const LOG_FIRST As Long = 5
Private Const LOG_LAST As Long = 1004
Private Const COL_FIRST As Long = 5
Private Const COL_LAST As Long = 204
Private Const STOCK_FIRST As Long = 5      ' Stock!A5:B7 = Small, Medium, Large
Private Const TITLE As String = "Glove Tracker"

Private Function Sh(ByVal nm As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(nm)
End Function

Private Function StockRow(ByVal sz As String) As Long
    Dim r As Long
    For r = STOCK_FIRST To STOCK_FIRST + 2
        If StrComp(Sh("Stock").Cells(r, 1).Value, sz, vbTextCompare) = 0 Then
            StockRow = r
            Exit Function
        End If
    Next r
    StockRow = 0
End Function

Private Function ColleagueRow(ByVal nm As String) As Long
    Dim r As Long, ws As Worksheet
    Set ws = Sh("Colleagues")
    For r = COL_FIRST To COL_LAST
        If StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), nm, vbTextCompare) = 0 Then
            ColleagueRow = r
            Exit Function
        End If
    Next r
    ColleagueRow = 0
End Function

Private Function LastLogRow() As Long
    Dim r As Long, ws As Worksheet
    Set ws = Sh("Issue Log")
    For r = LOG_LAST To LOG_FIRST Step -1
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" Then
            LastLogRow = r
            Exit Function
        End If
    Next r
    LastLogRow = 0
End Function

Private Function PairsHad(ByVal nm As String) As Double
    Dim r As Long, ws As Worksheet, total As Double
    Set ws = Sh("Issue Log")
    For r = LOG_FIRST To LOG_LAST
        If StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), nm, vbTextCompare) = 0 Then
            total = total + Val(ws.Cells(r, 4).Value)
        End If
    Next r
    PairsHad = total
End Function

'--------------------------------------------------------------
Public Sub IssueGloves()
    Dim wsD As Worksheet, wsC As Worksheet, wsL As Worksheet, wsS As Worksheet
    Dim nm As String, sz As String, qtyV As Variant, qty As Long
    Dim cRow As Long, sRow As Long, nr As Long, lastR As Long
    Dim onHand As Double, had As Double

    Set wsD = Sh("Dashboard"): Set wsC = Sh("Colleagues")
    Set wsL = Sh("Issue Log"): Set wsS = Sh("Stock")

    nm = Trim$(CStr(wsD.Range("C13").Value))
    If nm = "" Then
        MsgBox "Pick a name first.", vbExclamation, TITLE
        Exit Sub
    End If
    cRow = ColleagueRow(nm)
    If cRow = 0 Then
        MsgBox nm & " is not on the Colleagues list." & vbCrLf & _
               "Add them with ADD COLLEAGUE first.", vbExclamation, TITLE
        Exit Sub
    End If
    nm = Trim$(CStr(wsC.Cells(cRow, 2).Value))   ' use the spelling from the list

    sz = Trim$(CStr(wsD.Range("C14").Value))
    If sz = "" Then sz = Trim$(CStr(wsC.Cells(cRow, 3).Value))
    If sz = "" Then
        MsgBox "Pick a size (no usual size is set for " & nm & ").", vbExclamation, TITLE
        Exit Sub
    End If
    sRow = StockRow(sz)
    If sRow = 0 Then
        MsgBox "Size must be Small, Medium or Large.", vbExclamation, TITLE
        Exit Sub
    End If
    sz = wsS.Cells(sRow, 1).Value

    qtyV = wsD.Range("C15").Value
    If IsEmpty(qtyV) Or Trim$(CStr(qtyV)) = "" Then qtyV = 1
    If Not IsNumeric(qtyV) Then
        MsgBox "Pairs must be a whole number (1 or more).", vbExclamation, TITLE
        Exit Sub
    End If
    If CDbl(qtyV) < 1 Or CDbl(qtyV) <> Int(CDbl(qtyV)) Then
        MsgBox "Pairs must be a whole number (1 or more).", vbExclamation, TITLE
        Exit Sub
    End If
    qty = CLng(qtyV)

    onHand = Val(wsS.Cells(sRow, 2).Value)
    If qty > onHand Then
        MsgBox "Not enough " & sz & " gloves in stock." & vbCrLf & _
               "On hand: " & onHand & "   Asked for: " & qty & vbCrLf & vbCrLf & _
               "Update the Stock sheet if new gloves have arrived.", vbCritical, TITLE
        Exit Sub
    End If

    had = PairsHad(nm)
    If had + qty >= 3 Then
        If MsgBox("FLAG WARNING" & vbCrLf & vbCrLf & _
                  nm & " has already had " & had & " pair(s)." & vbCrLf & _
                  "This will be pair #" & (had + qty) & " and will be FLAGGED." & vbCrLf & vbCrLf & _
                  "Issue anyway?", vbYesNo + vbExclamation + vbDefaultButton2, TITLE) <> vbYes Then
            Exit Sub
        End If
    End If

    lastR = LastLogRow()
    If lastR = 0 Then nr = LOG_FIRST Else nr = lastR + 1
    If nr > LOG_LAST Then
        MsgBox "The Issue Log is full (" & (LOG_LAST - LOG_FIRST + 1) & " rows).", vbCritical, TITLE
        Exit Sub
    End If

    wsL.Unprotect
    wsS.Unprotect
    wsL.Cells(nr, 1).Value = Date
    wsL.Cells(nr, 2).Value = nm
    wsL.Cells(nr, 3).Value = sz
    wsL.Cells(nr, 4).Value = qty
    wsS.Cells(sRow, 2).Value = onHand - qty
    wsL.Protect
    wsS.Protect

    wsD.Range("C13").MergeArea.ClearContents
    wsD.Range("C14").MergeArea.ClearContents
    wsD.Range("C15").Value = 1

    MsgBox "Issued " & qty & " pair(s) of " & sz & " to " & nm & "." & vbCrLf & _
           "That is pair #" & (had + qty) & " for them." & vbCrLf & _
           sz & " left in stock: " & (onHand - qty), _
           IIf(had + qty >= 3, vbExclamation, vbInformation), TITLE
End Sub

'--------------------------------------------------------------
Public Sub UndoLastIssue()
    Dim wsL As Worksheet, wsS As Worksheet
    Dim r As Long, sRow As Long, qty As Double
    Dim nm As String, sz As String, dt As String

    Set wsL = Sh("Issue Log"): Set wsS = Sh("Stock")
    r = LastLogRow()
    If r = 0 Then
        MsgBox "The Issue Log is empty. Nothing to undo.", vbInformation, TITLE
        Exit Sub
    End If
    nm = CStr(wsL.Cells(r, 2).Value)
    sz = CStr(wsL.Cells(r, 3).Value)
    qty = Val(wsL.Cells(r, 4).Value)
    If IsDate(wsL.Cells(r, 1).Value) Then dt = Format$(wsL.Cells(r, 1).Value, "dd-mmm-yyyy")

    If MsgBox("Undo the last hand-out?" & vbCrLf & vbCrLf & _
              dt & "   " & nm & "   " & sz & "   " & qty & " pair(s)" & vbCrLf & vbCrLf & _
              "The pairs will be put back into stock.", vbYesNo + vbQuestion + vbDefaultButton2, TITLE) <> vbYes Then
        Exit Sub
    End If

    wsL.Unprotect
    wsS.Unprotect
    sRow = StockRow(sz)
    If sRow > 0 Then wsS.Cells(sRow, 2).Value = Val(wsS.Cells(sRow, 2).Value) + qty
    wsL.Range(wsL.Cells(r, 1), wsL.Cells(r, 4)).ClearContents
    wsL.Protect
    wsS.Protect

    MsgBox "Removed. " & qty & " pair(s) of " & sz & " put back into stock.", vbInformation, TITLE
End Sub

'--------------------------------------------------------------
Public Sub AddColleague()
    Dim wsD As Worksheet, wsC As Worksheet
    Dim nm As String, sz As String, r As Long, freeR As Long

    Set wsD = Sh("Dashboard"): Set wsC = Sh("Colleagues")
    nm = Trim$(CStr(wsD.Range("I13").Value))
    sz = Trim$(CStr(wsD.Range("I14").Value))

    If nm = "" Then
        MsgBox "Type the colleague's name first.", vbExclamation, TITLE
        Exit Sub
    End If
    If sz <> "" And StockRow(sz) = 0 Then
        MsgBox "Size must be Small, Medium or Large (or leave it blank).", vbExclamation, TITLE
        Exit Sub
    End If
    If ColleagueRow(nm) > 0 Then
        MsgBox nm & " is already on the Colleagues list.", vbExclamation, TITLE
        Exit Sub
    End If

    For r = COL_FIRST To COL_LAST
        If Trim$(CStr(wsC.Cells(r, 2).Value)) = "" Then
            freeR = r
            Exit For
        End If
    Next r
    If freeR = 0 Then
        MsgBox "The Colleagues list is full (" & (COL_LAST - COL_FIRST + 1) & " rows).", vbCritical, TITLE
        Exit Sub
    End If

    wsC.Unprotect
    wsC.Cells(freeR, 2).Value = nm
    If sz <> "" Then wsC.Cells(freeR, 3).Value = Sh("Stock").Cells(StockRow(sz), 1).Value
    wsC.Protect

    wsD.Range("I13").MergeArea.ClearContents
    wsD.Range("I14").ClearContents
    MsgBox nm & " added to the Colleagues list.", vbInformation, TITLE
End Sub

'--------------------------------------------------------------
' Run once after importing this module (Alt+F8 > SetupButtons > Run)
Public Sub SetupButtons()
    Dim ws As Worksheet
    Set ws = Sh("Dashboard")
    ws.Unprotect
    MakeButton ws, "btnIssue", ws.Range("B17:D18"), "ISSUE GLOVES", "IssueGloves", RGB(192, 0, 0)
    MakeButton ws, "btnUndo", ws.Range("E17:F18"), "UNDO LAST ISSUE", "UndoLastIssue", RGB(89, 89, 89)
    MakeButton ws, "btnAdd", ws.Range("H16:K17"), "ADD COLLEAGUE", "AddColleague", RGB(192, 0, 0)
    ws.Protect
    MsgBox "Buttons are ready on the Dashboard." & vbCrLf & vbCrLf & _
           "Now save as 'Excel Macro-Enabled Workbook (*.xlsm)'.", vbInformation, TITLE
End Sub

Private Sub MakeButton(ws As Worksheet, ByVal shpName As String, rng As Range, _
                       ByVal caption As String, ByVal macroName As String, ByVal colour As Long)
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(shpName).Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddShape(5, rng.Left + 3, rng.Top + 3, rng.Width - 6, rng.Height - 6) ' 5 = rounded rectangle
    s.Name = shpName
    s.Fill.ForeColor.RGB = colour
    s.Line.Visible = False
    With s.TextFrame.Characters
        .Text = caption
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = "Arial"
        .Font.Color = RGB(255, 255, 255)
    End With
    s.TextFrame.HorizontalAlignment = xlHAlignCenter
    s.TextFrame.VerticalAlignment = xlVAlignCenter
    s.Placement = xlMove
    s.OnAction = macroName
End Sub
