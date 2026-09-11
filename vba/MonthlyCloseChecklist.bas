Attribute VB_Name = "MonthlyCloseChecklist"
Option Explicit
' MonthlyCloseChecklist.bas — budget-workbench
'
' Runs the monthly close checks against THIS workbook and writes a dated
' pass/fail log to the "CloseLog" tab (newest entry first).
'
' Checks (each one a named constant so the checklist is auditable):
'   1. DATA_DATE present on Data!A1 and not older than 35 days
'   2. Exceptions tab has no MATERIAL or MISSING rows open
'   3. Variance tab has no unflagged |var %| > 25% rows (flag column says OK)
'   4. Execution tab: pacing = obligated / BA is between 0 and 1.2 (sanity)
'   5. UOB identity: unliquidated balance = obligated - outlayed within $1
'
' Usage: Alt+F8 -> MonthlyCloseChecklist  (log goes to the CloseLog tab)

Public Const MAX_AGE_DAYS As Long = 35

Public Sub MonthlyCloseChecklist()
    On Error GoTo Cleanup
    Dim wsData As Worksheet, wsLog As Worksheet
    Dim results() As String, n As Long
    Dim nowStr As String

    Set wsData = ThisWorkbook.Worksheets("Data")
    Set wsLog = EnsureTab(ThisWorkbook, "CloseLog")
    nowStr = Format(Now, "yyyy-mm-dd hh:nn")

    ReDim results(1 To 5)
    n = 0

    ' 1. data date
    Dim d As Date
    Dim ok As Boolean
    ok = False
    If IsDate(wsData.Range("A1").Value) Then
        d = CDate(wsData.Range("A1").Value)
        ok = (Date - d) <= MAX_AGE_DAYS
    End If
    n = n + 1
    results(n) = IIf(ok, "PASS", "FAIL") & " | data date on Data!A1 = " & _
                 IIf(IsDate(wsData.Range("A1").Value), CStr(wsData.Range("A1").Value), "<missing>") & _
                 " (max age " & MAX_AGE_DAYS & " days)"

    ' 2. open material exceptions
    ok = True
    If WorksheetExists(ThisWorkbook, "Exceptions") Then
        Dim r As Long, last As Long
        last = ThisWorkbook.Worksheets("Exceptions").Cells(Rows.Count, 8).End(xlUp).Row
        For r = 2 To last
            If ThisWorkbook.Worksheets("Exceptions").Cells(r, 8).Value = "MATERIAL" Or _
               ThisWorkbook.Worksheets("Exceptions").Cells(r, 8).Value = "MISSING" Then
                ok = False
            End If
        Next r
    End If
    n = n + 1
    results(n) = IIf(ok, "PASS", "FAIL") & " | no open MATERIAL/MISSING exceptions"

    ' 3. variance flags consistent
    ok = True
    If WorksheetExists(ThisWorkbook, "Variance") Then
        Dim vws As Worksheet, vr As Long, vLast As Long
        Set vws = ThisWorkbook.Worksheets("Variance")
        vLast = vws.Cells(Rows.Count, 1).End(xlUp).Row
        For vr = 2 To vLast
            If vws.Cells(vr, 6).Value > 25 And vws.Cells(vr, 7).Value = "OK" Then ok = False
            If vws.Cells(vr, 6).Value < -25 And vws.Cells(vr, 7).Value = "OK" Then ok = False
        Next vr
    End If
    n = n + 1
    results(n) = IIf(ok, "PASS", "FAIL") & " | all |var %| > 25 rows flagged"

    ' 4. pacing sanity (Execution tab: obligated / budget authority)
    ok = True
    If WorksheetExists(ThisWorkbook, "Execution") Then
        Dim ews As Worksheet, er As Long, eLast As Long, pace As Double
        Set ews = ThisWorkbook.Worksheets("Execution")
        eLast = ews.Cells(Rows.Count, 1).End(xlUp).Row
        For er = 2 To eLast
            If IsNumeric(ews.Cells(er, 4).Value) And IsNumeric(ews.Cells(er, 3).Value) Then
                If CDbl(ews.Cells(er, 3).Value) > 0 Then
                    pace = CDbl(ews.Cells(er, 4).Value) / CDbl(ews.Cells(er, 3).Value)
                    If pace < 0 Or pace > 1.2 Then ok = False
                End If
            End If
        Next er
    End If
    n = n + 1
    results(n) = IIf(ok, "PASS", "FAIL") & " | execution pacing within 0-1.2"

    ' 5. UOB identity
    ok = True
    If WorksheetExists(ThisWorkbook, "Execution") Then
        Dim uws As Worksheet, ur As Long, uLast As Long, lhs As Double, rhs As Double
        Set uws = ThisWorkbook.Worksheets("Execution")
        uLast = uws.Cells(Rows.Count, 1).End(xlUp).Row
        For ur = 2 To uLast
            If IsNumeric(uws.Cells(ur, 4).Value) And IsNumeric(uws.Cells(ur, 5).Value) _
               And IsNumeric(uws.Cells(ur, 6).Value) Then
                lhs = CDbl(uws.Cells(ur, 6).Value)
                rhs = CDbl(uws.Cells(ur, 4).Value) - CDbl(uws.Cells(ur, 5).Value)
                If Abs(lhs - rhs) > 1 Then ok = False
            End If
        Next ur
    End If
    n = n + 1
    results(n) = IIf(ok, "PASS", "FAIL") & " | UOB = obligated - outlayed within $1"

    ' --- write log (newest first)
    If wsLog.Cells(1, 1).Value <> "Monthly Close Log" Then
        wsLog.Cells.Clear
        wsLog.Range("A1").Value = "Monthly Close Log"
        wsLog.Range("A1").Font.Bold = True
        wsLog.Range("A1").Font.Size = 14
    End If
    wsLog.Rows(2).Insert
    wsLog.Cells(2, 1).Value = nowStr
    For i = 1 To n
        wsLog.Cells(1 + i + 1, 1).Value = results(i)
    Next i
    wsLog.Columns("A").ColumnWidth = 90

    Dim allOk As Boolean
    allOk = InStr(Join(results, " "), "FAIL") = 0
    If allOk Then
        MsgBox "Monthly close: ALL CHECKS PASS (" & nowStr & ").", vbInformation
    Else
        MsgBox "Monthly close: FAILURES FOUND — see CloseLog tab (" & nowStr & ").", vbExclamation
    End If
    Exit Sub
Cleanup:
    MsgBox "MonthlyCloseChecklist error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Function WorksheetExists(wb As Workbook, name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function
