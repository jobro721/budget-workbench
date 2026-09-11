Attribute VB_Name = "BuildVarianceReport"
Option Explicit
' BuildVarianceReport.bas — budget-workbench
'
' Budget-vs-actuals variance report (synthetic data only).
' Reads the Budget and Actuals blocks from the "Data" sheet (same key:
' Line + ObjectClass) and writes the "Variance" tab with $ and % variance
' plus threshold flags: |var %| > 10% amber, > 25% red.
'
' Mirror of python/variance.py — math is identical by design.
'
' Usage: press Alt+F8 -> BuildVarianceReport.

Public Const AMBER_PCT As Double = 10#
Public Const RED_PCT As Double = 25#

Public Sub BuildVarianceReport()
    On Error GoTo Cleanup

    Dim wsData As Worksheet, wsVar As Worksheet
    Dim nB As Long, i As Long
    Dim lineA() As String, ocA() As String, bud() As Double
    Dim lineB() As String, ocB() As String, act() As Double
    Dim varAmt As Double, varPct As Double, flag As String
    Dim nFlag As Long, rowOut As Long

    Set wsData = ThisWorkbook.Worksheets("Data")
    Set wsVar = EnsureTab(ThisWorkbook, "Variance")

    nB = ReadBudget(wsData, "BUDGET", lineA, ocA, bud)
    Dim nA As Long
    nA = ReadBudget(wsData, "ACTUALS", lineB, ocB, act)
    If nB = 0 Or nA = 0 Then
        MsgBox "Budget or Actuals block not found on the Data sheet.", vbCritical
        Exit Sub
    End If

    wsVar.Cells.Clear
    wsVar.Range("A1:I1").Value = Array("Line", "ObjectClass", "Budget", "Actual", _
                                       "Variance $", "Variance %", "Flag", "OC Name", "Note")
    FormatHeader wsVar, 9
    rowOut = 2
    For i = 1 To nB
        varAmt = FindAmount2(lineB, ocB, act, lineA(i), ocA(i)) - bud(i)
        varPct = varAmt / bud(i) * 100
        If Abs(varPct) > RED_PCT Then
            flag = "RED"
        ElseIf Abs(varPct) > AMBER_PCT Then
            flag = "AMBER"
        Else
            flag = "OK"
        End If
        If flag <> "OK" Then nFlag = nFlag + 1

        wsVar.Cells(rowOut, 1).Value = lineA(i)
        wsVar.Cells(rowOut, 2).Value = ocA(i)
        wsVar.Cells(rowOut, 3).Value = bud(i)
        wsVar.Cells(rowOut, 4).Value = bud(i) + varAmt
        wsVar.Cells(rowOut, 5).Value = varAmt
        wsVar.Cells(rowOut, 6).Value = varPct
        wsVar.Cells(rowOut, 7).Value = flag
        wsVar.Cells(rowOut, 3).NumberFormat = "#,##0"
        wsVar.Cells(rowOut, 4).NumberFormat = "#,##0"
        wsVar.Cells(rowOut, 5).NumberFormat = "#,##0"
        wsVar.Cells(rowOut, 6).NumberFormat = "0.0"
        Select Case flag
            Case "AMBER": wsVar.Range(wsVar.Cells(rowOut, 1), wsVar.Cells(rowOut, 7)).Interior.Color = RGB(246, 238, 221)
            Case "RED": wsVar.Range(wsVar.Cells(rowOut, 1), wsVar.Cells(rowOut, 7)).Interior.Color = RGB(243, 225, 228)
        End Select
        rowOut = rowOut + 1
    Next i

    wsVar.Columns("A:I").AutoFit
    MsgBox "Variance report complete: " & nFlag & " of " & nB & " rows flagged.", vbInformation
    Exit Sub
Cleanup:
    MsgBox "BuildVarianceReport error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' Reads the block whose header cell (column A) equals the given marker
' ("BUDGET" or "ACTUALS"). Layout: marker row, then data rows
' (Line, LineName, ObjectClass, OCName, Amount).
Private Function ReadBudget(ws As Worksheet, header As String, _
                            ByRef line() As String, ByRef oc() As String, _
                            ByRef amt() As Double) As Long
    Dim r As Long, n As Long
    For r = 1 To ws.Rows.Count
        If LCase(Trim(ws.Cells(r, 1).Value)) = LCase(header) Then
            ' read following data rows
            Do While Trim(ws.Cells(r + 1, 1).Value) <> ""
                n = n + 1
                ReDim Preserve line(n): oc(n): amt(n)
                line(n) = Trim(ws.Cells(r + 1, 1).Value)
                oc(n) = Trim(ws.Cells(r + 1, 3).Value)
                amt(n) = CDbl(ws.Cells(r + 1, 5).Value)
                r = r + 1
            Loop
            ReadBudget = n
            Exit Function
        End If
    Next r
    ReadBudget = 0
End Function

Private Function FindAmount2(lineB() As String, ocB() As String, amtB() As Double, _
                             keyLine As String, keyOc As String) As Double
    Dim i As Long
    For i = 1 To UBound(lineB)
        If lineB(i) = keyLine And ocB(i) = keyOc Then
            FindAmount2 = amtB(i): Exit Function
        End If
    Next i
End Function

Public Sub FormatHeader(ws As Worksheet, ncols As Long)
    Dim c As Long
    For c = 1 To ncols
        With ws.Cells(1, c)
            .Font.Bold = True
            .Font.Color = RGB(30, 75, 56)
            .Borders(xlEdgeBottom).LineStyle = xlDouble
            .Borders(xlEdgeBottom).Color = RGB(166, 138, 91)
        End With
    Next c
    ws.Rows(1).Font.Name = "Georgia"
End Sub
