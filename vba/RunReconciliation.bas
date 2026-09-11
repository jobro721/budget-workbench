Attribute VB_Name = "RunReconciliation"
Option Explicit
' RunReconciliation.bas — budget-workbench
'
' Two-way ledger reconciliation (synthetic data only — see python/make_data.py).
' Reads the two ledger blocks from the "Data" sheet (file A and file B),
' outer-joins on (FundControlPoint, ObjectClass), bands the difference against
' the TOLERANCE named constant, and writes the "Exceptions" and "Summary" tabs.
'
' Mirror of python/recon.py — the banding logic is identical by design so a
' reviewer can diff the two implementations line for line.
'
' Usage: press Alt+F8 -> RunReconciliation.

Public Const TOLERANCE As Double = 5000#

Public Sub RunReconciliation()
    On Error GoTo Cleanup

    Dim wsData As Worksheet, wsExc As Worksheet, wsSum As Worksheet
    Dim i As Long, j As Long
    Dim nA As Long, nB As Long
    Dim fpA() As String, ocA() As String, amtA() As Double
    Dim fpB() As String, ocB() As String, amtB() As Double
    Dim outFp As String, outOc As String
    Dim diff As Double, inA As Boolean, inB As Boolean
    Dim nMatch As Long, nTol As Long, nMat As Long, nMiss As Long
    Dim totDiff As Double
    Dim rowOut As Long

    Set wsData = ThisWorkbook.Worksheets("Data")
    Set wsExc = EnsureTab(ThisWorkbook, "Exceptions")
    Set wsSum = EnsureTab(ThisWorkbook, "Summary")

    ' --- load ledger A (block 1) and B (block 2) from Data
    nA = ReadLedger(wsData, "A", fpA, ocA, amtA)
    nB = ReadLedger(wsData, "B", fpB, ocB, amtB)
    If nA = 0 Or nB = 0 Then
        MsgBox "Ledger A or B not found on the Data sheet.", vbCritical
        Exit Sub
    End If

    ' --- outer join + banding
    wsExc.Cells.Clear
    wsExc.Range("A1:H1").Value = Array("FundControlPoint", "ObjectClass", "Description", _
                                       "LedgerA", "LedgerB", "Diff (A-B)", "Status", "Band")
    rowOut = 2
    totDiff = 0
    For i = 1 To nA
        outFp = fpA(i): outOc = ocA(i): inA = True: inB = False
        For j = 1 To nB
            If fpB(j) = outFp And ocB(j) = outOc Then inB = True: Exit For
        Next j
        diff = 0
        If inA And inB Then diff = amtA(i) - FindAmount(fpB, ocB, amtB, outFp, outOc)
        Call WriteException(wsExc, rowOut, outFp, outOc, "", amtA(i), _
                            IIf(inB, FindAmount(fpB, ocB, amtB, outFp, outOc), Empty), _
                            diff, inA, inB, nMatch, nTol, nMat, nMiss, totDiff)
        rowOut = rowOut + 1
    Next i
    For i = 1 To nB
        outFp = fpB(i): outOc = ocB(i): inA = False: inB = True
        Dim matched As Boolean: matched = False
        For j = 1 To nA
            If fpA(j) = outFp And ocA(j) = outOc Then matched = True: Exit For
        Next j
        If Not matched Then
            Call WriteException(wsExc, rowOut, outFp, outOc, "", Empty, amtB(i), _
                                Empty, False, True, nMatch, nTol, nMat, nMiss, totDiff)
            rowOut = rowOut + 1
        End If
    Next i

    ' --- summary
    wsSum.Cells.Clear
    wsSum.Range("A1").Value = "Reconciliation Summary - synthetic ledgers"
    wsSum.Range("A1").Font.Size = 14
    wsSum.Range("A1").Font.Bold = True
    wsSum.Range("A3").Value = "Tolerance band: $" & Format(TOLERANCE, "#,##0")
    wsSum.Range("A5:B9").Value = Array( _
        Array("MATCH", nMatch), Array("WITHIN TOLERANCE", nTol), _
        Array("MATERIAL", nMat), Array("MISSING", nMiss), _
        Array("TOTAL", nMatch + nTol + nMat + nMiss))
    wsSum.Range("C6").Value = totDiff
    wsSum.Range("C6").NumberFormat = "#,##0"

    MsgBox "Reconciliation complete: " & (nTol + nMat + nMiss) & " exceptions of " & _
           (nMatch + nTol + nMat + nMiss) & " rows.", vbInformation
    Exit Sub
Cleanup:
    MsgBox "RunReconciliation error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Function ReadLedger(ws As Worksheet, f As String, _
                            ByRef fp() As String, ByRef oc() As String, _
                            ByRef amt() As Double) As Long
    Dim r As Long, c As Long, n As Long
    ' find the row where column A = f (A header block, then B header block)
    For r = 1 To ws.Rows.Count
        If Trim(ws.Cells(r, 1).Value) = f Then Exit For
    Next r
    If r = ws.Rows.Count Then ReadLedger = 0: Exit Function
    ' header row = r, data starts r+1
    Do While Trim(ws.Cells(r + 1, 2).Value) <> ""
        n = n + 1
        ReDim Preserve fp(n): oc(n): amt(n)
        fp(n) = Trim(ws.Cells(r + 1, 2).Value)
        oc(n) = Trim(ws.Cells(r + 1, 3).Value)
        amt(n) = CDbl(ws.Cells(r + 1, 6).Value)
        r = r + 1
    Loop
    ReadLedger = n
End Function

Private Function FindAmount(fpB() As String, ocB() As String, amtB() As Double, _
                            keyFp As String, keyOc As String) As Double
    Dim i As Long
    For i = 1 To UBound(fpB)
        If fpB(i) = keyFp And ocB(i) = keyOc Then
            FindAmount = amtB(i): Exit Function
        End If
    Next i
End Function

Private Sub WriteException(ws As Worksheet, rowOut As Long, fp As String, oc As String, _
                           desc As String, aVal As Variant, bVal As Variant, _
                           diff As Double, inA As Boolean, inB As Boolean, _
                           ByRef nMatch As Long, ByRef nTol As Long, _
                           ByRef nMat As Long, ByRef nMiss As Long, ByRef totDiff As Double)
    Dim band As String, status As String
    status = IIf(inA And inB, "ok", IIf(inA, "MISSING IN B", "MISSING IN A"))
    If inA And inB Then
        If Abs(diff) < 0.005 Then band = "MATCH"
        ElseIf Abs(diff) < TOLERANCE Then band = "WITHIN TOLERANCE"
        Else band = "MATERIAL"
    Else
        band = "MISSING"
    End If
    Select Case band
        Case "MATCH": nMatch = nMatch + 1
        Case "WITHIN TOLERANCE": nTol = nTol + 1
        Case "MATERIAL": nMat = nMat + 1
        Case "MISSING": nMiss = nMiss + 1
    End Select
    If inA And inB Then totDiff = totDiff + Abs(diff)

    ws.Cells(rowOut, 1).Value = fp
    ws.Cells(rowOut, 2).Value = oc
    ws.Cells(rowOut, 3).Value = desc
    If inA Then ws.Cells(rowOut, 4).Value = aVal
    If inB Then ws.Cells(rowOut, 5).Value = bVal
    If inA And inB Then ws.Cells(rowOut, 6).Value = diff
    ws.Cells(rowOut, 7).Value = status
    ws.Cells(rowOut, 8).Value = band
    ws.Cells(rowOut, 4).NumberFormat = "#,##0"
    ws.Cells(rowOut, 5).NumberFormat = "#,##0"
    ws.Cells(rowOut, 6).NumberFormat = "#,##0"

    Dim fillColor As Long
    Select Case band
        Case "MATCH": fillColor = RGB(227, 238, 231)
        Case "WITHIN TOLERANCE": fillColor = RGB(246, 238, 221)
        Case Else: fillColor = RGB(243, 225, 228)
    End Select
    If band <> "MATCH" Then ws.Range(ws.Cells(rowOut, 1), ws.Cells(rowOut, 8)).Interior.Color = fillColor
End Sub

Public Function EnsureTab(wb As Workbook, name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set EnsureTab = ws
End Function
