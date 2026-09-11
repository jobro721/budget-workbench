Attribute VB_Name = "LoadFileD_CSV"
Option Explicit
' LoadFileD_CSV.bas — budget-workbench
'
' Ingests a USAspending File D (account balances) CSV into the "Execution"
' tab of this workbook — the bridge between the public data pipeline
' (federal-budget-data repo) and the Excel toolkit.
'
' Accepted columns (header names, case-insensitive):
'   federal_account_code | object_class_code | fiscal_period
'   obligated_amt | outlay_amt | unliquidated_balance_amt
' (the column names used by python/make_data.py's filed_sample.csv and by
' Ingests a USAspending File D (account balances) CSV into the "FileD"
' detail tab of this workbook — the bridge between the public data pipeline
' (federal-budget-data repo) and the Excel toolkit. The canonical grid on
' the "Execution" tab (checked by MonthlyCloseChecklist) is untouched.
' Usage: Alt+F8 -> LoadFileD_CSV

Public Sub LoadFileD_CSV()
    On Error GoTo Cleanup
    Dim f As Variant, path As String
    Dim ws As Worksheet
    Dim rowIn As Long, rowOut As Long
    Dim header() As String, nCols As Long
    Dim c As Long, mapCode As Long, mapOc As Long, mapP As Long
    Dim mapOb As Long, mapOl As Long, mapUob As Long
    Dim vals As Variant
    Dim nRows As Long

    f = Application.GetOpenFilename("CSV files (*.csv), *.csv", , "Pick the File D CSV")
    If f = False Then Exit Sub
    path = CStr(f)

    ' read the whole file into memory (File D extracts are at most ~500k rows;
    ' for bigger files, use the python pipeline instead)
    Open path For Input As #1
    Dim allText As String, line As String
    Do While Not EOF(1)
        Line Input #1, line
        allText = allText & line & vbCrLf
    Loop
    Close #1
    Dim lines As String()
    lines = Split(allText, vbCrLf)

    ' map header names -> column indexes
    nCols = 0
    Dim hRow As String()
    hRow = Split(lines(0), ",")
    For c = 0 To UBound(hRow)
        nCols = nCols + 1
    Next c
    ReDim header(1 To nCols)
    mapCode = 0: mapOc = 0: mapP = 0: mapOb = 0: mapOl = 0: mapUob = 0
    For c = 1 To nCols
        header(c) = LCase(Trim(hRow(c - 1)))
        Select Case header(c)
            Case "federal_account_code", "account_code", "tas": mapCode = c
            Case "object_class_code", "object_class": mapOc = c
            Case "fiscal_period", "period": mapP = c
            Case "obligated_amt", "gross_obligated_amt", "obligated": mapOb = c
            Case "unliquidated_balance_amt", "unliquidated_adjments_amt", "uob": mapUob = c
        End Select
    Next c
    If mapCode = 0 Or mapP = 0 Or mapOb = 0 Then
        MsgBox "Required columns not found (need federal_account_code, fiscal_period, obligated_amt).", vbCritical
        Exit Sub
    End If

    Set ws = EnsureTab(ThisWorkbook, "FileD")
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("Account", "ObjectClass", "Period", "Obligated", "Outlay", "UOB")
    rowOut = 2
    nRows = 0
    For rowIn = 1 To UBound(lines)
        If Trim(lines(rowIn)) = "" Then GoTo SkipRow
        vals = Split(lines(rowIn), ",")
        If UBound(vals) + 1 < mapCode Then GoTo SkipRow
        ws.Cells(rowOut, 1).Value = Trim(vals(mapCode - 1))
        If mapOc > 0 And UBound(vals) + 1 >= mapOc Then ws.Cells(rowOut, 2).Value = Trim(vals(mapOc - 1))
        ws.Cells(rowOut, 3).Value = CInt(Trim(vals(mapP - 1)))
        ws.Cells(rowOut, 4).Value = CDbl(NzVal(vals(mapOb - 1)))
        If mapOl > 0 And UBound(vals) + 1 >= mapOl Then ws.Cells(rowOut, 5).Value = CDbl(NzVal(vals(mapOl - 1)))
        If mapUob > 0 And UBound(vals) + 1 >= mapUob Then ws.Cells(rowOut, 6).Value = CDbl(NzVal(vals(mapUob - 1)))
        ws.Cells(rowOut, 4).NumberFormat = "#,##0"
        ws.Cells(rowOut, 5).NumberFormat = "#,##0"
        ws.Cells(rowOut, 6).NumberFormat = "#,##0"
        rowOut = rowOut + 1
        nRows = nRows + 1
SkipRow:
    Next rowIn

    ws.Columns("A:F").AutoFit
    EnforceLedgerFormat "Execution"
    MsgBox "Loaded " & nRows & " rows from File D into the Execution tab.", vbInformation
    Exit Sub
Cleanup:
    MsgBox "LoadFileD_CSV error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Function NzVal(v As String) As Double
    If Trim(v) = "" Or LCase(Trim(v)) = "null" Or LCase(Trim(v)) = "none" Then
        NzVal = 0
    Else
        NzVal = CDbl(Trim(v))
    End If
End Function
