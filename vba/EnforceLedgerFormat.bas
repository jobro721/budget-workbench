Attribute VB_Name = "EnforceLedgerFormat"
Option Explicit
' EnforceLedgerFormat.bas — budget-workbench
'
' Applies the house format to any sheet in this workbook:
'   - row 1 header: Georgia bold, ledger-green, double gold bottom rule
'   - tabular numerals + #,##0 on currency columns
'   - column widths auto-fit
'   - freeze top row
' Run on the active sheet, or pass a sheet name.
'
' Usage: Alt+F8 -> EnforceLedgerFormat  (works on the active sheet)

Public Sub EnforceLedgerFormat(Optional sheetName As String = "")
    On Error GoTo Cleanup
    Dim ws As Worksheet, lastCol As Long, lastRow As Long, c As Long
    Dim isCurrency As Boolean

    If sheetName = "" Then
        Set ws = ActiveSheet
    Else
        Set ws = ThisWorkbook.Worksheets(sheetName)
    End If

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastCol < 1 Or lastRow < 1 Then Exit Sub

    ' header
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Name = "Georgia"
        .Font.Bold = True
        .Font.Color = RGB(30, 75, 56)
        .Borders(xlEdgeBottom).LineStyle = xlDouble
        .Borders(xlEdgeBottom).Color = RGB(166, 138, 91)
        .RowHeight = 20
    End With

    ' body
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.ThemeColor = -1   ' window text — keep the default ink
    End With

    ' currency columns: anything whose header contains $, amt, amount, balance, budget, actual
    For c = 1 To lastCol
        Dim h As String
        h = LCase(Trim(ws.Cells(1, c).Value))
        isCurrency = InStr(h, "$") > 0 Or InStr(h, "amt") > 0 Or InStr(h, "amount") > 0 _
                    Or InStr(h, "balance") > 0 Or InStr(h, "budget") > 0 Or InStr(h, "actual") > 0
        If isCurrency Then
            ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).NumberFormat = "#,##0"
            ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).HorizontalAlignment = xlRight
        End If
    Next c

    ws.Columns.AutoFit
    ws.Rows(1).RowHeight = 20
    ws.ActiveCell = ws.Range("A2")
    ws.FreezePanes = False
    ws.Range("A2").Select
    If ws.WindowFrame Then ws.FreezePanes = True

    MsgBox "Format enforced on '" & ws.Name & "' (" & lastRow - 1 & " rows).", vbInformation
    Exit Sub
Cleanup:
    MsgBox "EnforceLedgerFormat error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
