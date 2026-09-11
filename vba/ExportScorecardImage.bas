Attribute VB_Name = "ExportScorecardImage"
Option Explicit
' ExportScorecardImage.bas — budget-workbench
'
' Exports the chart on the "Scorecard" sheet (or the active sheet) as a PNG
' to the workbook's folder — for pasting into PPT, email, or the monthly
' close package.
'
' Usage: Alt+F8 -> ExportScorecardImage

Public Sub ExportScorecardImage()
    On Error GoTo Cleanup
    Dim ws As Worksheet
    Dim ch As Chart
    Dim outPath As String
    Dim baseName As String

    baseName = ThisWorkbook.Name
    baseName = Left(baseName, InStrRev(baseName, ".") - 1)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Scorecard")
    On Error GoTo Cleanup
    If ws Is Nothing Then Set ws = ActiveSheet

    If ws.ChartObjects.Count = 0 Then
        MsgBox "No chart found on '" & ws.Name & "'. Add a chart first.", vbExclamation
        Exit Sub
    End If
    ' take the first (only) chart object
    Set ch = ws.ChartObjects(1).Chart

    outPath = ThisWorkbook.Path & "\" & baseName & "_scorecard_" & _
              Format(Now, "yyyymmdd_hhnn") & ".png"
    ch.Export outPath

    MsgBox "Chart exported:" & vbNewLine & outPath, vbInformation
    Exit Sub
Cleanup:
    MsgBox "ExportScorecardImage error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
