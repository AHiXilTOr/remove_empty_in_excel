Аналог через VBA 2007
```bash
Sub ProcessFilesInFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim tableRange As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = msoAutomationSecurityLow

    folderPath = ""

    fileName = Dir(folderPath)

    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)

        For Each ws In wb.Sheets
            On Error Resume Next
            ws.Unprotect
            ws.Cells.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp

            lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

            Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
            tableRange.Borders.LineStyle = xlContinuous
        Next ws

        wb.Close SaveChanges:=True
        fileName = Dir
    Loop

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.AutomationSecurity = msoAutomationSecurityByUI
End Sub
