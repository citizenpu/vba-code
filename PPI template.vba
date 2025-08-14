Sub TransformData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Change "Sheet1" to your actual sheet name
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim outputWs As Worksheet
    lastcolumn = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Set outputWs = ThisWorkbook.Sheets.Add(After:=ws)
    Dim k As Long
    For k = 2 To lastcolumn
        Dim parts As Variant
        parts = Split(ws.Cells(2, k).Value, ":")
        outputWs.Name = parts(2)
        ' Create month headers
        For i = 1 To 12
            outputWs.Cells(1, i + 1).Value = MonthName(i)
        Next i
        outputWs.Cells(1, 1).Value = "Year"
        
        Dim currentRow As Long: currentRow = 2
        Dim yearDict As Object
        Set yearDict = CreateObject("Scripting.Dictionary")
        
        Dim j As Long
        For j = 7 To lastRow ' Assuming data starts from row 7
            Dim dt As Date
            dt = ws.Cells(j, 1).Value
            
            Dim yr As Integer: yr = Year(dt)
            Dim mn As Integer: mn = Month(dt)
            Dim val As Variant: val = ws.Cells(j, k).Value
            
            If Not yearDict.exists(yr) Then
                yearDict.Add yr, currentRow
                outputWs.Cells(currentRow, 1).Value = yr
                currentRow = currentRow + 1
            End If
            
            Dim targetRow As Long: targetRow = yearDict(yr)
            outputWs.Cells(targetRow, mn + 1).Value = val
        Next j
        Set outputWs = ThisWorkbook.Sheets.Add(After:=outputWs)
    Next k
End Sub
