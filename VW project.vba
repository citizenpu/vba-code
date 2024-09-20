'step 1.this macro is built upon Jun's original code and enables automatic deletion of sheets that have 'na" while keeping the remaining. ws_start can be replaced by ActiveWorkbook.Worksheets("Index").Index
Sub MacroNA()
Dim ws_count As Integer
Dim I As Integer
ws_count = ActiveWorkbook.Worksheets.Count
ws_start = ActiveWorkbook.ActiveSheet.Index
Application.DisplayAlerts = False
For I = ws_start To ws_count
   ActiveWorkbook.Worksheets(I).Activate
   ActiveWorkbook.Worksheets(I).Calculate 
   If IsError(ActiveWorkbook.Worksheets(I).Cells(6, 5)) Then
            ActiveWorkbook.Worksheets(I).Delete
            I = I - 1
   Else
            Cells.Select
            Selection.Copy
            Cells.Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If        
Next I    
End Sub
'··························································································································································································································
'step2 this macro can be used to delete the CN** sheets before the "index" sheet
Sub Macrodelete()
Dim I As Integer
J = 100 
I = 1
Application.DisplayAlerts = False
While J > 1
   ActiveWorkbook.Worksheets(I).Delete
   J = ActiveWorkbook.Worksheets("Index").Index
Wend
End Sub
.......................................................................................................................
'step3 this macro is used to write the filename of each sheet into a single sheet (the "index" sheet)'s column (D corresponds to cell(.,4))
Sub Macroname()
Dim ws As Worksheet
Dim x As Integer
x = 1
For Each ws In Activeworkbook.Worksheets
     Sheets("Index").Cells(x, 4) = right(ws.Name,4)
     x = x + 1
Next ws
End Sub
.......................................................................................................................
'step 4 this macro is used to copy the unmatched sheet from the workbook in previous years to the current worbook. two workbooks must be in the same folder. when the program ends, it will alert"out of range"
sub Macroconnect()
For Each cel in Workbooks("2023_LabourWageLandUse.xlsx").Worksheets("Index").range("B:B").Cells
    If iserror(Application.match(cel.value,Workbooks("2024_LabourWageLandUse.xlsx").Worksheets("Index").range("B:B").Cells,0)) then
    Workbooks("2023_LabourWageLandUse.xlsx").Worksheets(cel.value).copy After:=Workbooks("2024_LabourWageLandUse.xlsx").Worksheets(Workbooks("2024_LabourWageLandUse.xlsx").Worksheets.count)
    end if
next cel
End sub   
