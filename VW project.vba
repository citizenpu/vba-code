Preparation step: move/copy the entire focus province/city into the workbook to be updated. They should be inserted before the index sheet in the destination workbook.
For province, replace all the CNXZ with CNTB in the CNXZ sheet, and the sheet name should aslo be replaced with CNTB!!!
#step 1.by Tianzeng (Range operation runs much faster)
'opearte by cell
Sub vw3()
    Dim cell As Range, Line As Range
    Dim I As Long
    Dim startyear As Long, endyear As Long
    ws_count = ActiveWorkbook.Worksheets.Count
    ws_start = ActiveWorkbook.ActiveSheet.Index
    'For Each cell In Worksheets("Index").Range("B2", Worksheets("Index").Range("B2").End(xlDown)).Cells
        'Worksheets(cell.Value).Activate
    For c = ws_start To ws_count
        For Each Line In Worksheets(c).Range("A6", Worksheets(c).Range("A6").End(xlDown)).Cells
            iden = Worksheets(c).Name
            startyear = Worksheets(c).Range("E5").Value
            endyear = 2035
            For I = startyear To endyear
                ' Fixed the range reference - assuming you want a named range or cell reference
                Dim sourcecell As String
                sourcecell = Line.Value & iden
                Var = Application.Match(sourcecell, Worksheets(Line.Value).Columns(7), 0)
                Line.Offset(0, 4 + I - startyear).Value = Worksheets(Line.Value).Cells(Var, 7).Offset(0, I - 1999 + 1).Value
            Next I
        Next Line
    Next c
    'Next cell
End Sub


'operate by range instead of cell
Sub vw4()
    Dim cell As Range, Line As Range
    Dim I As Long
    Dim startyear As Long, endyear As Long
    ws_count = ActiveWorkbook.Worksheets.Count
    ws_start = ActiveWorkbook.ActiveSheet.Index
    'For Each cell In Worksheets("Index").Range("B2", Worksheets("Index").Range("B2").End(xlDown)).Cells
        'Worksheets(cell.Value).Activate
    For c = ws_start To ws_count
        For Each Line In Worksheets(c).Range("A6", Worksheets(c).Range("A6").End(xlDown)).Cells
            iden = Worksheets(c).Name
            startyear = Worksheets(c).Range("E5").Value
            endyear = 2035 'sometimes it is 2030
	    sourcestartyear=1952 'it is 1999 for city, 1952 for province
            'For I = startyear To endyear
                ' Fixed the range reference - assuming you want a named range or cell reference
                Dim sourcecell As String
                sourcecell = Line.Value & iden
                Var = Application.Match(sourcecell, Worksheets(Line.Value).Columns(7), 0)
                'Worksheets(c).Range(Line.Offset(0, 4), Worksheets(c).Line.Offset(0, 4 + endyear - startyear)).Value = Worksheets(Line.Value).Range(Cells(Var, 7).Offset(0, startyear - 1999 + 1), Worksheets(Line.Value).Cells(Var, 7).Offset(0, endyear - 1999 + 1)).Value
                Dim sourceRange As Range
                Set sourceRange = Worksheets(Line.Value).Range(Worksheets(Line.Value).Cells(Var, 7).Offset(0, startyear - sourcestartyear + 1), Worksheets(Line.Value).Cells(Var, 7).Offset(0, endyear - sourcestartyear + 1))
                Line.Offset(0, 4).Resize(1, endyear - startyear + 1).Value = sourceRange.Value

            'Next I
        Next Line
    Next c
    'Next cell
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
