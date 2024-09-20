'to write cells from different sheets into one sheet
Sub macrosheet()
j = ActiveWorkbook.Worksheets.Count
Application.DisplayAlerts = False
For i = 2 To j
ActiveWorkbook.Worksheets(i).Activate
ActiveSheet.Range("G2:AM2").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
ActiveWorkbook.Worksheets(i).Range("G2:AM2").Copy Destination:=ActiveWorkbook.Worksheets(1).Range("H2:AD2").Offset(i - 2, 0)
Next i
End Sub
