Sub macroopen()
Const sPath = "G:\global\china forecasting service\Data\Provinces\Demographics\"
sfil = Dir(sPath & "CN*_POP.xls*")
i = 1
Application.DisplayAlerts = False
'Workbooks("Book1.xlsx").Worksheets.Add
Do While sfil <> "" 
Workbooks.Open sPath & sfil, "0"
For n = 3 To 4 'source sheet index
ActiveWorkbook.Worksheets(n).Activate
prov = ActiveSheet.Range("A1")
age = ActiveSheet.Range("B75:B87")
Set target = Workbooks("Book1.xlsx").Worksheets(n - 2).Range("A1:A13")
For j = 1 To 2 ' gender index
m = 13 * (j - 1) + 26 * (i - 1)
ActiveSheet.Range("BS167:CH179").Offset(92 * (j - 1), 0).Copy
Workbooks("Book1.xlsx").Worksheets(n - 2).Range("A1:R13").Offset(m, 3).PasteSpecial Paste:=xlPasteValues
target.Offset(m, 0) = prov
target.Offset(m, 1) = j
target.Offset(m, 2) = age
Next j
Next n
ActiveWorkbook.Close (False)
sfil = Dir
i = i + 1
Loop
'Workbooks("Book1.xlsx").Close Savechanges:=True
End Sub
