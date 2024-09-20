Sub macrosheet()
    Dim j As Integer
    Dim i As Integer
    Dim ws As Worksheet
    Dim destRange As Range

    j = ActiveWorkbook.Worksheets.Count
    Application.DisplayAlerts = False

    For i = 2 To j
        Set ws = ActiveWorkbook.Worksheets(i)
        With ws.Range("G2:AM2") "use with method to avoid select
            .Copy
            .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Set destRange = ActiveWorkbook.Worksheets(1).Range("H2:AD2").Offset(i - 2, 0)
            .Copy Destination:=destRange
        End With
    Next i

    Application.CutCopyMode = False
    Application.DisplayAlerts = True
End Sub
