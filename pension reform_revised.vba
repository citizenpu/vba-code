Sub pension()
    Const sPath = "G:\global\china forecasting service\Data\Provinces\Demographics\"
    Dim sfil As String
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim prov As String
    Dim age As Range
    Dim target As Range
    Dim i As Integer, n As Integer, j As Integer, m As Integer

    sfil = Dir(sPath & "CN*_POP.xls*")
    i = 1
    Application.DisplayAlerts = False

    ' Set reference to the target workbook
    Set wbTarget = Workbooks("Book1.xlsx")

    Do While sfil <> ""
        ' Open the source workbook
        Set wbSource = Workbooks.Open(sPath & sfil, False)

        For n = 3 To 4 'source sheet index for urban/rural area
            ' Set reference to the source worksheet
            Set wsSource = wbSource.Worksheets(n)
            prov = wsSource.Range("A1").Value
            Set age = wsSource.Range("B75:B87") 'from 50 to 62 years old

            ' Set reference to the target worksheet
            Set wsTarget = wbTarget.Worksheets(n - 2)
            Set target = wsTarget.Range("A1:A13")

            For j = 1 To 2 ' gender index
                m = 13 * (j - 1) + 26 * (i - 1)
                wsSource.Range("BS167:CH179").Offset(92 * (j - 1), 0).Copy
                wsTarget.Range("A1:R13").Offset(m, 3).PasteSpecial Paste:=xlPasteValues
                target.Offset(m, 0).Value = prov
                target.Offset(m, 1).Value = j
                target.Offset(m, 2).Value = age.Value
            Next j
        Next n

        wbSource.Close False
        sfil = Dir
        i = i + 1
    Loop

    'wbTarget.Close SaveChanges:=True
End Sub
