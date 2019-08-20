Function lastRow(ByRef Col As String)
'---Finding Last Row in ActiveSheet
    lastRow = ActiveSheet.Cells(Rows.Count, Col).End(xlUp).Row
End Function

Function LastCol(ByRef r As Integer)
'---Finding Last Column in ActiveSheet
    LastCol = Split(Cells(, ActiveSheet.Cells(r, Columns.Count).End(xlToLeft).Column).Address, "$")(1)
End Function

Sub Sorting(WS As Worksheet, _
        SortRange As String, _
        S1ColRange As String, _
        S1Order As Byte, _
        Optional ByVal S2ColRange As String, _
        Optional ByVal S2Order As Byte, _
        Optional ByVal S3ColRange As String, _
        Optional ByVal S3Order As Byte)
        
'---Custom Sorting data upto Last Column in ActiveSheet
    'xlAscending   =  1
    'xlDescending  =  2

    With WS
        .Activate
        .Sort.SortFields.Clear
        .Range(SortRange).Activate
        If S1ColRange <> Empty Then _
            .Sort.SortFields.Add2 Key:=Range(S1ColRange), SortOn:=xlSortOnValues, Order:=S1Order, DataOption:=xlSortNormal
        If S2ColRange <> Empty Then _
            .Sort.SortFields.Add2 Key:=Range(S2ColRange), SortOn:=xlSortOnValues, Order:=S2Order, DataOption:=xlSortNormal
        If S3ColRange <> Empty Then _
            .Sort.SortFields.Add2 Key:=Range(S3ColRange), SortOn:=xlSortOnValues, Order:=S3Order, DataOption:=xlSortNormal
        
        With WS.Sort
            .SetRange WS.Range(SortRange)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub

Sub test()
    Call Sorting(ActiveSheet, _
                "A4" & ":" & LastCol(4) & lastRow("A"), _
                "D5:D" & lastRow("A"), 1, _
                "C5:C" & lastRow("A"), 1)
End Sub
