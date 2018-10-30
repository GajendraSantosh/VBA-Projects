Sub Copy2Temp()
'
'Copying data to Temp sheet
'

'
Dim RNG As String

If Sheet3.Range("AP2").Value <> Empty Then Sheet3.Range("AP2:AW" & Sheet3.Cells(Rows.Count, "AP").End(xlUp).Row).Clear
With Sheet1

	On Error Resume Next
    If .AutoFilterMode Then .ShowAllData
    If Not .AutoFilterMode Then .Range("A19").AutoFilter

    A = .Cells(Rows.Count, "A").End(xlUp).Row 'Last Row
    .Range("$A$19:$AE$" & A).AutoFilter Field:=25, Criteria1:="<>"  'Removing Blank cells in filters
    RNG = "A20:A" & A & ",Y20:AE" & A
    Range(RNG).SpecialCells(xlCellTypeVisible).Copy Sheet3.Range("AP2")
    ActiveSheet.ShowAllData 'Clear Filters
    .Range("A19").CurrentRegion.Offset(1).Resize(Selection.CurrentRegion.Rows.Count - 1).Clear 'Clearing Data	
End With
End Sub
