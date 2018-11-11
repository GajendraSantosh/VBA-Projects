Sub MergeRows()
'
'It will merge all rows data into one cell.
'

'
    
Lr = Sheet2.Range("A" & Rows.Count).End(xlUp).Row
Lr2 = Sheet1.Range("A" & Rows.Count).End(xlUp).Row

frmProgressBar.Show vbModeless
frmProgressBar.LabelProgress.Width = 0
frmProgressBar.lbtime.Caption = "The Process may take aprox 1 Hour"
frmProgressBar.lbStatus.Caption = "Processing...."

For i = 2 To Lr
    Sheet1.Range("$A$1:$C$" & Lr2).AutoFilter Field:=1, Criteria1:=Sheet2.Cells(i, "A").Value
    FR = Sheet1.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 1).Row
'    lp = Val(Sheet1.Range("L1").Value)
    lp = Sheet1.Range("A" & Rows.Count).End(xlUp).Row
    
    c = Empty
    d = Empty
    For J = FR To lp
        c = Sheet1.Cells(J, "C").Value
        d = d & c & Chr(10)
    Next J
    Sheet2.Cells(i, "B").Value = d
    Sheet2.Cells(i, "B").WrapText = False
    
        frmProgressBar.Caption = Format(PercentComplete, "0%") & "  Complete"
        PercentComplete = i / Lr
        frmProgressBar.LabelProgress.Width = PercentComplete * 336
        DoEvents
Next i
MsgBox "Done"
End Sub
