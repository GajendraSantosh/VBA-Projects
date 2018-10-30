Sub Worksheets_Comparison()
'
'Comparing Worksheets with another worksheets
'Source URl :- https://www.exceltip.com/files-workbook-and-worksheets-in-vba/determine-if-a-sheet-exists-in-a-workbook-using-vba-in-microsoft-excel.html

'
Dim infobox As Integer
Dim MyOtherWB As Workbook
Dim MyOtherSht As Worksheet

infobox = MsgBox("Are you sure you want to Compare Worksheets?", vbYesNo + vbQuestion, "Worksheets Comparison")

If infobox = vbYes Then

    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .ButtonName = "Open"
        .Title = "Worksheets Comparison"
        .Filters.Clear
        .Filters.Add "Excel Macro-Enabled Workbook", "*.xlsm" 'Add extensions
        
			If .Show = -1 Then
                varFileName = Mid(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") + 1, Len(.SelectedItems(1))) 'File Name of selected Workbook
                Workbooks.Open Filename:=.SelectedItems(1) 'Open Selected Workbook
                Set MyOtherWB = ActiveWorkbook
				
                For Each MyOtherSht In MyOtherWB.Sheets
                    If ThisWorkbook.Sheets(MyOtherSht.Name).Name = MyOtherSht.Name Then 
						'your code if sheet found
					End If
                Next MyOtherSht
				
                Workbooks(varFileName).Close savechanges = False 'Close Selected Workbook
			End If
    End With
End If
End Sub
