Sub RestoreSheets()
'
'It will Restor all Sheets data from previous workbook to new workbook.
'Note:- Both workbooks contain same sheet names, Columns. 

'
Dim infobox As Integer
Dim MyOtherWB As Workbook
Dim MyOtherSht As Worksheet

infobox = MsgBox("Are you sure you want to Restore your previous Status/Allocated file?", vbYesNo + vbQuestion, "Restore File")

If infobox = vbYes Then

    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .ButtonName = "Open"
        .Title = "Select your previous Status/Allocated file"
        .Filters.Clear
        .Filters.Add "Excel Macro-Enabled Workbook", "*.xlsm"
        
            If .Show = -1 Then
                varFileName = Mid(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") + 1, Len(.SelectedItems(1))) 'File Name of selected Workbook
                Workbooks.Open Filename:=.SelectedItems(1) 'Open Selected Workbook
                Set MyOtherWB = ActiveWorkbook
                
                For Each MyOtherSht In MyOtherWB.Sheets
                    With ThisWorkbook.Sheets(MyOtherSht.Name)
                        If .Name = MyOtherSht.Name And _
                            .Name <> "Summary" And MyOtherWB.Sheets("Update Allocation").Range("A2").Value = Sheet1.Range("A2").Value Then
                            
                            'Apply and Clear Filters
                                On Error Resume Next
                                If .AutoFilterMode Then
                                    .ShowAllData
                                Else
                                    .Range("A19").AutoFilter
                                End If
    
                                If MyOtherSht.AutoFilterMode Then
                                    MyOtherSht.ShowAllData
                                Else
                                    MyOtherSht.Range("A19").AutoFilter
                                End If
                                
                            'Clear Existing Data if any
                                If .Range("A20").Value <> Empty Then _
                                .Range("A19").CurrentRegion.ClearContents
                                
                            'Copy Data
                                MyOtherSht.Range("A19").CurrentRegion.Copy
                                .Range("A19").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

                        End If
                    End With
                Next MyOtherSht
                Workbooks(varFileName).Close savechanges = False 'Close Selected Workbook
            End If
    End With
End If
End Sub
