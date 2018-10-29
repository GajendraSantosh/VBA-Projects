'Display the Filterd Coloum name in Message box
'Note:- Before running macro sheet must contain filter.

Sub Filterd_Column_Name()
	Dim Sht As Worksheet
	Dim i As Long
   
	Set Sht = ActiveSheet
   
	With Sht.AutoFilter
		For i = 1 To .Filters.Count
			If .Filters(i).On Then
				MsgBox .Range(1, i).Value
			End If
		Next i
	End With
End Sub
