'Copying the Selected Range and Pasting into Microsoft Outlook as New Mail.
'Note:- Declare range as 'UserMail'

Sub NewOutlookMail()
Dim infobox As Byte
Dim OutlookApp, OutlookMail As Object

	infobox = MsgBox("Do you want to send a mail?", vbYesNo + vbQuestion, "Sending Status")
	If infobox = vbYes And Sheet1.Range("A1").Value <> Empty Then
		'Call OptimizeCode_Begin
			Set OutlookApp = CreateObject("Outlook.Application")
			Set OutlookMail = OutlookApp.CreateItem(0)
			ActiveWorkbook.RefreshAll
			Sheet1.Range("UserMail").Copy
			With OutlookMail
				.to = ""
				.Subject = "Daily Status"
				
				On Error GoTo eh
				.GetInspector().WordEditor.Range.Paste
				.display
			End With
eh:
			Application.CutCopyMode = False
			ActiveWorkbook.RefreshAll
		'Call OptimizeCode_End
	End If
End Sub
