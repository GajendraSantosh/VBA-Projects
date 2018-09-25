'Copying the Selected Range and Pasting as Image into Microsoft Outlook as New Mail.
'Note:- Declare range as 'UserMail'

Sub cmdSendMail()
Dim infobox As Byte
	infobox = MsgBox("Do you want to send a mail?", vbYesNo + vbQuestion, "Last Updated Date")

	If infobox = vbYes Then
		Call Create_UserMailImage
		Call NewOutlookMail
	End If
End Sub

Sub NewOutlookMail()
Dim OutlookApp, OutlookMail As Object
Dim DTable As Range
	Set OutlookApp = CreateObject("Outlook.Application")
	Set OutlookMail = OutlookApp.CreateItem(0)

	ActiveWorkbook.RefreshAll

	UserName = "Test_Mail_Logo"
	stFileName = Environ$("UserProfile") & "\My Documents\My Pictures\" & UserName & ".jpg"

	With Sheet2
		With OutlookMail
			.To = ""
			.Subject = "Daily Production"
			.htmlbody = "<img src= '" & stFileName & "'/img>"
			.display
		End With
	End With

	ActiveWorkbook.RefreshAll
End Sub

Sub Create_UserMailImage()
Dim wb As Workbook, ws As Worksheet, rng As Range, ch As Chart
	Set wb = ThisWorkbook
	ActiveWorkbook.RefreshAll

	Set ws = wb.Sheets("Sheet2")
	Set rng = ws.Range("UserMail")
	Set wb = Workbooks.Add
	Set ch = Charts.Add

	ch.Location xlLocationAsObject, "Sheet1"
	Set ch = ActiveChart
	UserName = "Temp_Mail_Logo"

	ActiveChart.Parent.Name = UserName
	ActiveSheet.ChartObjects(UserName).Height = rng.Height
	ActiveSheet.ChartObjects(UserName).Width = rng.Width
	ActiveSheet.ChartObjects(UserName).Border.LineStyle = xlNone
	rng.CopyPicture xlScreen, xlBitmap
	ch.Paste
	ActiveWorkbook.RefreshAll
	ch.Export Environ$("UserProfile") & "\My Documents\My Pictures\" & UserName & ".jpg"
	ActiveSheet.ChartObjects(UserName).Delete
	ActiveWorkbook.Close savechanges:=False
End Sub
