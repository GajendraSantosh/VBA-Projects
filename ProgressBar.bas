'Display the ProgressBar while loop is runing
'Note:- Design the VBA Form according to your requirement will increase the lable Width.
'Make sure UserForm name must be frmProgressBar, Lable name must be LabelProgress, Lable width should be 336.

Sub TestCode1()
	Dim PercentComplete As Single
	
	'Call OptimizeCode_Begin
	
		frmProgressBar.Show vbModeless
		frmProgressBar.LabelProgress.Width = 0
		frmProgressBar.lbtime.Caption = "The Process may take aprox 1:00 Minutes"
		frmProgressBar.lbStatus.Caption = "Loop is running..."
		
		For I =1 to 10000	
			Sheet1.Cells(I,1).Value = I
		
        frmProgressBar.Caption = Format(PercentComplete, "0%") & "  Complete"
        PercentComplete = I / 10000
        frmProgressBar.LabelProgress.Width = PercentComplete * 336
		Next I
		
	'Call OptimizeCode_End
End Sub

