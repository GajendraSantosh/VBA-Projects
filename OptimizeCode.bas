Option Explicit
'Optimize Code Runtime
Private CalcState As Long
Private EventState As Boolean
Private PageBreakState As Boolean

Sub OptimizeCode_Begin()
	Application.ScreenUpdating = False

	EventState = Application.EnableEvents
	Application.EnableEvents = False

	CalcState = Application.Calculation
	Application.Calculation = xlCalculationManual

	PageBreakState = ActiveSheet.DisplayPageBreaks
	ActiveSheet.DisplayPageBreaks = False
End Sub

Sub OptimizeCode_End()
	ActiveSheet.DisplayPageBreaks = PageBreakState
	Application.Calculation = CalcState
	Application.EnableEvents = EventState
	Application.ScreenUpdating = True
End Sub

Sub TestCode1()
	Call OptimizeCode_Begin
	
		For I =1 to 10000	
			Sheet1.Cells(I,1).Value = I
		Next I
		
	Call OptimizeCode_End
End Sub
