Function isPrime(Number)
'
' This function will find the given number is Prime or not.
' It will return the Boolean value.
'
	div = 0
	for i=1 to Number
		if Number mod i = 0 Then
			div = div + 1
		End If
	Next

	If div = 2 Then
		isPrime = True
	Else
		isPrime = False
	End if
End Function

