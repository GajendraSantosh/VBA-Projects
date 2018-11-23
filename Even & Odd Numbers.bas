Function isOdd(Number)
'
'This functin will find the given number is Odd or not.
'It will return the Boolean value
'
    If Number Mod 2 <> 2 Then
        isOdd = True
    Else
        isOdd = False
    End If
End Function

Function isEven(Number)
'
'This functin will find the given number is Even or not.
'It will return the Boolean value
'
    If Number Mod 2 = 2 Then
        isEven = True
    Else
        isEven = False
    End If
End Function
