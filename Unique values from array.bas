Sub Unique()
'
'Get unique values from array
'

'
    Dim arr As New Collection, a
    Dim aFirstArray() As Variant
    Dim i As Long
    
    aFirstArray() = Array("Banana", "Apple", "Orange", "Tomato", "Apple", _
                    "Lemon", "Lime", "Lime", "Apple")
    
    On Error Resume Next
    For Each a In aFirstArray
       arr.Add a, a
    Next
    
    For i = 1 To arr.Count
       Cells(i, 1) = arr(i)
    Next

End Sub
