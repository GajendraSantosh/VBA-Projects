Sub DocsFindAndReplace()
    Dim wdApp As New Word.Application
    
    With wdApp
      'Opening Word Document
        .Documents.Open ("C:\Users\gajendra.santosh\Desktop\Template 1.docx")
        .Visible = True
        .Activate
        
      'Find and Replace All
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            
            .Text = "{Supplier Company Name}"                   '--Find Text
            .Replacement.Text = "Supplier Company Name"         '--Replace Text
            
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll			'--Replace All
        End With
        
      'Closing the open Word Document
        .Quit SaveChanges:=False
        
      'Release the external variables from the memory
        Set wdApp = Nothing
    End With
End Sub

