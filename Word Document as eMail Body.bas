Sub eMailFromDoc()
    Dim wdApp As New Word.Application
    Dim editor As Object
    Dim olApp As New Outlook.Application
    Dim olMail As MailItem
    
    'Opening Word Document
      Set wdDoc = wdApp.Documents.Open("C:\Users\gajendra.santosh\Desktop\Mail Body.docx")
    
    'Copying entire data
      wdDoc.Content.Copy
    
    'Pasting in Mail body
      Set olMail = olApp.CreateItem(olMailItem)
      With olMail
        .Display
        .To = Empty
        .CC = Empty
        .BodyFormat = olFormatRichText
        Set editor = .GetInspector.WordEditor
        editor.Content.Paste
      End With
    
    'Closing the open Word Document
      wdDoc.Close
    
    'Release the external variables from the memory
      Set wdApp = Nothing
      Set wdDoc = Nothing
      Set olApp = Nothing
      Set olMail = Nothing
End Sub

