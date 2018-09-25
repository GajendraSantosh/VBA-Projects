'Change the picture using file browser in the picturebox.

Sub cmdChangeImage()
Dim infobox As Integer
infobox = MsgBox("Are you sure you want to Change Picture?", vbYesNo + vbQuestion, "Last Updated Date")

If infobox = vbYes Then
Application.ScreenUpdating = False

  With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .ButtonName = "Submit"
    .Title = "Select an image file"
    .Filters.Clear
    .Filters.Add "JPG", "*.JPG"
    .Filters.Add "JPEG File Interchange Format", "*.JPEG"
    .Filters.Add "Graphics Interchange Format", "*.GIF"
    .Filters.Add "Portable Network Graphics", "*.PNG"
    .Filters.Add "Tag Image File Format", "*.TIFF"
    .Filters.Add "All Pictures", "*.*"

    If .Show = -1 Then
        Sheet1.PictureBox1.Picture = LoadPicture(.SelectedItems(1))
        Sheet1.PictureBox1.PictureSizeMode = fmPictureSizeModeStretch
    End If
  End With
  
 Application.ScreenUpdating = True
End If
End Sub
