Sub CreateMotherBook()
    Dim motherBook As Workbook, sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim folderPath As String, fileName As String, sheetName As String
    Dim fileDialog As FileDialog
    
    ' Ask the user to select a folder
    Set fileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If fileDialog.Show = -1 Then
        folderPath = fileDialog.SelectedItems(1) & "\"
    Else
        Exit Sub
    End If
    
    ' Create a new workbook for the Mother Book
    Set motherBook = Workbooks.Add
    
    ' Loop through all Excel files in the selected folder
    fileName = Dir(folderPath & "*.xlsx") ' Change to "*.xlsm" if working with macros
    Do While fileName <> ""
        ' Open each source workbook
        Set sourceBook = Workbooks.Open(folderPath & fileName)
        Set sourceSheet = sourceBook.Sheets(1) ' Copy only the first sheet
        
        ' Copy the sheet to the Mother Book
        sourceSheet.Copy After:=motherBook.Sheets(motherBook.Sheets.Count)
        
        ' Rename the copied sheet using the file name (without extension)
        sheetName = Left(fileName, InStrRev(fileName, ".") - 1)
        On Error Resume Next ' Prevents errors if the sheet name is too long or invalid
        motherBook.Sheets(motherBook.Sheets.Count).Name = sheetName
        On Error GoTo 0
        
        ' Close the source workbook without saving changes
        sourceBook.Close False
        
        ' Get the next file
        fileName = Dir
    Loop
    
    ' Delete default empty sheets if needed
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In motherBook.Sheets
        If ws.UsedRange.Cells.Count = 1 And ws.Cells(1, 1).Value = "" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' Save the Mother Book
    motherBook.SaveAs folderPath & "MotherBook.xlsx"
    MsgBox "Mother Book created successfully!", vbInformation
End Sub
