Sub CopyRowsByDeviceProtocol()
    Dim ws As Worksheet, newBook As Workbook, newSheet As Worksheet
    Dim lastRow As Long, copyRow As Long, i As Integer
    Dim userBookName As String, userRange As String
    Dim protocols As Variant, rangeParts As Variant
    Dim cell As Range
    
    ' Ask for workbook name
    userBookName = InputBox("Enter the new workbook name:")
    If userBookName = "" Then Exit Sub
    
    ' Ask for protocol ranges
    userRange = InputBox("Enter the Device Protocol ranges (e.g., 1200-1239,1600-1699):")
    If userRange = "" Then Exit Sub
    
    ' Split user input into an array
    protocols = Split(userRange, ",")
    
    ' Create a new workbook and sheet
    Set newBook = Workbooks.Add
    Set newSheet = newBook.Sheets(1)
    newSheet.Name = "Filtered Data"
    
    ' Copy headers from the first sheet in the original workbook
    ThisWorkbook.Sheets(1).Rows(1).Copy Destination:=newSheet.Rows(1)
    copyRow = 2 ' Start copying from row 2
    
    ' Loop through each sheet
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row in the sheet
        lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row ' Assuming 'Device Protocol' is in column D
        
        ' Loop through each row in the sheet
        For Each cell In ws.Range("D2:D" & lastRow) ' Adjust column if needed
            ' Check if the Device Protocol matches any user-defined range
            For i = LBound(protocols) To UBound(protocols)
                rangeParts = Split(Trim(protocols(i)), "-")
                
                If UBound(rangeParts) = 1 Then
                    If IsNumeric(rangeParts(0)) And IsNumeric(rangeParts(1)) Then
                        If cell.Value >= CLng(rangeParts(0)) And cell.Value <= CLng(rangeParts(1)) Then
                            ' Copy the entire row to the new sheet
                            cell.EntireRow.Copy Destination:=newSheet.Rows(copyRow)
                            copyRow = copyRow + 1
                            Exit For ' Move to the next row after copying
                        End If
                    End If
                End If
            Next i
        Next cell
    Next ws
    
    ' Save the new workbook
    newBook.SaveAs ThisWorkbook.Path & "\" & userBookName & ".xlsx"
    MsgBox "Rows copied successfully to " & userBookName & "!", vbInformation
End Sub
