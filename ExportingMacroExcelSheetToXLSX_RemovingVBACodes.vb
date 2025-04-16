Sub ExportSheet1ToXLSX_RemoveVBA()
    Dim ws As Worksheet
    Dim newWorkbook As Workbook
    Dim savePath As String

    ' Prompt the user to choose a save location
    savePath = Application.GetSaveAsFilename(FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    
    ' Check if the user canceled the Save As dialog
    If savePath = "False" Then Exit Sub

    ' Create a new workbook
    Set newWorkbook = Application.Workbooks.Add

    ' Copy only Sheet1 from the current workbook to the new workbook
    ThisWorkbook.Sheets(1).Copy Before:=newWorkbook.Sheets(1)

    ' Delete default sheets from the new workbook (e.g., Sheet1, Sheet2, etc.)
    Dim defaultSheet As Worksheet
    For Each defaultSheet In newWorkbook.Sheets
        If defaultSheet.Name <> ThisWorkbook.Sheets(1).Name Then
            Application.DisplayAlerts = False
            defaultSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next defaultSheet

    ' Save the new workbook as an .xlsx file
    Application.DisplayAlerts = False
    newWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' Close the new workbook
    newWorkbook.Close SaveChanges:=False

    ' Notify the user
    MsgBox "Sheet1 exported successfully to " & savePath & ". VBA modules were removed.", vbInformation
End Sub



