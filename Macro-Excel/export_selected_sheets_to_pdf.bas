Sub SaveSelectedSheetsAsPDF()
    Dim pdfPath As Variant
    Dim ws As Worksheet
    Dim defaultName As String
    
    ' Use the first selected sheet's name for the default filename
    defaultName = ActiveSheet.Name & ".pdf"
    
    ' Ask user with a Save As dialog
    pdfPath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Selected Sheet(s) as PDF")
    
    ' If user cancels, exit
    If pdfPath = False Then Exit Sub
    
    ' Export the selected sheet(s) as PDF
    ActiveSheet.Parent.Sheets(ActiveWindow.SelectedSheets(1).Name).Parent.Sheets(ActiveWindow.SelectedSheets(1).Name).ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    MsgBox "PDF saved as: " & pdfPath, vbInformation, "Done"
End Sub
