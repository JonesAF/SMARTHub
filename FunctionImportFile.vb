Option Compare Database
Option Explicit

Public Function ImportFile(tblName As String) As Boolean

    ' return is a boolean that indicates whether the import was successful or cancelled
    ' parameter is the temp table name that is passed in from the calling function
    ' Get the file path and name
    ' https://stackoverflow.com/questions/4813598/using-the-browse-for-file-dialog-in-access-vba
    Dim path As Office.FileDialog
    Set path = Application.FileDialog(msoFileDialogOpen)
    path.AllowMultiSelect = False
    path.Filters.Clear
    path.Filters.Add "Excel Files", "*.xlsx, *.xls"
    
    If path.Show Then
        
        ' Turn off warnings and import to temp table from file path selected above
        ' https://docs.microsoft.com/en-us/office/vba/api/access.docmd.transferspreadsheet
        With DoCmd
            .SetWarnings False
            .TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, tblName, path.SelectedItems(1), True
        End With
        
        ImportFile = True
        Exit Function
    Else
        MsgBox "Import Cancelled"
        ImportFile = False
        Exit Function
    End If
    
    
End Function
