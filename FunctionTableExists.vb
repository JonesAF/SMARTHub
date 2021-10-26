Option Compare Database
Option Explicit

Public Function TblExists(strTableName As String) As Boolean

    If DCount("[Name]", "MSysObjects", "[Name] = '" & strTableName & "'") Then
        TblExists = True
    End If
    
End Function
