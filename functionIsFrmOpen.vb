Option Compare Database


' GLOBAL FUNCTIONS

'---------------------------------------------------------------------------------------
' Procedure : IsFrmOpen
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Determine whether a form is open or not
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFrmName  : Name of the form to check if it is open or not
'
' Usage Example:
' ~~~~~~~~~~~~~~~~
' IsFrmOpen("Form1")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2010-05-26              Initial Release
' 2         2018-02-10              Minor code simplification (per Søren M. Petersen)
'                                   Updated Error Handling
'                                   Updated Copyright
'---------------------------------------------------------------------------------------
Function IsFrmOpen(sFrmName As String) As Boolean
    On Error GoTo Error_Handler
 
    IsFrmOpen = Application.CurrentProject.AllForms(sFrmName).IsLoaded
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: IsFrmOpen" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function
