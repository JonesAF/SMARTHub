Option Compare Database
Option Explicit


Private Declare Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function fDODID() As String
' Returns the network login name
Dim lngLen As Long, lngX As Long
Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If (lngX > 0) Then
        fDODID = Left$(strUserName, lngLen - 1)
    Else
        fDODID = vbNullString
    End If
End Function
