Function fDODID() As String
' Returns the network login name
' Network login will be DODID followed by identifier (A=Active, E=civilian etc)
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
