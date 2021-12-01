' creates a location check to check for a filename in the directory. This just sets the parameters as soon as the form opens.
Private Sub Form_Open(Cancel As Integer)
    Dim strFileName As String
    Dim strFilePath As String
    Dim strLockout As String
    strFileName = DLookup("[Filename]", "[tblLinks]", "Description = 'Lockout File'")
    strFilePath = DLookup("[Address]", "[tblLinks]", "Description = 'Lockout File'")
    strLockout = Dir(strFilePath & strFileName)
    ' Set Count Down variable to false
    ' on the initial opening of the form.
    boolCountDown = False
    Me.Text0.Value = strLockout
End Sub
' This function on the same form uses a predetermined timer to check for the file regularly. If file is missing, or renamed, Access will exit after warning user
Private Sub Form_Timer()
On Error GoTo Err_Form_Timer
    Dim strFileName As String
    Dim strFilePath As String
    Dim strLockout As String
    strFileName = DLookup("[Filename]", "[tblLinks]", "Description = 'Lockout File'")
    strFilePath = DLookup("[Address]", "[tblLinks]", "Description = 'Lockout File'")
    strLockout = Dir(strFilePath & strFileName)
    If boolCountDown = False Then
        ' Do nothing unless the check file is missing.
        If strLockout <> strFileName Then
            ' The check file is not found so
            ' set the count down variable to true and
            ' number of minutes until this session
            ' of Access will be shut down.
            boolCountDown = True
            intCountDownMinutes = 2
        End If
    Else
        ' Count down variable is true so warn
        ' the user that the application will be shut down
        ' in X number of minutes.  The number of minutes
        ' will be 1 less than the initial value of the
        ' intCountDownMinutes variable because the form timer
        ' event is set to fire every 60 seconds
        intCountDownMinutes = intCountDownMinutes - 1
        DoCmd.OpenForm "frmAppShutDownWarn"
        If intCountDownMinutes < 1 Then
            ' Shut down Access if the countdown is zero,
            ' saving all work by default.
            Application.Quit acQuitSaveAll
        End If
    End If

Exit_Form_Timer:
    Exit Sub

Err_Form_Timer:
    Resume Next
End Sub
