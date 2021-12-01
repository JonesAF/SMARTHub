' This code creates an email using pre-filled admin email address to request access, pulling the members network user name and prefilling it into the email
' This eliminates the chance members 'fat finger' their username, ensuring ease of access. Multiple email addresses can be used by using a semicolon (;) as a delineator
    Dim olLook As Object 'Start MS Outlook
    Dim olNewEmail As Object 'New email in Outlook
    Dim strContactEmail As String 'Contact email address

    Set olLook = CreateObject("Outlook.Application")
    Set olNewEmail = olLook.createitem(0)
    strEmailSubject = "Mobility Database Access Request"
    strEmailText = "PLEASE NOTE: Request may take up to 24 hours. You will receive a confirmation email once your request has been processed." & vbCrLf & vbCrLf & "Username:" & vbCrLf & userRole.fDODID & vbCrLf & "Requested Role: (User)(PROJO)" & vbCrLf & vbCrLf & "Please provide brief justification below, i.e. 'Required to assign members to TDY'; 'Required to perform PROJO duties for Red Flag Rescue', etc."
    strContactEmail = DLookup("[Address]", "[tblLinks]", "Description = 'Admin Email'") 'this allows you to change the default email from a form input at any time

    With olNewEmail 'Attach template
    .To = strContactEmail
    '.CC = strCc
    .body = strEmailText
    .subject = strEmailSubject
    '.attachments.Add ("S:Email templateInformation Templatev3.xls")
    .display
    DoCmd.Quit acQuitSaveAll
    End With
End Sub
