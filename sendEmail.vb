Option Compare Database

Private Sub btnSendIt_Click()
    Dim Sent As Boolean
       
    Sent = SendEmailWithOutlook("jackie.adair@us.af.mil", "Test", "Test")
    
End Sub


Public Function SendEmailWithOutlook(MessageTo As String, Subject As String, MessageBody As String)

    DoCmd.SendObject , , , MessageTo, , , Subject, MessageBody, False, False

End Function
