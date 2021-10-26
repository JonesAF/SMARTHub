Private Sub Form_Timer()
    Me!lblClock.Caption = Format(Now, "dddd, d mmm yyyy, hhmm") + "  [JD:" + Format(Now, "y") + "]"
End Sub
