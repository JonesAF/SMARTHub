Option Compare Database
Option Explicit


'Set an unchangeable variable to the amount (10% for example) to increase or
'decrease the font size with each zoom, in or out.
Const FONT_ZOOM_PERCENT_CHANGE = 0.1


'Create the fontZoom and ctrlKeyIsPressed variables outside of
'the sub definitions so they can be shared between subs
Private fontZoom As Double
Private ctrlKeyIsPressed As Boolean


'Create an enum so we can use it later when pulling the data out of the "Tag" property
Private Enum ControlTag
    FromLeft = 0
    FromTop
    ControlWidth
    ControlHeight
    OriginalFontSize
    OriginalControlHeight
End Enum


Private Sub Form_Load()
    'Set the font zoom setting to the default of 100% (represented by a 1 below).
    'This means that the fonts will appear initially at the proportional size
    'set during design time. But they can be made smaller or larger at run time
    'by holding the "Shift" key and hitting the "+" or "-" key at the same time,
    'or by holding the "Ctrl" key and scrolling the mouse wheel up or down.
    fontZoom = 1

    'When the form loads, we need to find the relative position of each control
    'and save it in the control's "Tag" property so the resize event can use it
    SaveControlPositionsToTags Me
End Sub


Private Sub Form_Resize()
    'Set the height of the header and footer before calling RepositionControls
    'since it caused problems changing their heights from inside that sub.
    'The Tag property for the header and footer is set inside the SaveControlPositionsToTags sub
    Me.Section(acHeader).Height = Me.WindowHeight * CDbl(Me.Section(acHeader).Tag)
    Me.Section(acFooter).Height = Me.WindowHeight * CDbl(Me.Section(acFooter).Tag)

    'Call the RepositionControls Sub and pass this form as a parameter
    'and the fontZoom setting which was initially set when the form loaded and then
    'changed if the user holds the "Shift" key and hits the "+" or "-" key
    'or holds the "Ctrl" key and scrolls the mouse wheel up or down.
    RepositionControls Me, fontZoom
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'PURPOSE: Make the text on the form bigger if "Shift" and "+" are pressed
    'at the same time and smaller if "Shift" and "-" are pressed at the same time.
    'NOTE: Using the "Ctrl" key instead of the "Shift" key conflicts with Access's
    'default behavior of using "Ctrl -" to delete a record, so "Shift" is used instead

    'Was the "Shift" key being held down while the Key was pressed?
    Dim shiftKeyPressed As Boolean
    shiftKeyPressed = (Shift And acShiftMask) > 0

    'If so, check to see if the user pressed the "+" or the "-" button at the
    'same time as the "Shift" key. If so, then make the font bigger/smaller
    'by the percentage specificed in the FONT_ZOOM_PERCENT_CHANGE variable.
    If shiftKeyPressed Then

        Select Case KeyCode
            Case vbKeyAdd
                fontZoom = fontZoom + FONT_ZOOM_PERCENT_CHANGE
                RepositionControls Me, fontZoom

                'Set the KeyCode back to zero to prevent the "+" symbol from
                'showing up if a textbox or similar control has the focus
                KeyCode = 0

            Case vbKeySubtract
                fontZoom = fontZoom - FONT_ZOOM_PERCENT_CHANGE
                RepositionControls Me, fontZoom

                'Set the KeyCode back to zero to prevent the "-" symbol from
                'showing up if a textbox or similar control has the focus
                KeyCode = 0

        End Select

    End If

    'Detect if the "Ctrl" key was pressed. This variable
    'will be used later when we detect a mouse wheel scroll event.
    If (Shift And acCtrlMask) > 0 Then
        ctrlKeyIsPressed = True
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Change the ctrlKeyIsPressed variable to false when
    'any key is let up. This will make sure the form text does
    'not continue to grow/shrink when the mouse wheel is
    'scrolled after the ctrl key is pressed and let up.
    ctrlKeyIsPressed = False
End Sub


Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    'If the "Ctrl" key is also being pressed, then zoom the form in or out
    If ctrlKeyIsPressed Then
        Debug.Print ctrlKeyIsPressed
        'The user scrolled up, so make the text larger
        If Count < 0 Then

            'Make the font bigger by the percentage specificed
            'in the FONT_ZOOM_PERCENT_CHANGE variable
            fontZoom = fontZoom + FONT_ZOOM_PERCENT_CHANGE
            RepositionControls Me, fontZoom

        'The user scrolled down, so make the text smaller
        ElseIf Count > 0 Then

            'Make the font smaller by the percentage specificed
            'in the FONT_ZOOM_PERCENT_CHANGE variable
            fontZoom = fontZoom - FONT_ZOOM_PERCENT_CHANGE
            RepositionControls Me, fontZoom
        End If

    End If
End Sub


Public Sub SaveControlPositionsToTags(frm As Form)
On Error Resume Next

    Dim ctl As Control

    Dim ctlLeft As String
    Dim ctlTop As String
    Dim ctlWidth As String
    Dim ctlHeight As String
    Dim ctlOriginalFontSize As String
    Dim ctlOriginalControlHeight As String

    For Each ctl In frm.Controls

        'Find the relative position of this control in design view
        'e.g.- This control is 5% from the left, 10% from the top, etc.
        'Those percentages can then be saved in the Tag property for this control
        'and used later in the form's resize event
        ctlLeft = CStr(Round(ctl.Left / frm.Width, 4))
        ctlTop = CStr(Round(ctl.Top / frm.Section(ctl.Section).Height, 4))
        ctlWidth = CStr(Round(ctl.Width / frm.Width, 4))
        ctlHeight = CStr(Round(ctl.Height / frm.Section(ctl.Section).Height, 4))

        'If this control has a FontSize property, then capture the
        'control's original font size and the control's original height from design-time
        'These will be used later to calculate what the font size should be when the form is resized
        Select Case ctl.ControlType
            Case acLabel, acCommandButton, acTextBox, acComboBox, acListBox, acTabCtl, acToggleButton
                ctlOriginalFontSize = ctl.FontSize
                ctlOriginalControlHeight = ctl.Height
        End Select

        'Add all this data to the Tag property of the current control, separated by colons
        ctl.Tag = ctlLeft & ":" & ctlTop & ":" & ctlWidth & ":" & ctlHeight & ":" & ctlOriginalFontSize & ":" & ctlOriginalControlHeight

    Next

    'Set the Tag properties for the header and the footer to their proportional height
    'in relation to the height of the whole form (header + detail + footer)
    frm.Section(acHeader).Tag = CStr(Round(frm.Section(acHeader).Height / (frm.Section(acHeader).Height + frm.Section(acDetail).Height + frm.Section(acFooter).Height), 4))
    frm.Section(acFooter).Tag = CStr(Round(frm.Section(acFooter).Height / (frm.Section(acHeader).Height + frm.Section(acDetail).Height + frm.Section(acFooter).Height), 4))

End Sub


Public Sub RepositionControls(frm As Form, fontZoom As Double)
On Error Resume Next

    Dim formDetailHeight As Long
    Dim tagArray() As String

    'Since "Form.Section(acDetail).Height" usually returns the same value (unless the detail section is tiny)
    'go ahead and calculate the detail section height ourselves and store it in a variable
    formDetailHeight = frm.WindowHeight - frm.Section(acHeader).Height - frm.Section(acFooter).Height

    Dim ctl As Control

    'Loop through all the controls on the form
    For Each ctl In frm.Controls

        'An extra check to make sure the Tag property has a value
        If ctl.Tag <> "" Then

            'Split the Tag property into an array
            tagArray = Split(ctl.Tag, ":")

            If ctl.Section = acDetail Then
                'This is the Detail section of the form so use our "formDetailHeight" variable from above
                ctl.Move frm.WindowWidth * (CDbl(tagArray(ControlTag.FromLeft))), _
                                   formDetailHeight * (CDbl(tagArray(ControlTag.FromTop))), _
                                   frm.WindowWidth * (CDbl(tagArray(ControlTag.ControlWidth))), _
                                   formDetailHeight * (CDbl(tagArray(ControlTag.ControlHeight)))
            Else
                ctl.Move frm.WindowWidth * (CDbl(tagArray(ControlTag.FromLeft))), _
                                   frm.Section(ctl.Section).Height * (CDbl(tagArray(ControlTag.FromTop))), _
                                   frm.WindowWidth * (CDbl(tagArray(ControlTag.ControlWidth))), _
                                   frm.Section(ctl.Section).Height * (CDbl(tagArray(ControlTag.ControlHeight)))
            End If

            'Now we need to change the font sizes on the controls.
            'If this control has a FontSize property, then find the ratio of
            'the current height of the control to the form-load height of the control.
            'So if form-load height was 1000 (twips) and the current height is 500 (twips)
            'then we multiply the original font size * (500/1000), or 50%.
            'Then we multiply that by the fontZoom setting in case the user wants to
            'increase or decrease the font sizes while viewing the form.
            Select Case ctl.ControlType
                Case acLabel, acCommandButton, acTextBox, acComboBox, acListBox, acTabCtl, acToggleButton
                    ctl.FontSize = Round(CDbl(tagArray(ControlTag.OriginalFontSize)) * CDbl(ctl.Height / tagArray(ControlTag.OriginalControlHeight))) * fontZoom
            End Select

        End If

    Next

End Sub

