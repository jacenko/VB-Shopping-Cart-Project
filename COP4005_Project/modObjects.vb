Module modObjects
    Public Function SafeDate(strIn As String) As DateTime
        If strIn Is Nothing Then
            Return NULL_DATE
        End If
        Dim dt As Date = NULL_DATE
        If Date.TryParse(strIn, dt) Then
            Return dt
        Else
            'probably should give an indication of a bad date
            Return dt 'should return a null date
        End If
    End Function

    Public Sub ClearScreenControls(ByVal container As Control)
        'this procedure will clear all controls on the form that is passed in 
        'as the argument
        Dim obj As Control  'generic control object
        Dim strControlType As String    'will hold the type name of the control
        'Loop through all the control on the form using the Forms.Control collection
        'For each control, determine its type, and then clear it in the appropriate manner

        For Each obj In container.Controls
            strControlType = TypeName(obj)  'TypeName returns the class name of control
            Select Case strControlType
                Case "TextBox"
                    Dim cntrl As TextBox
                    cntrl = DirectCast(obj, TextBox)
                    cntrl.Text = vbNullString 'or .clear()
                Case "Checkbox"
                    Dim cntrl As CheckBox
                    cntrl = DirectCast(obj, CheckBox)
                    cntrl.Checked = False
                Case "ComboBox"
                    Dim cntrl As ComboBox
                    cntrl = DirectCast(obj, ComboBox)
                    cntrl.SelectedIndex = -1
                Case "ListBox"
                    Dim cntrl As ListBox
                    cntrl = DirectCast(obj, ListBox)
                    cntrl.SelectedIndex = -1
                Case "RadioButton"
                    Dim cntrl As RadioButton
                    cntrl = DirectCast(obj, RadioButton)
                    cntrl.Checked = False
                Case "MaskedTextBox"
                    Dim cntrl As MaskedTextBox
                    cntrl = DirectCast(obj, MaskedTextBox)
                    cntrl.Clear()
                Case "GroupBox"
                    Dim cntrl As GroupBox
                    cntrl = DirectCast(obj, GroupBox)
                    'recursively call this routine to cycle through all controlls in the groupbox
                    ClearScreenControls(cntrl)
                Case "DateTimePicker"
                    Dim cntrl As DateTimePicker
                    cntrl = DirectCast(obj, DateTimePicker)
                    cntrl.Value = Today
                Case Else 'for anything we did not include
                    'no action
            End Select
        Next
    End Sub
End Module
