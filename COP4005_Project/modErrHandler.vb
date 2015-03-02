Module modErrHandler
    Public Function ValidateTextBoxLength(ByVal obj As TextBox, ByVal errP As ErrorProvider) As Boolean
        'This procedure validates that a textbox is not empty.  The caller must pass in
        'the textbox and the error provider controls as arguments.
        If obj.Text.Length = 0 Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleRight)
            errP.SetError(obj, "You must enter a value here!")
            obj.Focus()
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function

    Public Function ValidateCombos(ByVal obj As ComboBox, ByVal errP As ErrorProvider) As Boolean
        'This procedure validates that combo selection has been made
        If obj.SelectedIndex = -1 Then 'no selection was made
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleRight)
            errP.SetError(obj, "YOu must select a value here!")
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function

    Public Function ValidateTextBoxNumeric(ByVal obj As TextBox, ByVal errP As ErrorProvider) As Boolean
        If Not IsNumeric(obj.Text) Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleRight)
            errP.SetError(obj, "You must enter a numeric value here!")
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If
    End Function

    Public Function ValidateTextBoxDate(ByVal obj As TextBox, ByVal errP As ErrorProvider)
        If Not IsDate(obj.Text) Then
            errP.SetIconAlignment(obj, ErrorIconAlignment.MiddleRight)
            errP.SetError(obj, "You must enter a valid date value here!")
            Return False
        Else
            errP.SetError(obj, "")
            Return True
        End If

    End Function
End Module