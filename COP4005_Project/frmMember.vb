'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Imports System.Data.SqlClient

Public Class frmMember

    Private Members As CMembers
    Private Programs As CProgams
    Private sqlReader As SqlDataReader
    Private blnClearing As Boolean
    Private blnNew As Boolean

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Members = New CMembers
        Programs = New CProgams
    End Sub

    Private Sub tsbContact_Click(sender As Object, e As EventArgs) Handles tsbContact.Click
        intNextAction = ACTION_CONTACT
        Me.Hide()
    End Sub

    Private Sub tsbHelp_Click(sender As Object, e As EventArgs) Handles tsbHelp.Click
        intNextAction = ACTION_HELP
        Me.Hide()
    End Sub

    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        'do nothing, we are already on this form

    End Sub

    Private Sub tsbHome_Click(sender As Object, e As EventArgs) Handles tsbHome.Click
        intNextAction = ACTION_HOME
        Me.Hide()
    End Sub

    Private Sub tsbProgram_Click(sender As Object, e As EventArgs) Handles tsbProgram.Click
        intNextAction = ACTION_PROGRAM
        Me.Hide()
    End Sub

    Private Sub tsbReturn_Click(sender As Object, e As EventArgs) Handles tsbReturn.Click
        intNextAction = ACTION_NONE
        Me.Hide()
    End Sub

    Private Sub tsbShop_Click(sender As Object, e As EventArgs) Handles tsbShop.Click
        intNextAction = ACTION_SHOP
        Me.Hide()
    End Sub

    Private Sub tsbEmployee_Click(sender As Object, e As EventArgs) Handles tsbEmployee.Click
        intNextAction = ACTION_EMPLOYEE
        Me.Hide()
    End Sub

    Private Sub tsbSupplier_Click(sender As Object, e As EventArgs) Handles tsbSupplier.Click
        intNextAction = ACTION_SUPPLIER
        Me.Hide()
    End Sub

    Private Sub tsbPO_Click(sender As Object, e As EventArgs) Handles tsbPO.Click
        intNextAction = ACTION_PO
        Me.Hide()
    End Sub

    Private Sub tsbProxy_MouseEnter(sender As Object, e As EventArgs) Handles tsbContact.MouseEnter, tsbReturn.MouseEnter, tsbHelp.MouseEnter,
        tsbHome.MouseEnter, tsbMember.MouseEnter, tsbProgram.MouseEnter, tsbShop.MouseEnter, tsbEmployee.MouseEnter, tsbPO.MouseEnter, tsbSupplier.MouseEnter
        Dim tsbProxy As ToolStripButton
        tsbProxy = CType(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Text
    End Sub

    Private Sub tsbProxy_MouseLeave(sender As Object, e As EventArgs) Handles tsbContact.MouseLeave, tsbReturn.MouseLeave, tsbHelp.MouseLeave,
        tsbHome.MouseLeave, tsbMember.MouseLeave, tsbProgram.MouseLeave, tsbShop.MouseLeave, tsbEmployee.MouseLeave, tsbPO.MouseLeave, tsbSupplier.MouseLeave
        Dim tsbProxy As ToolStripButton
        tsbProxy = CType(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Image
    End Sub

    Private Sub LoadMemberCombo()
        'load the name search combo
        sqlReader = Members.GetMemberList
        cboName.Items.Clear() 'clear out current list in order to reload
        While sqlReader.Read
            cboName.Items.Add(sqlReader.Item("LName") & ", " & sqlReader.Item("FName"))
        End While
        sqlReader.Close()
    End Sub

    Private Sub LoadProgramCombo()
        'load the program search combo
        sqlReader = Programs.GetAllProgramIDs
        cboProgram.Items.Clear() 'clear out current list in order to reload
        While sqlReader.Read
            cboProgram.Items.Add(sqlReader.Item("ProgID"))
        End While
        sqlReader.Close()
    End Sub

    Private Sub LoadPhoneList()
        'load the phone search combo
        sqlReader = Members.GetMemberPhoneList
        cboPhone.Items.Clear() 'clear out current list in order to reload
        While sqlReader.Read
            cboPhone.Items.Add(sqlReader.Item("Phone"))
        End While
        sqlReader.Close()
    End Sub

    Private Sub frmMember_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        blnClearing = True
        ClearScreenControls(Me)
        blnClearing = False
        LoadMemberCombo()
        LoadProgramCombo()
        LoadPhoneList()
    End Sub

    Private Sub LoadMemberData(aMember As CMember)
        'takes the values from the cmember object and displays them on the form
        If Not aMember Is Nothing Then
            With aMember
                txtID.Text = .MbrID
                txtFName.Text = .FName
                txtLName.Text = .LName
                txtAddress.Text = .Address
                txtCity.Text = .City
                txtState.Text = .State
                mskZip.Text = .Zip
                mskPhone.Text = .Phone
                txtEmail.Text = .Email
                dtmJoined.Value = .DateJoined
                cboProgram.SelectedIndex = cboProgram.FindStringExact(.ProgID)
            End With
        End If
    End Sub

    Private Sub LoadMemberFormByName(strFullName As String)
        Dim strNames() As String 'a string array
        strNames = strFullName.Split(","c) 'break out the pieces of the name
        Members.GetMemberByName(Trim(strNames(0)), Trim(strNames(1))) 'this loads CurrentObject
        LoadMemberData(Members.CurrentObject)
    End Sub

    'name has been changed
    Private Sub cboName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboName.SelectedIndexChanged
        If Not blnClearing Then
            'grab the selected name and fill the form
            LoadMemberFormByName(cboName.SelectedItem.ToString)
            cboPhone.SelectedIndex = cboName.SelectedIndex
        End If
    End Sub

    'phone has been changed
    Private Sub cboPhone_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboPhone.SelectedIndexChanged
        cboName.SelectedIndex = cboPhone.SelectedIndex
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim blnError As Boolean = False
        Dim intResult As Integer
        'First do input data validation - left as an exercise for student

        'after validation
        If blnError Then 'don't try to process, data is not good
            Exit Sub
        End If
        'now load the CMmeber object with form data
        With Members.CurrentObject
            .MbrID = Trim(txtID.Text)
            .FName = Trim(txtFName.Text)
            .LName = Trim(txtLName.Text)
            .Address = Trim(txtAddress.Text)
            .City = Trim(txtCity.Text)
            .State = Trim(txtState.Text)
            .Zip = mskZip.Text
            .Phone = mskPhone.Text
            .Email = Trim(txtEmail.Text)
            .DateJoined = dtmJoined.Value.Date
            .ProgID = cboProgram.SelectedItem.ToString
        End With
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            intResult = Members.Save
            'success!
            sslStatus.Text = "Record Saved"
        Catch ex As Exception
            If intResult = -1 Then    'ID is not unique
                MessageBox.Show("Member ID must be unique. Unable to save record.",
                                "Database error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else    'some other error
                MessageBox.Show("Unable to save the record: " & ex.ToString, "Database error",
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        End Try
    End Sub

    Private Sub chkNew_CheckedChanged(sender As Object, e As EventArgs) Handles chkNew.CheckedChanged
        If chkNew.CheckState Then
            blnClearing = True
            ClearScreenControls(Me)
            blnClearing = False
            Members.Clear()
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        blnClearing = True
        ClearScreenControls(Me)
        blnClearing = False
    End Sub

    'user wants to search for member by last partial name
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Dim search As String = ""
        cboName.Items.Clear()
        cboPhone.Items.Clear()
        search = txtSearch.Text.ToLower
        If search = "" Then
            LoadMemberCombo()
            LoadPhoneList()
        Else
            sqlReader = Members.GetSearchResults(search)
            'fill the searched last name's full name and phone combo boxes
            While sqlReader.Read
                cboName.Items.Add(sqlReader.Item("LName") & ", " & sqlReader.Item("FName"))
                cboPhone.Items.Add(sqlReader.Item("PHONE"))
            End While
            sqlReader.Close()
        End If
    End Sub
End Class