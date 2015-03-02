'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Public Class frmMain
    Private MemberInfo As frmMember
    Private SupplierList As frmSupplierList
    Private LoginScreen As frmLogin
    Private ShopCart As frmShoppingCart

    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        Me.Hide()
        MemberInfo.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub tsbExit_Click(sender As Object, e As EventArgs) Handles tsbExit.Click
        EndProgram()
    End Sub

    Private Sub tsbPO_Click(sender As Object, e As EventArgs) Handles tsbPO.Click
        MessageBox.Show("Not yet implemented")
    End Sub

    Private Sub tsbSupplier_Click(sender As Object, e As EventArgs) Handles tsbSupplier.Click
        SupplierList.ShowDialog()
    End Sub

    Private Sub tsbEmployee_Click(sender As Object, e As EventArgs) Handles tsbEmployee.Click
        MessageBox.Show("Not yet implemented")
    End Sub

    Private Sub tsbContact_Click(sender As Object, e As EventArgs) Handles tsbContact.Click
        MessageBox.Show("Not yet implemented")
    End Sub

    Private Sub tsbShop_Click(sender As Object, e As EventArgs) Handles tsbShop.Click
        If shopAccess = True Then
            Me.Hide()
            ShopCart.ShowDialog()
            PerformNextAction()
        Else
            MessageBox.Show("You don't have sufficient authorization to use the shopping cart")
        End If
    End Sub

    Private Sub tsbProxy_MouseEnter(sender As Object, e As EventArgs) Handles tsbContact.MouseEnter, tsbExit.MouseEnter, tsbHelp.MouseEnter,
        tsbHome.MouseEnter, tsbMember.MouseEnter, tsbProgram.MouseEnter, tsbShop.MouseEnter, tsbEmployee.MouseEnter, tsbPO.MouseEnter, tsbSupplier.MouseEnter
        Dim tsbProxy As ToolStripButton
        tsbProxy = CType(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Text
    End Sub

    Private Sub tsbProxy_MouseLeave(sender As Object, e As EventArgs) Handles tsbContact.MouseLeave, tsbExit.MouseLeave, tsbHelp.MouseLeave,
        tsbHome.MouseLeave, tsbMember.MouseLeave, tsbProgram.MouseLeave, tsbShop.MouseEnter, tsbEmployee.MouseLeave, tsbPO.MouseLeave, tsbSupplier.MouseLeave
        Dim tsbProxy As ToolStripButton
        tsbProxy = CType(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Image
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        'instantiate an object for each other form in the application
        MemberInfo = New frmMember
        LoginScreen = New frmLogin
        ShopCart = New frmShoppingCart
        SupplierList = New frmSupplierList

        'open the database
        Try
            myDB.OpenDB()
        Catch ex As Exception
            MessageBox.Show("Unable to open database. Connection string = " & gstrConn & " Program will end.",
                            "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            EndProgram()
        End Try

        If frmLogin.blnLogin = False Then
            Me.Hide()
            LoginScreen.ShowDialog()
            Me.Show()
        End If
    End Sub

    Private Sub EndProgram()
        'close every form except main
        Dim f As Form
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        For Each f In Application.OpenForms
            If f.Name <> Me.Name Then
                If Not f Is Nothing Then
                    f.Close()
                End If
            End If
        Next
        'close the database connection
        If Not myDB.sqlConn Is Nothing Then
            myDB.sqlConn.Close()
            myDB.sqlConn.Dispose()
        End If
        Windows.Forms.Cursor.Current = Cursors.Default
        End 'kill the program
    End Sub

    Private Sub PerformNextAction()
        'get the next action selected on the child form, and then simulate the toolstrip button click here
        Select Case intNextAction
            Case ACTION_HOME
                tsbHome.PerformClick()
            Case ACTION_CONTACT
                tsbContact.PerformClick()
            Case ACTION_HELP
                tsbHelp.PerformClick()
            Case ACTION_MEMBER
                tsbMember.PerformClick()
            Case ACTION_PROGRAM
                tsbProgram.PerformClick()
            Case ACTION_SHOP
                tsbShop.PerformClick()
            Case Else
                'do nothing
        End Select
    End Sub
End Class
