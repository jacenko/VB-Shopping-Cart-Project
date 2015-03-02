'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Imports System.Data.SqlClient

Public Class frmLogin
    'declare variables
    Private Security As CSecurity
    Private sqlReader As SqlDataReader
    Private strUser As String
    Private strPass As String
    Public Shared blnMatch As Boolean = False
    Public Shared blnLogin As Boolean = False

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Security = New CSecurity
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

    'clicking the login button
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        strUser = txtUserName.Text
        strPass = txtPassword.Text
        'check credentials
        sqlReader = Security.CheckLogin(strUser, strPass)
        If sqlReader.HasRows Then
            'set global variables
            While sqlReader.Read
                employeeID = sqlReader.Item("EMPID")
                employeeInfo = sqlReader.Item("EMPID") + "/" + sqlReader.Item("USERID")
                employeeFName = sqlReader.Item("FNAME")
            End While
            sqlReader.Close()
            'check if user has access to the shopping cart
            sqlReader = Security.CheckAccess(strUser, strPass)
            If sqlReader.HasRows Then
                shopAccess = True
            End If
            sqlReader.Close()
            Me.Close()
            blnLogin = True
        Else    'user entered the wrong information. clear all fields and notify the user
            txtUserName.Text = ""
            txtPassword.Text = ""
            sqlReader.Close()
            sslLoginStatus.Text = "Incorrect Username/Password"
        End If
    End Sub

    'user does not want to use the application
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        End
    End Sub

    'user presses the Enter key from any field in the form
    Private Sub txtBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUserName.KeyPress, txtPassword.KeyPress
        If e.KeyChar = Chr(13) Then 'Chr(13) is the Enter Key 
            'Runs the Click Event 
            btnLogin_Click(Me, EventArgs.Empty)
        End If
    End Sub
End Class