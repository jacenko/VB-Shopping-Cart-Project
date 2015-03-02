'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Imports System.Data.SqlClient
Public Class CSecurity

    'check if user ID and password are correct
    Public Function CheckLogin(user As String, pass As String) As SqlClient.SqlDataReader
        Dim params As New ArrayList 'an aray list for the SQL parameters
        Dim param1 As New SqlParameter("username", user)
        params.Add(param1)
        Dim param2 As New SqlParameter("password", pass)
        params.Add(param2)
        Return myDB.GetDataReaderBySP("dbo.sp_CheckUserIDAndPassword", params)
    End Function

    'check if user has access to the shopping cart
    Public Function CheckAccess(user As String, pass As String) As SqlClient.SqlDataReader
        Dim params As New ArrayList 'an aray list for the SQL parameters
        Dim param1 As New SqlParameter("username", user)
        params.Add(param1)
        Dim param2 As New SqlParameter("password", pass)
        params.Add(param2)
        Return myDB.GetDataReaderBySP("dbo.sp_CheckUserAccess", params)
    End Function
End Class
