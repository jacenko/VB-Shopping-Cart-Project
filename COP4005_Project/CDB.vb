Imports System.Data.SqlClient
Public Class CDB

    Public sqlConn As SqlConnection

    Public Function OpenDB() As Boolean
        Dim blnResult As Boolean
        blnResult = False
        If sqlConn Is Nothing Then
            Try
                sqlConn = New SqlConnection(gstrConn)  'instantiate the connection object
                sqlConn.Open()
                blnResult = True
            Catch exOpenConnError As Exception
                MessageBox.Show(exOpenConnError.ToString, "Connection error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                MessageBox.Show("Cannot open database - program will end.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                blnResult = False 'Error log this better
            End Try
            Return blnResult
        Else
            If sqlConn.State = ConnectionState.Open Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Public Sub CloseDB()
        Try
            'objSQLCommand = Nothing
            sqlConn.Close()
        Catch ex As Exception
            MessageBox.Show("Error when attempting to close database: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function GetDataReaderBySP(ByVal strSP As String, ByRef params As ArrayList) As SqlDataReader
        If Not OpenDB() Then
            'Error log this better
            Return Nothing
        End If
        Dim sqlComm As New SqlCommand(strSP, sqlConn)
        sqlComm.CommandType = CommandType.StoredProcedure
        If Not params Is Nothing Then
            For Each p As SqlParameter In params
                sqlComm.Parameters.Add(p)
            Next
        End If
        Try
            Return sqlComm.ExecuteReader()
        Catch ex As Exception
            'exception handle here
        End Try
        Return Nothing
    End Function

    Public Function GetDataReaderBySP(ByVal strSP As String, ByRef param As SqlParameter) As SqlDataReader
        'Returns a dataread object filled with the records requested in the SQL statement passed in.
        Dim al As ArrayList
        If Not param Is Nothing Then
            al = New ArrayList()
            al.Add(param)
        Else
            al = Nothing
        End If
        Return GetDataReaderBySP(strSP, al)

    End Function

    Public Function GetDataReaderBySP(strSP As String)
        Dim p As SqlParameter = Nothing
        Return GetDataReaderBySP(strSP, p)
    End Function

    Public Function GetDataReader(ByVal strSQL As String) As SqlDataReader
        'Returns a dataread object filled with the records requested in the SQL statement passed in.
        If Not OpenDB() Then
            'Error Log Connection
            Return Nothing
        End If
        Dim sqlCom As New SqlCommand(strSQL, sqlConn)
        Try
            Return sqlCom.ExecuteReader()
        Catch ex As Exception
            'Exception Handle here
        End Try
        Return Nothing
    End Function

    Public Function ExecSP(strSP As String, params As ArrayList) As Integer
        If Not OpenDB() Then
            Return -1
        End If

        Dim sqlComm As New SqlCommand(strSP, sqlConn)
        sqlComm.CommandType = CommandType.StoredProcedure
        Try
            If Not params Is Nothing Then
                For Each p As SqlParameter In params
                    sqlComm.Parameters.Add(p)
                Next
            End If
            Return sqlComm.ExecuteNonQuery() 'Should wrap in an error code handler
            'Return 0
        Catch ex As Exception

            Return -1
        End Try
    End Function

    Public Function GetSingleValueFromSP(ByVal strSP As String, ByRef params As ArrayList) As String
        Dim dr As SqlDataReader = GetDataReaderBySP(strSP, params)
        If Not dr Is Nothing Then
            If dr.Read() Then
                Return dr.GetValue(0).ToString()
            Else
                Return -1 'no data
            End If
        End If
        Return -2 'failed to connect to db
    End Function

    Public Function GetSingleValueFromSP(strSP As String, param As SqlParameter) As String
        Dim al As ArrayList
        If param Is Nothing Then
            al = Nothing
        Else
            al = New ArrayList()
            al.Add(param)
        End If
        Return GetSingleValueFromSP(strSP, al)
    End Function

    Public Function GetSingleValueFromSP(strSP As String) As String
        Dim al As ArrayList = Nothing
        Return GetSingleValueFromSP(strSP, al)
    End Function

End Class
