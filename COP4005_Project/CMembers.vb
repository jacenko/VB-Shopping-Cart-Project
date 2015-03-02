Imports System.Data.SqlClient

Public Class CMembers
    Private _Member As CMember
    'constructor
    Public Sub New()
        _Member = New CMember
    End Sub
    Public ReadOnly Property CurrentObject() As CMember
        Get
            Return _Member
        End Get
    End Property

    Public Sub CreateNewMember() 'call me when you are clearing the screen to add a new record
        Clear()
        _Member.IsNewMember = True
    End Sub

    Public Sub Clear()
        _Member = New CMember()
    End Sub


    Public Function Save() As Integer
        Return _Member.Save
    End Function

    Public Function GetMemberList() As SqlClient.SqlDataReader
        Return myDB.GetDataReaderBySP("dbo.sp_GetMemberList")
    End Function

    Public Function GetMemberPhoneList() As SqlClient.SqlDataReader
        Return myDB.GetDataReaderBySP("dbo.sp_GetMemberPhoneList")
    End Function

    Public Function GetMemberByID(strID As String) As CMember
        Dim aParam As New SqlParameter("id", strID)
        FillObject(myDB.GetDataReaderBySP("dbo.sp_GetMemberByID", aParam))
        Return _Member
    End Function

    Public Function GetMemberByName(strLName As String, strFName As String) As CMember
        Dim params As New ArrayList 'an aray list for the SQL parameters
        Dim param1 As New SqlParameter("lname", strLName)
        params.Add(param1)
        Dim param2 As New SqlParameter("fname", strFName)
        params.Add(param2)
        FillObject(myDB.GetDataReaderBySP("dbo.sp_GetMemberByName", params))
        Return _Member
    End Function

    Public Function FillObject(sqlDR As SqlDataReader) As CMember
        Using sqlDR
            If sqlDR.Read Then
                With _Member
                    .MbrID = sqlDR.Item("MBRID")    'value in parens must be a column name
                    .LName = sqlDR.Item("LNAME")
                    .FName = sqlDR.Item("FNAME")
                    .Address = sqlDR.Item("ADDRESS")
                    .City = sqlDR.Item("CITY")
                    .Zip = sqlDR.Item("ZIP")
                    .Phone = sqlDR.Item("PHONE")
                    .DateJoined = SafeDate(sqlDR.Item("DateJoined").ToString())
                    .Email = sqlDR.Item("EMAIL")
                    .ProgID = sqlDR.Item("PROGID")
                End With
            End If
        End Using
        Return _Member
    End Function

    Public Function GetSearchResults(search As String) As SqlClient.SqlDataReader
        Dim aParam As New SqlParameter("search", search)
        Return myDB.GetDataReaderBySP("dbo.sp_GetMemberSearchResults", aParam)
    End Function
End Class