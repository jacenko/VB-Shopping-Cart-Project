'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Imports System.Data.SqlClient

Public Class CInvoice

    Private sqlReader As SqlDataReader

    'get the newest invoice ID
    Public Function GetLastInvoiceID() As String
        sqlReader = myDB.GetDataReaderBySP("dbo.sp_GetLastInvoiceID")
        Dim lastInvID As String = ""
        While sqlReader.Read
            lastInvID = sqlReader.Item("INVID")
        End While
        'Dim lastInvID As String = myDB.GetSingleValueFromSP("dbo.sp_GetLastInvoiceID")
        If lastInvID = "" Then
            lastInvID = "I0000"
        End If
        sqlReader.Close()
        Return lastInvID
    End Function

    'save the product invoice after checking out
    Public Function SaveProductInvoice(invID As String, mbrID As String, empID As String, _
                                       total As Double, subtotal As Double, totalTax As Double, _
                                       todaysDate As String) As Integer
        Dim execute As Integer
        Dim params As New ArrayList
        params.Add(New SqlClient.SqlParameter("invID", invID))
        params.Add(New SqlClient.SqlParameter("mbrID", mbrID))
        params.Add(New SqlClient.SqlParameter("empID", empID))
        params.Add(New SqlClient.SqlParameter("total", total))
        params.Add(New SqlClient.SqlParameter("subtotal", subtotal))
        params.Add(New SqlClient.SqlParameter("totalTax", totalTax))
        params.Add(New SqlClient.SqlParameter("todaysDate", todaysDate))
        execute = myDB.ExecSP("sp_SaveProductInvoice", params)
        Return execute
    End Function

    'save information for each item in a receipt
    Public Function SaveInvoiceItem(invID As String, prodID As String, qty As String, price As String) As Integer
        Dim execute As Integer
        Dim params As New ArrayList
        params.Add(New SqlClient.SqlParameter("invID", invID))
        params.Add(New SqlClient.SqlParameter("prodID", prodID))
        params.Add(New SqlClient.SqlParameter("qty", qty))
        params.Add(New SqlClient.SqlParameter("price", price))
        execute = myDB.ExecSP("sp_SaveInvoiceItem", params)
        Return execute
    End Function

    'retrieve product description, price, and quantity for the data grid
    Public Function GetInvoiceInfo(invID As String) As SqlClient.SqlDataReader
        Dim aParam As New SqlParameter("invID", invID)
        Return myDB.GetDataReaderBySP("dbo.sp_GetInvoiceInfo", aParam)
    End Function

    'retrieve information for the receipt form labels
    Public Function GetInvoiceFormInfo(invID As String) As SqlClient.SqlDataReader
        Dim aParam As New SqlParameter("invID", invID)
        Return myDB.GetDataReaderBySP("dbo.sp_GetInvoiceFormInfo", aParam)
    End Function
End Class