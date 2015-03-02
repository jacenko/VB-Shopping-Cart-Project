'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Imports System.Data.SqlClient

Public Class CProducts
    Private _Product As CProduct

    'constructor
    Public Sub New()
        _Product = New CProduct
    End Sub
    Public ReadOnly Property CurrentObject() As CProduct
        Get
            Return _Product
        End Get
    End Property


    Public Sub Clear()
        _Product = New CProduct()
    End Sub

    Public Function FillObject(sqlDR As SqlDataReader) As CProduct
        Using sqlDR
            If sqlDR.Read Then
                With _Product
                    .ProdID = sqlDR.Item("PRODID")    'value in parens must be a column name
                    .ProdDesc = sqlDR.Item("PRODDESC")
                    .WhCost = sqlDR.Item("WHCOST")
                    .RetPrice = sqlDR.Item("RETPRICE")
                    .Taxable = sqlDR.Item("TAXABLE")
                    .SupID = sqlDR.Item("SUPID")
                End With
            End If
        End Using
        Return _Product
    End Function

    'returns product descriptions with matching substrings
    Public Function GetSearchResults(search As String) As SqlClient.SqlDataReader
        Dim aParam As New SqlParameter("search", search)
        Return myDB.GetDataReaderBySP("dbo.sp_GetProductSearchResults", aParam)
    End Function

    'returns information to be used in the shopping cart DataGridView
    Public Function GetAddToCartItems(desc As String) As SqlClient.SqlDataReader
        Dim aParam As New SqlParameter("desc", desc)
        Return myDB.GetDataReaderBySP("dbo.sp_GetAddToCartItems", aParam)
    End Function

End Class
