'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Public Class CProduct
    'represents a single record in the Product table
    'declare the private instance variables
    Private _mstrProdID As String
    Private _mstrProdDesc As String
    Private _mstrWhCost As String
    Private _mstrRetPrice As String
    Private _mstrTaxable As String
    Private _mstrSupID As String

    Public Property ProdID() As String
        Get
            Return _mstrProdID
        End Get
        Set(strProdID As String)
            _mstrProdID = strProdID
        End Set
    End Property

    Public Property ProdDesc() As String
        Get
            Return _mstrProdDesc
        End Get
        Set(strProdDesc As String)
            _mstrProdDesc = strProdDesc
        End Set
    End Property

    Public Property WhCost() As String
        Get
            Return _mstrWhCost
        End Get
        Set(strWhCost As String)
            _mstrWhCost = strWhCost
        End Set
    End Property

    Public Property RetPrice() As String
        Get
            Return _mstrRetPrice
        End Get
        Set(strRetPrice As String)
            _mstrRetPrice = strRetPrice
        End Set
    End Property

    Public Property Taxable() As String
        Get
            Return _mstrTaxable
        End Get
        Set(strTaxable As String)
            _mstrTaxable = strTaxable
        End Set
    End Property

    Public Property SupID() As String
        Get
            Return _mstrSupID
        End Get
        Set(strSupID As String)
            _mstrSupID = strSupID
        End Set
    End Property

    Public ReadOnly Property GetSaveParameteres() As ArrayList
        Get
            Dim paramList As New ArrayList
            paramList.Add(New SqlClient.SqlParameter("ProdID", _mstrProdID))
            paramList.Add(New SqlClient.SqlParameter("ProdDesc", _mstrProdDesc))
            paramList.Add(New SqlClient.SqlParameter("WhCost", _mstrWhCost))
            paramList.Add(New SqlClient.SqlParameter("RetPrice", _mstrRetPrice))
            paramList.Add(New SqlClient.SqlParameter("Taxable", _mstrTaxable))
            paramList.Add(New SqlClient.SqlParameter("SupID", _mstrSupID))

            Return paramList
        End Get

    End Property

End Class