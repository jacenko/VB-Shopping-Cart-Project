Public Class CMember
    'represents a single record in the Member table
    'declare the private instance variables
    Private _mstrMbrID As String
    Private _mstrLName As String
    Private _mstrFName As String
    Private _mstrAddress As String
    Private _mstrCity As String
    Private _mstrState As String
    Private _mstrZip As String
    Private _mstrPhone As String
    Private _mdtmDateJoined As DateTime
    Private _mstrEmail As String
    Private _mstrProgID As String
    Private _isNewMember As Boolean
    Private _hasChanges As Boolean
    'constructor
    Public Sub New()
        'initialize class instance variables
        _mdtmDateJoined = New Date(Now.Year, Now.Month, Now.Day) 'set it to today
    End Sub

    Public Property MbrID() As String
        Get
            Return _mstrMbrID
        End Get
        Set(strMbrID As String)
            _mstrMbrID = strMbrID
        End Set
    End Property

    Public Property LName() As String
        Get
            Return _mstrLName
        End Get
        Set(strLName As String)
            _mstrLName = strLName
        End Set
    End Property

    Public Property FName() As String
        Get
            Return _mstrFName
        End Get
        Set(strFName As String)
            _mstrFName = strFName
        End Set
    End Property

    Public Property Address() As String
        Get
            Return _mstrAddress
        End Get
        Set(strAddress As String)
            _mstrAddress = strAddress
        End Set
    End Property

    Public Property City() As String
        Get
            Return _mstrCity
        End Get
        Set(strCity As String)
            _mstrCity = strCity
        End Set
    End Property

    Public Property State() As String
        Get
            Return _mstrState
        End Get
        Set(strState As String)
            _mstrState = strState
        End Set
    End Property

    Public Property Zip() As String
        Get
            Return _mstrZip
        End Get
        Set(strZip As String)
            _mstrZip = strZip
        End Set
    End Property

    Public Property Phone() As String
        Get
            Return _mstrPhone
        End Get
        Set(strPhone As String)
            _mstrPhone = strPhone
        End Set
    End Property

    Public Property Email() As String
        Get
            Return _mstrEmail
        End Get
        Set(strEmail As String)
            _mstrEmail = strEmail
        End Set
    End Property

    Public Property ProgID() As String
        Get
            Return _mstrProgID
        End Get
        Set(strProgID As String)
            _mstrProgID = strProgID
        End Set
    End Property

    Public Property DateJoined() As DateTime
        Get
            Return _mdtmDateJoined
        End Get
        Set(dtmDateJoined As DateTime)
            _mdtmDateJoined = dtmDateJoined
        End Set
    End Property

    Public Property IsNewMember() As Boolean
        Get
            Return _isNewMember
        End Get
        Set(blnVal As Boolean)
            _isNewMember = blnVal
        End Set
    End Property

    Public Property hasChanges() As Boolean
        Get
            Return _hasChanges
        End Get
        Set(blnVal As Boolean)
            _hasChanges = blnVal
        End Set
    End Property

    Public ReadOnly Property GetSaveParameteres() As ArrayList
        Get
            Dim paramList As New ArrayList
            paramList.Add(New SqlClient.SqlParameter("mbrid", _mstrMbrID))
            paramList.Add(New SqlClient.SqlParameter("lname", _mstrLName))
            paramList.Add(New SqlClient.SqlParameter("fname", _mstrFName))
            paramList.Add(New SqlClient.SqlParameter("address", _mstrAddress))
            paramList.Add(New SqlClient.SqlParameter("city", _mstrCity))
            paramList.Add(New SqlClient.SqlParameter("state", _mstrState))
            paramList.Add(New SqlClient.SqlParameter("zip", _mstrZip))
            paramList.Add(New SqlClient.SqlParameter("phone", _mstrPhone))
            paramList.Add(New SqlClient.SqlParameter("datejoined", _mdtmDateJoined))
            paramList.Add(New SqlClient.SqlParameter("email", _mstrEmail))
            paramList.Add(New SqlClient.SqlParameter("progid", _mstrProgID))

            Return paramList
        End Get

    End Property

    Public Function Save() As Integer
        'return -1 if the ID already exists in the database (and we cannot create a new record)
        If IsNewMember Then
            Dim strRes As String = myDB.GetSingleValueFromSP _
                               ("sp_CheckMemberIDExists", New SqlClient.SqlParameter("memberID", _mstrMbrID))
            If Not strRes = 0 Then
                Return -1 'ID not unique
            End If
        End If
        'if not a new record and ID is unique then do the save (an update or an insert)
        Return myDB.ExecSP("sp_SaveMember", GetSaveParameteres())
    End Function
End Class
