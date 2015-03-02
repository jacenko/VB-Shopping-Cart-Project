'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Module modGlobal
    Public gstrConn As String = "Data Source=(LocalDB)\v11.0;AttachDbFilename='C:\Users\Deniss\Documents\Visual Studio 2013\Projects\COP4005_Project\COP4005ProjectDB.mdf';Integrated Security=True;Timeout=30"
    Public myDB As New CDB
    Public ReadOnly NULL_DATE As Date = New Date(1900, 1, 1)
    Public shopAccess As Boolean = False        'checks if user has access to the shopping cart
    Public employeeID As String = ""            'stores the currently logged in employee's ID
    Public employeeInfo As String = ""          'stores the currently logged in employee's information
    Public employeeFName As String = ""         'stores the currently logged in employee's first name
    Public invoiceID As String = ""             'stores the currently observed invoice ID for use in the receipt

    'constants that need to be accessed in more than one form
    Public Const ACTION_NONE As Integer = 0
    Public Const ACTION_HOME As Integer = 1
    Public Const ACTION_MEMBER As Integer = 2
    Public Const ACTION_PROGRAM As Integer = 3
    Public Const ACTION_SHOP As Integer = 4
    Public Const ACTION_CONTACT As Integer = 5
    Public Const ACTION_HELP As Integer = 6
    Public Const ACTION_SUPPLIER As Integer = 7
    Public Const ACTION_EMPLOYEE As Integer = 8
    Public Const ACTION_PO As Integer = 9
    Public Const TAX As Double = 0.07           'Dade County standard tax

    Public intNextAction As Integer

End Module
