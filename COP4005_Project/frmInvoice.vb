'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Imports System.Data.SqlClient

Public Class frmInvoice
    'declare variables
    Private sqlReader As SqlDataReader
    Private Invoice As CInvoice
    Private ShopCart As frmShoppingCart

    'user wants to go to the main form
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        intNextAction = ACTION_NONE
        ShopCart.Show()
        Me.Close()
    End Sub


    Private Sub frmInvoice_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Focus()
        LoadDataGridView()

        ShopCart = New frmShoppingCart
        Invoice = New CInvoice
        sqlReader = Invoice.GetInvoiceInfo(invoiceID)
        'load the data grid with the latest invoice info through a data table
        Dim dt As New DataTable
        dt.Load(sqlReader)
        dgvItems.DataSource = dt 'bind datagridview to the datatable we created
        dgvItems.Refresh()
        sqlReader.Close()
        sqlReader = Invoice.GetInvoiceFormInfo(invoiceID)
        'update all labels
        While sqlReader.Read()
            lblMemberID.Text = sqlReader.Item("MBRID")
            lblDate.Text = sqlReader.Item("INVDATE")
            lblSubtotal.Text = FormatCurrency(sqlReader.Item("PRODUCTTOTAL"))
            lblTax.Text = FormatCurrency(sqlReader.Item("TAXTOTAL"))
            lblTotal.Text = FormatCurrency(sqlReader.Item("INVTOTAL"))
        End While
        sqlReader.Close()
        lblEmpName.Text = employeeFName
        lblEmployeeID.Text = employeeID
        lblInvoiceID.Text = invoiceID
    End Sub

    Private Sub LoadDataGridView()
        'build with DB table fields as the column headings
        dgvItems.AutoGenerateColumns = True
        dgvItems.ReadOnly = True
        dgvItems.MultiSelect = False
        dgvItems.RowHeadersVisible = False
        dgvItems.DefaultCellStyle.SelectionBackColor = Color.White
        dgvItems.DefaultCellStyle.SelectionForeColor = Color.Black
        'to create your own headings
        dgvItems.ColumnCount = 3
        dgvItems.Columns(0).Name = "Description"
        dgvItems.Columns(0).Width = 230
        dgvItems.Columns(1).Name = "Price"
        dgvItems.Columns(1).Width = 40
        dgvItems.Columns(2).Name = "QTY"
        dgvItems.Columns(2).Width = 38
        'must now associate each column with a table column
        dgvItems.Columns(0).DataPropertyName = "PRODDESC"
        dgvItems.Columns(1).DataPropertyName = "PRICE"
        dgvItems.Columns(2).DataPropertyName = "QTY"
    End Sub
End Class