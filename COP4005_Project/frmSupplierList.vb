Imports System.Data.SqlClient

Public Class frmSupplierList
    Private sqlReader As SqlDataReader

    Private Sub LoadDataGridView()
        sqlReader = myDB.GetDataReaderBySP("sp_GetSupplierList")
        'cannot bind a datareader to a grid view directly
        'must load a data table with the reader, and load the datagrid view from the data table
        Dim dt As New DataTable
        dt.Load(sqlReader)
        'build with DB table fields as the column headings
        'dgvSupplier.AutoGenerateColumns = True
        'to create your own headings
        dgvSupplier.ColumnCount = 8
        dgvSupplier.Columns(0).Name = "ID"
        dgvSupplier.Columns(1).Name = "Name"
        dgvSupplier.Columns(2).Name = "Address"
        dgvSupplier.Columns(3).Name = "City"
        dgvSupplier.Columns(4).Name = "State"
        dgvSupplier.Columns(5).Name = "Zip"
        dgvSupplier.Columns(6).Name = "Phone"
        dgvSupplier.Columns(7).Name = "Rep"
        'must now associate each column with a table column
        dgvSupplier.Columns(0).DataPropertyName = "SUPID"
        dgvSupplier.Columns(1).DataPropertyName = "SUPNAME"
        dgvSupplier.Columns(2).DataPropertyName = "SUPADDRESS"
        dgvSupplier.Columns(3).DataPropertyName = "SUPCITY"
        dgvSupplier.Columns(4).DataPropertyName = "SUPSTATE"
        dgvSupplier.Columns(5).DataPropertyName = "SUPZIP"
        dgvSupplier.Columns(6).DataPropertyName = "SUPPHONE"
        dgvSupplier.Columns(7).DataPropertyName = "SUPREPNAME"

        dgvSupplier.DataSource = dt 'bind datagridview to the datatable we created
        dgvSupplier.Refresh()
    End Sub

    Private Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click
        Me.Hide()
    End Sub

    Private Sub frmSupplierList_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'clear out datagridview by disconnecting it from its datasource
        dgvSupplier.DataSource = Nothing
        LoadDataGridView()
    End Sub

    
    Private Sub dgvSupplier_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvSupplier.CellMouseDoubleClick
        If e.RowIndex = -1 Then 'user clicked on a column heading
            Exit Sub
        End If
        Dim strID As String
        strID = dgvSupplier.SelectedRows(0).Cells("ID").Value.ToString
        MessageBox.Show("You selected Supplier #" & strID)
    End Sub
End Class