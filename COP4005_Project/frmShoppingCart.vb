'Affirmation of Authorship:
'Name: Deniss Jacenko
'Date: 6/18/14
'I affirm that this program was created by me. It is solely my work and does not include any work done by anyone else.

Imports System.Data.SqlClient

Public Class frmShoppingCart
    'declare variables
    Private arrSearchResults() As String
    Private strSearch As String
    Private blnClearing As Boolean
    Private dblSubtotal As Double
    Private dblTotalTax As Double
    Private dblTotal As Double
    Private Receipt As frmInvoice
    Private sqlReader As SqlDataReader
    Private Members As CMembers
    Private Products As CProducts
    Private Main As frmMain
    Private Invoice As CInvoice

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Members = New CMembers
        Products = New CProducts
    End Sub

    'fill the data grid with headers
    Private Sub LoadDataGridView()
        'build with DB table fields as the column headings
        dgvCart.AutoGenerateColumns = True
        dgvCart.ReadOnly = True
        dgvCart.MultiSelect = False
        dgvCart.RowHeadersVisible = False
        dgvCart.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvCart.ScrollBars = ScrollBars.Vertical
        'to create your own headings
        dgvCart.ColumnCount = 5
        dgvCart.Columns(0).Name = "Description"
        dgvCart.Columns(0).Width = 280
        dgvCart.Columns(1).Name = "Price"
        dgvCart.Columns(1).Width = 69
        dgvCart.Columns(2).Name = "QTY"
        dgvCart.Columns(2).Width = 40
        dgvCart.Columns(3).Name = "Taxable"
        dgvCart.Columns(3).Width = 0
        dgvCart.Columns(4).Name = "Product ID"
        dgvCart.Columns(4).Width = 0
        'must now associate each column with a table column
        dgvCart.Columns(0).DataPropertyName = "PRODDESC"
        dgvCart.Columns(1).DataPropertyName = "RETPRICE"
        dgvCart.Columns(3).DataPropertyName = "TAXABLE"
        dgvCart.Columns(4).DataPropertyName = "PRODID"
    End Sub

    'user goes back to the main form
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        intNextAction = ACTION_NONE
        Main = New frmMain
        Main.Show()
        Me.Close()
    End Sub

    'load member names
    Private Sub LoadMemberCombo()
        'load the name search combo
        sqlReader = Members.GetMemberList
        cboName.Items.Clear() 'clear out current list in order to reload
        While sqlReader.Read
            cboName.Items.Add(sqlReader.Item("LName") & ", " & sqlReader.Item("FName"))
        End While
        sqlReader.Close()
    End Sub

    'load member IDs
    Private Sub LoadIDCombo()
        'load the name search combo
        sqlReader = Members.GetMemberList
        cboMemberID.Items.Clear() 'clear out current list in order to reload
        While sqlReader.Read
            cboMemberID.Items.Add(sqlReader.Item("MbrID"))
        End While
        sqlReader.Close()
    End Sub

    Private Sub frmShoppingCart_Load(sender As Object, e As EventArgs) Handles Me.Load

        lblLoggedIn.Text = "Logged In as " & employeeInfo       'displayed at the top
        LoadDataGridView()
        'disable controls until a member is selected
        txtSearch.Enabled = False
        btnSearchProduct.Enabled = False
        cboResults.Enabled = False
        btnAdd.Enabled = False
        btnRemove.Enabled = False
        dgvCart.Enabled = False
        btnClear.Enabled = False
        btnCheckout.Enabled = False
        btnLess.Enabled = False
        btnMore.Enabled = False
        sslStatus.Text = "Please choose a member"
    End Sub

    'prepare the shopping cart screen
    Private Sub frmShoppingCart_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        btnClear_Click(Me, EventArgs.Empty)
        blnClearing = True
        ClearScreenControls(Me)
        blnClearing = False
        LoadMemberCombo()
        LoadIDCombo()
    End Sub

    'perform the sale
    Private Sub btnCheckout_Click(sender As Object, e As EventArgs) Handles btnCheckout.Click
        If dgvCart.Rows.Count > 0 Then      'if something is in the cart
            Dim intResult1 As Integer       'user was able to save the invoice information
            Dim intResult2 As Integer       'user was able to save the invoice product information
            Dim lastInvID As String         'the newest invoice ID in the database
            Dim newInvID As String          '+1
            Dim lastInvIDNumber As Integer  'integer for parsing to the following string
            Dim newInvIDNumber As String    'the newly created invoice number without "I"
            Dim todaysDate As String = String.Format("{0:MM/dd/yyyy}", DateTime.Now)
            Invoice = New CInvoice

            lastInvID = Invoice.GetLastInvoiceID()

            'Take the old invoice ID number out
            lastInvIDNumber = CInt(lastInvID.Substring(1, 4))

            'needs to be a number 4 characters long (more only whe invoice number is greater than 9999)
            Select Case lastInvIDNumber
                Case Is < 10
                    newInvIDNumber = "000" & (lastInvIDNumber + 1)
                Case Is < 100
                    newInvIDNumber = "00" & (lastInvIDNumber + 1)
                Case Is < 1000
                    newInvIDNumber = "0" & (lastInvIDNumber + 1)
                Case Else
                    newInvIDNumber = (lastInvIDNumber + 1)
            End Select

            'concatenate "i" with the new ID number
            newInvID = "I" & newInvIDNumber

            'save the invoice to the database
            intResult1 = Invoice.SaveProductInvoice(newInvID, cboMemberID.SelectedItem.ToString, employeeID, _
                                                   dblTotal, dblSubtotal, dblTotalTax, _
                                                   todaysDate)

            'save each product item to the database
            For i = 0 To dgvCart.Rows.Count - 1
                intResult2 = Invoice.SaveInvoiceItem(newInvID, dgvCart.Rows(i).Cells(4).Value.ToString, _
                                                     dgvCart.Rows(i).Cells(2).Value.ToString, _
                                                     dgvCart.Rows(i).Cells(1).Value.ToString)
            Next

            'check if both saves worked
            If intResult1 = -1 Or intResult2 = -1 Then
                'alert that a connection problem occured
                sslStatus.Text = "There was a problem connecting to the database"
            Else
                'show the invoice page
                invoiceID = newInvID
                Me.Hide()
                Receipt = New frmInvoice
                Receipt.Show()
            End If

        Else
            sslStatus.Text = "Please add items to cart before checking out"
        End If
    End Sub

    'if the name is selected or changed, the controls are enabled and the member ID changes accordingly
    Private Sub cboName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboName.SelectedIndexChanged
        cboMemberID.SelectedIndex = cboName.SelectedIndex

        txtSearch.Enabled = True
        btnSearchProduct.Enabled = True
        cboResults.Enabled = True
        btnAdd.Enabled = True
        btnRemove.Enabled = True
        dgvCart.Enabled = True
        btnClear.Enabled = True
        btnCheckout.Enabled = True
        btnLess.Enabled = True
        btnMore.Enabled = True
        btnExit.Enabled = True
        sslStatus.Text = "Ready to search and add products"
    End Sub

    'if the ID is selected or changed, the controls are enabled and the member name changes accordingly
    Private Sub cboMemberID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMemberID.SelectedIndexChanged
        cboName.SelectedIndex = cboMemberID.SelectedIndex
    End Sub

    'search a product by substring (not case-sensitive)
    Private Sub btnSearchProduct_Click(sender As Object, e As EventArgs) Handles btnSearchProduct.Click
        cboResults.Items.Clear()
        strSearch = txtSearch.Text.ToLower
        sqlReader = Products.GetSearchResults(strSearch)
        While sqlReader.Read
            cboResults.Items.Add(sqlReader.Item("PRODDESC"))
        End While
        sqlReader.Close()
        If cboResults.Items.Count = 0 Then
            sslStatus.Text = "No products found"
        Else
            sslStatus.Text = "Found " & cboResults.Items.Count & " products"
        End If
    End Sub

    'show which product was selected in the status bar
    Private Sub cboResults_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboResults.SelectedIndexChanged
        sslStatus.Text = cboResults.SelectedItem.ToString & " selected"
    End Sub

    'add or increment the currently selected result
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If cboResults.SelectedIndex > -1 Then
            sqlReader = Products.GetAddToCartItems(cboResults.SelectedItem.ToString)

            Dim blnProductMatch As Boolean = False

            For i = 0 To dgvCart.Rows.Count - 1
                If dgvCart.Rows(i).Cells(0).Value.ToString = cboResults.SelectedItem.ToString Then
                    blnProductMatch = True
                    Dim currQTY = CInt(dgvCart.Rows(i).Cells(2).Value)
                    Dim newQTY As Integer = currQTY + 1
                    dgvCart.Rows(i).Cells(2).Value = CStr(newQTY)
                    sslStatus.Text = "Added one more " & cboResults.SelectedItem.ToString & " to cart"
                    Exit For
                End If
            Next

            If blnProductMatch = False Then
                Dim price As Double = 0
                Dim taxable As String = ""
                Dim prodID As String = ""
                While sqlReader.Read
                    price = sqlReader.Item("RETPRICE")
                    taxable = sqlReader.Item("TAXABLE")
                    prodID = sqlReader.Item("PRODID")
                End While
                dgvCart.Rows.Add(cboResults.Text, price, 1, taxable, prodID)
                sslStatus.Text = "Added " & cboResults.SelectedItem.ToString & " to cart"
            End If

        Else
            sslStatus.Text = "Please select a search result first"
        End If
        sqlReader.Close()
        UpdateTotals()
    End Sub

    'remove or decrement the currently selected result
    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        If (dgvCart.Rows.Count > 0) And (dgvCart.SelectedRows.Count = 1) Then
            For Each item In dgvCart.SelectedRows
                sslStatus.Text = "Removed " & dgvCart.Rows(item.Index).Cells(0).Value.ToString & " from cart"
                dgvCart.Rows.RemoveAt(item.Index)
            Next
        Else
            sslStatus.Text = "There is nothing to remove"
        End If
        UpdateTotals()
    End Sub

    'increment the currently selected row
    Private Sub btnMore_Click(sender As Object, e As EventArgs) Handles btnMore.Click
        If (dgvCart.Rows.Count > 0) And (dgvCart.SelectedRows.Count = 1) Then
            For Each item In dgvCart.SelectedRows
                Dim currQTY = CInt(dgvCart.Rows(item.Index).Cells(2).Value)
                Dim newQTY As Integer = currQTY + 1
                dgvCart.Rows(item.Index).Cells(2).Value = CStr(newQTY)
                sslStatus.Text = "Added one more " & dgvCart.Rows(item.Index).Cells(0).Value.ToString & " to cart"
            Next
        Else
            sslStatus.Text = "Nothing selected"
        End If
        UpdateTotals()
    End Sub

    'decrement the currently selected row
    Private Sub btnLess_Click(sender As Object, e As EventArgs) Handles btnLess.Click
        If (dgvCart.Rows.Count > 0) And (dgvCart.SelectedRows.Count = 1) Then
            For Each item In dgvCart.SelectedRows
                If dgvCart.Rows(item.Index).Cells(2).Value = "1" Then
                    sslStatus.Text = "Removed " & dgvCart.Rows(item.Index).Cells(0).Value.ToString & " from cart"
                    dgvCart.Rows.RemoveAt(item.Index)
                Else
                    sslStatus.Text = "Removed one " & dgvCart.Rows(item.Index).Cells(0).Value.ToString & " from cart"
                    Dim currQTY = CInt(dgvCart.Rows(item.Index).Cells(2).Value)
                    Dim newQTY As Integer = currQTY - 1
                    dgvCart.Rows(item.Index).Cells(2).Value = CStr(newQTY)
                End If
            Next
        Else
            sslStatus.Text = "Nothing selected"
        End If
        UpdateTotals()
    End Sub

    'change the totals labels accordingly after updates
    Private Sub UpdateTotals()
        dblSubtotal = 0
        dblTotalTax = 0
        dblTotal = 0
        If dgvCart.Rows.Count > 0 Then
            For i = 0 To dgvCart.Rows.Count - 1
                dblSubtotal += (CDbl(dgvCart.Rows(i).Cells(1).Value)) * (CInt(dgvCart.Rows(i).Cells(2).Value))
                If CStr(dgvCart.Rows(i).Cells(3).Value) = "True" Then
                    dblTotalTax += ((CDbl(dgvCart.Rows(i).Cells(1).Value)) * (CDbl(dgvCart.Rows(i).Cells(2).Value))) * TAX
                End If
            Next
            dblTotal = dblSubtotal + dblTotalTax
            lblSubtotal.Text = FormatCurrency(dblSubtotal)
            lblTax.Text = FormatCurrency(dblTotalTax)
            lblTotal.Text = FormatCurrency(dblTotal)
        Else
            lblSubtotal.Text = "$0.00"
            lblTax.Text = "$0.00"
            lblTotal.Text = "$0.00"
            dblSubtotal = 0
            dblTotalTax = 0
            dblTotal = 0
        End If
    End Sub

    'clear the shopping cart and search areas
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        dgvCart.RowCount = 0
        cboResults.Items.Clear()
        txtSearch.Clear()
        UpdateTotals()
        sslStatus.Text = "Cart emptied"
    End Sub
End Class