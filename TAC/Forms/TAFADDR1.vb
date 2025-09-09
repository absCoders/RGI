Public Class TAFADDR1

    Public frmAddressMatches As List(Of WHCSHIP1.AddressMatchDetail)

    Public Sub New(ByRef AddressMatches As List(Of WHCSHIP1.AddressMatchDetail))

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        frmAddressMatches = AddressMatches
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim tblAddresses As DataTable = Nothing
        Dim rowAddresses As DataRow

        tblAddresses = New DataTable

        tblAddresses.Columns.Add("AddressIndex", GetType(System.Int32))
        tblAddresses.Columns.Add("Name", GetType(System.String))
        tblAddresses.Columns.Add("Company", GetType(System.String))
        tblAddresses.Columns.Add("AddressLine1", GetType(System.String))
        tblAddresses.Columns.Add("AddressLine2", GetType(System.String))
        tblAddresses.Columns.Add("City", GetType(System.String))
        tblAddresses.Columns.Add("State", GetType(System.String))
        tblAddresses.Columns.Add("PostalCode", GetType(System.String))
        'tblAddresses.Columns.Add("PostalCodeExtended", GetType(System.String))
        tblAddresses.Columns.Add("CountryCode", GetType(System.String))

        For ictr As Integer = 0 To frmAddressMatches.Count - 1
            Dim address As TAC.WHCSHIP1.AddressMatchDetail = frmAddressMatches(ictr)
            address.isSelected = False
            rowAddresses = tblAddresses.NewRow
            rowAddresses.Item("AddressIndex") = ictr
            'rowAddresses.Item("Name") = address.Name
            rowAddresses.Item("Company") = address.Company
            rowAddresses.Item("AddressLine1") = address.Address1
            rowAddresses.Item("AddressLine2") = address.Address2
            rowAddresses.Item("City") = address.City
            rowAddresses.Item("State") = address.State
            rowAddresses.Item("PostalCode") = address.ZipCode
            'rowAddresses.Item("PostalCodeExtended") = address.Zip4
            rowAddresses.Item("CountryCode") = address.Country
            tblAddresses.Rows.Add(rowAddresses)
        Next

        grdADDRESS.DataSource = tblAddresses
        grdADDRESS.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click

        If Me.grdADDRESS.Selected.Rows.Count = 0 Then
            MessageBox.Show("You must select an address or click 'Cancel'.", "Select", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim addressIndex As Integer = Me.grdADDRESS.Selected.Rows(0).Cells("AddressIndex").Value
        Dim addr As TAC.WHCSHIP1.AddressMatchDetail = frmAddressMatches(addressIndex)
        addr.isSelected = True
        'frmAddressMatches(addressIndex) = addr

        Me.Close()
    End Sub

    Private Sub grdADDRESS_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdADDRESS.DoubleClickRow
        If grdADDRESS.Selected.Rows.Count = 1 Then
            Me.cmdSelect_Click(Nothing, Nothing)
        End If
    End Sub

End Class