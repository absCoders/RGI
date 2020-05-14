Public Class SOTCARR1

    'alter table sotcarr1 add SHIP_3PY_COUNTRY varchar2(3);
    'alter table sotcarr1 add SHIP_3PY_ZIPCODE varchar2(15);
    'alter table sotcarr1 add SHIP_ACCT_NO varchar2(9);

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "SOTCARR2", "*", 1)
            Create_TDA(.Tables.Add, "SOTCARR3", "*", 1)
            .Tables("SOTCARR3").Columns.Add("CUST_NAME", GetType(System.String))
            Create_TDA(.Tables.Add, "SOTCARR4", "*", 1)
            Create_TDA(.Tables.Add, "SOTCARR5", "*", 1)
        End With

        Create_Relation("SOTCARR3", "SOTCARR5", "CARRIER_CODE,DIVISION_CODE,CARRIER_ACCOUNT_NO")

        grdSOTCARR2.DataSource = dst.Tables("SOTCARR2")
        grdSOTCARR3.DataSource = dst.Tables("SOTCARR3")
        grdSOTCARR4.DataSource = dst.Tables("SOTCARR4")

        ASCMAIN1.Add_Value_List(grdSOTCARR2, "SERVICE_CODE", Nothing, New String() {":", "D:Domestic", "I:International"}, 0)
        ASCMAIN1.Add_Value_List(grdSOTCARR2, "TRACKING_ID_TYPE", Nothing, New String() {":", "0:Fedex Express", "1:Fedex Ground", "2:USPS", "3:N/A"}, 0)

        Create_Lookup("ARTCUST1")
        Create_Lookup("GLTACCT1")
        Create_Lookup("GLTSEGM1")

        grdSOTCARR3.DisplayLayout.Bands(1).Columns("ACCOUNT_PHONE").MaskInput = String.Empty
        grdSOTCARR3.DisplayLayout.Bands(1).Columns("ACCOUNT_PHONE").CellDisplayStyle = UltraWinGrid.CellDisplayStyle.Default

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"

                grdSOTCARR2.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdSOTCARR3.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdSOTCARR4.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""
        Update_Record_TDA("SOTCARR2", "CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
        Update_Record_TDA("SOTCARR3", "CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
        Update_Record_TDA("SOTCARR4", "CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
        Update_Record_TDA("SOTCARR5", "CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()


        EnforceConstraints(False)
        Fill_Records("SOTCARR2", New String() {Absx1.txtFor("CARRIER_CODE").Text})

        ASCMAIN1.sql = "SELECT SOTCARR3.*, ARTCUST1.CUST_NAME FROM SOTCARR3, ARTCUST1 WHERE SOTCARR3.SHIPPER_DIVISION_CODE = ARTCUST1.CUST_CODE (+) AND CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'"
        Fill_Records("SOTCARR3", String.Empty, True, ASCMAIN1.sql)

        Fill_Records("SOTCARR4", New String() {Absx1.txtFor("CARRIER_CODE").Text})
        Fill_Records("SOTCARR5", New String() {Absx1.txtFor("CARRIER_CODE").Text})
        EnforceConstraints(True)

        grdSOTCARR2.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTCARR3.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTCARR4.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        ASCMAIN1.Add_Value_List(grdSOTCARR3, "CARRIER_PROD_CODE", "SELECT CARRIER_PROD_CODE, CARRIER_PROD_DESC FROM SOTCARR2 WHERE CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")


        If EntryMode = "New" Then
            optCARRIER_PPA_TYPE.Value = "L"
            optCARRIER_SURCHARGE_BASE.Value = "L"
        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            dst.EnforceConstraints = False
            dst.Tables("SOTCARR2").Rows.Clear()
            dst.Tables("SOTCARR3").Rows.Clear()
            dst.Tables("SOTCARR4").Rows.Clear()
            dst.EnforceConstraints = False
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        tabOther.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
    End Sub

#End Region

    Private Sub grdSOTCARR4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCARR4.BeforeRowUpdate
        e.Row.Cells("CARRIER_CODE").Value = Absx1.txtFor("CARRIER_CODE").Text

        e.Row.Cells("PACKAGE_CODE").Value = (e.Row.Cells("PACKAGE_CODE").Value & String.Empty).ToString.Trim
        e.Row.Cells("PACKAGE_DESC").Value = (e.Row.Cells("PACKAGE_DESC").Value & String.Empty).ToString.Trim

        If e.Row.Cells("PACKAGE_CODE").Value.ToString.Length = 0 OrElse _
            e.Row.Cells("PACKAGE_DESC").Value.ToString.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Package Code and Description are required.")
        End If
    End Sub

    Private Sub grdSOTCARR3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCARR3.BeforeRowUpdate

        Select Case e.Row.Band.Key
            Case grdSOTCARR3.DisplayLayout.Bands(0).Key
                e.Row.Cells("CARRIER_CODE").Value = Absx1.txtFor("CARRIER_CODE").Text

                If e.Row.Cells("DIVISION_CODE").Value & String.Empty = String.Empty Then
                    e.Row.Cells("DIVISION_CODE").Value = ASCMAIN1.CLIENT
                End If

                Dim CUST_CODE As String = e.Row.Cells("SHIPPER_DIVISION_CODE").Value & String.Empty
                CUST_CODE = CUST_CODE.Trim
                If CUST_CODE.Length > 0 Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        MessageBox.Show("Invalid entry for Customer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        e.Cancel = True
                        Exit Sub
                    Else
                        e.Row.Cells("CUST_NAME").Value = rowARTCUST1.Item("CUST_NAME") & String.Empty
                    End If
                End If
            Case grdSOTCARR3.DisplayLayout.Bands(1).Key
                e.Row.Cells("CARRIER_CODE").Value = e.Row.ParentRow.Cells("CARRIER_CODE").Value & String.Empty
                e.Row.Cells("DIVISION_CODE").Value = e.Row.ParentRow.Cells("DIVISION_CODE").Value & String.Empty
                e.Row.Cells("CARRIER_ACCOUNT_NO").Value = e.Row.ParentRow.Cells("CARRIER_ACCOUNT_NO").Value & String.Empty

                ' Minimum Requirements
                If e.Row.Cells("ACCOUNT_NAME").Value & String.Empty = String.Empty _
                    OrElse e.Row.Cells("ACCOUNT_PHONE").Value & String.Empty = String.Empty _
                    OrElse e.Row.Cells("ACCOUNT_ADDR1").Value & String.Empty = String.Empty _
                    OrElse e.Row.Cells("ACCOUNT_CITY").Value & String.Empty = String.Empty _
                    OrElse e.Row.Cells("ACCOUNT_STATE").Value & String.Empty = String.Empty _
                    OrElse e.Row.Cells("ACCOUNT_ZIP_CODE").Value & String.Empty = String.Empty Then

                    MessageBox.Show("Account requires Name, Address Line 1, City, State, Zip and Phone.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    e.Cancel = True
                    Exit Sub
                End If
        End Select


        'Dim ACCT_CODE As String = e.Row.Cells("ACCT_CODE").Value.ToString.Trim
        'Dim SEG2_CODE As String = e.Row.Cells("SEG2_CODE").Value.ToString.Trim
        'Dim SEG3_CODE As String = e.Row.Cells("SEG3_CODE").Value.ToString.Trim
        'Dim SEG4_CODE As String = e.Row.Cells("SEG4_CODE").Value.ToString.Trim

        'If ACCT_CODE.Length > 0 Then
        '    If LookUp("GLTACCT1", ACCT_CODE) Is Nothing Then
        '        MessageBox.Show("Invalid entry for account Code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        e.Cancel = True
        '        Exit Sub
        '    End If
        'End If

        'If ACCT_CODE.Length > 0 Then
        '    If LookUp("GLTSEGM1", New String() {SEG2_CODE, "2"}) Is Nothing Then
        '        MessageBox.Show("Invalid entry for Seg 2 Code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        e.Cancel = True
        '        Exit Sub
        '    End If

        '    If LookUp("GLTSEGM1", New String() {SEG3_CODE, "3"}) Is Nothing Then
        '        MessageBox.Show("Invalid entry for Seg 3 Code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        e.Cancel = True
        '        Exit Sub
        '    End If

        '    If LookUp("GLTSEGM1", New String() {SEG4_CODE, "4"}) Is Nothing Then
        '        MessageBox.Show("Invalid entry for Seg 4 Code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        e.Cancel = True
        '        Exit Sub
        '    End If

        'ElseIf SEG2_CODE.Length > 0 OrElse SEG3_CODE.Length > 0 OrElse SEG4_CODE.Length > 0 Then
        '    If MessageBox.Show("There are account segment values with no Account Code. Do you want to clear the segment codes?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
        '        e.Cancel = True
        '        Exit Sub
        '    End If
        '    e.Row.Cells("SEG2_CODE").Value = String.Empty
        '    e.Row.Cells("SEG3_CODE").Value = String.Empty
        '    e.Row.Cells("SEG4_CODE").Value = String.Empty
        'End If
    End Sub

    Private Sub grdSOTCARR2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCARR2.BeforeRowUpdate
        e.Row.Cells("CARRIER_CODE").Value = MyBase.Absx1.txtFor("CARRIER_CODE").Text

        If (e.Row.Cells("CARRIER_PROD_DESC").Value).ToString.Trim.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Product Description is required.", "Update Error", MessageBoxButtons.OK)
            Exit Sub
        End If
    End Sub

    Private Sub grdSOTCARR3_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSOTCARR3.ClickCellButton
        With e.Cell.Row
            Dim sql_where As String = ""

            Select Case e.Cell.Column.Key
                Case "CUST_CODE"
                    grdClickCellButton(grdSOTCARR3, sql_where)
            End Select
        End With
    End Sub
End Class