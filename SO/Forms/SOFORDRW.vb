Public Class SOFORDRW

    Private tblErrors As DataTable

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst

            ASCMAIN1.sql = "SELECT * FROM SOTORDRX WHERE (ORDR_NO, ORDR_LNO) IN (SELECT ORDR_NO, MIN(ORDR_LNO) ORDR_LNO FROM SOTORDRX WHERE ORDR_STATUS <> 'O' GROUP BY ORDR_NO) "
            Create_TDA(.Tables.Add, "SOTORDRXH", ASCMAIN1.sql, 0, False, String.Empty, 0)

            ASCMAIN1.sql = "SELECT * FROM SOTORDRX WHERE ORDR_STATUS <> 'O'"
            Create_TDA(.Tables.Add, "SOTORDRXA", ASCMAIN1.sql, 0, False, String.Empty, 0)
            .Tables("SOTORDRXA").Columns.Add("ERROR_CODES", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTORDRX", "*", 2)
            .Tables("SOTORDRX").Columns.Add("ERROR_CODES", GetType(System.String))

            ASCMAIN1.sql = "SELECT * FROM SOTORDRX WHERE ORDR_SOURCE = :PARM1 AND ORDR_NO = :PARM2 AND ORDR_LNO = :PARM3 "
            Create_TDA(.Tables.Add, "SOTORDRX1", ASCMAIN1.sql, 0, False, "VVN")

            Create_TDA(.Tables.Add, "ASTTASK1", "*")
            Create_TDA(.Tables.Add, "ASTTASK2", "*")

            .Relations.Add("ASTTASK1_ASTTASK2", .Tables("ASTTASK1").Columns("TASK_NO"), .Tables("ASTTASK2").Columns("TASK_NO"))

        End With

        grdSOTORDRXH.DataSource = dst.Tables("SOTORDRXH")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdASTTASK1.DataSource = dst.Tables("ASTTASK1")

        Create_Summary(grdSOTORDRXH, "ORDR_NO", "Count")

        Create_Summary(grdSOTORDRX, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDRX, "ORDR_QTY", "Sum")

        tblErrors = New DataTable
        tblErrors.Columns.Add("SEL", GetType(System.String))
        tblErrors.Columns.Add("CODE", GetType(System.String))
        tblErrors.Columns.Add("DESC", GetType(System.String))

        tblErrors.Rows.Add(New Object() {"0", "A", "Sold To"})
        tblErrors.Rows.Add(New Object() {"0", "B", "Ship To"})
        tblErrors.Rows.Add(New Object() {"0", "C", "Routing Code"})
        tblErrors.Rows.Add(New Object() {"0", "D", "Item Code"})
        tblErrors.Rows.Add(New Object() {"0", "E", "Pricing"})
        tblErrors.Rows.Add(New Object() {"0", "F", "Order Quantity"})
        tblErrors.Rows.Add(New Object() {"0", "G", "Sales Tax"})
        tblErrors.Rows.Add(New Object() {"0", "H", "Freight"})

        grdErrors.DataSource = tblErrors

        tabOrders.Parent = tab.Parent
        spl.Parent = tab.Parent
        tab.Visible = False

        TABLE_NAME = "SOTORDRX1"

        dte0.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        dte1.MaxDate = dte0.MaxDate

        dte0.MinDate = DateAdd(DateInterval.Year, -1, DateTime.Now)
        dte1.MinDate = dte0.MinDate

    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFORDRI")

        'With UltraExplorerBar1.Groups("Screen Control")
        '    .Items("New").Visible = Not InquiryMode
        '    .Items("Edit").Visible = Not InquiryMode
        '    .Items("Update").Visible = Not InquiryMode
        '    .Items("Cancel").Visible = Not InquiryMode
        '    .Items("Delete").Visible = Not InquiryMode
        'End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

            Case "Edit", "Load"
                If grdSOTORDRXH.Rows.Count = 0 Then
                    EMsg &= vbCr & "There are no sales order import errors to correct"
                ElseIf grdSOTORDRXH.ActiveRow Is Nothing Then
                    EMsg &= vbCr & "Please select a sales order from the provided list"
                End If

            Case "Update"
                RecordHasErrors(Nothing, Nothing)
                If tblErrors.Select("SEL = '1'").Length > 0 Then
                    Dim zMsg As String = "The sales order still has errors. The order will not be processed until all errors are cleared."
                    zMsg &= Environment.NewLine & Environment.NewLine
                    zMsg &= "Update anyway?"
                    If MessageBox.Show(zMsg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                    dst.Tables("SOTORDRX1").Rows(0).Item("ORDR_STATUS") = "H"
                Else
                    dst.Tables("SOTORDRX1").Rows(0).Item("ORDR_STATUS") = "O"
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"

            Case "Cancel Order"
                Dim zMsg As String = "The sales order will be permanently removed for the import table."
                zMsg &= Environment.NewLine & Environment.NewLine
                zMsg &= "Cancel Order?"
                If MessageBox.Show(zMsg, "Cancel Order", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Edit", "View"
                If eItemKey = "View" Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Call Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Cancel Order"
                Cancel_Order()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")

                If EntryMode = "V" And Not InquiryMode Then
                    .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If

                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Cancel Order").Settings.Enabled = iScreenMode
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode

                .Items("Edit").Visible = (Not InquiryMode)
                .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")
                .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")
                .Items("Cancel Order").Visible = (EntryMode = "E")

            End With
        End With

        tabOrders.Visible = Not tf
        spl.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTORDRX", "SOTORDRX1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        lblDuplicate.Visible = False

        Fill_Records("SOTORDRXH")
        Fill_Records("SOTORDRXA")

        btnFetch_Click(Nothing, Nothing)

        Sort_grdColumns(grdErrors, "CODE")
        Sort_grdColumns(grdASTTASK1, "TASK_NO")
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)
        EnforceConstraints(False)

        Try
            ASCMAIN1.Progress("Now Loading Data")
            Me.Cursor = Cursors.WaitCursor

            Dim ORDR_SOURCE As String = grdSOTORDRXH.Selected.Rows(0).Cells("ORDR_SOURCE").Value
            Dim ORDR_NO As String = grdSOTORDRXH.Selected.Rows(0).Cells("ORDR_NO").Value
            Dim ORDR_LNO As String = grdSOTORDRXH.Selected.Rows(0).Cells("ORDR_LNO").Value

            Fill_Records("SOTORDRX", New Object() {ORDR_SOURCE, ORDR_NO})
            Fill_Records("SOTORDRX1", New Object() {ORDR_SOURCE, ORDR_NO, ORDR_LNO})

            lblDuplicate.Visible = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_SOURCE = '" & ORDR_SOURCE & "' AND ORDR_NO_WEB = '" & ORDR_NO & "'") IsNot Nothing

            RecordHasErrors(Nothing, Nothing)
            grdSOTORDRX.Refresh()

            grdSOTORDRX.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.False

            If EntryMode = "E" Then
                Set_Read_Only(spl, False)
                Set_Read_Only(grpSoldTo, ORDR_SOURCE = "W")
                Set_Read_Only(grpShipTo, ORDR_SOURCE = "W")
                Set_Read_Only_for_ctl(txtRoutingCode, False)
            Else
                Set_Read_Only(spl, True)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty)
        End Try

        EnforceConstraints(False)
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()

            ASCMAIN1.Progress("Updating Data")
            Me.Cursor = Cursors.WaitCursor

            ' Update the fields the user can modify to all records
            Dim fields As String = " CUST_CODE,CUST_STORE_NO,ORDR_DATE,ORDR_DPD,SHIP_VIA_CODE,"
            fields &= "CUST_STORE_NAME,CUST_STORE_ADDR1,CUST_STORE_ADDR2,CUST_STORE_ADDR3,"
            fields &= "CUST_STORE_CITY,CUST_STORE_STATE,CUST_STORE_ZIP_CODE,CUST_STORE_PHONE,"
            fields &= "CUST_STORE_COUNTRY,CUST_STORE_FAX,CUST_STORE_EMAIL,CUST_NAME,"
            fields &= "CUST_ADDR1,CUST_ADDR2,CUST_ADDR3,CUST_CITY,CUST_STATE,CUST_ZIP_CODE,"
            fields &= "CUST_PHONE,CUST_COUNTRY,CUST_FAX,CUST_EMAIL,ORDR_STATUS,ROUTING_CODE,ORDR_STATUS"

            For Each row As DataRow In dst.Tables("SOTORDRX").Rows
                For Each fieldName As String In fields.Split(",")
                    fieldName = fieldName.Trim
                    row.Item(fieldName) = dst.Tables("SOTORDRX1").Rows(0).Item(fieldName)
                Next
            Next

            Update_Record_TDA("SOTORDRX")

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)

        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty)
        End Try

    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTORDRXH_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDRXH.DoubleClickRow
        If grdSOTORDRXH.ActiveRow Is Nothing Then Exit Sub
        Click_Command("View")
    End Sub

    Private Sub grdSOTORDRX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDRX.AfterRowActivate
        If EntryMode = "E" AndAlso (grdSOTORDRX.ActiveRow.Cells("ERROR_CODES").Value & String.Empty).ToString.Contains("D") Then
            grdSOTORDRX.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.True
        Else
            grdSOTORDRX.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.False
        End If
    End Sub

    Private Sub grdSOTORDRX_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDRX.ClickCellButton

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                grdClickCellButton(grdSOTORDRX, String.Empty, False, "ITEM_CODE")
        End Select

    End Sub

    Private Sub grdSOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow

        If e.Row.Cells("ERROR_CODES").Value.ToString.Contains("D") Then
            e.Row.Cells("ITEM_CODE").Appearance.BackColor = Drawing.Color.Red
        Else
            e.Row.Cells("ITEM_CODE").Appearance.BackColor = Drawing.Color.White
        End If

        If e.Row.Cells("ERROR_CODES").Value.ToString.Contains("E") Then
            e.Row.Cells("ORDR_UNIT_PRICE").Appearance.BackColor = Drawing.Color.Red
        Else
            e.Row.Cells("ORDR_UNIT_PRICE").Appearance.BackColor = Drawing.Color.White
        End If

        If e.Row.Cells("ERROR_CODES").Value.ToString.Contains("F") Then
            e.Row.Cells("ORDR_QTY").Appearance.BackColor = Drawing.Color.Red
        Else
            e.Row.Cells("ORDR_QTY").Appearance.BackColor = Drawing.Color.White
        End If
    End Sub

    Private Sub btnFetch_Click(sender As System.Object, e As System.EventArgs) Handles btnFetch.Click

        Dim startDate As String = dte0.DateTime.ToString("dd-MMM-yyyy")
        Dim endDate As String = dte1.DateTime.ToString("dd-MMM-yyyy")

        Dim sql As String = "Select * from ASTTASK1 WHERE TRUNC(START_TIME) between '" & startDate & "' and '" & endDate & "'"
        Fill_Records("ASTTASK1", String.Empty, True, sql)

        sql = "SELECT * FROM ASTTASK2 WHERE TASK_NO IN (Select TASK_NO from ASTTASK1 WHERE TRUNC(START_TIME) between '" & startDate & "' and '" & endDate & "')"
        Fill_Records("ASTTASK2", String.Empty, True, sql)

    End Sub

    Private Sub tabOrders_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabOrders.SelectedTabChanged
        With UltraExplorerBar1
            .Groups("Errors").Visible = tabOrders.ActiveTab.Index = 0
        End With
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub RecordHasErrors(sender As Object, E As System.EventArgs) Handles txtCustCode.ValueChanged, txtCustStoreNo.ValueChanged, txtRoutingCode.ValueChanged, grdSOTORDRX.AfterRowUpdate, grdSOTORDRXH.AfterRowActivate

        If SELECTION_NO = 0 Then Exit Sub

        Dim CUST_CODE As String = String.Empty
        Dim CUST_STORE_NO As String = String.Empty
        Dim ROUTING_CODE As String = String.Empty

        Dim ORDR_FREIGHT As Decimal = Val(MyBase.Absx1.numFor("ORDR_FREIGHT").Value & String.Empty)
        Dim ORDR_STAX As Decimal = Val(MyBase.Absx1.numFor("ORDR_STAX").Value & String.Empty)

        Dim tbl As DataTable = dst.Tables("SOTORDRX")

        Try
            Dim evalCodes As String = String.Empty

            If sender Is Nothing Then
                ' Force global check
                evalCodes = "ABCD"
                CUST_CODE = MyBase.Absx1.txtFor("CUST_CODE").Text
                CUST_STORE_NO = MyBase.Absx1.txtFor("CUST_STORE_NO").Text
                ROUTING_CODE = MyBase.Absx1.txtFor("ROUTING_CODE").Text
            Else
                Select Case sender.name
                    Case txtCustCode.Name
                        CUST_CODE = MyBase.Absx1.txtFor("CUST_CODE").Text
                        evalCodes = "A"
                    Case txtCustStoreNo.Name
                        CUST_CODE = MyBase.Absx1.txtFor("CUST_CODE").Text
                        CUST_STORE_NO = MyBase.Absx1.txtFor("CUST_STORE_NO").Text
                        evalCodes = "B"
                    Case txtRoutingCode.Name
                        ROUTING_CODE = MyBase.Absx1.txtFor("ROUTING_CODE").Text
                        evalCodes = "C"
                    Case grdSOTORDRX.Name
                        evalCodes = "D"
                    Case grdSOTORDRXH.Name
                        evalCodes = "ABCD"
                        CUST_CODE = grdSOTORDRXH.ActiveRow.Cells("CUST_CODE").Value & String.Empty
                        CUST_STORE_NO = grdSOTORDRXH.ActiveRow.Cells("CUST_STORE_NO").Value & String.Empty
                        ROUTING_CODE = grdSOTORDRXH.ActiveRow.Cells("ROUTING_CODE").Value & String.Empty

                        ORDR_FREIGHT = Val(grdSOTORDRXH.ActiveRow.Cells("ORDR_FREIGHT").Value & String.Empty)
                        ORDR_STAX = Val(grdSOTORDRXH.ActiveRow.Cells("ORDR_STAX").Value & String.Empty)

                        Dim ORDR_SOURCE As String = grdSOTORDRXH.ActiveRow.Cells("ORDR_SOURCE").Value & String.Empty
                        Dim ORDR_NO As String = grdSOTORDRXH.ActiveRow.Cells("ORDR_NO").Value & String.Empty

                        tbl = dst.Tables("SOTORDRXA").Select("ORDR_SOURCE = '" & ORDR_SOURCE & "' AND ORDR_NO = '" & ORDR_NO & "'").CopyToDataTable
                    Case Else
                        Exit Sub
                End Select
            End If

            ' These are done in both Modes, althought currently they cannot chnage these values
            If ORDR_STAX < 0 Then
                tblErrors.Select("CODE = 'G'")(0).Item("SEL") = "1"
            Else
                tblErrors.Select("CODE = 'G'")(0).Item("SEL") = "0"
            End If

            If ORDR_FREIGHT < 0 Then
                tblErrors.Select("CODE = 'H'")(0).Item("SEL") = "1"
            Else
                tblErrors.Select("CODE = 'H'")(0).Item("SEL") = "0"
            End If

            ' Evaluate Sold To, Ship To, Routing Code
            If evalCodes.Contains("A") Then
                If LookUp("ARTCUST1", CUST_CODE) Is Nothing Then
                    tblErrors.Select("CODE = 'A'")(0).Item("SEL") = "1"
                Else
                    tblErrors.Select("CODE = 'A'")(0).Item("SEL") = "0"
                End If
            End If

            If evalCodes.Contains("B") Then
                If LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO}) Is Nothing Then
                    tblErrors.Select("CODE = 'B'")(0).Item("SEL") = "1"
                Else
                    tblErrors.Select("CODE = 'B'")(0).Item("SEL") = "0"
                End If
            End If

            If evalCodes.Contains("C") Then
                If LookUp("SOTROUT1", ROUTING_CODE) Is Nothing Then
                    tblErrors.Select("CODE = 'C'")(0).Item("SEL") = "1"
                Else
                    tblErrors.Select("CODE = 'C'")(0).Item("SEL") = "0"
                End If
            End If

            If evalCodes.Contains("D") Then
                tblErrors.Select("CODE = 'D'")(0).Item("SEL") = "0"
                tblErrors.Select("CODE = 'E'")(0).Item("SEL") = "0"
                tblErrors.Select("CODE = 'F'")(0).Item("SEL") = "0"

                For Each row As DataRow In tbl.Select
                    Dim ERROR_CODES As String = String.Empty

                    If LookUp("ICTITEM1", row.Item("ITEM_CODE")) Is Nothing Then
                        tblErrors.Select("CODE = 'D'")(0).Item("SEL") = "1"
                        ERROR_CODES &= "D"
                        row.Item("ITEM_DESC") = String.Empty
                    Else
                        row.Item("ITEM_DESC") = cdr.Item("ITEM_DESC") & String.Empty
                    End If

                    If Val(row.Item("ORDR_UNIT_PRICE") & String.Empty) < 0 Then
                        tblErrors.Select("CODE = 'E'")(0).Item("SEL") = "1"
                        ERROR_CODES &= "E"
                    End If

                    If Val(row.Item("ORDR_QTY") & String.Empty) < 0 Then
                        tblErrors.Select("CODE = 'F'")(0).Item("SEL") = "1"
                        ERROR_CODES &= "F"
                    End If

                    row.Item("ERROR_CODES") = ERROR_CODES
                Next
            End If

        Catch ex As Exception
            MessageBox.Show("Error evaluating Data. Do not Update changes. Error as follows: " & ex.Message, "Error", MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub Cancel_Order()

        Try
            BeginTrans()
            Dim ORDR_SOURCE As String = grdSOTORDRXH.Selected.Rows(0).Cells("ORDR_SOURCE").Value
            Dim ORDR_NO As String = grdSOTORDRXH.Selected.Rows(0).Cells("ORDR_NO").Value

            ASCDATA1.ExecuteSQL("Delete from SOTORDRX WHERE ORDR_SOURCE = '" & ORDR_SOURCE & "' AND ORDR_NO = '" & ORDR_NO & "'")
            CommitTrans("Sales Order Cancelled")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region

End Class