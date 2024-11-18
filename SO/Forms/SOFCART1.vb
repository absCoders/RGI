Public Class SOFCART1

#Region "Declarations"
    Dim SHIP_BOL_NOs As New List(Of String)
    Dim ORDR_GROUP_NO As String
    Dim rowSOTSHIP0 As DataRow
    Dim rowSOTSHIP0_ORIG As DataRow
    Dim sqlSOTPICK1 As String
    Dim sqlSOTPICK2 As String
    Dim sqlSOTSHIPX As String
    Dim sqlSOTCART2 As String
    Dim SOTSHIP0 As String
    Dim SOTSHIPX As String
    Dim SOTCONF1 As String
    Dim SOTCONF2 As String
    ' Dim CART_NO_new As Int64 = 0
    Dim expSOTPICK1 As New Dictionary(Of String, String)

    Dim expirationDateMode As Boolean = False

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        expirationDateMode = ASCMAIN1.MENU_ITEM_OBJECT = "SOFCARTE"

        SOTSHIP0 = ASCMAIN1.Temp_Table("Select SHIP_BOL_NO from SOTSHIP1 where ROWNUM < 1")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIP0 & " Add Primary Key (SHIP_BOL_NO)")

        Set_Configs(" AND ROWNUM < 1")

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            sqlSOTSHIPX = "Select SOTSHIP1.*" _
                & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" _
                & " from SOTSHIP1,SOTORDR0" _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf
            sqlSOTSHIPX &= "" _
                & "   and SOTSHIP1.SHIP_STATUS = 'P'"
            ASCMAIN1.sql = sqlSOTSHIPX
            Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "", 1)

            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "", 1)
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP0", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTCART1", "*")

            sqlSOTCART2 = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
                & " from SOTCART2,SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO"
            ASCMAIN1.sql = sqlSOTCART2 & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0)

            ASCMAIN1.sql = "SELECT STYLE_CODE, STYLE_CODE COLOR_CODE, STYLE_DESC FROM ICTSTYL1"
            Create_TDA(.Tables.Add, "ICTSTYLE", ASCMAIN1.sql, 0, False, String.Empty, 0)

            With .Tables.Add("SOTCARTX")
                .Columns.Add("PICK_NO")
                .Columns.Add("ORDR_NO")
                .Columns.Add("ORDR_LNO", GetType(System.Int64))
                .Columns.Add("PICK_QTY_CONF", GetType(System.Int64), "")
                .Columns.Add("QTY_PACKED", GetType(System.Int64), "")
                .PrimaryKey = New DataColumn() {.Columns("PICK_NO"), .Columns("ORDR_NO"), .Columns("ORDR_LNO")}
            End With
            'CONFIG_NO was added to SOTPICK1 which is why doug had problem loading screen on email from 5/28/15
            'sqlSOTPICK1 = "Select SOTPICK1.*, SOTCONF1.CONFIG_NO" & vbCrLf _
            sqlSOTPICK1 = "Select SOTPICK1.*" & vbCrLf _
           & ", SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
           & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST" & vbCrLf _
           & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
           & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
           & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND, SOTCONF1.CONFIG_NO CONFIG_NO_NEW" & vbCrLf _
           & " from SOTPICK1,SOTORDR1,SOTSHIP1," & SOTCONF1 & " SOTCONF1" & vbCrLf _
           & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
           & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
           & "   and SOTCONF1.PICK_NO (+) = SOTPICK1.PICK_NO"
            ASCMAIN1.sql = sqlSOTPICK1 & " and ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICK1", "**")
            .Tables("SOTPICK1").Columns.Add("OK")
            .Tables("SOTPICK1").Columns("OK").DefaultValue = "0"

            Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")

            sqlSOTPICK2 = "Select SOTPICK2.*, ICTSTYL1.STYLE_WEIGHT" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC" & vbCrLf _
                & ", SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                & " from SOTPICK2,SOTPICK1,SOTORDR2,SOTSHIP1,ICTSTYL1" & vbCrLf _
                & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE"
            ASCMAIN1.sql = sqlSOTPICK2 & " and ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICK2", "**")

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            dst.Tables("SOTPICK1").Columns.Add("PICK_QTY", GetType(System.Int64))
            dst.Tables("SOTPICK2").Columns.Add("OK", GetType(System.String), "PARENT(SOTPICK1_SOTPICK2).OK")

            Create_Relation("SOTPICK2", "SOTCART2", "ORDR_NO,ORDR_LNO")
            dst.Tables("SOTPICK2").Columns.Add("QTY_PACKED", GetType(System.Int64), "SUM(CHILD(SOTPICK2_SOTCART2).QTY_PACKED)")
            dst.Tables("SOTCART2").Columns.Add("STYLE_WEIGHT", GetType(System.Decimal), "PARENT(SOTPICK2_SOTCART2).STYLE_WEIGHT")
            dst.Tables("SOTCART2").Columns.Add("STYLE_TOTAL_WEIGHT_CALC", GetType(System.Decimal), "ISNULL(QTY_PACKED,0) * ISNULL(STYLE_WEIGHT,0)")

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")

            Create_Relation("SOTCARTX", "SOTPICK2", "PICK_NO,ORDR_NO,ORDR_LNO")
            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_CALC", GetType(System.Int64))
            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_ORIG", GetType(System.Int64))
            dst.Tables("SOTCART1").Columns.Add("STYLE_TOTAL_WEIGHT_CALC", GetType(System.Int64), "SUM(CHILD.STYLE_TOTAL_WEIGHT_CALC)")

            Create_Relation("SOTCARTX", "SOTCART2", "PICK_NO,ORDR_NO,ORDR_LNO")

            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
            dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_WGT_CALC", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns.Add("PICK_CNT_CARTONS_CALC", GetType(System.Int64))
            dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_UNITS_CALC", GetType(System.Int64))

        End With

        Bind_Controls(UltraGroupBox1, "SOTSHIP1")

        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")
        grdICTSTYLE.DataSource = dst.Tables("ICTSTYLE")

        grdSOTSHIPX.DisplayLayout.UseFixedHeaders = True
        With grdSOTSHIPX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTPICK1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"OK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                If New String() {"PICK_NO", "CUST_STORE_NO"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ' gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
        End With

        With grdSOTPICK2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PICK_LNO", "STYLE_CODE", "STYLE_DESC", "PICK_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
        End With

        With grdSOTCART1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"CART_NO"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
        End With

        With grdSOTCART2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"ORDR_LNO", "STYLE_CODE", "QTY_PACKED"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"QTY_PACKED"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")

        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")
        Create_Summary(grdSOTPICK1, New String() {"PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT", "PICK_TOTAL_WGT_CALC", "PICK_CNT_CARTONS_CALC", "PICK_TOTAL_UNITS_CALC"})
        Create_Summary(grdSOTPICK1, New String() {"PICK_QTY"})

        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICK2, New String() _
            {"PICK_QTY"})

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART1, New String() _
            {"CART_FREIGHT", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL", "CART_TOTAL_UNITS_ORIG", "CART_TOTAL_UNITS_CALC"})

        Create_Summary(grdSOTCART2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTCART2, New String() _
            {"QTY_PACKED", "QTY_PACKED_ORIG"})

        Show_Filter(grdSOTSHIPX, True)
        grdSOTSHIPX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS", Nothing, New String() {":", "P:In Pick", "F:Shipped", "D:Deleted", "C:Cancelled"})

        dteExpire.MinDate = DateAdd(DateInterval.Month, -1, DateTime.Now)
        dteExpire.MaxDate = DateAdd(DateInterval.Year, 20, DateTime.Now)
        dteExpire.DateTime = DateTime.Now

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"

                ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIP0)
                SHIP_BOL_NOs.Clear()

                Dim SHIP_STATUS As String = ""

                If Absx1.txtFor("SHIP_BOL_NO").Text = "" Then
                    EMsg &= vbCr & "You must First Select a Shipment No"
                Else
                    Dim SHIP_BOL_NO As String = Absx1.txtFor("SHIP_BOL_NO").Text
                    Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                    If rowSOTSHIP1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Shipment No (" & SHIP_BOL_NO & ")"
                    Else
                        SHIP_STATUS = rowSOTSHIP1.Item("SHIP_STATUS")
                        ORDR_GROUP_NO = rowSOTSHIP1.Item("ORDR_GROUP_NO")
                        If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)

                        If optShipmentSelection.Value = "S" Then
                            If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
                            SHIP_BOL_NOs.Add(SHIP_BOL_NO)
                            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP0 & " (SHIP_BOL_NO) values ('" & SHIP_BOL_NO & "')")
                        Else
                            ASCMAIN1.sql = "Select SHIP_BOL_NO from SOTSHIP1 " _
                            & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                            & "   and SHIP_STATUS = '" & SHIP_STATUS & "'"
                            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                                Dim SHIP_BOL_NO2 As String = row.Item("SHIP_BOL_NO")
                                If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO2) Then Exit Sub
                                Me.SHIP_BOL_NOs.Add(SHIP_BOL_NO2)
                                ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP0 & " (SHIP_BOL_NO) values ('" & SHIP_BOL_NO2 & "')")
                            Next
                        End If
                    End If
                End If

                If EMsg = "" Then
                    For Each SHIP_BOL_NO As String In SHIP_BOL_NOs
                        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                        If Not InquiryMode Then
                            If rowSOTSHIP1.Item("SHIP_STATUS") <> SHIP_STATUS Then
                                EMsg &= vbCr & "Shipment Status Changed for Shipment " & SHIP_BOL_NO
                            End If
                            If rowSOTSHIP1.Item("SHIP_BOL_NO_REV") & "" <> "" Then
                                EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " is a Part of a Shipment/Invoice Reversal"
                            End If
                            If SHIP_STATUS = "P" Then
                            Else
                                EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " is No Longer in Pick"
                            End If
                        End If
                    Next
                End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Copy to Like Configs"
                Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
                Dim CONFIG_NO As String = grdSOTPICK1.ActiveRow.Cells("CONFIG_NO").Value & ""
                Dim sqlp As String = " and PICK_NO = '" & PICK_NO & "'"
                EMsg = Check_Carton_Pack_Integrity(sqlp)

                If EMsg = "" Then
                    Dim CONFIG_COUNT As Integer = Val(dst.Tables("SOTPICK1").Compute("COUNT(PICK_NO)", "CONFIG_NO = '" & CONFIG_NO & "'") & "")
                    If CONFIG_COUNT = 1 Then
                        EMsg &= vbCrLf & "There are no other Pick Tickets with the same Configuration"
                    Else
                        If MsgBox("This option will copy the current Carton Pack" _
                                & vbCrLf & " defined to the Pick Ticket " & PICK_NO _
                                & vbCrLf & " to the " & CStr(CONFIG_COUNT - 1) & " other Pick Tickets with the same" _
                                & vbCrLf & " Item/Qty Distribution Configuration (" & CONFIG_NO & ")" _
                                & vbCrLf & vbCrLf & "OK to Proceed?", _
                            MsgBoxStyle.YesNo, _
                            "Verification to Copy Carton Pack") = MsgBoxResult.No Then Exit Sub
                    End If
                End If

            Case "Update"

                If expirationDateMode Then
                    If MessageBox.Show("Do you want to Update modified Expiration Dates?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                Else
                    EMsg = Check_Carton_Pack_Integrity("")

                    If EMsg = "" Then
                        If MsgBox("This option will Replace all Carton Labels records" _
                                  & vbCrLf & " connected to this shipment with new Carton Label Control Numbers." _
                                  & vbCrLf & vbCrLf & "You MUST destroy all Carton Labels printed previously for this Shipment." _
                                  & vbCrLf & vbCrLf & "OK to Proceed?", _
                              MsgBoxStyle.YesNo, _
                              "Verification to Copy Carton Pack") = MsgBoxResult.No Then Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("Are you sure that you want to Cancel?", _
                      MsgBoxStyle.YesNo, _
                      "Verification to Cancel working with this Record") = MsgBoxResult.No Then
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
            Case "Select"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Done", "Cancel"
                Mode_Settings(False)

            Case "Add Carton"
                Add_Carton()

            Case "Copy to Like Configs"
                Copy_to_Like_Configs()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Select").Settings.Enabled = not_iScreenMode
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode

                .Items("Select").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
            End With

            .Groups("Special Operations").Visible = ScreenMode And (EntryMode <> "L") AndAlso Not InquiryMode AndAlso Not expirationDateMode
            .Groups("Shipment Selection").Visible = Not ScreenMode
            .Groups("Expiration Dates").Visible = ScreenMode AndAlso Not InquiryMode AndAlso dst.Tables("ICTSTYLE").Rows.Count > 0 AndAlso expirationDateMode
        End With

        tabSelect.Visible = Not tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CODE"), True)
        Set_Read_Only_for_ctl(Absx1.optFor("SHIP_ADDR_TYPE"), True)
        Set_Read_Only_for_ctl(Absx1.txtFor("SHIP_ADDR_CODE"), True)

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                {grdSOTPICK1, grdSOTPICK2, grdSOTCART1, grdSOTCART2}
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                End With
            Next

            If Not expirationDateMode Then
                grdSOTCART1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdSOTCART2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTPICK1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        Absx1.txtFor("PICK_NO").Text = ""
        ORDR_GROUP_NO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIP1", "SOTPICK1", "SOTPICK2", _
             "SOTCART1", "SOTCART2", "SOTSHIP0", "SOTCARTX", "ICTSTYLE"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_SOTSHIPX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ToggleDataTableExpressions(False)

        If EntryMode = "N" Then
        Else
            Dim sqlwhere_SOTSHIP1 As String = "" _
                & "   and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & ")" & vbCrLf _
                & "   and SOTSHIP1.SHIP_STATUS = 'P'"

            ASCMAIN1.sql = sqlSOTSHIPX & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)
            If dst.Tables("SOTSHIP1").Rows.Count <> SHIP_BOL_NOs.Count Then Stop ' NEED AN ABORT LOAD FEATURE IN STDS

            Set_Configs(sqlwhere_SOTSHIP1)

            ASCMAIN1.sql = sqlSOTPICK1 & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)


            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                Dim CONFIG_NO_NEW As String = rowSOTPICK1.Item("CONFIG_NO_NEW")
                rowSOTPICK1.Item("CONFIG_NO") = CONFIG_nO_NEW
            Next




                Dim row As DataRow = dst.Tables("SOTSHIP1").Rows(0)
            rowSOTSHIP0 = dst.Tables("SOTSHIP0").NewRow
            For i As Integer = 0 To dst.Tables("SOTSHIP0").Columns.Count - 1
                rowSOTSHIP0.Item(i) = row.Item(i)
            Next
            dst.Tables("SOTSHIP0").Rows.Add(rowSOTSHIP0)

            rowSOTSHIP0_ORIG = dst.Tables("SOTSHIP0").NewRow
            rowSOTSHIP0_ORIG.ItemArray = rowSOTSHIP0.ItemArray

            ASCMAIN1.sql = sqlSOTPICK2 & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
                & " from SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTCART1", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = sqlSOTCART2 & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTCART2", "", True, ASCMAIN1.sql)

            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

            If expirationDateMode AndAlso rowARTCUST1 IsNot Nothing AndAlso rowARTCUST1.Item("EDI_REQUIRES_EXP_DATE") & String.Empty = "1" Then
                ASCMAIN1.sql = "SELECT XX.STYLE_CODE, XX.COLOR_CODE, ICTSTYL1.STYLE_DESC" _
                    & " FROM ICTSTYL1, " _
                    & " (SELECT DISTINCT STYLE_CODE, COLOR_CODE FROM (" & sqlSOTCART2 & sqlwhere_SOTSHIP1 & ")) XX" _
                    & " WHERE XX.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                    & " AND ICTSTYL1.REQUIRES_EXP_DATE = '1'"
                Fill_Records("ICTSTYLE", "", True, ASCMAIN1.sql)
                Sort_grdColumns(grdICTSTYLE, "STYLE_CODE,COLOR_CODE")
            End If

            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                Dim CART_NO As String = rowSOTCART1.Item("CART_NO")
                rowSOTCART1.Item("CART_TOTAL_UNITS_ORIG") = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED_ORIG)", "CART_NO = '" & CART_NO & "'") & "")
            Next
        End If

        ToggleDataTableExpressions(True)

        dst.Tables("SOTPICK1").AcceptChanges()
        dst.Tables("SOTPICK2").AcceptChanges()

        Sort_grdColumns(grdSOTPICK1, "PICK_NO")
        Setup_SOTPICK1()

        ' CART_NO_new = 0

        Display_Totals()

        dst.Tables("SOTCARTX").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTPICK2"), New String() {"PICK_NO", "ORDR_NO", "ORDR_LNO"}).Rows
            dst.Tables("SOTCARTX").Rows.Add(New Object() {row.Item("PICK_NO"), row.Item("ORDR_NO"), row.Item("ORDR_LNO")})
        Next

        Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Updating ...")
            BeginTrans()

            If expirationDateMode Then
                Update_Record_TDA("SOTCART2")
                CommitTrans("Update Complete")
                Exit Sub
            End If

            'INIT_LAST("SOTSHIP1", False, , True)
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                ASCMAIN1.Progress("Updating Shipment " & SHIP_BOL_NO, "")
                Delete_Records(SHIP_BOL_NO)
            Next

            For Each TABLE_NAME As String In New String() _
                {"SOTCART1", "SOTCART2"}
                dst.Tables(TABLE_NAME).AcceptChanges()
                For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                    row.SetAdded()
                    If TABLE_NAME = "SOTCART1" Then
                        'Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                        'row.Item("CART_NO") = CART_NO
                        row.Item("CART_TOTAL_UNITS") = row.Item("CART_TOTAL_UNITS_CALC")
                        row.Item("CART_TOTAL_WGT_CALC") = row.Item("STYLE_TOTAL_WEIGHT_CALC")
                    End If
                Next
                Update_Record_TDA(TABLE_NAME)
            Next

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("ISNULL(QTY_PACKED,0)<>ISNULL(PICK_QTY_CONF,0)")
                rowSOTPICK2.Item("PICK_QTY_CONF") = rowSOTPICK2.Item("QTY_PACKED")
                rowSOTPICK2.Item("PICK_QTY_CANC") = Val(rowSOTPICK2.Item("PICK_QTY") & "") - Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")
            Next
            Update_Record_TDA("SOTPICK2")

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(PICK_CNT_CARTONS,0)<>ISNULL(PICK_CNT_CARTONS_CALC,0)")
                rowSOTPICK1.Item("PICK_CNT_CARTONS") = rowSOTPICK1.Item("PICK_CNT_CARTONS_CALC")
                rowSOTPICK1.Item("PICK_TOTAL_WGT") = rowSOTPICK1.Item("PICK_TOTAL_WGT_CALC")
            Next
            Update_Record_TDA("SOTPICK1")

            Dim SHIP_BOL_NOsz As String = Join(SHIP_BOL_NOs.ToArray, "','")

            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) " _
                     & " Select 'SOTORDR1', ORDR_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'CARTPK','Carton Pack Changes', NULL from SOTPICK1" _
                     & " where SHIP_BOL_NO in ('" & SHIP_BOL_NOsz & "')"
            ASCDATA1.ExecuteSQL()

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "SHIP_BOL_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTSHIP1.SHIP_STATUS = 'P' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTORDR0.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and SOTORDR0.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPICK1, "B", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTPICK2, "B", "Item Status Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key
            Case "Item Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTSHIPX()
                End If

            Case "SHIP_BOL_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("Select")
                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTSHIPX()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_SOTSHIPX()

            Case "SHIPMENT_NO"
                Click_Command("Select")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub

#End Region

    Sub Load_SOTSHIPX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If InquiryMode Then
            ASCMAIN1.sql = sqlSOTSHIPX _
                & IIf(CUST_CODE = "", "", " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'")
            ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
            grdSOTSHIPX.Text = "Shipments In Pick"

            If CUST_CODE <> "" Then grdSOTSHIPX.Text &= " associated with " & CUST_CODE
            Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)

            Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
        Else
            If CUST_CODE = "" Then
                Fill_Records("SOTSHIPX")
                grdSOTSHIPX.Text = "Unconfirmed Shipments"
                Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
            Else
                ASCMAIN1.sql = sqlSOTSHIPX _
                    & " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'"
                Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)
                grdSOTSHIPX.Text = "Unconfirmed Shipments associated with " & CUST_CODE
                Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
            End If
        End If

        grdSOTSHIPX.Visible = True
    End Sub

    Private Sub grdSOTSHIPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSHIPX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("SHIP_BOL_NO").Text = e.Row.Cells("SHIP_BOL_NO").Value
            Click_Command("Select")
        End If
    End Sub

    Sub Display_Totals()

    End Sub

#Region "grdSOTPICK1"

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowActivate
        Setup_SOTPICK1()
    End Sub

    Private Sub grdSOTPICK1_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICK1.BeforeRowActivate
        If grdSOTCART1.ActiveRow IsNot Nothing AndAlso grdSOTCART1.ActiveRow.DataChanged Then
            grdSOTCART1.ActiveRow.Update()
        End If
        If grdSOTPICK2.ActiveRow IsNot Nothing AndAlso grdSOTPICK2.ActiveRow.DataChanged Then
            grdSOTPICK2.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdSOTPICK1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICK1.InitializeRow
        
        If Val(e.Row.Cells("PICK_QTY").Value & "") <> Val(e.Row.Cells("PICK_TOTAL_UNITS_CALC").Value & "") Then
            e.Row.Cells("PICK_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("PICK_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If

    End Sub

#End Region

#Region "grdSOTPICK2"

    Sub SOTPICK1_Expressions(remove_expressions As Boolean)
        If remove_expressions Then
            expSOTPICK1.Clear()
            For Each fCOLUMN_NAME As String In New String() {"PICK_QTY", "PICK_TOTAL_WGT_CALC", "PICK_CNT_CARTONS_CALC", "PICK_TOTAL_UNITS_CALC"}
                expSOTPICK1.Add(fCOLUMN_NAME, dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression)
                dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression = ""
            Next
        Else
            For Each fCOLUMN_NAME As String In expSOTPICK1.Keys
                dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression = expSOTPICK1(fCOLUMN_NAME)
            Next
        End If
    End Sub

    Private Sub grdSOTPICK2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPICK2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "ADD_TO_CARTON"
                    Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value

                    If dst.Tables("SOTCART2").Select($"CART_NO = '{CART_NO}' and ORDR_NO = '{grdSOTPICK2.ActiveRow.Cells("ORDR_NO").Value}' and ORDR_LNO = '{grdSOTPICK2.ActiveRow.Cells("ORDR_LNO").Value}'").Length > 0 Then
                        MsgBox($"Line {grdSOTPICK2.ActiveRow.Cells("ORDR_LNO").Value} already in Carton", MsgBoxStyle.Critical, "Cannot Add Line")
                        Exit Sub
                    End If
                    If grdSOTCART1.ActiveRow IsNot Nothing And grdSOTPICK2.ActiveRow IsNot Nothing Then
                        Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow


                        rowSOTCART2.Item("CART_NO") = CART_NO
                        rowSOTCART2.Item("CART_LNO") = Val(dst.Tables("SOTCART2").Compute("MAX(CART_LNO)", "CART_NO = '" & CART_NO & "'") & "") + 1
                        rowSOTCART2.Item("ORDR_NO") = grdSOTPICK2.ActiveRow.Cells("ORDR_NO").Value
                        rowSOTCART2.Item("ORDR_LNO") = grdSOTPICK2.ActiveRow.Cells("ORDR_LNO").Value
                        Dim QTY_PACKED As Int64 = Val(grdSOTPICK2.ActiveRow.Cells("PICK_QTY").Value & "") - Val(grdSOTPICK2.ActiveRow.Cells("QTY_PACKED").Value & "")
                        If QTY_PACKED < 0 Then QTY_PACKED = 0
                        rowSOTCART2.Item("QTY_PACKED") = QTY_PACKED

                        Dim STYLE_CODE As String = grdSOTPICK2.ActiveRow.Cells("STYLE_CODE").Value
                        Dim COLOR_CODE As String = grdSOTPICK2.ActiveRow.Cells("COLOR_CODE").Value
                        Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", STYLE_CODE)

                        If rowICTSTYC1 IsNot Nothing Then
                            rowSOTCART2.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE") & String.Empty
                        Else
                            rowSOTCART2.Item("UPC_CODE") = DBNull.Value
                        End If
                        rowSOTCART2.Item("STYLE_CODE") = STYLE_CODE
                        rowSOTCART2.Item("COLOR_CODE") = COLOR_CODE
                        rowSOTCART2.Item("PICK_NO") = grdSOTCART1.ActiveRow.Cells("PICK_NO").Value

                        dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                    End If
            End Select
        End With

    End Sub

    Private Sub grdSOTPICK2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICK2.InitializeRow
        If Val(e.Row.Cells("PICK_QTY").Value & "") <> Val(e.Row.Cells("QTY_PACKED").Value & "") Then
            e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If

        e.Row.Cells("ADD_TO_CARTON").Value = "->"
    End Sub

#End Region

#Region "grdSOTCART1"

    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCART1.AfterRowActivate
        Setup_SOTCART2_from_SOTCART1()
    End Sub

    Private Sub grdSOTCART1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.AfterRowUpdate

    End Sub

    Private Sub grdSOTCART1_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.BeforeRowActivate
        If grdSOTCART2.ActiveRow IsNot Nothing AndAlso grdSOTCART2.ActiveRow.DataChanged Then
            grdSOTCART2.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdSOTCART1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART1.InitializeRow
        If Val(e.Row.Cells("CART_TOTAL_UNITS_CALC").Value & "") <> Val(e.Row.Cells("CART_TOTAL_UNITS_ORIG").Value & "") Then
            ' e.Row.Cells("CART_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            ' e.Row.Cells("CART_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

#End Region

    Private Sub grdSOTCART2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART2.InitializeRow
        If Val(e.Row.Cells("QTY_PACKED").Value & "") <> Val(e.Row.Cells("QTY_PACKED_ORIG").Value & "") Then
            ' e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            ' e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Setup_SOTPICK1()
        If grdSOTPICK1.ActiveRow Is Nothing Then
            grdSOTPICK2.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
            Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
            Dim CUST_STORE_NO As String = grdSOTPICK1.ActiveRow.Cells("CUST_STORE_NO").Value
            Dim dvw As DataView = DirectCast(grdSOTPICK2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
            grdSOTPICK2.Text = "Item Details for Pick Ticket No " & PICK_NO & ", Store " & CUST_STORE_NO
            Sort_grdColumns(grdSOTPICK2, "PICK_LNO")

            dvw = DirectCast(grdSOTCART1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
            grdSOTCART1.Text = "Cartons for Pick Ticket No " & PICK_NO
            Sort_grdColumns(grdSOTCART1, "CART_NO")
            Setup_SOTCART2_from_SOTCART1()
            grdSOTPICK2.Visible = True
        End If
    End Sub

    Sub Setup_SOTCART2_from_SOTCART1()
        If grdSOTCART1.ActiveRow Is Nothing Then
            grdSOTCART2.Visible = False
        Else
            Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value
            Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "CART_NO = '" & CART_NO & "'"
            grdSOTCART2.Text = "Contents of Carton " & CART_NO
            Sort_grdColumns(grdSOTCART2, "ORDR_LNO")
            grdSOTCART2.Visible = True
        End If
    End Sub

    Sub Setup_SOTCART2_from_SOTPICK2()
        If grdSOTPICK2.ActiveRow Is Nothing Then
            grdSOTCART2.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK2.ActiveRow.Cells("PICK_NO").Value
            Dim PICK_LNO As Int32 = Val(grdSOTPICK2.ActiveRow.Cells("PICK_LNO").Value & "")
            Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "' and ORDR_LNO = " & CStr(PICK_LNO)
            grdSOTCART2.Text = "Cartons Contents for Pick Ticket " & PICK_NO & ", Line " & CStr(PICK_LNO)
            grdSOTCART2.Visible = True
        End If
    End Sub

    Sub Delete_Records(SHIP_BOL_NO As String)

        Dependent_Updates(-1, SHIP_BOL_NO)

        Dim sqlw As String = "where CART_NO in (" _
            & " Select CART_NO from SOTCART1 where PICK_NO in (" _
            & " Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'))"
        ASCDATA1.ExecuteSQL("Delete from SOTCART2 " & sqlw)
        ASCDATA1.ExecuteSQL("Delete from SOTCART1 " & sqlw)

    End Sub

    Sub Dependent_Updates(S As Integer, SHIP_BOL_NO As String)

    End Sub

    Sub ToggleDataTableExpressions(ByVal tf As Boolean)

        With dst.Tables("SOTPICK2")
        End With

        With dst.Tables("SOTCARTX")
            .Columns("QTY_PACKED").Expression = IIf(Not tf, "", "SUM(CHILD(SOTCARTX_SOTCART2).QTY_PACKED)")
        End With

        With dst.Tables("SOTCART1")
            .Columns("CART_TOTAL_UNITS_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTCART1_SOTCART2).QTY_PACKED)")
        End With

        With dst.Tables("SOTPICK1")
            .Columns("PICK_TOTAL_WGT_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_ACTUAL)")
            .Columns("PICK_CNT_CARTONS_CALC").Expression = IIf(Not tf, "", "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
            .Columns("PICK_TOTAL_UNITS_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_UNITS_CALC)")
            .Columns("PICK_QTY").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY)")
        End With

    End Sub

    Sub Add_Carton()
        If grdSOTPICK1.ActiveRow IsNot Nothing Then
            Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
            'CART_NO_new += 1
            Dim CART_NO As String = String.Empty
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                CART_NO = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
            Else
                CART_NO = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
            End If

            rowSOTCART1.Item("CART_NO") = CART_NO ' "NEW" & Format(CART_NO_new, "0000000")
            rowSOTCART1.Item("PICK_NO") = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
            dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)
        End If
    End Sub

    Sub Set_Configs(sqlwhere_SOTSHIP1 As String)

        Dim C As String = "SOTPICK2.PICK_QTY"

        ASCMAIN1.sql = "Select SOTPICK2.PICK_NO, SOTORDR2.STYLE_CODE" & vbCrLf _
            & ", SUM (" & C & ") QTY" & vbCrLf _
            & " from SOTPICK1,SOTPICK2,SOTSHIP1,SOTORDR2" & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO " & vbCrLf _
            & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & sqlwhere_SOTSHIP1 & vbCrLf _
            & " group by SOTPICK2.PICK_NO, SOTORDR2.STYLE_CODE"
        If SOTCONF2 = "" Then
            SOTCONF2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTCONF2 & " Add Primary Key (PICK_NO, STYLE_CODE)")

            ASCMAIN1.sql = "Select Distinct PICK_NO from " & SOTCONF2
            SOTCONF1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTCONF1 & " Add Primary Key (PICK_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTCONF1 & " Add CONFIG_NO VARCHAR2(6)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTCONF2)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCONF2 & " (PICK_NO, STYLE_CODE, QTY) " & ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTCONF1)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCONF1 & " (PICK_NO) Select Distinct PICK_NO from " & SOTCONF2)
        End If

        Dim CONFIG_NO As Integer = 0
        Dim CONFIG_NOs As New Dictionary(Of Integer, String)
        ASCMAIN1.sql = "Select Distinct PICK_NO from " & SOTCONF2
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim CONFIG_NO_for_this_PICK_NO As Integer = 0
            Dim PICK_NO As String = row.Item("PICK_NO")
            If CONFIG_NO > 0 Then
                For I As Integer = 1 To CONFIG_NO
                    ASCMAIN1.sql = "" _
                        & "(" _
                        & "Select STYLE_CODE, QTY from " & SOTCONF2 _
                        & " where PICK_NO = '" & PICK_NO & "'" _
                        & " minus " _
                        & "Select STYLE_CODE, QTY from " & SOTCONF2 _
                        & " where PICK_NO = '" & CONFIG_NOs(I) & "'" _
                        & ") union (" _
                        & "Select STYLE_CODE, QTY from " & SOTCONF2 _
                        & " where PICK_NO = '" & CONFIG_NOs(I) & "'" _
                        & " minus " _
                        & "Select STYLE_CODE, QTY from " & SOTCONF2 _
                        & " where PICK_NO = '" & PICK_NO & "'" _
                        & ")"
                    ASCMAIN1.sql = "Select Count (*) from (" & ASCMAIN1.sql & ")"
                    Dim RECORDS_DIFFERENCE As Integer = Val(ASCDATA1.GetDataValue & "")
                    If RECORDS_DIFFERENCE = 0 Then
                        CONFIG_NO_for_this_PICK_NO = I
                        Exit For
                    End If
                Next
            End If
            If CONFIG_NO_for_this_PICK_NO = 0 Then
                CONFIG_NO += 1
                CONFIG_NO_for_this_PICK_NO = CONFIG_NO
                CONFIG_NOs.Add(CONFIG_NO, PICK_NO)
            End If
            ASCMAIN1.sql = "Update " & SOTCONF1 _
                & " Set CONFIG_NO = '" & Format(CONFIG_NO_for_this_PICK_NO, "000000") & "'" _
                & " where PICK_NO = '" & PICK_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next

    End Sub

    Function Check_Carton_Pack_Integrity(sqlp As String) As String
        Dim EMsg As String = ""

        Dim rows() As DataRow = Nothing

        ' EMPTY CARTONS MAY BE REQUIRED FOR SEARS
        rows = dst.Tables("SOTCART1").Select("ISNULL(CART_TOTAL_UNITS_CALC,0) = 0" & sqlp)
        If rows.Length <> 0 Then
            EMsg &= vbCr & "Some Cartons are Empty - You must Delete Empty Cartons"
            EMsg &= vbCr & " (See Pick Ticket No " & rows(0).Item("PICK_NO") & ")"
        End If

        ' ZERO QTY IN CARTON DETAILS MAY BE REQUIRED FOR SEARS
        rows = dst.Tables("SOTCART2").Select("ISNULL(QTY_PACKED,0) = 0" & sqlp)
        If rows.Length <> 0 Then
            EMsg &= vbCr & "Some Cartons have Details with 0 Qty - You must Delete Details with 0 Qty Packed"
            EMsg &= vbCr & " (See Pick Ticket No " & rows(0).Item("PICK_NO") & ", Carton " & rows(0).Item("CART_NO") & ")"
        End If

        rows = dst.Tables("SOTPICK2").Select("OK <> '1' and ISNULL(PICK_QTY,0) <> ISNULL(QTY_PACKED,0)" & sqlp)
        If rows.Length <> 0 Then
            EMsg &= vbCr & "Some Items are not in Balance with Carton Packing Changes"
            EMsg &= vbCr & " (See Pick Ticket No " & rows(0).Item("PICK_NO") & ", Line No " & rows(0).Item("ORDR_LNO") & ")"
        End If

        rows = dst.Tables("SOTPICK1").Select("OK <> '1' and ISNULL(PICK_QTY,0) <> ISNULL(PICK_TOTAL_UNITS_CALC,0)" & sqlp)
        If rows.Length <> 0 Then
            EMsg &= vbCr & "Some Pick Tickets are not in Balance with Carton Packing Changes"
            EMsg &= vbCr & " (See Pick Ticket No " & rows(0).Item("PICK_NO") & ")"
        End If

        rows = dst.Tables("SOTPICK2").Select("ISNULL(QTY_PACKED,0)>ISNULL(PICK_QTY,0)") 
        If rows.Length <> 0 Then
            EMsg &= vbCr & "Some Pick Ticket Details are Packed with a Qty > Qty Released"
            EMsg &= vbCr & " (See Pick Ticket No " & rows(0).Item("PICK_NO") & ")"
        End If

        Return EMsg
    End Function

    Sub Copy_to_Like_Configs()
        Dim PICK_NO_gold As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
        Dim CONFIG_NO As String = grdSOTPICK1.ActiveRow.Cells("CONFIG_NO").Value
        Dim rowSOTPICK1_gold As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO_gold)

        EnforceConstraints(False)
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("CONFIG_NO = '" & CONFIG_NO & "' and PICK_NO <> '" & PICK_NO_gold & "'")
            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
            ASCDATA1.DeleteRows(dst.Tables("SOTCART1"), "PICK_NO = '" & PICK_NO & "'")

            For Each rowSOTCART1_gold As DataRow In rowSOTPICK1_gold.GetChildRows("SOTPICK1_SOTCART1")

                Dim CART_NO As String = String.Empty
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    CART_NO = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                Else
                    CART_NO = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                End If

                'Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
                rowSOTCART1.ItemArray = rowSOTCART1_gold.ItemArray
                rowSOTCART1.Item("CART_NO") = CART_NO
                rowSOTCART1.Item("PICK_NO") = PICK_NO
                dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)
                For Each rowSOTCART2_gold As DataRow In rowSOTCART1_gold.GetChildRows("SOTCART1_SOTCART2")
                    Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                    rowSOTCART2.ItemArray = rowSOTCART2_gold.ItemArray
                    rowSOTCART2.Item("CART_NO") = CART_NO
                    rowSOTCART2.Item("PICK_NO") = PICK_NO
                    rowSOTCART2.Item("ORDR_NO") = rowSOTPICK1.Item("ORDR_NO")
                    'rowSOTCART2.Item("ORDR_LNO") = 0
                    dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                Next
            Next
        Next
        EnforceConstraints(True)
    End Sub

    Private Sub btnSetExpireDate_Click(sender As System.Object, e As System.EventArgs) Handles btnSetExpireDate.Click
        If dst.Tables("ICTSTYLE").Rows.Count = 0 Then
            Exit Sub
        End If

        If grdICTSTYLE.Selected.Rows.Count = 0 Then
            MessageBox.Show("You must select a Style / Color.", "Expire Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim sqlSOTCART1 As String = String.Empty
        Dim sqlSOTCART2 As String = String.Empty
        Dim userMessage As String = "Do you want to Update the Expiration Date for " _
                                    & grdICTSTYLE.Selected.Rows(0).Cells("STYLE_CODE").Value & "/" _
                                    & grdICTSTYLE.Selected.Rows(0).Cells("COLOR_CODE").Value & " " _
                                    & grdICTSTYLE.Selected.Rows(0).Cells("STYLE_DESC").Value

        Select Case optExpire.Value

            Case "S"
                userMessage &= " for the shipment"

            Case "P"
                If grdSOTPICK1.Selected.Rows.Count = 0 Then
                    MessageBox.Show("You must select a Pick Ticket.", "Expire Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTPICK1.Selected.Rows
                    sqlSOTCART1 &= " or PICK_NO = '" & grdRow.Cells("PICK_NO").Value & "'"
                Next
                sqlSOTCART1 = sqlSOTCART1.Substring(3).Trim
                userMessage &= " for the selected pick ticket"

            Case "C"
                If grdSOTCART1.Selected.Rows.Count = 0 Then
                    MessageBox.Show("You must select a Carton.", "Expire Date", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTCART1.Selected.Rows
                    sqlSOTCART1 &= " or CART_NO = '" & grdRow.Cells("CART_NO").Value & "'"
                Next
                sqlSOTCART1 = sqlSOTCART1.Substring(3).Trim
                userMessage &= " for the selected carton"
        End Select

        userMessage &= " using date: " & dteExpire.DateTime.ToString("MM/dd/yyyy") & "?"

        If MessageBox.Show(userMessage, "Expire Date", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim STYLE_CODE As String = grdICTSTYLE.Selected.Rows(0).Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grdICTSTYLE.Selected.Rows(0).Cells("COLOR_CODE").Value

        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select(sqlSOTCART1)
            sqlSOTCART2 = "CART_NO = '" & rowSOTCART1.Item("CART_NO") & "' AND STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"
            For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select(sqlSOTCART2)
                rowSOTCART2.Item("ITEM_EXP_DATE") = dteExpire.DateTime.ToShortDateString
                Exit For
            Next
        Next

    End Sub
End Class