Imports System.Drawing
Imports System.Drawing.Printing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFPACK1
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow
    Dim ShipLoc As String = ""
    Dim C As WHC.WHCRF000

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "WHFPACKI" Then
            InquiryMode = True
        End If

        Get_PARM("SOTPARM1")
        With dst

            ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
                & ", SOTORDR0.CUST_CODE, nvl(SOTORDR5.CUST_NAME, ARTCUST1.CUST_NAME) CUST_NAME " & vbCrLf _
                & ", SOTSHIP1.ORDR_GROUP_NO, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTSHIP1.SHIP_STATUS" & vbCrLf _
                & ", bol.BILL_OF_LADING_NO, GRP_CNT, STARTED, LAST, PACK_FRST, PACK_LAST" & vbCrLf _
                & " from SOTPICK1,SOTORDR0,ARTCUST1,SOTSHIP1,SOTORDR5," & vbCrLf _
                & " (SELECT BILL_OF_LADING_NO, COUNT(1) GRP_CNT from SOTSHIP1 group by BILL_OF_LADING_NO) bol, " & vbCrLf _
                & " (SELECT PICK_NO, MIN(INIT_DATE) STARTED, MAX(INIT_DATE) LAST FROM SOTPICK5 GROUP BY PICK_NO) pick5, " & vbCrLf _
                & " (SELECT PICK_NO, MIN(CART_PACKED) PACK_FRST, MAX(CART_PACKED) PACK_LAST from WHTCART1 GROUP BY PICK_NO) cart1 " & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and pick5.PICK_NO(+) = SOTPICK1.PICK_NO " & vbCrLf _
                & "   and cart1.PICK_NO(+) = SOTPICK1.PICK_NO " & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
                & "   and SOTORDR5.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and bol.BILL_OF_LADING_NO(+) = SOTSHIP1.BILL_OF_LADING_NO" & vbCrLf _
                & "   and SOTORDR5.CUST_ADDR_TYPE = 'ST'" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and SOTORDR0.ORDR_TYPE_CODE <> 'B2C'" & vbCrLf _
                & "   and SOTSHIP1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and NVL(SOTPICK1.PACK_STATUS,'0') = :PARM2"
            Create_TDA(.Tables.Add, "SOTPACKX", "**", 0, False, "VV", 1)

            ASCMAIN1.sql = "Select ICTWHSE1.WHSE_CODE, ICTWHSE1.WHSE_DESC" & vbCrLf _
                & ", X.SHIPS" & vbCrLf _
                & "  from ICTWHSE1" & vbCrLf _
                & ", (Select SOTSHIP1.WHSE_CODE, Count (*) SHIPS" & vbCrLf _
                & "  from SOTSHIP1" & vbCrLf _
                & " where SOTSHIP1.SHIP_STATUS IN ('P', 'H')" & vbCrLf _
                & " group by SOTSHIP1.WHSE_CODE) X" & vbCrLf _
                & " where X.WHSE_CODE (+) = ICTWHSE1.WHSE_CODE" & vbCrLf _
                & "   and X.SHIPS <> 0"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False)
            .Tables("ICTWHSEX").Columns("SHIPS").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select * from SOTPICK1 where SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE,  ICTSTYC1.UPC_CODE" _
                & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" _
                & ", SUM (SOTPICK2.PICK_QTY_CONF) PICK_QTY_CONF" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC) PICK_QTY_CANC" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK) PICK_QTY_BACK" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK_REL) PICK_QTY_BACK_REL" _
                & ", MAX (SOTORDR2.STYLE_DESC) STYLE_DESC" _
                & ", MAX(ICTCOLR1.COLOR_DESC) COLOR_DESC" _
                & " from SOTPICK2,SOTORDR2,SOTPICK1,ICTCOLR1,ICTSTYC1 " _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
                & " and ICTSTYC1.STYLE_CODE = SOTORDR2.STYLE_CODE and ICTSTYC1.COLOR_CODE = SOTORDR2.COLOR_CODE" _
                & " and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO and SOTPICK1.SHIP_BOL_NO = :PARM1" _
                & " and SOTPICK1.ORDR_NO = :PARM2" _
                & " and ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYC1.UPC_CODE"
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False, "VV", 0)
            With .Tables("SOTPICKX")
                .Columns.Add("PICKED_QTY", GetType(System.Int64))
                .Columns.Add("PACKED_QTY", GetType(System.Int64))
                .Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PICKED_QTY,0)-ISNULL(PACKED_QTY,0)")
                .Columns.Add("SHORTAGE", GetType(System.Int64), "ISNULL(PICK_QTY,0)-ISNULL(PICKED_QTY,0)")
                .Columns.Add("LOCATION_CODE")
            End With

            ASCMAIN1.sql = "Select * from SOTPICK5 where PICK_NO = :PARM1 AND PICK_STATUS = 'P'"
            Create_TDA(.Tables.Add, "SOTPICK5", "**", 0, True, "V")

            Create_TDA(dst.Tables.Add, "TATCNTRY", "*", 0, False)

            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select * from SOTORDR5 where ORDR_NO = :PARM1 and CUST_ADDR_TYPE = 'ST'"
            Create_TDA(.Tables.Add, "SOTORDR5", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from SOTORDR0 where ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTSHIP1.*, SREP_NAME from SOTSHIP1, SOTSREP1 where SOTSREP1.SREP_CODE = SOTSHIP1.SREP_CODE AND SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select * from ICTWHSE1 where WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from ARTCUST1 where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from SOTSVIA1 where SHIP_VIA_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTSVIA1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from WHTMOVE1 where WHSE_TRAN_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTMOVE1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select * from WHTMOVE2 where WHSE_TRAN_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTMOVE2", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "SELECT INIT_OPER, ENTITY_TYPE, APP_DESC " & vbCrLf _
                & "FROM ASTMTSK2, WHTGUNA1  " & vbCrLf _
                & "WHERE PROCEDURE_NAME = MENU_ITEM_OBJECT " & vbCrLf _
                & "AND ENTITY_TYPE LIKE '99-G%-A' " & vbCrLf _
                & "AND (PROCEDURE_NAME, APP_ID) IN ( " & vbCrLf _
                & "SELECT PROCEDURE_NAME, MIN(APP_ID) APP_ID FROM WHTGUNA1 " & vbCrLf _
                & "GROUP BY PROCEDURE_NAME)"
            Create_TDA(.Tables.Add, "WHTGUNAS", "**", 0, False)

            ASCMAIN1.sql = ""
            Create_TDA(.Tables.Add, "SOTCART1", "*")
            Create_TDA(.Tables.Add, "SOTCART2", "*")
            Create_TDA(.Tables.Add, "ASTATTA2", "*")
            Create_TDA(.Tables.Add, "WHTPICKS", "*")

            ASCMAIN1.sql = "Select WHTCART1.*" & vbCrLf _
                & " from WHTCART1 " & vbCrLf _
                & " where WHTCART1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTCART1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select WHTCART2.*, STYLE_DESC, ICTCOLR1.COLOR_DESC COLOR_CODE_LONG, WHTCART1.CART_SEQ" & vbCrLf _
                & " from WHTCART2,WHTCART1,SOTORDR2, ICTCOLR1" & vbCrLf _
                & " where WHTCART2.CART_NO = WHTCART1.CART_NO" & vbCrLf _
                & " and SOTORDR2.ORDR_NO = WHTCART2.ORDR_NO and SOTORDR2.ORDR_LNO = WHTCART2.ORDR_LNO" & vbCrLf _
                & " and ICTCOLR1.COLOR_CODE = WHTCART2.COLOR_CODE" & vbCrLf _
                & "   and WHTCART1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTCART2", "**", 0, True, "V", 2)

            Create_Relation("WHTCART1", "WHTCART2", "CART_NO")

            Create_TDA(.Tables.Add, "WHTPKGM1", "*")
            Fill_Records("WHTPKGM1", String.Empty, True, "SELECT * FROM WHTPKGM1")


            ASCMAIN1.sql = "SELECT ORDR_NO, STYLE_CODE, UPC_CODE, STYLE_DESC, COLOR_DESC, TTL_CTNS, TTL_UNITS, QTY_PACKED
                                FROM
                                (
                                SELECT ORDR_NO, STYLE_CODE, STYLE_DESC, COLOR_DESC, UPC_CODE, QTY_PACKED,
                                COUNT(*) TTL_CTNS, 
                                SUM(QTY_PACKED) TTL_UNITS
                                FROM
                                (
                                SELECT SOTPICK1.ORDR_NO, SOTCART2.STYLE_CODE, SOTCART2.UPC_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, SOTCART2.QTY_PACKED
                                FROM SOTCART1, SOTCART2, ICTSTYL1, ICTCOLR1, SOTPICK1
                                WHERE SOTCART1.CART_NO = SOTCART2.CART_NO
                                AND SOTCART2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)
                                AND SOTCART2.COLOR_CODE = ICTCOLR1.COLOR_CODE (+)
                                AND SOTCART1.PICK_NO = SOTPICK1.PICK_NO
                                AND SOTPICK1.ORDR_NO = :PARM1
                                AND SOTPICK1.PICK_STATUS = 'P'
                                )
                                GROUP BY ORDR_NO, STYLE_CODE, STYLE_DESC, COLOR_DESC, UPC_CODE, QTY_PACKED
                                )"
            Create_TDA(.Tables.Add, "WHTPACKC", ASCMAIN1.sql, 0, False, "V", 0)

            Dim rows() As DataRow = ASCDATA1.GetDataTable("SELECT *  FROM WHTLPRT1").Select("")
            For Each row As DataRow In rows
                cbxLabelPrinter.Items.Add(row.Item("LABEL_PRINTER_ID"))
            Next
            cbxLabelPrinter.SelectedIndex = 0

            Dim settings As New PrinterSettings
            For Each printer As String In PrinterSettings.InstalledPrinters
                If printer.ToLower.Contains("zebra") Or printer.ToLower.Contains("upc") Or printer.ToLower.Contains("microsoft") _
                    Or printer.ToLower.Contains("brother ql") Or printer.ToLower.Contains("pdf") Then
                Else
                    settings.PrinterName = printer
                    Debug.Print(printer)
                    If settings.DefaultPageSettings.PaperSize.PaperName = "Letter" Then
                        cbxReportPrinter.Items.Add(printer)
                        If settings.IsDefaultPrinter Then
                            cbxReportPrinter.SelectedIndex = cbxReportPrinter.Items.IndexOf(printer)
                        End If
                    End If
                End If
            Next

        End With

        With ultraComboPackage.DisplayLayout.Bands(0)

            ultraComboPackage.Font = grdWHTCART1.Font
            ultraComboPackage.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Default
            ultraComboPackage.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDownList

            .Columns.Add("PKG_CODE")
            .Columns("PKG_CODE").Header.Caption = "Code"
            .Columns("PKG_CODE").Width = 75

            .Columns.Add("PKG_DESC")
            .Columns("PKG_DESC").Header.Caption = "Desc"
            .Columns("PKG_DESC").Width = 75

            .Columns.Add("PKG_D")
            .Columns("PKG_D").Header.Caption = "L x W x H"
            .Columns("PKG_D").Width = 200

        End With

        ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1")
        ultraComboPackage.ValueMember = "PKG_CODE"
        ultraComboPackage.DisplayMember = "PKG_DESC"
        grdWHTCART1.DisplayLayout.Bands(0).Columns("PKG_CODE").EditorComponent = ultraComboPackage

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdSOTPACKX.DataSource = dst.Tables("SOTPACKX")

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")

        grdWHTCART1.DataSource = dst.Tables("WHTCART1")
        ASCMAIN1.Add_Value_List(grdWHTCART1, "PROCESS_STATUS", , New String() {":", "0:Open", "1:Closed", "2:Printed"})

        grdWHTCART2.DataSource = dst.Tables("WHTCART2")

        grdWHTGUNAS.DataSource = dst.Tables("WHTGUNAS")

        Fill_Records("ICTWHSEX")
        Fill_Records("TATCNTRY")

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")
        Create_Summary(grdICTWHSEX, "SHIPS")

        Create_Summary(grdSOTPACKX, "PICK_NO", "Count")

        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")

        Create_Summary(grdSOTPICKX, "STYLE_CODE", "Count")
        Create_Summary(grdSOTPICKX, New String() _
                       {"PICK_QTY", "PICK_QTY_CONF", "PICKED_QTY", "PACKED_QTY" _
                       , "VARIANCE", "SHORTAGE"})

        Sort_grdColumns(grdWHTCART2, "CART_SEQ")

        'grdSOTPACKX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        'For Each C As UltraWinGrid.UltraGridColumn In grdSOTPACKX.DisplayLayout.Bands(0).Columns
        '    If C.Key = "SEL" Then
        '        C.CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Else
        '        C.CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'Next

        With grdSOTPACKX.DisplayLayout.Bands("SOTPACKX")
            .Columns("PICK_NO").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
            For Each COLUMN_NAME As String In New String() {"INIT_DATE", "STARTED", "LAST", "PACK_FRST", "PACK_LAST"}
                .Columns(COLUMN_NAME).Format = "MM/dd/yy HH:mm"
            Next
        End With
        With grdSOTPICKX.DisplayLayout.Bands("SOTPICKX")
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdWHTCART1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"CART_SEQ", "PKG_CODE", "CART_TOTAL_WGT_ACTUAL"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With
        Show_Filter(grdSOTPACKX)


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            'Case "View"
            '    If Absx1.txtFor("WHSE_CODE").Text = "" Then
            '        EMsg &= vbCrLf & "You must specify a Warehouse"
            '    Else
            '        rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
            '        If rowICTWHSE1 Is Nothing Then
            '            EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
            '        End If
            '    End If
            '    WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")

            Case "Load", "View"

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")
                    If Not ASCMAIN1.Logical_Open("WHTPACK1", WHSE_CODE) Then Exit Sub

                End If

            Case "Print"
                If cbxReportPrinter.SelectedItem Is Nothing Then
                    EMsg &= vbCrLf & "Select Report Printer"
                End If

            Case "Print Cons by P.O."
                If cbxReportPrinter.SelectedItem Is Nothing Then
                    EMsg &= vbCrLf & "Select Report Printer"
                End If

            Case "Update"
                If cbxReportPrinter.SelectedItem Is Nothing Then
                    EMsg &= vbCrLf & "Select Report Printer"
                End If
                If dst.Tables("WHTCART1").Compute("Count(PROCESS_STATUS)", "PROCESS_STATUS = 0") Then
                    EMsg &= vbCrLf & "Found Open Cartons"
                End If
                If EMsg = "" And dst.Tables("SOTPICKX").Compute("SUM(VARIANCE)", "") <> 0 Then
                    If MsgBox("Select 'Yes' to force close this shipment", vbYesNoCancel, "Not in Balance") <> vbYes Then
                        EMsg &= vbCrLf & "Shipment out of balance"
                    End If
                End If
                If EMsg = "" Then
                    If (MsgBox("Continue with Close and Print?", MsgBoxStyle.YesNo, grdSOTPACKX.ActiveRow.Cells("CUST_NAME").Value & "") <> MsgBoxResult.Yes) Then
                        Exit Sub
                    End If
                    chkAutoRefresh.Checked = False
                    chkAutoRefresh_CheckedChanged(New Object, New EventArgs)
                    Dim SHIP_BOL_NO As String = grdSOTPACKX.ActiveRow.Cells("SHIP_BOL_NO").Value & ""
                    Dim PICK_NO As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO").Value & ""
                    Dim ORDR_GROUP_NO As String = grdSOTPACKX.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""

                    If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOTRDR0", ORDR_GROUP_NO) Then Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            'Case "View"
            '    EntryMode = "V"
            '    Load_Record()
            '    Mode_Settings(True)

            Case "Load", "View"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                pnlCloseNPrint.Enabled = False
                Update_Record()
                Print_Packing_List()
                ASCMAIN1.MultiTask_Release()
                Mode_Settings(False)
                Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
                Click_Command("Load")

            Case "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Packing_List()

            Case "Print Cons by P.O."
                Print_Packing_List_Consolidated()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    If EntryMode = "V" And Not InquiryMode Then
                        .Items("Load").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Load").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Update").Visible = False

                    If Not tf Then
                        .Items("Print").Visible = False
                        .Items("Print Cons by P.O.").Visible = False
                    End If

                    .Items("Load").Visible = Not InquiryMode
                End With

                .Groups("Packing Filters").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        splShipments.Visible = ScreenMode

        With grdWHTCART1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            If EntryMode = "L" Then
                .AllowUpdate = DefaultableBoolean.True
            Else
                .AllowUpdate = DefaultableBoolean.False
            End If
        End With

        If ScreenMode Then
            splPicks.Panel2Collapsed = True
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        chkAutoRefresh.Checked = False

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTPACKX", "SOTPICK1", "SOTPICKX", "ICTWHSEX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        optPickFilter.Value = "P"
        txtPICK_CNT_PALLETS.Value = ""
        txtPICK_TOTAL_WGT.Value = ""
        txtFLOOR_LOC.Value = ""

        Fill_Records("ICTWHSEX")
        Sort_grdColumns(grdICTWHSEX, "WHSE_CODE")
        Setup_tab0()


    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        'Load_SOTPACKX()
        optPending_ValueChanged(optPickFilter, New EventArgs)
        Fill_Records("ICTWHSE1", WHSE_CODE)

        ShipLoc = ASCDATA1.GetDataValue("select WHSE_LOC_SHP from ICTWHSE1 where whse_code = '" & WHSE_CODE & "'")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Update")
        Dim SHIP_BOL_NO As String = grdSOTPACKX.ActiveRow.Cells("SHIP_BOL_NO").Value & ""
        Dim PICK_NO As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO").Value & ""
        Dim ORDR_NO As String = grdSOTPACKX.ActiveRow.Cells("ORDR_NO").Value & ""
        Dim ORDR_GROUP_NO As String = grdSOTPACKX.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
        Dim PICK_NO_CONS As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO_CONS").Value & ""
        Dim rowSOTCART1 As DataRow
        Dim rowSOTCART2 As DataRow
        Dim rowWHTCART1 As DataRow
        Dim rowWHTCART2 As DataRow
        Dim PACK_PACKERS As String = ""

        dst.Tables("SOTCART1").Rows.Clear()
        dst.Tables("SOTCART2").Rows.Clear()
        dst.Tables("SOTPICK1").Rows.Clear()
        dst.Tables("SOTSHIP1").Rows.Clear()
        BeginTrans()

        Fill_Records("SOTPICK1", PICK_NO, True, "Select * from SOTPICK1 where PICK_NO = :PARM1")
        Fill_Records("SOTSHIP1", SHIP_BOL_NO, True)

        ASCMAIN1.sql = "DELETE SOTCART2 WHERE CART_NO in (SELECT CART_NO FROM SOTCART1 where PICK_NO = :PARM1)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_NO)

        ASCMAIN1.sql = "DELETE SOTCART1 WHERE PICK_NO = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_NO)

        If dst.Tables("SOTPICKX").Compute("SUM(VARIANCE)", "") <> 0 And PICK_NO_CONS = "" Then
            'forced closed warning
            PACK_PACKERS = ",FORCED"
            Dim Extra_CART_NO = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))

            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Select("").FirstOrDefault

            rowWHTCART1 = dst.Tables("WHTCART1").NewRow
            With rowWHTCART1
                .Item("CART_NO") = Extra_CART_NO
                .Item("PICK_NO") = PICK_NO
                .Item("CART_PACKER") = "False"
                .Item("CART_PACKED") = DATETIME_STAMP
                .Item("GUN_ID") = "G99"
                .Item("PRINTER") = ""
                .Item("PROCESS_STATUS") = "2"
                .Item("CART_SEQ") = Val(rowSOTPICK1.Item("CART_SEQ_CTR") & "") + 1
                .Item("CART_TOTAL_UNITS") = dst.Tables("SOTPICKX").Compute("SUM(VARIANCE)", "")
                .Item("CARTONS_PER_UNIT") = 0
                .Item("CART_NO_CONS") = ""
            End With
            dst.Tables("WHTCART1").Rows.Add(rowWHTCART1)
            Update_Record_TDA("WHTCART1")

            For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("VARIANCE <> 0")
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("STYLE_CODE = '" & rowSOTPICKX.Item("STYLE_CODE") & "' and COLOR_CODE = '" & rowSOTPICKX.Item("COLOR_CODE") & " '")
                    rowWHTCART2 = dst.Tables("WHTCART2").NewRow
                    rowWHTCART2.Item("CART_NO") = Extra_CART_NO
                    rowWHTCART2.Item("CART_LNO") = Val(dst.Tables("WHTCART2").Compute("MAX(CART_LNO)", "CART_NO = '" & Extra_CART_NO & "'") & "") + 1
                    rowWHTCART2.Item("ORDR_NO") = ORDR_NO
                    rowWHTCART2.Item("ORDR_LNO") = rowSOTORDR2.Item("ORDR_LNO")
                    rowWHTCART2.Item("QTY_PACKED") = Math.Min(Val(rowSOTPICKX.Item("PICK_QTY")), Math.Abs(Val(rowSOTPICKX.Item("VARIANCE")))) ' This doesn't handle consalidated well, nor same style/color in multiple lines for order
                    rowWHTCART2.Item("UPC_CODE") = rowSOTPICKX.Item("UPC_CODE")
                    rowWHTCART2.Item("STYLE_CODE") = rowSOTPICKX.Item("STYLE_CODE")
                    rowWHTCART2.Item("COLOR_CODE") = rowSOTPICKX.Item("COLOR_CODE")
                    dst.Tables("WHTCART2").Rows.Add(rowWHTCART2)
                Next
            Next
            Update_Record_TDA("WHTCART2")
        End If

        If dst.Tables("SOTPICKX").Compute("SUM(SHORTAGE)", "") <> 0 And PICK_NO_CONS = "" Then
            dst.Tables("WHTPICKS").Rows.Clear()

            For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select("SHORTAGE <> 0")
                Dim rowWHTPICKS As DataRow = dst.Tables("WHTPICKS").NewRow
                With rowWHTPICKS
                    .Item("PICK_NO") = PICK_NO
                    .Item("STYLE_CODE") = rowSOTPICKX.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = rowSOTPICKX.Item("COLOR_CODE")
                    .Item("STYLE_DESC") = rowSOTPICKX.Item("STYLE_DESC")
                    .Item("COLOR_DESC") = rowSOTPICKX.Item("COLOR_DESC")
                    .Item("LOCATION_CODE") = rowSOTPICKX.Item("LOCATION_CODE")
                    .Item("SHORTAGE") = rowSOTPICKX.Item("SHORTAGE")
                    .Item("STATUS") = "O"
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                End With
                dst.Tables("WHTPICKS").Rows.Add(rowWHTPICKS)
            Next
            Update_Record_TDA("WHTPICKS", $"PICK_NO = '{PICK_NO}'")

            ASCMAIN1.sql = "INSERT INTO ASTNOTEM " &
                       "Select 'SHORTAGES' NOTE_CODE, " &
                       "NVL((SELECT max(SEND_LNO) FROM ASTNOTEM WHERE NOTE_CODE = 'SHORTAGES'), 0) + 1 SEND_LNO, " &
                       "'Shortage In Pick Tckt: " & PICK_NO & "' NOTE_MEMO " &
                       "from DUAL"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        End If


        For Each row As DataRow In dst.Tables("WHTCART1").DefaultView.ToTable(True, "CART_PACKER").Select("")
            PACK_PACKERS = PACK_PACKERS & "," & row.Item("CART_PACKER")
        Next
        If PACK_PACKERS.Length > 1 Then
            PACK_PACKERS = PACK_PACKERS.Substring(1, Math.Min(PACK_PACKERS.Length - 1, 20))
        End If

        Dim Prompt As String = "Pack Slip Printed"
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
            rowSOTSHIP1.Item("SHIP_CNT_PALLETS") = Val(txtPICK_CNT_PALLETS.Value)
            rowSOTSHIP1.Item("SHIP_TOTAL_WGT") = Val(txtPICK_TOTAL_WGT.Value)
            rowSOTSHIP1.Item("SHIP_CNT_CARTONS") = dst.Tables("WHTCART1").Select("").Length
            rowSOTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSOTSHIP1.Item("LAST_DATE") = DATETIME_STAMP
            rowSOTSHIP1.Item("SHIP_APPT_NO") = Prompt

            Dim BILL_OF_LADING_NO As String = rowSOTSHIP1.Item("BILL_OF_LADING_NO") & String.Empty
            If BILL_OF_LADING_NO.Length > 0 Then
                ASCMAIN1.sql = "UPDATE SOTSHIPB SET SHIP_APPT_NO = :PARM1 WHERE BOL_NO = :PARM2"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {Prompt, BILL_OF_LADING_NO})
            End If
        Next
        Update_Record_TDA("SOTSHIP1")

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("PICK_CNT_PALLETS") = Val(txtPICK_CNT_PALLETS.Value)
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = Val(txtPICK_TOTAL_WGT.Value)
            rowSOTPICK1.Item("PACK_PACKER") = PACK_PACKERS
            rowSOTPICK1.Item("PACK_SUPERVISOR") = ASCMAIN1.USER_ID
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = dst.Tables("WHTCART1").Select("").Length
            rowSOTPICK1.Item("PACK_STATUS") = "F"
            rowSOTPICK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSOTPICK1.Item("LAST_DATE") = DATETIME_STAMP
            rowSOTPICK1.Item("FLOOR_LOC") = txtFLOOR_LOC.Value
        Next
        Update_Record_TDA("SOTPICK1")

        For Each rowWHTCART1 In dst.Tables("WHTCART1").Select("")
            rowSOTCART1 = dst.Tables("SOTCART1").NewRow
            With rowSOTCART1
                .Item("CART_NO") = rowWHTCART1.Item("CART_NO")
                .Item("PICK_NO") = rowWHTCART1.Item("PICK_NO")
                .Item("CART_PACKER") = rowWHTCART1.Item("CART_PACKER")
                .Item("CART_PACKED") = rowWHTCART1.Item("CART_PACKED")
                .Item("CART_TOTAL_UNITS") = rowWHTCART1.Item("CART_TOTAL_UNITS")
                .Item("CART_SEQ") = Val(rowWHTCART1.Item("CART_SEQ") & "")
                .Item("PKG_CODE") = rowWHTCART1.Item("PKG_CODE") & ""
                .Item("PKG_L") = Val(rowWHTCART1.Item("PKG_L") & "")
                .Item("PKG_W") = Val(rowWHTCART1.Item("PKG_W") & "")
                .Item("PKG_H") = Val(rowWHTCART1.Item("PKG_H") & "")
                .Item("CART_TOTAL_WGT_ACTUAL") = Val(rowWHTCART1.Item("CART_TOTAL_WGT_ACTUAL") & "")
                .Item("CART_TOTAL_WGT_CALC") = Val(rowWHTCART1.Item("CART_TOTAL_WGT_ACTUAL") & "")
            End With
            dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)
        Next
        Update_Record_TDA("SOTCART1")

        For Each rowWHTCART2 In dst.Tables("WHTCART2").Select("")
            rowSOTCART2 = dst.Tables("SOTCART2").NewRow
            With rowSOTCART2
                .Item("CART_NO") = rowWHTCART2.Item("CART_NO")
                .Item("CART_LNO") = rowWHTCART2.Item("CART_LNO")
                .Item("ORDR_NO") = rowWHTCART2.Item("ORDR_NO")
                .Item("ORDR_LNO") = rowWHTCART2.Item("ORDR_LNO")
                .Item("QTY_PACKED") = rowWHTCART2.Item("QTY_PACKED")
                .Item("UPC_CODE") = rowWHTCART2.Item("UPC_CODE")
                .Item("STYLE_CODE") = rowWHTCART2.Item("STYLE_CODE")
                .Item("COLOR_CODE") = rowWHTCART2.Item("COLOR_CODE")
            End With
            dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
        Next
        Update_Record_TDA("SOTCART2")

        CommitTrans("Update Complete")

        If chkUCC128.Checked Then
            For Each rowWHTCART1 In dst.Tables("WHTCART1").Select("")
                Dim CART_NO As String = rowWHTCART1.Item("CART_NO")
                Dim cartonLabel As New TAC.CartonLabel(CART_NO)
                cartonLabel.PrintLabel()
            Next
        End If

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWHSEX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPACKX, "SB", "Show Filter", "Open Shipment", "Pick Shipment")
        Load_Popup_Menu(grdSOTPICKX, "BB", "Void Pick Ln", "Change Pick Qty")
        Load_Popup_Menu(grdSOTPICK1, "B", "Pick")
        'Load_Popup_Menu(grdSOTPACKX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins" _
        '                , "Select All", "De-Select All", "Select All in Group", "Recall Shipment")
        Load_Popup_Menu(grdWHTCART1, "BBB", "Print Label", "Close Carton", "Delete Carton", "Open Carton", "Delete Line")
        Load_Popup_Menu(grdWHTCART2, "S", "Show Filter")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTPACKX"
                    chkAutoRefresh.Checked = False
                    chkAutoRefresh_CheckedChanged(New Object, New EventArgs)
                    tlb_pop.Tools("Open Shipment").SharedProps.Visible = optPickFilter.Value = "F" And EntryMode = "L"
                    tlb_pop.Tools("Pick Shipment").SharedProps.Visible = optPickFilter.Value = "0" And EntryMode = "L" And tabShipment.SelectedTab.Text = "Pick Tickets"
                Case "grdSOTPICKX"
                    tlb_pop.Tools("Void Pick Ln").SharedProps.Visible = optPickFilter.Value = "P" And EntryMode = "L" And grd.ActiveRow.Band.Index = 0 AndAlso (grd.ActiveRow.Cells("PACKED_QTY").Text = "" OrElse grd.ActiveRow.Cells("PACKED_QTY").Value = "0") AndAlso (grd.ActiveRow.Cells("PiCKED_QTY").Text <> "")
                    tlb_pop.Tools("Change Pick Qty").SharedProps.Visible = optPickFilter.Value = "P" And EntryMode = "L" And grd.ActiveRow.Band.Index = 0 AndAlso (grd.ActiveRow.Cells("PACKED_QTY").Text = "" OrElse grd.ActiveRow.Cells("PICKED_QTY").Value > grd.ActiveRow.Cells("PACKED_QTY").Value) AndAlso (grd.ActiveRow.Cells("PiCKED_QTY").Text <> "")
                Case "grdSOTPICK1"
                    tlb_pop.Tools("Pick").SharedProps.Visible = optPickFilter.Value = "0" And EntryMode = "L" And grd.ActiveRow.Band.Index = 0 'AndAlso ASCMAIN1.Running_in_VS
                Case "grdWHTCART1"
                    chkAutoRefresh.Checked = False
                    chkAutoRefresh_CheckedChanged(New Object, New EventArgs)
                    tlb_pop.Tools("Print Label").SharedProps.Visible = grd.ActiveRow.Band.Index = 0 AndAlso grd.ActiveRow.Cells("PROCESS_STATUS").Value <> "0"
                    tlb_pop.Tools("Close Carton").SharedProps.Visible = grd.ActiveRow.Band.Index = 0 AndAlso grd.ActiveRow.Cells("PROCESS_STATUS").Value = "0"
                    tlb_pop.Tools("Delete Carton").SharedProps.Visible = grd.ActiveRow.Band.Index = 0 AndAlso grd.ActiveRow.Cells("PROCESS_STATUS").Value <> "0" And optPickFilter.Value = "P" And EntryMode = "L"
                    tlb_pop.Tools("Open Carton").SharedProps.Visible = grd.ActiveRow.Band.Index = 0 AndAlso grd.ActiveRow.Cells("PROCESS_STATUS").Value <> "0" And optPickFilter.Value = "P"
                    tlb_pop.Tools("Delete Line").SharedProps.Visible = grd.ActiveRow.Band.Index = 1 AndAlso grd.ActiveRow.ParentRow.Cells("PROCESS_STATUS").Value <> "0" And optPickFilter.Value = "P" And EntryMode = "L"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All"

                'For Each grow As UltraWinGrid.UltraGridRow In grdSOTPACKX.Rows
                '    If grow.IsDataRow Then
                '        grow.Cells("SEL").Value = "1"
                '    End If
                'Next
                'grdSOTPACKX.Update()

                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTPACKX").Rows
                    rowSOTWSHIPX.Item("SEL") = "1"
                Next

                MsgBox("You have selected " & dst.Tables("SOTPACKX").Select("SEL = '1'").Length & " Records by Selecting All", MsgBoxStyle.OkOnly, "Verification")

            Case "De-Select All"

                'For Each grow As UltraWinGrid.UltraGridRow In grdSOTPACKX.Rows
                '    If grow.IsDataRow Then
                '        grow.Cells("SEL").Value = "0"
                '    End If
                'Next
                'grdSOTPACKX.Update()

                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTPACKX").Rows
                    rowSOTWSHIPX.Item("SEL") = "0"
                Next
            Case "Open Shipment"
                Dim SHIP_BOL_NO As String = grdSOTPACKX.ActiveRow.Cells("SHIP_BOL_NO").Value & ""
                Dim PICK_NO As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO").Value & ""
                Dim ORDR_NO As String = grdSOTPACKX.ActiveRow.Cells("ORDR_NO").Value & ""

                If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub

                Fill_Records("SOTPICK1", SHIP_BOL_NO, True)
                Dim row As DataRow = dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'").First
                If row.Item("PICK_STATUS") <> "P" Then
                    MsgBox("Shipment Cannot be Open", vbInformation, "Cannot Open")
                    Exit Sub
                End If
                row.Item("PACK_STATUS") = "P"
                Update_Record_TDA("SOTPICK1")
                Load_SOTPACKX()
                ASCMAIN1.Record_Event("SOTORDR1", ORDR_NO, "", Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "PCK_OPN", "Open Shipment - Packing Slip Recalled", "")
                ASCMAIN1.MultiTask_Release()
            Case "Void Pick Ln"
                Dim PICK_NO As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO").Value & ""
                Dim PICK_NO_CONS As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO_CONS").Value & ""
                If PICK_NO_CONS Is "" Then
                    Fill_Records("SOTPICK5", PICK_NO, True)
                Else
                    Fill_Records("SOTPICK5", PICK_NO_CONS, True)
                    PICK_NO = PICK_NO_CONS 'need to send to void sub
                End If
                Dim PickCnt = dst.Tables("SOTPICK5").Compute("count(UPC_CODE)", "UPC_CODE = '" & grdSOTPICKX.ActiveRow.GetCellValue("UPC_CODE") & "'").ToString
                If PickCnt = "1" Then
                    VoidPickLine(PICK_NO, grdSOTPICKX.ActiveRow.GetCellValue("UPC_CODE"))
                Else
                    MsgBox("Multiple Picks for this line, conflict - use gun", MsgBoxStyle.Critical, "Cannot Void Line")
                End If
            Case "Change Pick Qty"
                'tlb_pop.Tools("Change Pick Qty").SharedProps.Visible = optPickFilter.Value = "P" And EntryMode = "L" And grd.ActiveRow.Band.Index = 0 AndAlso 
                '(grd.ActiveRow.Cells("PACKED_QTY").Text = "" OrElse grd.ActiveRow.Cells("PICKED_QTY").Value > grd.ActiveRow.Cells("PACKED_QTY").Value)
                Dim newQty = InputBox("Enter new Pick Qty for Line", "Change Pick Qty")
                If newQty <> "" And Val(newQty) > 0 Then
                    If grd.ActiveRow.Cells("PACKED_QTY").Text <> "" AndAlso Val(newQty) < grd.ActiveRow.Cells("PACKED_QTY").Value Then
                        MsgBox("Action Canceled, Qty Packed is greater than " & newQty, MsgBoxStyle.Critical, "Change Pick Qty")
                        Exit Select
                    ElseIf Val(newQty) >= grdSOTPICKX.ActiveRow.Cells("PICKED_QTY").Value Then
                        MsgBox("Action Canceled, New Qty Picked not allowed to increase", MsgBoxStyle.Critical, "Change Pick Qty")
                        Exit Select
                    Else
                        Dim PICK_NO As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO").Value & ""
                        Dim PICK_NO_CONS As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO_CONS").Value & ""
                        If PICK_NO_CONS Is "" Then
                            Fill_Records("SOTPICK5", PICK_NO, True)
                        Else
                            Fill_Records("SOTPICK5", PICK_NO_CONS, True)
                            PICK_NO = PICK_NO_CONS 'need to send to void sub
                        End If
                        Dim ChgQty = grdSOTPICKX.ActiveRow.Cells("PICKED_QTY").Value - Val(newQty)
                        ChangePickQty(PICK_NO, grdSOTPICKX.ActiveRow.GetCellValue("UPC_CODE"), ChgQty)
                    End If
                Else
                    MsgBox("Action Canceled, invalid qty", MsgBoxStyle.Critical, "Change Pick Qty")
                End If
            Case "Print Label"
                PrintCartonLabel(grd.ActiveRow.Cells("CART_NO").Value)

            Case "Close Carton"
                Dim PICK_NO As String = grdSOTPACKX.ActiveRow.GetCellValue("PICK_NO")
                Dim CART_NO As String = grd.ActiveRow.Cells("CART_NO").Value
                'If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SOTCART1", CART_NO) Then Exit Sub
                Dim row As DataRow = dst.Tables("WHTCART1").Select("CART_NO = '" & CART_NO & "'").First

                row.Item("PROCESS_STATUS") = 1
                Update_Record_TDA("WHTCART1")
                Load_SOTPACKX()

                ASCMAIN1.MultiTask_Release()

            Case "Open Carton"
                Dim PICK_NO As String = grdSOTPACKX.ActiveRow.GetCellValue("PICK_NO")
                Dim CART_NO As String = grd.ActiveRow.Cells("CART_NO").Value
                'If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SOTCART1", CART_NO) Then Exit Sub

                Dim openCnt = dst.Tables("WHTCART1").Compute("count(PROCESS_STATUS)", "PROCESS_STATUS = 0").ToString
                If openCnt <> "0" Then
                    MsgBox("Only one open carton per pick", vbInformation, "Cannot Open")
                    Exit Sub
                End If
                Dim row As DataRow = dst.Tables("WHTCART1").Select("CART_NO = '" & CART_NO & "'").First
                If row.Item("CART_TOTAL_UNITS") < 2 Then
                    MsgBox("Carton is Single Unit or Multipack", vbInformation, "Cannot Open")
                    Exit Sub
                End If

                row.Item("PROCESS_STATUS") = 0
                Update_Record_TDA("WHTCART1")
                Load_SOTPACKX()

                ASCMAIN1.MultiTask_Release()
                Setup_SOTPACKX()

            Case "Delete Carton"
                Dim PICK_NO As String = grdSOTPACKX.ActiveRow.GetCellValue("PICK_NO")
                Dim CART_NO As String = grd.ActiveRow.Cells("CART_NO").Value
                If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SOTCART1", CART_NO) Then Exit Sub

                If vbYes = MsgBox("Delete Carton " & grd.ActiveRow.Cells("CART_SEQ").Value, vbYesNoCancel, "Delete") Then
                    DeleteCarton(CART_NO)
                End If

                ASCMAIN1.MultiTask_Release()
                Setup_SOTPACKX()

            Case "Delete Line"
                Dim CART_NO As String = grd.ActiveRow.Cells("CART_NO").Value
                Dim CART_LNO As String = grd.ActiveRow.Cells("CART_LNO").Value
                Dim PICK_NO As String = grdSOTPACKX.ActiveRow.GetCellValue("PICK_NO")
                If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SOTCART1", CART_NO) Then Exit Sub

                If vbYes = MsgBox("Delete Line " & grd.ActiveRow.Cells("STYLE_CODE").Value & " - " & grd.ActiveRow.Cells("COLOR_CODE").Value & " Qty " & grd.ActiveRow.Cells("QTY_PACKED").Value, vbYesNoCancel, "Delete") Then
                    DeleteLine(CART_NO, CART_LNO)
                    'PrintCartonLabel(CART_NO)
                End If

                ASCMAIN1.MultiTask_Release()
                Setup_SOTPACKX()

            Case "Pick"
                Start_App(grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value)
                Load_SOTPACKX()

            Case "Pick Shipment"
                Dim SHIP_BOL_NO As String = grdSOTPACKX.ActiveRow.Cells("SHIP_BOL_NO").Value & ""
                For Each row As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PACK_STATUS is Null")
                    Dim PICK_NO As String = row.Item("PICK_NO")
                    Start_App(PICK_NO)
                    If cmdClosePicks.Visible Then

                        Automate_Pick()

                        splPicks.Panel2Collapsed = True
                        cmdClosePicks.Visible = False
                        btnAutomate.Visible = False
                        C = Nothing
                    End If
                Next
                Close_Picks()
                Load_SOTPACKX()

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Select All in Group"
                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTPACKX").Select("ORDR_GROUP_NO = '" & grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & "'")
                    rowSOTWSHIPX.Item("SEL") = "1"
                Next
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("View")
        End Select
    End Sub
#End Region

    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("View")
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()

    End Sub

    Private Sub grdSOTPACKX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPACKX.AfterRowActivate
        Setup_SOTPACKX()
    End Sub

    Sub Setup_SOTPACKX()
        If grdSOTPACKX.ActiveRow Is Nothing OrElse Not grdSOTPACKX.ActiveRow.IsDataRow Then
            tabShipment.Visible = False
            UltraExplorerBar1.Groups("Screen Control").Items("Print").Visible = False
            UltraExplorerBar1.Groups("Screen Control").Items("Print Cons by P.O.").Visible = False
        Else
            tabShipment.Visible = True
            UltraExplorerBar1.Groups("Screen Control").Items("Print").Visible = (optPickFilter.Value = "F")
            UltraExplorerBar1.Groups("Screen Control").Items("Print Cons by P.O.").Visible = (optPickFilter.Value = "F")
            Dim SHIP_BOL_NO As String = grdSOTPACKX.ActiveRow.Cells("SHIP_BOL_NO").Value & ""
            Dim PICK_NO As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO").Value & ""
            Dim ORDR_NO As String = grdSOTPACKX.ActiveRow.Cells("ORDR_NO").Value & ""
            Dim ORDR_GROUP_NO As String = grdSOTPACKX.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
            Dim CUST_CODE As String = grdSOTPACKX.ActiveRow.Cells("CUST_CODE").Value & ""
            Dim PACK_PACKERS As String = ""
            Dim PICK_NO_CONS As String = grdSOTPACKX.ActiveRow.Cells("PICK_NO_CONS").Value & ""

            grdSOTPICK1.Text = "Pick Tickets for Shipment No " & SHIP_BOL_NO
            Fill_Records("SOTPICK1", SHIP_BOL_NO, True)
            Sort_grdColumns(grdSOTPICK1, "PICK_NO")
            grdSOTPICKX.Text = "Style/Color Summary for Shipment No " & SHIP_BOL_NO
            Fill_Records("SOTPICKX", New String() {SHIP_BOL_NO, ORDR_NO}, True)
            Sort_grdColumns(grdSOTPICKX, "STYLE_CODE,COLOR_CODE")
            If PICK_NO_CONS Is "" Then
                Fill_Records("SOTPICK5", PICK_NO, True)
            Else
                Fill_Records("SOTPICK5", PICK_NO_CONS, True)
            End If

            Fill_Records("SOTORDR1", ORDR_NO, True)
            Fill_Records("SOTORDR2", ORDR_NO, True)
            Fill_Records("SOTORDR5", ORDR_NO, True)
            Fill_Records("SOTORDR0", ORDR_GROUP_NO, True)
            Fill_Records("SOTSHIP1", SHIP_BOL_NO, True)
            Fill_Records("ARTCUST1", CUST_CODE, True)
            Fill_Records("WHTGUNAS", "", True)

            For Each row As DataRow In dst.Tables("SOTSHIP1").Select("")
                Fill_Records("SOTSVIA1", row.Item("SHIP_VIA_CODE"))
            Next

            dst.EnforceConstraints = False
            dst.Tables("WHTCART1").Rows.Clear()
            dst.Tables("WHTCART2").Rows.Clear()
            If (optPickFilter.Value = "P" Or optPickFilter.Value = "F") Then
                Fill_Records("WHTCART1", PICK_NO, True)
                Sort_grdColumns(grdWHTCART1, "CART_NO")
                Fill_Records("WHTCART2", PICK_NO, True)
            End If
            dst.EnforceConstraints = True

            For Each row As DataRow In dst.Tables("WHTCART1").DefaultView.ToTable(True, "CART_PACKER").Select("")
                PACK_PACKERS = PACK_PACKERS & "," & row.Item("CART_PACKER")
            Next
            If Not IsNothing(PACK_PACKERS) AndAlso PACK_PACKERS.Length > 1 Then
                PACK_PACKERS = PACK_PACKERS.Substring(1, Math.Min(PACK_PACKERS.Length - 1, 20))
            End If

            For Each row As DataRow In dst.Tables("SOTPICKX").Select("")
                row.Item("PICKED_QTY") = dst.Tables("SOTPICK5").Compute("SUM(PICK_QTY)", "UPC_CODE = '" & row.Item("UPC_CODE") & "'")
                row.Item("PACKED_QTY") = dst.Tables("WHTCART2").Compute("SUM(QTY_PACKED)", "STYLE_CODE = '" & row.Item("STYLE_CODE") & "' and COLOR_CODE = '" & row.Item("COLOR_CODE") & "'")
                row.Item("LOCATION_CODE") = GetLocation(row.Item("STYLE_CODE"), row.Item("COLOR_CODE"), row.Item("PICK_QTY"))
            Next

            tabShipment.Tabs("Pick Tickets").Visible = Not (optPickFilter.Value = "P" Or optPickFilter.Value = "F")
            tabShipment.Tabs("Cartons").Visible = (optPickFilter.Value = "P" Or optPickFilter.Value = "F")
            tabShipment.Tabs("Packed by style").Visible = (optPickFilter.Value = "P" Or optPickFilter.Value = "F")

            pnlCloseNPrint.Visible = (optPickFilter.Value <> "F" And EntryMode = "L")
            pnlCloseNPrint.Enabled = (optPickFilter.Value <> "F" And EntryMode = "L")

        End If
    End Sub

    Private Sub optPending_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPickFilter.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub

        chkAutoRefresh.Checked = False
        chkAutoRefresh.Visible = (optPickFilter.Value = "P")
        Load_SOTPACKX()
    End Sub

    Sub Load_SOTPACKX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Shipments Queue")

        Dim PACK_STATUS As String = ""
        Dim GridRow As UltraWinGrid.UltraGridRow
        Dim PICK_NO As String = ""

        If grdSOTPACKX.ActiveRow IsNot Nothing AndAlso grdSOTPACKX.ActiveRow.IsDataRow Then
            PICK_NO = grdSOTPACKX.ActiveRow.GetCellValue("PICK_NO")
        End If

        If optPickFilter.Value = "0" Then
            PACK_STATUS = "0"
            grdSOTPACKX.Text = "Pick Tickets in Pick nd Not in Pack for Warehouse " & WHSE_CODE

        ElseIf optPickFilter.Value = "P" Then

            PACK_STATUS = "P"
            grdSOTPACKX.Text = "Pick Tickets in Pick and also in Pack for Warehouse " & WHSE_CODE
            'If chkAutoRefresh.Checked Then
            '    grdSOTPACKX.Text &= " Last Checked " & Format(Now, "MM/dd/yyyyy HH:mm:ss")
            '    '  selGridRow = grdSOTPACKX.ActiveRow
            '    If grdSOTPACKX.ActiveRow IsNot Nothing Then
            '        PICK_NO = grdSOTPACKX.ActiveRow.GetCellValue("PICK_NO")
            '    End If
            'End If

        Else
            PACK_STATUS = "F"
            grdSOTPACKX.Text = "Pick Tickets Packed and Not Shipped for Warehouse " & WHSE_CODE

        End If

        UltraExplorerBar1.Groups("Screen Control").Items("Print").Visible = False
        UltraExplorerBar1.Groups("Screen Control").Items("Print Cons by P.O.").Visible = False

        Fill_Records("SOTPACKX", New String() {WHSE_CODE, PACK_STATUS}, True)
        ASCMAIN1.Progress("Now Sorting Data")
        Sort_grdColumns(grdSOTPACKX, "PICK_NO".ToLower)
        If PICK_NO <> "" And PICK_NO IsNot Null Then
            ASCMAIN1.Progress("Finding previous row")
            For intRow As Integer = 0 To grdSOTPACKX.Rows.Count - 1
                GridRow = grdSOTPACKX.Rows(intRow)
                If GridRow.Band.Index = 0 AndAlso GridRow.Cells("PICK_NO").Value = PICK_NO Then
                    grdSOTPACKX.ActiveRow = GridRow
                End If
            Next
        End If

        ASCMAIN1.Progress("Now Setting up Data")
        Setup_SOTPACKX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Default_Printer()
        Dim settings As New PrinterSettings
        For Each printer As String In PrinterSettings.InstalledPrinters
            settings.PrinterName = printer
            If settings.IsDefaultPrinter Then
                Return printer
            End If
        Next
        Return String.Empty
    End Function

    Private Sub chkAutoRefresh_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoRefresh.CheckedChanged
        Timer1.Enabled = chkAutoRefresh.Checked
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Load_SOTPACKX()
    End Sub

    Sub Print_Packing_List()

        'NOTE THAT THIS PRINT ROUTINE WAS USING THE DATA LAYER & DST THAT IS ASSOCIATED WITH THIS FORM   
        Dim sql As String = ""
        Dim BILL_OF_LADING_NO As String = grdSOTPACKX.ActiveRow.Cells("BILL_OF_LADING_NO").Value & ""
        Dim ORDR_NO As String = grdSOTPACKX.ActiveRow.Cells("ORDR_NO").Value & ""
        Dim PICK_NO As String = ""
        Dim REPORT_NAME As String = String.Empty

        REPORT_NAME = String.Format("WHAFPACK1_{0}", ASCMAIN1.Next_Control_No("WHFPACK1.PACKING_LIST"))
        ' Entity to attach report to
        ENTITY.TABLE_NAME = "SOTORDR1"
        ENTITY.COLUMN_NAME = "ORDR_NO"
        ENTITY.CODE_VALUE = ORDR_NO

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = "WHRPACK1"

        Generate_Report(RPT, "Packing List", , , "PDF", REPORT_NAME, True)
        Generate_Report(RPT, "Packing List")

        'FileCopy(ASCMAIN1.Folders("Temp") & REPORT_NAME & ".PDF", ASCMAIN1.Folders("Archive") & REPORT_NAME & ".PDF")
        Attach_File(ASCMAIN1.Folders("Temp") & REPORT_NAME & ".PDF", "Packing List")
        ASCMAIN1.Record_Event(ENTITY.TABLE_NAME, ORDR_NO, "", Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "PCK_CLS", "Packing Slip Printed", "")
        If grdSOTPACKX.ActiveRow.Cells("GRP_CNT").Value & "" > "1" Then
            If MsgBox("Print Summary Report", vbYesNo, "Combined shipment") = vbYes Then

                sql = "SELECT SOTPICK1.* FROM SOTPICK1, SOTSHIP1 " _
                    & " WHERE SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO " _
                    & " And SOTSHIP1.BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'"
                Fill_Records("SOTPICK1", , True, sql)

                sql = "SELECT DISTINCT SOTORDR1.* FROM SOTORDR1, SOTSHIP1 " _
                    & " WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO " _
                    & " And SOTSHIP1.BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'"
                Fill_Records("SOTORDR1", , True, sql)

                sql = "SELECT DISTINCT SOTORDR2.* FROM SOTORDR2, SOTORDR1, SOTSHIP1 " _
                    & " WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO " _
                    & " and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
                    & " And SOTSHIP1.BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'"
                Fill_Records("SOTORDR2", , True, sql)

                sql = "SELECT DISTINCT SOTORDR5.* FROM SOTORDR5, SOTORDR1, SOTSHIP1 " _
                    & " WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO " _
                    & " and SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO and SOTORDR5.CUST_ADDR_TYPE = 'ST'" _
                    & " And SOTSHIP1.BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'"
                Fill_Records("SOTORDR5", , True, sql)

                sql = "SELECT DISTINCT SOTORDR0.* FROM SOTORDR0, SOTSHIP1 " _
                    & " WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO " _
                    & " And SOTSHIP1.BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'"
                Fill_Records("SOTORDR0", , True, sql)

                sql = "SELECT * FROM SOTSHIP1 " _
                    & " WHERE SOTSHIP1.BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'"
                Fill_Records("SOTSHIP1", , True, sql)

                dst.EnforceConstraints = False
                dst.Tables("WHTCART1").Clear()
                dst.Tables("WHTCART2").Clear()
                For Each row As DataRow In dst.Tables("SOTPICK1").Select("")
                    PICK_NO = row.Item("PICK_NO")
                    Fill_Records("WHTCART1", PICK_NO, False)
                    Fill_Records("WHTCART2", PICK_NO, False)
                Next
                dst.EnforceConstraints = True
                Generate_Report("WHRPACK2", "Packing List Summary", , , , , False)
            End If
        End If
        If ASCMAIN1.Running_in_VS Then
            Print_Report_End()
        Else
            Print_Report_End(False, , cbxReportPrinter.SelectedItem, 2)
        End If
    End Sub

    Sub Print_Packing_List_Consolidated()
        Try
            Dim ORDR_NO As String = grdSOTPACKX.ActiveRow.Cells("ORDR_NO").Value & String.Empty
            Fill_Records("WHTPACKC", New String() {ORDR_NO})

            Print_Report_Begin()
            CR_params.Add("SUBT", "")
            Generate_Report("WHRPACKC", "Packing List by Cartons", , , , , False)

            If ASCMAIN1.Running_in_VS Then
                Print_Report_End()
            Else
                Print_Report_End(False, , cbxReportPrinter.SelectedItem, 2)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Print Pack Slip", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCloseNPrint_Click(sender As Object, e As EventArgs) Handles btnCloseNPrint.Click
        Click_Command("Update")

    End Sub
    Private Sub PrintCartonLabel(ByVal CART_NO As String)
        Try
            Dim label As String = ""
            Dim sql As String
            Dim Printer = cbxLabelPrinter.SelectedItem
            Dim qty As Integer
            Dim lines As Integer

            If Printer Is Null Then
                MsgBox("Select Label Printer to print to", vbInformation, "Label Printer")
                Exit Sub
            End If

            sql = "SELECT PRINTER, CART_NO, CART_SEQ, SOTPICK1.PICK_NO, nvl(ORDR_CUST_PO, ' ') ORDR_CUST_PO, CASE CUST_STORE_NO WHEN NULL THEN '' WHEN '000000' THEN '' ELSE LTRIM(CUST_STORE_NO,'0') || '-' END || nvl(SOTORDR5.CUST_NAME,' ') CUST_NAME," & vbCrLf _
                    & " CUST_CONTACT, CUST_ADDR1, CUST_ADDR2, CUST_ADDR3, CUST_CITY, CUST_STATE, CUST_ZIP_CODE, NVL(CART_NO_CONS, CART_NO) CART_NO_CONS" & vbCrLf _
                    & " FROM WHTCART1, SOTPICK1, SOTORDR1, SOTORDR5" & vbCrLf _
                    & " WHERE SOTPICK1.PICK_NO = WHTCART1.PICK_NO" & vbCrLf _
                    & " AND SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                    & " AND SOTORDR5.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                    & " AND SOTORDR5.CUST_ADDR_TYPE = 'ST'" & vbCrLf _
                    & " AND WHTCART1.CART_NO = '" & CART_NO & "'"
            For Each ROW As DataRow In ASCDATA1.GetDataTable(sql).Select("")
                label = "NEW|CartonLabel.lbx|" & Printer & "|" & ROW.Item("CART_SEQ") & "|" & ROW.Item("CART_NO") & "|" & ROW.Item("PICK_NO") & "|" _
                    & ROW.Item("ORDR_CUST_PO") & "|" & ROW.Item("CUST_NAME") & "~" & If(ROW.Item("CUST_CONTACT") IsNot Null, ROW.Item("CUST_CONTACT") & "~", "") & ROW.Item("CUST_ADDR1") & "~" _
                    & If(ROW.Item("CUST_ADDR2") IsNot Null, ROW.Item("CUST_ADDR2") & "~", "") & If(ROW.Item("CUST_ADDR3") IsNot Null, ROW.Item("CUST_ADDR3") & "~", "") _
                    & ROW.Item("CUST_CITY") & ", " & ROW.Item("CUST_STATE") & " " & ROW.Item("CUST_ZIP_CODE") & "|"
                lines = 0
                For Each rowDtl As DataRow In ASCDATA1.SelectDistinct(dst.Tables("WHTCART2").Select("CART_NO = '" & ROW.Item("CART_NO_CONS") & "'"), "STYLE_CODE", "COLOR_CODE", "STYLE_UOM").Rows
                    lines += 1
                    If lines > 6 Then
                        label = label & "See Packing list..."
                        Exit For
                    End If
                    qty = dst.Tables("WHTCART2").Compute("SUM(QTY_PACKED)", "STYLE_CODE = '" & rowDtl.Item("STYLE_CODE") & "' and COLOR_CODE = '" & rowDtl.Item("COLOR_CODE") & "' and CART_NO = '" & ROW.Item("CART_NO_CONS") & "'")
                    label = label & rowDtl.Item("STYLE_CODE") & " " & rowDtl.Item("COLOR_CODE") & "  " & qty & "  " & rowDtl.Item("STYLE_UOM") & "~"
                Next

                Using ipp As New nsoftware.IPWorks.Ipport
                    ipp.RuntimeLicense = "31504E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004B4857525953375A4A5A375A0000"
                    If ASCMAIN1.Running_in_VS Then
                        ipp.Connect("192.168.1.3", "4444") 'ipp.Connect("192.168.120.67", "4444") '"192.168.4.117", "4444")
                    Else
                        ipp.Connect("192.168.110.223", "4444")
                    End If

                    ipp.SendLine(label)
                    ASCDATA1.ExecuteSQL("Update WHTCART1 set PROCESS_STATUS = '2' where CART_NO = '" & CART_NO & "'")
                    ipp.Disconnect()
                End Using
            Next
        Catch ex As Exception

        End Try
    End Sub
    Sub DeleteCarton(ByVal CART_NO As String)

        For Each row As DataRow In dst.Tables("WHTCART2").Select("CART_NO = '" & CART_NO & "'")
            Dim qty = row.Item("QTY_PACKED")
            For Each rowSOTPICK5 As DataRow In dst.Tables("SOTPICK5").Select("UPC_CODE = '" & row.Item("UPC_CODE") & "'")
                If Val(rowSOTPICK5.Item("PACK_QTY") & "") < qty Then
                    qty = qty - Val(rowSOTPICK5.Item("PACK_QTY") & "")
                    rowSOTPICK5.Item("PACK_QTY") = 0
                Else
                    rowSOTPICK5.Item("PACK_QTY") = rowSOTPICK5.Item("PACK_QTY") - qty
                    Exit For
                End If
            Next
        Next
        Update_Record_TDA("SOTPICK5")

        ASCDATA1.ExecuteSQL("Delete WHTCART2 Where CART_NO = '" & CART_NO & "'")
        ASCDATA1.ExecuteSQL("Delete WHTCART1 Where CART_NO = '" & CART_NO & "'")
        Setup_SOTPACKX()

    End Sub

    Sub DeleteLine(ByVal CART_NO As String, ByVal CART_LNO As String)

        Dim rowWHTCART1 As DataRow = dst.Tables("WHTCART1").Select("CART_NO = '" & CART_NO & "'").First

        For Each row As DataRow In dst.Tables("WHTCART2").Select("CART_NO = '" & CART_NO & "' and CART_LNO = '" & CART_LNO & "'")
            Dim qty = row.Item("QTY_PACKED")
            rowWHTCART1.Item("CART_TOTAL_UNITS") = rowWHTCART1.Item("CART_TOTAL_UNITS") - qty
            For Each rowSOTPICK5 As DataRow In dst.Tables("SOTPICK5").Select("UPC_CODE = '" & row.Item("UPC_CODE") & "'")
                If rowSOTPICK5.Item("PACK_QTY") < qty Then
                    qty = qty - rowSOTPICK5.Item("PACK_QTY")
                    rowSOTPICK5.Item("PACK_QTY") = 0
                Else
                    rowSOTPICK5.Item("PACK_QTY") = rowSOTPICK5.Item("PACK_QTY") - qty
                    Exit For
                End If
            Next
        Next
        Update_Record_TDA("WHTCART1")
        Update_Record_TDA("SOTPICK5")

        ASCDATA1.ExecuteSQL("Delete WHTCART2 Where CART_NO = '" & CART_NO & "' and CART_LNO = '" & CART_LNO & "'")
        Setup_SOTPACKX()

    End Sub
    Sub VoidPickLine(ByVal PICK_NO As String, ByVal UPC_CODE As String)
        Dim WHSE_TRAN_NO As String

        'Need to allow managers to reverse a line picked by gun - WHCRF016 
        BeginTrans()

        For Each OrigRow As DataRow In dst.Tables("SOTPICK5").Select("UPC_CODE = '" & UPC_CODE & "'")

            With OrigRow
                .Item("PICK_STATUS") = "V"
                .Item("PICK_CASES") = 0
                .Item("PICK_UNITS") = 0
                .Item("PICK_QTY") = 0
                WHSE_TRAN_NO = .Item("WHSE_TRAN_NO")
            End With
            Update_Record_TDA("SOTPICK5")

            ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                           New Object() {WHSE_TRAN_NO, 0, -1},
                           New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        Next
        CommitTrans()
        Setup_SOTPACKX()

    End Sub

    Sub ChangePickQty(ByVal pick_no As String, ByVal UPC_CODE As String, ByVal ChgQty As Integer)
        Dim WHSE_TRAN_NO As String
        Dim rowSOTPICK5 As DataRow
        'Need to allow managers to chage pick qty for a line picked by gun - WHCRF016 
        BeginTrans()
        Dim qty = ChgQty
        Dim newPick As Integer = 0
        For Each OrigRow As DataRow In dst.Tables("SOTPICK5").Select("UPC_CODE = '" & UPC_CODE & "'")
            rowSOTPICK5 = OrigRow
            With OrigRow
                If .Item("PICK_QTY") < qty Then
                    newPick = 0
                    qty = qty - .Item("PICK_QTY")
                Else
                    newPick = .Item("PICK_QTY") - qty
                    qty = 0
                End If

                .Item("PICK_STATUS") = "V"
                .Item("PICK_CASES") = 0
                .Item("PICK_UNITS") = 0
                .Item("PICK_QTY") = 0
                WHSE_TRAN_NO = .Item("WHSE_TRAN_NO")
            End With

            ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                           New Object() {WHSE_TRAN_NO, 0, -1},
                           New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

            If newPick > 0 Then
                Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
                Dim rowWHTMOVE2 As DataRow
                Dim newWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
                Dim newSOTPICK5 As DataRow = dst.Tables("SOTPICK5").NewRow
                Fill_Records("WHTMOVE1", WHSE_TRAN_NO)
                Fill_Records("WHTMOVE2", WHSE_TRAN_NO)

                Dim new_WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

                With rowWHTMOVE1
                    .Item("WHSE_TRAN_NO") = new_WHSE_TRAN_NO
                    .Item("WHSE_TRAN_TYPE") = "M"
                    .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("STATUS") = "U"
                End With
                dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
                Update_Record_TDA("WHTMOVE1")

                rowWHTMOVE2 = dst.Tables("WHTMOVE2")(0)
                newWHTMOVE2.ItemArray = rowWHTMOVE2.ItemArray.Clone()
                With newWHTMOVE2
                    .Item("WHSE_TRAN_NO") = new_WHSE_TRAN_NO
                    .Item("WHSE_TRAN_QTY") = newPick
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("STATUS") = "U"
                End With
                dst.Tables("WHTMOVE2").Rows.Add(newWHTMOVE2)
                Update_Record_TDA("WHTMOVE2")

                newSOTPICK5.ItemArray = rowSOTPICK5.ItemArray.Clone()
                With newSOTPICK5
                    .Item("WHSE_TRAN_NO") = new_WHSE_TRAN_NO
                    .Item("PICK_CASES") = 0
                    .Item("PICK_UNITS") = newPick
                    .Item("PICK_QTY") = newPick
                    .Item("PICK_STATUS") = "P"
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                End With
                dst.Tables("SOTPICK5").Rows.Add(newSOTPICK5)

                ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                           New Object() {new_WHSE_TRAN_NO, 0, 1},
                           New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})
            End If
            If qty = 0 Then Exit For
        Next
        Update_Record_TDA("SOTPICK5")

        CommitTrans()
        Setup_SOTPACKX()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        'If optPickFilter.Value = "0" Then
        Load_SOTPACKX()
        'Else
        '    Setup_SOTPACKX()
        'End If

    End Sub

    Private Sub grdWHTCART1_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdWHTCART1.BeforeCellUpdate
        If e.Cell.Column.Key = "CART_SEQ" Then
            Dim PICK_NO As String = grdWHTCART1.ActiveRow.Cells("PICK_NO").Value
            If grdWHTCART1.ActiveRow.Cells("PROCESS_STATUS").Value = 0 Then
                MsgBox("Carton Is Open, wait for close.", vbInformation, "Cannot Update")
                e.Cancel = True
            ElseIf ASCDATA1.GetDataValue("Select Count(1) FROM WHTCART1 where PICK_NO = '" & PICK_NO & "' and CART_SEQ = " & e.NewValue.ToString) > 0 Then
                MsgBox("Value in use, check again", vbInformation, "Cannot Update")
                e.Cancel = True
            End If
        End If

    End Sub

    Private Sub grdWHTCART1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdWHTCART1.AfterRowUpdate
        Dim CART_NO As String = grdWHTCART1.ActiveRow.Cells("CART_NO").Value
        'ASCDATA1.ExecuteSQL("Update WHTCART1 Set CART_SEQ = " & grdWHTCART1.ActiveRow.Cells("CART_SEQ").Value & " where CART_NO = '" & CART_NO & "'")
        Update_Record_TDA("WHTCART1")
        'PrintCartonLabel(CART_NO)
        Setup_SOTPACKX()

    End Sub
    Private Sub grdWHTCART1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTCART1.AfterCellUpdate

        Dim displayBoxAttributes As Boolean = False

        'If e.Cell.Column.Key = "PKG_CODE" Then

        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTCART1.Rows
            If row.Cells("PKG_CODE").Text = "OTHER" Then
                displayBoxAttributes = True
                Exit For
            End If
        Next

        If displayBoxAttributes Then
            grdWHTCART1.DisplayLayout.Bands(0).Columns("PKG_W").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdWHTCART1.DisplayLayout.Bands(0).Columns("PKG_L").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdWHTCART1.DisplayLayout.Bands(0).Columns("PKG_H").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdWHTCART1.DisplayLayout.Bands(0).Columns("PKG_W").CellActivation = UltraWinGrid.Activation.NoEdit
            grdWHTCART1.DisplayLayout.Bands(0).Columns("PKG_L").CellActivation = UltraWinGrid.Activation.NoEdit
            grdWHTCART1.DisplayLayout.Bands(0).Columns("PKG_H").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
        'End If
    End Sub

    Private Sub grdWHTCART1_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdWHTCART1.BeforeRowUpdate

        'If ASCMAIN1.CLIENT = "RGI" Then
        '    If e.Row.Cells("REFERENCE1").Value & String.Empty = String.Empty Then
        '        e.Row.Cells("REFERENCE1").Value = Val(e.Row.Cells("PICK_NO").Value & String.Empty)
        '    End If
        '    e.Row.Cells("REFERENCE2").Value = e.Row.Cells("CART_NO").Value & String.Empty
        'End If

        Dim rowWHTPKGM1 As DataRow = Nothing
        Dim PKG_CODE As String = e.Row.Cells("PKG_CODE").Value & String.Empty

        If dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE & "' AND PKG_CODE <> 'OTHER'").Length > 0 Then
            rowWHTPKGM1 = dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE & "'")(0)
            e.Row.Cells("PKG_L").Value = rowWHTPKGM1.Item("PKG_L")
            e.Row.Cells("PKG_W").Value = rowWHTPKGM1.Item("PKG_W")
            e.Row.Cells("PKG_H").Value = rowWHTPKGM1.Item("PKG_H")
            Exit Sub
        End If

        ' Sort the values by length, width, height
        Dim PKG_L As Decimal = Val(e.Row.Cells("PKG_L").Value & String.Empty)
        Dim PKG_W As Decimal = Val(e.Row.Cells("PKG_W").Value & String.Empty)
        Dim PKG_H As Decimal = Val(e.Row.Cells("PKG_H").Value & String.Empty)

        If PKG_L <= 0 OrElse PKG_W <= 0 OrElse PKG_H < 0 Then
            MessageBox.Show("All dimensions must be greater than 0", "Update", MessageBoxButtons.OK)
            e.Cancel = True
            Exit Sub
        End If

        Dim dimList As New List(Of Decimal)
        dimList.Add(PKG_L)
        dimList.Add(PKG_W)
        dimList.Add(PKG_H)
        dimList.Sort()
        PKG_L = dimList(2)
        PKG_W = dimList(1)
        PKG_H = dimList(0)

        e.Row.Cells("PKG_L").Value = PKG_L
        e.Row.Cells("PKG_W").Value = PKG_W
        e.Row.Cells("PKG_H").Value = PKG_H

    End Sub

    Private Sub grdSOTPACKX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPACKX.InitializeRow
        With e.Row.Cells("GRP_CNT")
            If .Value & "" > "1" Then
                e.Row.Appearance.BackColor = System.Drawing.Color.LightGreen
            Else
                e.Row.Appearance.BackColor = System.Drawing.Color.Empty
            End If
        End With
    End Sub

    Function GetLocation(ByVal Style As String, ByVal Color As String, ByVal PICK_QTY As Int32) As String
        Dim LOCATION_CODE As String = ""
        'Verify logic against WHCRF016 - we should use same SQL logic
        ASCMAIN1.sql = " select b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE " & vbCrLf _
            & " from whtlocb1 b1 " & vbCrLf _
            & "  join whtlocm1 m1 on b1.LOCATION_CODE = m1.LOCATION_CODE and b1.WHSE_CODE = m1.WHSE_CODE " & vbCrLf _
            & "  where b1.STYLE_CODE = :PARM1 and b1.COLOR_CODE = :PARM2 " & vbCrLf _
            & "  and  nvl(m1.LOCATION_USE,'A') = 'A' " & vbCrLf _
            & "  and m1.WHSE_CODE = :PARM3" & vbCrLf _
            & "  order by b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE"

        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VVV", New String() {Style, Color, WHSE_CODE}).Select("")
            LOCATION_CODE = row.Item("LOCATION_CODE")
            If row("LOCATION_QTY") >= PICK_QTY Then
                Exit For
            End If
        Next

        Return LOCATION_CODE

    End Function

    Sub Start_App(PICK_NO As String)

        ASCMAIN1.sql = "Select APP_ID from WHTGUNA1 where PROCEDURE_NAME = 'WHCRF016' and PICK_TYPE = 'N' and USE_CLASS = '1'"
        Dim row As DataRow = ASCDATA1.GetDataRow
        If row Is Nothing Then
            MsgBox("Problem Identifying Pick Application", MsgBoxStyle.OkOnly, "Cannot Pick")
            Exit Sub
        End If

        Dim APP_ID As String = row.Item("APP_ID")

        Dim rowWHTGUNA1 As DataRow = LookUp("WHTGUNA1", APP_ID)
        Dim APP_DESC As String = rowWHTGUNA1.Item("APP_DESC")
        Dim PROCEDURE_NAME As String = rowWHTGUNA1.Item("PROCEDURE_NAME") & ""
        Dim GUN_PARAM As String = rowWHTGUNA1.Item("PICK_TYPE") & ""

        splPicks.Panel2Collapsed = False
        cmdClosePicks.Visible = True
        btnAutomate.Visible = True
        btnAutomate.Enabled = True
        txtPickChat.Text = ""
        Dim GUN_LOC As String = "99-G00-A"

        'Dim C As WHC.WHCRF000
        C = WHC.WHCFACT1.CreateWhcClass(PROCEDURE_NAME, New WHC.GunEnvironment With
            {.DBS_COMPANY = ASCMAIN1.DBS_COMPANY, .DBS_SERVER = ASCMAIN1.DBS_SERVER, .DBS_PASSWORD = ASCMAIN1.DBS_PASSWORD,
             .THREAD_NO = 0, .APP_ID = APP_ID, .APP_DESC = APP_DESC,
             .USER_ID = ASCMAIN1.USER_ID, .GUN_LOC = GUN_LOC, .PICK_TYPE = GUN_PARAM, .WHSE_CODE = WHSE_CODE})

        AddHandler C.RespondToScan, AddressOf Display_Text
        txt2.Focus()

        Display_Text(0, C.Hello)
        txt2.Text = "A" & PICK_NO
        'End If

    End Sub

    Sub Display_Text(THREAD_NO As Integer, TXT As String)
        txtPickChat.Text &= vbCrLf & "Thread " & THREAD_NO & ":" & TXT
        txtPickChat.SelectionStart = txtPickChat.Text.Length - 1
        txtPickChat.ScrollToCaret()

        txt2.Text = C.RESPONSE_anticipated_next
        txt2.SelectionStart = txt2.Text.Length
        txt2.ScrollToCaret()
    End Sub

    Sub txt2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt2.KeyDown

        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Try
                Dim txt2 As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
                'Dim THREAD_NO As Integer = txt2.Tag
                C.GetResponseToScan(txt2.Text)
                'txt2.Text = ""
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub cmdClosePicks_Click(sender As System.Object, e As System.EventArgs) Handles cmdClosePicks.Click
        Close_Picks()
    End Sub

    Sub Close_Picks()
        splPicks.Panel2Collapsed = True
        cmdClosePicks.Visible = False
        btnAutomate.Visible = False

        C = Nothing

        'Dim WAVE_NO As String = Me.WAVE_NO
        'Click_Command("Done")
        'Absx1.txtFor("WAVE_NO").Text = WAVE_NO
        'Click_Command("View")
    End Sub

    Private Sub btnAutomate_Click(sender As System.Object, e As System.EventArgs) Handles btnAutomate.Click
        Automate_Pick()
        If MsgBox("Automated Pick is Complete." & vbCrLf & vbCrLf & "Refresh Screen?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
            Close_Picks()
        End If
    End Sub

    Sub Automate_Pick()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now performing Automated Pick")

        Do While txt2.Text <> ""
            C.GetResponseToScan(txt2.Text)
        Loop

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
End Class