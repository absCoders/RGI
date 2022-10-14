Imports Infragistics.Win.UltraWinGrid
Imports System.Text

Public Class WHFPNPK1
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow
    Dim sqlWHTPNPS1 As String = ""

    ' KENNY SAYS TOM SHOULD SEE A MARK IN THE PNP STYLE STATUS SCREEN
    ' Add flag to override EDI Active Status for the warehouse to affect reports.

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "WHFPACKI" Then
            InquiryMode = True
        End If

        Get_PARM("SOTPARM1")
        With dst

            ASCMAIN1.sql = "Select * from ICTWHSE1 where WHSE_CODE in (Select Distinct WHSE_CODE from WHTLOCM1 where LOCATION_USE = 'E')"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select WHTPNPS1.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTCOLR1.COLOR_DESC, ICTSTYC1.UPC_CODE" & vbCrLf _
                & " from WHTPNPS1,ICTSTYL1,ICTCOLR1, ICTSTYC1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = WHTPNPS1.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = WHTPNPS1.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE" & vbCrLf
            sqlWHTPNPS1 = ASCMAIN1.sql
            ASCMAIN1.sql &= " and WHTPNPS1.WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "WHTPNPS1", "**", 0, True, "V", 3)
            With .Tables("WHTPNPS1")
                ' .Columns.Add("UPC_CODE", GetType(System.String))
                .Columns.Add("SET_QTY", GetType(System.Int64))
                .Columns.Add("ECOMM_LOC", GetType(System.String))
                .Columns.Add("WHSE_LOC", GetType(System.String))
                .Columns.Add("QTY_IN_PICK", GetType(System.Int64))
                .Columns.Add("QTY_IN_WHSE", GetType(System.Int64))
                .Columns.Add("QTY_IN_ECOM", GetType(System.Int64))
                .Columns.Add("QTY_SHIPPED", GetType(System.Int64))
                .Columns.Add("QTY_TY", GetType(System.Int64))
                .Columns.Add("QTY_LY", GetType(System.Int64))
                'Change AVAIL to trigger on Zero
                '.Columns.Add("AVAIL", GetType(System.Int64), "(ISNULL(QTY_IN_ECOM,0)-ISNULL(MIN_QTY_ECOM,ISNULL((QTY_IN_WHSE * PCT_QTY_ECOM /100),0))-ISNULL(QTY_IN_PICK,0))*iif(NOT_INSEASON='1',0,1)")
                .Columns.Add("AVAIL", GetType(System.Int64), "(ISNULL(QTY_IN_ECOM,00)-ISNULL(MIN_QTY_ECOM,0)-ISNULL(QTY_IN_PICK,0))*iif(NOT_INSEASON='1',0,1)")
                .Columns.Add("SHORTAGE", GetType(System.Int64)) ', "IIF(AVAIL<0,-1 * Math.Round(AVAIL / SET_QTY) * SET_QTY ,0)")
                .Columns.Add("EDI_STATUS", GetType(System.String))
                .Columns.Add("LOCATION_ROUTE_SEQ", GetType(System.Int64))
                .Columns.Add("LOCATION_ROUTE_2", GetType(System.Int64))
                '.Columns.Add("MAX_AVAIL", GetType(System.Int64), "(ISNULL(QTY_IN_ECOM,0)-ISNULL(MAX_QTY_ECOM,ISNULL((QTY_IN_WHSE * PCT_QTY_ECOM /100),0))-ISNULL(QTY_IN_PICK,0))*iif(NOT_INSEASON='1',0,1)")
                .Columns.Add("MAX_AVAIL", GetType(System.Int64), "(ISNULL(QTY_IN_ECOM,0)-ISNULL(QTY_IN_PICK,0))*iif(NOT_INSEASON='1',0,1)")
            End With
            .Tables("WHTPNPS1").Columns("NOT_INSEASON").DefaultValue = "0"
            .Tables("WHTPNPS1").Columns("SET_QTY").DefaultValue = 1
            .Tables("WHTPNPS1").Columns("LOCATION_ROUTE_SEQ").DefaultValue = 99999

            'ASCMAIN1.sql = "select STYLE_CODE, COLOR_CODE, sum(sotordr2.ORDR_QTY_SHIP) ORDR_QTY_SHIP from sotship1, sotordr1, sotordr2" & vbCrLf _
            '    & " where sotordr1.ordr_no = sotordr2.ordr_no" & vbCrLf _
            '    & " and sotordr1.ORDR_TYPE_CODE = 'B2C'" & vbCrLf _
            '    & " and sotordr1.WHSE_CODE = :PARM3" & vbCrLf _
            '    & " and sotship1.ORDR_GROUP_NO = sotordr1.ORDR_GROUP_NO " & vbCrLf _
            '    & " and sotship1.ship_date_shipped > (sysdate - 31)" & vbCrLf _
            '    & " Group by STYLE_CODE, COLOR_CODE "
            'Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "V")

            ASCMAIN1.sql = "SELECT  i2.style_code, i2.color_code, SUM(i2.ordr_qty_ship) as QTY_TY, LY.QTY as QTY_LY" & vbCrLf _
                & "     FROM sotinvh1 i1, sotinvh2 i2, sotordr1 o1, ICTSTYL3 s3, ictstyl1 s1," & vbCrLf _
                & "        (  SELECT i2.style_code, i2.color_code, SUM(i2.ordr_qty_ship) as QTY" & vbCrLf _
                & "             FROM sotinvh1 i1, sotinvh2 i2, sotordr1 o1, SYS.dual, ictstyl1 s1" & vbCrLf _
                & "             WHERE i1.inv_type = i2.inv_type" & vbCrLf _
                & "                AND i1.inv_no = i2.inv_no" & vbCrLf _
                & "                AND i1.inv_type = 'I'" & vbCrLf _
                & "                AND i1.inv_date < '01-JAN-'||to_char(sysdate, 'YYYY')" & vbCrLf _
                & "                AND i1.inv_date >= '01-JAN-'||to_char(trunc(sysdate, 'yyyy') - interval '1' year, 'YYYY')  " & vbCrLf _
                & "                AND s1.style_code = i2.style_code" & vbCrLf _
                & "                AND i1.ordr_no = o1.ordr_no" & vbCrLf _
                & "                AND o1.ordr_status = 'F'" & vbCrLf _
                & "                AND i1.ordr_type_code = 'B2C'" & vbCrLf _
                & "                AND i2.style_code is not null" & vbCrLf _
                & "                AND i1.whse_code =:PARM1" & vbCrLf _
                & "                group by s1.style_desc, i2.style_code, i2.color_code" & vbCrLf _
                & "                ORDER BY QTY desc ) LY" & vbCrLf _
                & "     WHERE i1.inv_type = i2.inv_type" & vbCrLf _
                & "        AND i1.inv_no = i2.inv_no" & vbCrLf _
                & "        AND i1.inv_type = 'I'" & vbCrLf _
                & "        AND i1.inv_date >= '01-JAN-'||to_char(sysdate, 'YYYY')" & vbCrLf _
                & "        AND s1.style_code = i2.style_code" & vbCrLf _
                & "        AND i1.ordr_no = o1.ordr_no" & vbCrLf _
                & "        AND o1.ordr_status = 'F'" & vbCrLf _
                & "        AND s1.style_code = s3.style_code" & vbCrLf _
                & "        AND i1.ordr_type_code = 'B2C'" & vbCrLf _
                & "        AND i2.style_code is not null" & vbCrLf _
                & "        AND i1.whse_code =:PARM1" & vbCrLf _
                & "        AND i2.style_code = LY.style_code" & vbCrLf _
                & "        AND i2.color_code = LY.color_code" & vbCrLf _
                & "        group by i2.style_code, i2.color_code, LY.QTY"
            Create_TDA(.Tables.Add, "SOTSHIPY", "**", 0, False, "V")

            ASCMAIN1.sql = "select STYLE_CODE, max(NVL(SET_QTY,1)) SET_QTY from ECTESTY1 group by STYLE_CODE"
            Create_TDA(.Tables.Add, "ECTESTY1", "**", 0, False)

            ASCMAIN1.sql = "select distinct ECTESTY2.STYLE_CODE,  ECTESTY2.COLOR_CODE" & vbCrLf _
                & " from ECTESTY2, ECTESTY1 " & vbCrLf _
                & " where ECTESTY1.STYLE_CODE = ECTESTY2.STYLE_CODE " & vbCrLf _
                & " AND NVL(ECTESTY1.SHIP_DROP,'0') = '1' " & vbCrLf _
                & " AND ECTESTY2.ECOM_STYLE_COLOR_STATUS = 'A'"
            Create_TDA(.Tables.Add, "ECTESTY2", "**", 0, False)

            ASCMAIN1.sql = "Select SOTPICK2.*,SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.INIT_DATE, SOTORDR1.ECOM_CODE" & vbCrLf _
                & " from SOTPICK2,SOTPICK1,SOTORDR1,SOTORDR2" & vbCrLf _
                & " where SOTORDR1.ORDR_STATUS = 'P' and SOTORDR1.ECOM_CODE is Not Null and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO and SOTORDR1.WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False, "V")

            'Create_TDA(dst.Tables.Add, "TATCNTRY", "*", 0, False)

            ASCMAIN1.sql = "Select WHTLOCB1.*, WHTLOCM1.LOCATION_USE" & vbCrLf _
                & " from WHTLOCB1,WHTLOCM1" & vbCrLf _
                & " where WHTLOCB1.WHSE_CODE = :PARM1 and WHTLOCB1.STYLE_CODE = :PARM2 and WHTLOCB1.COLOR_CODE = :PARM3 and WHTLOCB1.LOCATION_QTY <> 0" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = WHTLOCB1.WHSE_CODE and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "VVV")

            ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE from SOTPICK1,SOTORDR1" & vbCrLf _
                & " where SOTORDR1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and SOTORDR1.ORDR_STATUS = 'P' and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO and SOTORDR1.ORDR_TYPE_CODE = 'B2C'"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V")


            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC from SOTPICK2,SOTORDR2" & vbCrLf _
                & " where SOTPICK2.PICK_NO = :PARM1" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO " & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO "
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V")

            With .Tables.Add("SOTSTAT1")
                .Columns.Add("XDATE", GetType(System.DateTime))
                .Columns.Add("EDIA_COUNT", GetType(System.Int32))
                .Columns.Add("EDIA_UNITS", GetType(System.Int32))
                .Columns.Add("EDIA_SALES", GetType(System.Int32))
                .Columns.Add("EDIU_COUNT", GetType(System.Int32))
                .Columns.Add("EDIU_UNITS", GetType(System.Int32))
                .Columns.Add("EDIU_SALES", GetType(System.Int32))
                .Columns.Add("OPEN_COUNT", GetType(System.Int32))
                .Columns.Add("OPEN_UNITS", GetType(System.Int32))
                .Columns.Add("OPEN_SALES", GetType(System.Int32))
                .Columns.Add("PICK_COUNT", GetType(System.Int32))
                .Columns.Add("PICK_UNITS", GetType(System.Int32))
                .Columns.Add("PICK_SALES", GetType(System.Int32))
                .Columns.Add("CANC_COUNT", GetType(System.Int32))
                .Columns.Add("CANC_UNITS", GetType(System.Int32))
                .Columns.Add("CANC_SALES", GetType(System.Int32))
                .Columns.Add("SHIP_COUNT", GetType(System.Int32))
                .Columns.Add("SHIP_UNITS", GetType(System.Int32))
                .Columns.Add("SHIP_SALES", GetType(System.Int32))
            End With
            dst.Tables("SOTSTAT1").PrimaryKey = New DataColumn() {dst.Tables("SOTSTAT1").Columns("XDATE")}

            ASCMAIN1.sql = "select distinct o2.EDI_STYLE, o2.EDI_COLOR_CODE from edt846o1 o1, edt846o2 o2" & vbCrLf _
                & "where trunc(o1.edi_report_date) = (select max(trunc(edi_report_date)) from edt846o1 where edi_report_date < = trunc(sysdate))" & vbCrLf _
                & "And o2.EDI_OUTBOUND_DOC_NO = o1.EDI_OUTBOUND_DOC_NO" & vbCrLf _
                & "and o2.EDI_MAINT_TYPE_CODE = '001'"
            Create_TDA(.Tables.Add, "EDT846OX", "**", 0, False)
            dst.Tables("EDT846OX").PrimaryKey = New DataColumn() {dst.Tables("EDT846OX").Columns("EDI_STYLE"), dst.Tables("EDT846OX").Columns("EDI_COLOR_CODE")}

        End With

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdWHTPNPS1.DataSource = dst.Tables("WHTPNPS1")
        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")
        grdWHTLOCB1.DataSource = dst.Tables("WHTLOCB1")

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")

        grdSOTSTAT1.DataSource = dst.Tables("SOTSTAT1")

        Fill_Records("ICTWHSEX")

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")

        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")

        Create_Summary(grdWHTPNPS1, "STYLE_CODE", "Count")
        Create_Summary(grdWHTPNPS1, "QTY_IN_PICK", "Sum")
        Create_Summary(grdWHTPNPS1, "QTY_IN_WHSE", "Sum")
        Create_Summary(grdWHTPNPS1, "QTY_IN_ECOM", "Sum")
        'Create_Summary(grdWHTPNPS1, "QTY_SHIPPED", "Sum")
        Create_Summary(grdWHTPNPS1, "QTY_TY", "Sum")
        Create_Summary(grdWHTPNPS1, "QTY_LY", "Sum")
        'Create_Summary(grdWHTPNPS1, "AVAIL", "Sum")
        Create_Summary(grdWHTPNPS1, "SHORTAGE", "Sum")

        With grdWHTPNPS1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                If gcol.Key = "MIN_QTY_ECOM" Or gcol.Key = "MAX_QTY_ECOM" Or gcol.Key = "PCT_QTY_ECOM" Or gcol.Key = "NOT_INSEASON" Then
                    gcol.CellActivation = Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"STYLE_CODE", "STYLE_DESC", "CARTON_PACK_QTY", "INNER_PACK_QTY", "SET_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"COLOR_CODE", "COLOR_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"MIN_QTY_ECOM", "MAX_QTY_ECOM", "PCT_QTY_ECOM", "NOT_INSEASON"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"SHORTAGE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"MIN_QTY_ECOM"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"QTY_IN_PICK", "QTY_SHIPPED", "AAVAIL"}.Contains(gcol.Key) Then
                    gcol.Hidden = True
                End If
            Next

        End With

        With grdSOTPICK1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
        End With

        With grdSOTPICK2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
        End With

        With grdSOTSTAT1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "XDATE" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Header.Caption = "Date"
                    gcol.Format = "MM/dd"
                Else
                    Dim SFX As String = ""
                    If gcol.Key.EndsWith("COUNT") Then
                        gcol.Width = 50
                        SFX = "Orders"
                    ElseIf gcol.Key.EndsWith("UNITS") Then
                        gcol.Width = 60
                        SFX = "Units"
                    ElseIf gcol.Key.EndsWith("SALES") Then
                        gcol.Width = 70
                        SFX = "$Sales"
                    End If

                    gcol.Format = "#,##0"

                    If gcol.Key.StartsWith("EDIA") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                        gcol.Header.Caption = "All " & SFX
                    ElseIf gcol.Key.StartsWith("EDIU") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        gcol.Header.Caption = "EDIq " & SFX
                    ElseIf gcol.Key.StartsWith("OPEN") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        gcol.Header.Caption = "Open " & SFX
                    ElseIf gcol.Key.StartsWith("PICK") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                        gcol.Header.Caption = "Pick " & SFX
                    ElseIf gcol.Key.StartsWith("CANC") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                        gcol.Header.Caption = "Canc " & SFX
                    ElseIf gcol.Key.StartsWith("SHIP") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LimeGreen
                        gcol.Header.Caption = "Ship " & SFX
                    End If

                    Create_Summary(grdSOTSTAT1, gcol.Key)
                End If
            Next
        End With

        splStats.Panel2Collapsed = True
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    Else
                        ASCMAIN1.sql = "Select Count (*) from WHTLOCM1 where LOCATION_USE = 'E' and WHSE_CODE = '" & rowICTWHSE1.Item("WHSE_CODE") & "'"
                        Dim eCom_Locations_Count As Integer = (ASCDATA1.GetDataValue)
                        If eCom_Locations_Count = 0 Then
                            EMsg &= vbCrLf & "Invalid Value specified for Warehouse (no eCommerce Locations found)"
                        End If
                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")
                    '  If Not ASCMAIN1.Logical_Open("WHTPACK1", WHSE_CODE) Then Exit Sub
                End If

            Case "Update"

                'If dst.Tables("WHTCART1").Compute("Count(PROCESS_STATUS)", "PROCESS_STATUS = 0") Then
                '    EMsg &= vbCrLf & "Found Open Cartons"
                'End If
                'If EMsg = "" And dst.Tables("SOTPICKX").Compute("SUM(VARIANCE)", "") <> 0 Then
                '    If MsgBox("Select 'Yes' to force close this shipment", vbYesNoCancel, "Not in Balance") <> vbYes Then
                '        EMsg &= vbCrLf & "Shipment out of balance"
                '    End If
                'End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                'Update_Record()
                Mode_Settings(False)

            Case "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Replenishment()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode

                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Update").Visible = False

                End With

                .Groups("Refresh").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        splShipments.Visible = ScreenMode

        With grdWHTPNPS1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ICTWHSEX", "WHTPNPS1", "SOTPICKX", "SOTPICK1", "SOTPICK2", "SOTSTAT1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("ICTWHSEX")
        Sort_grdColumns(grdICTWHSEX, "WHSE_CODE")
        '     Setup_tab0()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Load_Data()

    End Sub

    Function Get_WHTPNPS1(STYLE_CODE As String, COLOR_CODE As String) As DataRow
        Dim rowWHTPNPS1 As DataRow = dst.Tables("WHTPNPS1").Rows.Find(New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
        If rowWHTPNPS1 Is Nothing Then
            rowWHTPNPS1 = dst.Tables("WHTPNPS1").Rows.Add(New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE}, True)
            ASCMAIN1.sql = "Select * from WHTLOCB1,WHTLOCM1 Where WHTLOCB1.WHSE_CODE =:PARM1 and WHTLOCB1.STYLE_CODE =:PARM2 and WHTLOCB1.COLOR_CODE = :PARM3" & vbCrLf _
                        & "and WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE and WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE and WHTLOCM1.LOCATION_USE = 'A' and WHTLOCB1.LOCATION_QTY <> 0"
            Dim rowWHTLOCB1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            rowWHTPNPS1.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
            rowWHTPNPS1.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")
            rowWHTPNPS1.Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
            rowWHTPNPS1.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
            rowWHTPNPS1.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE")
            If Not IsNothing(rowWHTLOCB1) Then
                rowWHTPNPS1.Item("WHSE_LOC") = rowWHTLOCB1.Item("LOCATION_CODE")
                rowWHTPNPS1.Item("LOCATION_ROUTE_SEQ") = rowWHTLOCB1.Item("LOCATION_ROUTE_SEQ")
            End If
            Dim rowICTSTAT2 As DataRow = LookUp("ICTSTAT2", New String() {STYLE_CODE, COLOR_CODE, WHSE_CODE})
            If rowICTSTAT2 IsNot Nothing Then
                rowWHTPNPS1.Item("QTY_IN_WHSE") = rowICTSTAT2.Item("WHSE_QTY_ON_HAND")
            End If
        End If
        'If IsDBNull(rowWHTPNPS1.Item("UPC_CODE")) Then
        '    Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE}, True)
        '    rowWHTPNPS1.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE")
        'End If
        Return rowWHTPNPS1
    End Function
    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Update")

        ' Update_Record_TDA("SOTCART2")

        CommitTrans("Update Complete")

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
        Load_Popup_Menu(grdWHTPNPS1, "SBBBB", "Show Filter", "Style Status Inquiry", "Location Inquiry", "Check Not InSeason", "UnCheck Not InSeason")
        Load_Popup_Menu(grdSOTPICK1, "B", "Show Pick Ticket")
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


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Location Inquiry"
                Dim KEY As String = ""
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                KEY = "S:" & STYLE_CODE

                Context_Launch("Select", KEY, e.Tool.Key, "WHFLOCS1")


            Case "Show Pick Ticket"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Printing Pick Ticket")

                Dim REPORT_NAME As String = "SORPICKE"

                Print_Report_Begin()
                Generate_Report(REPORT_NAME, "eCommerce Pick Ticket", "")
                Print_Report_End()

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

                'Case "Style Master"
                '    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                '    If rowICTSTYL1 IsNot Nothing Then
                '        Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                '    End If
            Case "Check Not InSeason"
                For Each grow As UltraWinGrid.UltraGridRow In grdWHTPNPS1.Selected.Rows
                    Dim row As DataRow = dst.Tables("WHTPNPS1").Rows.Find(New Object() {grow.Cells("WHSE_CODE").Value, grow.Cells("STYLE_CODE").Value, grow.Cells("COLOR_CODE").Value})
                    row.Item("NOT_INSEASON") = 1
                Next
                grdWHTPNPS1.Selected.Rows.Clear()
                Update_Record_TDA("WHTPNPS1")

            Case "UnCheck Not InSeason"
                For Each grow As UltraWinGrid.UltraGridRow In grdWHTPNPS1.Selected.Rows
                    Dim row As DataRow = dst.Tables("WHTPNPS1").Rows.Find(New Object() {grow.Cells("WHSE_CODE").Value, grow.Cells("STYLE_CODE").Value, grow.Cells("COLOR_CODE").Value})
                    row.Item("NOT_INSEASON") = 0
                Next
                grdWHTPNPS1.Selected.Rows.Clear()
                Update_Record_TDA("WHTPNPS1")
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
        Setup_tab1()
    End Sub

    Sub Setup_tab1()
        'lblPickFilter.Visible = (tab1.SelectedTab.Key = "Pick Tickets")
        'optPickFilter.Visible = (tab1.SelectedTab.Key = "Pick Tickets")
    End Sub

    Private Sub grdWHTPNPS1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTPNPS1.AfterRowActivate
        Dim STYLE_CODE As String = grdWHTPNPS1.ActiveRow.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = grdWHTPNPS1.ActiveRow.Cells("COLOR_CODE").Value & ""

        grdSOTPICKX.Text = "Pick Tickets for Style-Color " & STYLE_CODE & "-" & COLOR_CODE
        Dim dvw As DataView = DirectCast(grdSOTPICKX.DataSource, DataTable).DefaultView
        dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

        grdWHTLOCB1.Text = "Locations for Style-Color " & STYLE_CODE & "-" & COLOR_CODE
        Fill_Records("WHTLOCB1", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPICK1.AfterRowActivate
        Load_SOTPICK2()
    End Sub

    Sub Load_SOTPICK2()
        If grdSOTPICK1.ActiveRow Is Nothing OrElse grdSOTPICK1.ActiveRow.IsFilterRow OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdSOTPICK2.Visible = False
        Else
            grdSOTPICK2.Visible = True
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value & ""
            Fill_Records("SOTPICK2", PICK_NO)
            Sort_grdColumns(grdSOTPICK2, "PICK_LNO")
            grdSOTPICK2.Text = "Pick Ticket Details for Pick No " & PICK_NO
        End If
    End Sub

    Private Sub grdSOTSTAT1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTSTAT1.InitializeRow
        If e.Row.IsDataRow AndAlso Format(e.Row.Cells("XDATE").Value, "MM/dd/yyyy") = Format(Now.Date, "MM/dd/yyyy") Then
            'e.Row.Appearance.ForeColor = Drawing.Color.Blue
            'e.Row.Appearance.BackColor = Drawing.Color.Yellow

            For Each dcol As DataColumn In dst.Tables("SOTSTAT1").Columns
                e.Row.Cells(dcol.ColumnName).Appearance.ForeColor = Drawing.Color.Blue
                e.Row.Cells(dcol.ColumnName).Appearance.BackColor = Drawing.Color.Yellow
            Next
        End If
    End Sub


    Sub Refresh_Stats()

        Dim SQLD As String = ""
        Dim DT As String = "SOTORDR1.ORDR_DATE"

        dst.Tables("SOTSTAT1").Rows.Clear()
        For I As Integer = -11 To 11
            Dim XDATE As Date = Now.Date.AddDays(I)
            Dim X As String = Format(XDATE, "dd-MMM-yyyy")
            If I = -11 Then
                SQLD &= "CASE WHEN " & DT & " <= '" & X & "' THEN TO_DATE('" & X & "') ELSE "
            ElseIf I = 11 Then
                SQLD &= "CASE WHEN " & DT & " >= '" & X & "' THEN TO_DATE('" & X & "') ELSE "
            End If
            dst.Tables("SOTSTAT1").Rows.Add(New Object() {XDATE})
        Next
        Sort_grdColumns(grdSOTSTAT1, "XDATE")

        SQLD &= " " & DT & " END END"

        For Each T As String In New String() {"OPEN", "PICK", "EDIA", "EDIU"}
            Select Case T
                Case "OPEN"
                    ASCMAIN1.sql = "" & vbCrLf _
                        & ", Count (Distinct SOTORDR1.ORDR_NO) OPEN_COUNT, Sum (ORDR_QTY_OPEN) OPEN_UNITS, Sum (ORDR_QTY_OPEN * ORDR_UNIT_PRICE) OPEN_SALES" & vbCrLf _
                        & " from SOTORDR1,SOTORDR2" & vbCrLf _
                        & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                        & "   and SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
                        & "   and SOTORDR1.ORDR_TYPE_CODE = 'B2C'" & vbCrLf _
                        & "   and SOTORDR1.WHSE_CODE = '" & WHSE_CODE & "'"

                Case "PICK"
                    ASCMAIN1.sql = "" & vbCrLf _
                        & ", Count (Distinct SOTPICK1.PICK_NO) PICK_COUNT, Sum (PICK_QTY) PICK_UNITS, Sum (PICK_QTY * PICK_UNIT_PRICE) PICK_SALES" & vbCrLf _
                        & " from SOTPICK1,SOTPICK2,SOTORDR1" & vbCrLf _
                        & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                        & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                        & "   and SOTORDR1.ORDR_TYPE_CODE = 'B2C'" & vbCrLf _
                        & "   and SOTORDR1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                        & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"

                Case "EDIA", "EDIU"
                    ASCMAIN1.sql = "" & vbCrLf _
                        & ", Count (Distinct EDT850T1.EDI_DOC_SEQ_NO) EDIA_COUNT, Sum (EDI_TOTAL_QTY) EDIA_UNITS, Sum (EDI_TOTAL_QTY * EDI_PRICE) EDIA_SALES" _
                        & " from EDT850T1,EDT850T2" & vbCrLf _
                        & " where EDT850T1.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO" & vbCrLf _
                        & IIf(T = "EDIU",
                              "   and NVL(EDT850T1.EDI_PROCESS_IND,'0') = '0'",
                              "   and NVL(EDT850T1.EDI_PROCESS_IND,'0') IN ('0','1')") & vbCrLf _
                        & "   and TRIM(EDT850T1.EDI_PO_TYPE) = 'DS'" & vbCrLf _
                        & "   and TRIM(EDT850T1.EDI_SUPPLIER_NO) = '23249'"

                    If T = "EDIU" Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDIA", "EDIU")

                    SQLD = Replace(SQLD, DT, "EDT850T1.EDI_PO_DATE")
            End Select

            ASCMAIN1.sql = "Select " & SQLD & " XDATE" & ASCMAIN1.sql & vbCrLf & " group by " & SQLD

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim XDATE As Date = row.Item("XDATE")

                Dim rowSOTSTAT1 As DataRow = dst.Tables("SOTSTAT1").Rows.Find(XDATE)

                rowSOTSTAT1.Item(T & "_COUNT") = row.Item(T & "_COUNT")
                rowSOTSTAT1.Item(T & "_UNITS") = row.Item(T & "_UNITS")
                rowSOTSTAT1.Item(T & "_SALES") = row.Item(T & "_SALES")
            Next
        Next
    End Sub

    Private Sub grdWHTLOCB1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTLOCB1.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("LOCATION_USE").Value & "" = "E" Then
                e.Row.CellAppearance.BackColor = Drawing.Color.Yellow
            Else
                e.Row.CellAppearance.BackColor = Drawing.Color.Empty
            End If
        End If
    End Sub


    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Load_Data()
    End Sub

    Private Sub tab1_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab1.SelectedTabChanged
        Setup_tab1()
    End Sub

    Sub Load_Data()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Fill_Records("WHTPNPS1", WHSE_CODE)
        Fill_Records("SOTPICKX", WHSE_CODE)
        Fill_Records("SOTPICK1", WHSE_CODE)
        'Fill_Records("SOTSHIPX", WHSE_CODE)
        Fill_Records("SOTSHIPY", WHSE_CODE)
        Fill_Records("ECTESTY1")
        Fill_Records("ECTESTY2")
        Fill_Records("EDT846OX")
        Show_Filter(grdWHTPNPS1, True)
        For Each ROW As DataRow In dst.Tables("ECTESTY2").Select()
            Dim STYLE_CODE As String = ROW.Item("STYLE_CODE")
            Dim COLOR_CODE As String = ROW.Item("COLOR_CODE")
            Dim rowWHTPNPS1 As DataRow = Get_WHTPNPS1(STYLE_CODE, COLOR_CODE)
        Next

        For Each ROW As DataRow In dst.Tables("SOTPICKX").Select()
            Dim STYLE_CODE As String = ROW.Item("STYLE_CODE")
            Dim COLOR_CODE As String = ROW.Item("COLOR_CODE")
            Dim rowWHTPNPS1 As DataRow = Get_WHTPNPS1(STYLE_CODE, COLOR_CODE)
            rowWHTPNPS1.Item("QTY_IN_PICK") = Val(rowWHTPNPS1.Item("QTY_IN_PICK") & "") + Val(ROW.Item("PICK_QTY") & "")
        Next

        ASCMAIN1.sql = "Select WHTLOCB1.*, LOCATION_ROUTE_SEQ from WHTLOCB1,WHTLOCM1" & vbCrLf _
            & " where WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE" & vbCrLf _
            & "   and WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE" & vbCrLf _
            & "   and WHTLOCM1.LOCATION_USE = 'E' and WHTLOCB1.LOCATION_QTY <> 0"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim STYLE_CODE As String = ROW.Item("STYLE_CODE")
            Dim COLOR_CODE As String = ROW.Item("COLOR_CODE")
            Dim rowWHTPNPS1 As DataRow = Get_WHTPNPS1(STYLE_CODE, COLOR_CODE)
            rowWHTPNPS1.Item("QTY_IN_ECOM") = Val(rowWHTPNPS1.Item("QTY_IN_ECOM") & "") + Val(ROW.Item("LOCATION_QTY") & "")
            rowWHTPNPS1.Item("ECOMM_LOC") = ROW.Item("LOCATION_CODE")
            rowWHTPNPS1.Item("LOCATION_ROUTE_2") = ROW.Item("LOCATION_ROUTE_SEQ")
        Next

        ASCMAIN1.sql = "Select WHTLOCB1.*, LOCATION_ROUTE_SEQ from WHTLOCB1,WHTLOCM1,WHTPNPS1" & vbCrLf _
           & " where WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE" & vbCrLf _
           & "   and WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE" & vbCrLf _
           & "   and WHTLOCM1.LOCATION_USE = 'A' and WHTLOCB1.LOCATION_QTY <> 0" & vbCrLf _
           & "   and WHTLOCB1.WHSE_CODE = WHTPNPS1.WHSE_CODE" & vbCrLf _
           & "   and WHTLOCB1.STYLE_CODE = WHTPNPS1.STYLE_CODE" & vbCrLf _
           & "   and WHTLOCB1.COLOR_CODE = WHTPNPS1.COLOR_CODE" & vbCrLf
        Dim lstRec As String = ""
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
            If lstRec <> ROW.Item("STYLE_CODE") & ROW.Item("COLOR_CODE") Then
                Dim STYLE_CODE As String = ROW.Item("STYLE_CODE")
                Dim COLOR_CODE As String = ROW.Item("COLOR_CODE")
                Dim rowWHTPNPS1 As DataRow = Get_WHTPNPS1(STYLE_CODE, COLOR_CODE)
                rowWHTPNPS1.Item("WHSE_LOC") = ROW.Item("LOCATION_CODE")
                rowWHTPNPS1.Item("LOCATION_ROUTE_SEQ") = ROW.Item("LOCATION_ROUTE_SEQ")
            End If
        Next

        ASCMAIN1.sql = "select WHTPNPS1.WHSE_CODE, WHTPNPS1.STYLE_CODE, WHTPNPS1.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND " & vbCrLf _
            & " from WHTPNPS1, ICTSTAT2 " & vbCrLf _
            & " where WHTPNPS1.WHSE_CODE = ICTSTAT2.WHSE_CODE" & vbCrLf _
            & "   and WHTPNPS1.STYLE_CODE = ICTSTAT2.STYLE_CODE" & vbCrLf _
            & "   and WHTPNPS1.COLOR_CODE = ICTSTAT2.COLOR_CODE"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim STYLE_CODE As String = ROW.Item("STYLE_CODE")
            Dim COLOR_CODE As String = ROW.Item("COLOR_CODE")
            Dim rowWHTPNPS1 As DataRow = Get_WHTPNPS1(STYLE_CODE, COLOR_CODE)
            rowWHTPNPS1.Item("QTY_IN_WHSE") = Val(rowWHTPNPS1.Item("QTY_IN_WHSE") & "") + Val(ROW.Item("WHSE_QTY_ON_HAND") & "")
        Next

        'Dim loc_sql = "Select WHTLOCM1.LOCATION_CODE, WHTLOCM1.LOCATION_ROUTE_SEQ from WHTLOCB1,WHTLOCM1" & vbCrLf _
        '    & " where WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE" & vbCrLf _
        '    & "   and WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE" & vbCrLf _
        '    & "   and WHTLOCM1.LOCATION_USE = 'A' and WHTLOCB1.LOCATION_QTY <> 0" & vbCrLf _
        '    & "  and WHTLOCB1.STYLE_CODE = :PARM1 " & vbCrLf _
        '    & "  and WHTLOCB1.COLOR_CODE = :PARM2 " & vbCrLf _
        '    & "   and WHTLOCB1.WHSE_CODE = :PARM3 " & vbCrLf _
        '    & "  and rownum = 1"

        For Each ROW As DataRow In dst.Tables("WHTPNPS1").Select()
            Dim STYLE_CODE As String = ROW.Item("STYLE_CODE")
            Dim COLOR_CODE As String = ROW.Item("COLOR_CODE")
            Dim rowWHTPNPS1 As DataRow = Get_WHTPNPS1(STYLE_CODE, COLOR_CODE)
            'Dim rowICTSTAT2 As DataRow = LookUp("ICTSTAT2", New String() {STYLE_CODE, COLOR_CODE, WHSE_CODE}, True)
            'rowWHTPNPS1.Item("QTY_IN_WHSE") = Val(rowWHTPNPS1.Item("QTY_IN_WHSE") & "") + Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "")
            'rowWHTPNPS1.Item("QTY_SHIPPED") = dst.Tables("SOTSHIPX").Compute("SUM(ORDR_QTY_SHIP)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
            rowWHTPNPS1.Item("QTY_TY") = dst.Tables("SOTSHIPY").Compute("SUM(QTY_TY)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
            rowWHTPNPS1.Item("QTY_LY") = dst.Tables("SOTSHIPY").Compute("SUM(QTY_LY)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
            rowWHTPNPS1.Item("SET_QTY") = dst.Tables("ECTESTY1").Compute("SUM(SET_QTY)", "STYLE_CODE = '" & STYLE_CODE & "'")
            'Dim rowWHTLOCM1 As DataRow = ASCDATA1.GetDataRow(loc_sql, "VVV", New String() {STYLE_CODE, COLOR_CODE, WHSE_CODE})
            'If Not IsNothing(rowWHTLOCM1) Then
            '    rowWHTPNPS1.Item("WHSE_LOC") = rowWHTLOCM1.Item("LOCATION_CODE")
            '    rowWHTPNPS1.Item("LOCATION_ROUTE_SEQ") = rowWHTLOCM1.Item("LOCATION_ROUTE_SEQ")
            'Else
            '    rowWHTPNPS1.Item("WHSE_LOC") = ""
            '    rowWHTPNPS1.Item("LOCATION_ROUTE_SEQ") = 99999
            'End If

            Dim SET_QTY As Int64 = Val(rowWHTPNPS1.Item("SET_QTY") & "")
            If SET_QTY = 0 Then SET_QTY = 1
            If String.IsNullOrEmpty(rowWHTPNPS1.Item("MIN_QTY_ECOM") & "") Then
                Dim Inner As Int64 = Val(rowWHTPNPS1.Item("INNER_PACK_QTY") & "")
                If Inner > SET_QTY Then
                    rowWHTPNPS1.Item("MIN_QTY_ECOM") = SET_QTY
                    rowWHTPNPS1.Item("MAX_QTY_ECOM") = Inner
                Else
                    rowWHTPNPS1.Item("MIN_QTY_ECOM") = Inner
                    rowWHTPNPS1.Item("MAX_QTY_ECOM") = SET_QTY
                End If
                If rowWHTPNPS1.Item("MIN_QTY_ECOM") = 0 Then
                    rowWHTPNPS1.Item("MIN_QTY_ECOM") = rowWHTPNPS1.Item("MAX_QTY_ECOM")
                End If
            End If

            Dim AVAIL As Int64 = rowWHTPNPS1.Item("AVAIL")
            Dim SHORTAGE As Int64 = IIf(AVAIL < 0, -1 * Math.Round(AVAIL / SET_QTY) * SET_QTY, 0)
            rowWHTPNPS1.Item("SHORTAGE") = SHORTAGE

            Dim rowEDT846OX As DataRow = dst.Tables("EDT846OX").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If Not IsNothing(rowEDT846OX) Then
                rowWHTPNPS1.Item("EDI_STATUS") = "Active"
            Else
                rowWHTPNPS1.Item("EDI_STATUS") = "In-Active"
            End If
        Next
        UpdategrdWHTPNPS1View()

        Sort_grdColumns(grdWHTPNPS1, "AVAIL, STYLE_CODE,COLOR_CODE")
        'Refresh_Stats()

        If dst.Tables("WHTPNPS1").Compute("Count(STYLE_CODE)", "QTY_IN_PICK > 0 and NOT_INSEASON = '1'") > 0 Then
            MsgBox("Found Items with Picks flaged out of season", MsgBoxStyle.Exclamation, "Warning")
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_Replenishment()

        'WE want to get rid of nulls for QTY's
        For Each row As DataRow In dst.Tables("WHTPNPS1").Select("QTY_IN_ECOM is null or MAX_QTY_ECOM is null")
            'Debug.Print(row("QTY_IN_ECOM") & "-" & row("MAX_QTY_ECOM"))
            If String.IsNullOrEmpty(row("QTY_IN_ECOM").ToString()) Then
                row("QTY_IN_ECOM") = 0
            End If
            If String.IsNullOrEmpty(row("MAX_QTY_ECOM").ToString()) Then
                row("MAX_QTY_ECOM") = 0
            End If
        Next

        For Each row As DataRow In dst.Tables("WHTPNPS1").Select("(QTY_IN_ECOM <> 0 and SHORTAGE < 1) and (EDI_STATUS = 'In-Active' or NOT_INSEASON = '1' or QTY_IN_ECOM > MAX_QTY_ECOM)")
            ' Swap Route Sequences for putback
            row.Item("LOCATION_ROUTE_SEQ") = Val(ASCDATA1.GetDataValue("SELECT LOCATION_ROUTE_SEQ from WHTLOCM1 where LOCATION_CODE = '" & row.Item("ECOMM_LOC") & "'")) + 0
            row.Item("LOCATION_ROUTE_2") = Val(ASCDATA1.GetDataValue("SELECT LOCATION_ROUTE_SEQ from WHTLOCM1 where LOCATION_CODE = '" & row.Item("WHSE_LOC") & "'")) + 0
        Next

        'NOTE THAT THIS PRINT ROUTINE WAS USING THE DATA LAYER & DST THAT IS ASSOCIATED WITH THIS FORM   
        Dim sql As String = ""
        Dim SubTitle As String = IIf(chkPICKONLY.Checked, "Fill Pick Qty", "")

        Print_Report_Begin()
        CR_params.Add("SUBT", SubTitle)
        Dim RPT As String = "WHRPNPK1"

        If chkPICKONLY.Checked Then
            Generate_Report(RPT, "E-Comm Styles - Replenish", "", "{WHTPNPS1.QTY_IN_WHSE} > 4 and {WHTPNPS1.SHORTAGE} > 0 and {WHTPNPS1.EDI_STATUS} = 'Active' and {WHTPNPS1.QTY_IN_ECOM} + 1 < {WHTPNPS1.MAX_QTY_ECOM}")
        Else
            Generate_Report(RPT, "E-Comm Styles - Replenish", "", "{WHTPNPS1.QTY_IN_WHSE} > 4 and {WHTPNPS1.SHORTAGE} > 0 and {WHTPNPS1.EDI_STATUS} = 'Active' and {WHTPNPS1.QTY_IN_WHSE} > {WHTPNPS1.QTY_IN_ECOM}")

            Generate_Report(RPT, "E-Comm Styles - Putback", "", " ({WHTPNPS1.QTY_IN_ECOM} <> 0 and {WHTPNPS1.SHORTAGE} < 1 and {WHTPNPS1.AVAIL} >= 0) and ({WHTPNPS1.EDI_STATUS} = 'In-Active' or {WHTPNPS1.NOT_INSEASON} = '1' or {WHTPNPS1.QTY_IN_ECOM} > {WHTPNPS1.MAX_QTY_ECOM})")
        End If
        Print_Report_End()

    End Sub
    Private Sub grdWHTPNPS1_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdWHTPNPS1.BeforeCellUpdate
        If e.Cell.Column.Key = "CART_SEQ" Then
            Dim PICK_NO As String = grdWHTPNPS1.ActiveRow.Cells("PICK_NO").Value
            If grdWHTPNPS1.ActiveRow.Cells("PROCESS_STATUS").Value = 0 Then
                MsgBox("Carton Is Open, wait for close.", vbInformation, "Cannot Update")
                e.Cancel = True
            ElseIf ASCDATA1.GetDataValue("Select Count(1) FROM WHTCART1 where PICK_NO = '" & PICK_NO & "' and CART_SEQ = " & e.NewValue.ToString) > 0 Then
                MsgBox("Value in use, check again", vbInformation, "Cannot Update")
                e.Cancel = True
            End If
        End If

    End Sub

    Private Sub grdWHTPNPS1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdWHTPNPS1.AfterRowUpdate
        'Dim CART_NO As String = grdWHTPNPS1.ActiveRow.Cells("CART_NO").Value
        'ASCDATA1.ExecuteSQL("Update WHTCART1 Set CART_SEQ = " & grdWHTCART1.ActiveRow.Cells("CART_SEQ").Value & " where CART_NO = '" & CART_NO & "'")
        Update_Record_TDA("WHTPNPS1")

    End Sub
    Private Sub grdWHTPNPS1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTPNPS1.InitializeRow
        If e.Row.IsDataRow Then
            If Val(e.Row.Cells("SHORTAGE").Value & "") > 0 And e.Row.Cells("EDI_STATUS").Value & "" = "Active" And e.Row.Cells("NOT_INSEASON").Value & "" <> "1" Then
                e.Row.CellAppearance.BackColor = Drawing.Color.OrangeRed
            ElseIf e.Row.Cells("EDI_STATUS").Value & "" = "In-Active" Or e.Row.Cells("NOT_INSEASON").Value & "" = "1" Then
                e.Row.CellAppearance.BackColor = Drawing.Color.DarkSalmon
            Else
                e.Row.CellAppearance.BackColor = Drawing.Color.Empty
            End If
        End If
    End Sub
    Private Sub grdSOTPICK1_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSOTPICK1.InitializeLayout

    End Sub

    Private Sub grdSOTPICK1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPICK1.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("PICK_STATUS").Value & "" = "F" Then
                e.Row.CellAppearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.CellAppearance.BackColor = Drawing.Color.Empty
            End If
        End If
    End Sub

    Private Sub chkPICKONLY_CheckedChanged(sender As Object, e As EventArgs) Handles chkPICKONLY.CheckedChanged
        UpdategrdWHTPNPS1View()
    End Sub

    Private Sub chkNotInSSN_CheckedChanged(sender As Object, e As EventArgs) Handles chkNotInSSN.CheckedChanged
        UpdategrdWHTPNPS1View()
    End Sub

    Private Sub UpdategrdWHTPNPS1View()
        Dim dvw As DataView = DirectCast(grdWHTPNPS1.DataSource, DataTable).DefaultView
        If Not chkNotInSSN.Checked Then
            dvw.RowFilter = "QTY_IN_ECOM > 0 or (EDI_STATUS = 'Active' and QTY_IN_WHSE > 4)"
        Else
            dvw.RowFilter = "QTY_IN_ECOM > 0 or (EDI_STATUS = 'Active' and QTY_IN_WHSE > 4)  and NOT_INSEASON = 0"
        End If
    End Sub

End Class