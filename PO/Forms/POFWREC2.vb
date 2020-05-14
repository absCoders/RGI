Public Class POFWREC2

    Dim PO_SHIPMENT_NO As String
    Dim rowPOTSHIP1 As DataRow
    Dim POTSHIP3 As String
    Dim ICTSTYL1_Recent As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Get_PARM("POTPARM1")

        With dst
            ASCMAIN1.sql = "Select POTSHIP1.*" & vbCrLf _
            & ", POTSHIP2.CONTAINER_NO, POTSHIP2.PO_SHIP_STATUS" & vbCrLf _
            & ", POTSHIP2.PO_SHIPMENT_LNO, POTSHIP2.BOL_NO, POTSHIP2.PO_SHIP_CTNS, POTSHIP2.CBM" & vbCrLf _
            & " from POTSHIP1,POTSHIP2 " & vbCrLf _
            & " where POTSHIP2.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO" & vbCrLf _
            & IIf(ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA", " and POTSHIP1.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")", "") & vbCrLf _
            & "   and POTSHIP2.PO_SHIP_STATUS = 'O' "
            Create_TDA(.Tables.Add, "POTSHIPX", "**", 0, False, "", 0)
            .Tables("POTSHIPX").Columns.Add("CONTAINER_TYPE_CODE")
            .Tables("POTSHIPX").Columns.Add("CONTAINER_SEAL_NO")


            If ASCMAIN1.CLIENT = "RGI" Then
                ' ASCMAIN1.Temp_Table
                ASCMAIN1.sql = "Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & ", POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO, POTSHIP3.PO_QTY_SHP" & vbCrLf _
                & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR2.PO_QTY_UOM" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_UNIT_PACK_QTY" & vbCrLf _
                & ", ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                & ", ICTSTYL1.STYLE_WEIGHT , ICTSTYL1.STYLE_UOM, WHTPREC3.LOCATION_CODE STYLE_BIN, ICTSTYC1.UPC_CODE " & vbCrLf _
                & ", POTORDR1.PO_CARTON_MARKS, POTORDR1.CUST_CODE, ARTCUST1.CUST_NAME  " & vbCrLf _
                & " from POTSHIP3,POTORDR2,POTORDR1,ICTSTYL1, ICTSTYC1, ARTCUST1, WHTPREC3" & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = POTORDR2.COLOR_CODE" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR1.CUST_CODE = ARTCUST1.CUST_CODE (+)" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO (+)= :PARM1" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = :PARM2" & vbCrLf _
                & "   and WHTPREC3.PO_SHIPMENT_NO (+)= POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and WHTPREC3.PO_SHIPMENT_LNO (+)= POTSHIP3.PO_SHIPMENT_LNO " & vbCrLf _
                & "   and WHTPREC3.PO_ORDER_NO (+)= POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and  WHTPREC3.PO_ORDER_LNO (+)= POTSHIP3.PO_ORDER_LNO"
            Else
                ASCMAIN1.sql = "Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & ", POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO, POTSHIP3.PO_QTY_SHP" & vbCrLf _
                & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR2.PO_QTY_UOM" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_UNIT_PACK_QTY" & vbCrLf _
                & ", ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                & ", ICTSTYL1.STYLE_WEIGHT , ICTSTYL1.STYLE_UOM, STYLE_BIN, ICTSTYC1.UPC_CODE " & vbCrLf _
                & ", POTORDR1.CUST_CODE, ARTCUST1.CUST_NAME  " & vbCrLf _
                & " from POTSHIP3,POTORDR2,POTORDR1,ICTSTYL1, ICTSTYC1, ARTCUST1" & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = POTORDR2.COLOR_CODE" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR1.CUST_CODE = ARTCUST1.CUST_CODE (+)" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO (+)= :PARM1" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = :PARM2"
            End If

            Create_TDA(.Tables.Add, "POTSHIP3", "**", 0, False, "VN", 4)
            .Tables("POTSHIP3").Columns.Add("UNITS", GetType(System.Decimal), _
                "IIF(ISNULL(PO_QTY_UOM,0)=1,PO_QTY_SHP,PO_QTY_SHP * PO_QTY_UOM / IIF(ISNULL(SUB_UNIT_PACK_QTY,0)=0,1,ISNULL(SUB_UNIT_PACK_QTY,0)))")
            .Tables("POTSHIP3").Columns.Add("CARTON_NOS")
            .Tables("POTSHIP3").Columns.Add("CARTONS", GetType(System.Int32), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,PO_QTY_SHP/CARTON_PACK_QTY)")


            'If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            '    ASCMAIN1.sql = "Select * from ICTSTYL1"
            '    dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

            '    ASCMAIN1.sql = "Select * from ICTSTYC1"
            '    dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYC1"))
            'End If


            ASCMAIN1.sql = "Select POTSHIP7.*" & vbCrLf _
                & " from POTSHIP7" & vbCrLf _
                & " where POTSHIP7.PO_SHIPMENT_NO = :PARM1" & vbCrLf _
                & "   and POTSHIP7.PO_SHIPMENT_LNO = :PARM2"
            Create_TDA(.Tables.Add, "POTSHIP7", "**", 0, False, "VN")

            ASCMAIN1.sql = "Select POTSHIP8.*" & vbCrLf _
                & " from POTSHIP8" & vbCrLf _
                & " where POTSHIP8.PO_SHIPMENT_NO = :PARM1" & vbCrLf _
                & "   and POTSHIP8.PO_SHIPMENT_LNO = :PARM2"
            Create_TDA(.Tables.Add, "POTSHIP8", "**", 0, False, "VN")



            ASCMAIN1.sql = "SELECT POTORDR1.*,SOTORDR1.CUST_NAME,SOTORDR1.ORDR_CANCEL_DATE from POTORDR1,SOTORDR1 where SOTORDR1.ORDR_NO (+) = POTORDR1.ORDR_NO"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "", 0)
            With .Tables("POTORDRX")
                .Columns.Add("PO_QTY_ORD", GetType(System.Int64))
                .Columns.Add("PO_QTY_SHP", GetType(System.Int64))
                .Columns.Add("PO_QTY_REC", GetType(System.Int64))
                .Columns.Add("PO_QTY_OPN", GetType(System.Int64))
                .Columns.Add("PO_AMT_ORD", GetType(System.Decimal))
                .Columns.Add("PO_AMT_SHP", GetType(System.Decimal))
                .Columns.Add("PO_AMT_REC", GetType(System.Decimal))
                .Columns.Add("PO_AMT_OPN", GetType(System.Decimal))
                .Columns.Add("PO_LINES_CONF", GetType(System.Int64))
                .Columns.Add("PO_LINES", GetType(System.Int64))
                .Columns.Add("PO_CTNS_ORD", GetType(System.Decimal))
                .Columns.Add("PO_CTNS_SHP", GetType(System.Decimal))
                .Columns.Add("PO_CTNS_OPN", GetType(System.Decimal))
                .Columns.Add("PO_CUBE_ORD", GetType(System.Decimal))
                .Columns.Add("PO_CUBE_SHP", GetType(System.Decimal))
                .Columns.Add("PO_CUBE_OPN", GetType(System.Decimal))
                .Columns.Add("PO_DATE_SHIP_BY_MIN", GetType(System.DateTime))
                .Columns.Add("PO_DATE_ETA_MIN", GetType(System.DateTime))
                .Columns.Add("PO_DATE_SHIP_BY_MAX", GetType(System.DateTime))
                .Columns.Add("PO_DATE_ETA_MAX", GetType(System.DateTime))
            End With




            ASCMAIN1.sql = "Select ICTSTYL1.*,ICTBODY2.MASTER_BODY_CODE" & vbCrLf _
                & " from ICTSTYL1,ICTBODY2" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = :PARM1" & vbCrLf _
                & "   and ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE" & vbCrLf
            For Each TABLE_NAME As String In New String() {"ICTSTYL1_RECENT"}
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "V", 0)

                .Tables(TABLE_NAME).Columns.Add("LAST_ORDR_DATE", GetType(System.DateTime))
                .Tables(TABLE_NAME).Columns.Add("LAST_ORDR_NO")
                .Tables(TABLE_NAME).Columns.Add("LAST_ORDR_CUST_CODE")
                .Tables(TABLE_NAME).Columns.Add("LAST_ORDR_CUST_PO")

                .Tables(TABLE_NAME).Columns.Add("QTY_ONHD", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_ONPO", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_TRAN", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_OPEN", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_PICK", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_COMM", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_PROD", GetType(System.Int64))
                .Tables(TABLE_NAME).Columns.Add("QTY_NETA", GetType(System.Int64), "ISNULL(QTY_ONHD,0) + ISNULL(QTY_ONPO,0) + ISNULL(QTY_TRAN,0) - ISNULL(QTY_OPEN,0) - ISNULL(QTY_PICK,0) - ISNULL(QTY_COMM,0) + ISNULL(QTY_PROD,0)")
                .Tables(TABLE_NAME).Columns.Add("COLOR_CODE")
                .Tables(TABLE_NAME).Columns.Add("COLOR_DESC")
                .Tables(TABLE_NAME).Columns.Add("WHSE_CODE")
                .Tables(TABLE_NAME).Columns.Add("WHSE_DESC")
                .Tables(TABLE_NAME).Columns.Add("IMAGE", GetType(System.Byte()))
                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    .Tables(TABLE_NAME).Columns.Add("OPNQTY", GetType(System.Int64))
                    .Tables(TABLE_NAME).Columns.Add("OPNAMT", GetType(System.Decimal))
                    .Tables(TABLE_NAME).Columns.Add("OPNCST", GetType(System.Decimal), "IIF(OPNQTY=0,0,OPNAMT/OPNQTY)")
                    .Tables(TABLE_NAME).Columns.Add("SHPQTY", GetType(System.Int64))
                    .Tables(TABLE_NAME).Columns.Add("SHPAMT", GetType(System.Decimal))
                    .Tables(TABLE_NAME).Columns.Add("SHPCST", GetType(System.Decimal), "IIF(SHPQTY=0,0,SHPAMT/SHPQTY)")
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_LDP", GetType(System.Decimal))
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_LDP_CODE")
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_ELC", GetType(System.Decimal))
                    .Tables(TABLE_NAME).Columns.Add("STYLE_COST_CUM", GetType(System.Decimal))
                End If
                .Tables(TABLE_NAME).Columns.Add("STYLE_COST_EXT", GetType(System.Decimal), "ISNULL(STYLE_COST,0)*ISNULL(QTY_NETA,0)")
 
            Next


        End With

        grdPOTSHIPX.DataSource = dst.Tables("POTSHIPX")
        grdPOTSHIP3.DataSource = dst.Tables("POTSHIP3")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdICTSTYL1_Recent.DataSource = dst.Tables("ICTSTYL1_RECENT")

        '   Create_Summary(grdPOTSHIPX, "PO_SHIPMENT_NO", "Count")

        Create_Summary(grdPOTSHIP3, "PO_SHIPMENT_LNO", "Count")
        Create_Summary(grdPOTSHIP3, New String() {"PO_QTY_SHP", "UNITS"})
        'Create_Summary(grdPOTSHIP3, New String() {"PO_COST_FREIGHT_IN", "PO_COST_TRUCKING", "PO_COST_DUTY", "PO_COST_CUSTOMS" _
        '                                         , "PO_COST_LANDED", "PO_COST_VCOST", "PO_COST_MATLS", "PO_COST_VCOST_DZ" _
        '                                         , "PO_COST_MATLS_DZ", "PO_COST_OTHER", "PO_COST_OTHER_DZ" _
        '                                         , "FIRST_COST_TOTAL", "FIRST_COST_TOTAL_DZ", "COMMISION_COST", "COMMISION_COST_DZ" _
        '                                         , "PO_COST_QUOTA", "PO_COST_QUOTA_DZ", "PO_COST_QUOTA_DF", "PO_COST_QUOTA_DF_DZ"}, "Custom")


        Create_Summary(grdPOTORDRX, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRX, New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN", _
                                                  "PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN", _
                                                  "PO_CTNS_ORD", "PO_CTNS_SHP", "PO_CTNS_OPN", _
                                                  "PO_CUBE_ORD", "PO_CUBE_SHP", "PO_CUBE_OPN"}, , , "#,##0")

        Show_Filter(grdICTSTYL1_Recent, True)
        grdICTSTYL1_Recent.DisplayLayout.GroupByBox.Hidden = False
        Create_Summary(grdICTSTYL1_Recent, "STYLE_CODE", "Count")
        Create_Summary(grdICTSTYL1_Recent, "STYLE_COST_EXT")
        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Create_Summary(grdICTSTYL1_Recent, New String() {"OPNQTY", "OPNAMT", "SHPQTY", "SHPAMT"})
        End If
        Style_grdICTSTYL1_Recent()
        grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns("CARTONS_PER_UNIT").Hidden = True

        With grdPOTORDRX.DisplayLayout.Bands(0)
            .Columns("PO_ORDER_NO").Header.Fixed = True
            For Each COLUMN_NAME As String In New String() {"PO_STATUS", "FOB_CMT"}
                .Columns(COLUMN_NAME).Header.Appearance.TextHAlign = HAlign.Center
                .Columns(COLUMN_NAME).CellAppearance.TextHAlign = HAlign.Center
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"PO_CTNS_ORD", "PO_CTNS_SHP", "PO_CTNS_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"PO_CUBE_ORD", "PO_CUBE_SHP", "PO_CUBE_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_NO", "CUST_CODE", "CUST_NAME", "ORDR_CANCEL_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With
        grdPOTORDRX.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grdPOTORDRX, True)

        If ASCMAIN1.CLIENT = "NYA" Then
            With grdPOTSHIPX.DisplayLayout.Bands(0)
                .Columns("PO_SHIP_REF_NO").Hidden = True
                .Columns("PO_DATE_SHIPPED").Hidden = True
                .Columns("PORT_CODE").Hidden = True
                .Columns("BOL_NO").Hidden = True
            End With
        Else
            With grdPOTSHIP3.DisplayLayout.Bands(0)
                .Columns("CARTONS").Hidden = True
            End With
        End If

        With grdPOTSHIP3.DisplayLayout.Bands(0)
            .Columns("CARTON_NOS").Hidden = Not (ASCMAIN1.CLIENT = "RGI")
        End With


        ASCMAIN1.Add_Value_List(grdPOTORDRX, "PO_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "FOB_CMT", Nothing, New String() {":", "F:FOB", "C:NonInv-CMT", "I:Invty-CMT", "B:BTB"})
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "LABEL_RESP_CODE")
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "PO_APPR_PENDING", Nothing, New String() {":", "0:WIP", "1:Queued"})

        If ASCMAIN1.CLIENT = "NYA" Then
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_SPEC_ORDR_NO").Header.Caption = "Customer"
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY").Header.Caption = "ETD"
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY_MIN").Header.Caption = "ETD Min"
            grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY_MAX").Header.Caption = "ETD Max"

            With grdICTSTYL1_Recent.DisplayLayout.Bands(0)
                Dim C As Integer = .Columns("QTY_NETA").Header.VisiblePosition
                '"STYLE_COST", "STYLE_COST_EXT",
                For Each COLUMN_NAME As String In New String() {"CARTON_PACK_QTY", "INNER_PACK_QTY", "STYLE_COST_LDP", "STYLE_COST_ELC", "STYLE_COST_CUM", "SALES_DIVISION_CODE", "STYLE_GROUP_CODE", "VEND_CODE", "CUST_CODE"}
                    C += 1
                    .Columns(COLUMN_NAME).Header.SetVisiblePosition(C, False)
                Next

                C = .Columns("PURCH_NOTES").Header.VisiblePosition
                .Columns("COLOR_CODE").Header.SetVisiblePosition(C, False)
                .Columns("COLOR_DESC").Header.SetVisiblePosition(C, False)

                .Columns("STYLE_COST_LDP").Hidden = True
                .Columns("STYLE_COST_LDP_CODE").Hidden = True
                .Columns("STYLE_COST_ELC").Hidden = True
                .Columns("STYLE_COST_CUM").Hidden = True
                .Columns("SAFETY_STOCK").Hidden = True
                .Columns("EXCLUSIVE_STYLE").Hidden = True
                .Columns("IMAGE").Hidden = True
            End With

        End If

        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
        Else
            tabMain.Tabs("Open POs").Visible = False
            tabMain.Tabs("Inventory Status").Visible = False
        End If

        Show_Filter(grdPOTSHIPX, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                Validate_Code("PO_SHIPMENT_NO")

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                Load_POTSHIPX()
                If ASCMAIN1.CLIENT = "NYA" Then
                    Load_POTORDRX()
                    Setup_Recent()
                End If

            Case "Print"
                Print_Record()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '.Items("View").Settings.Enabled = not_iScreenMode
                    '.Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = iScreenMode

                    '.Items("View").Visible = (EntryMode = "L" Or Not ScreenMode)
                    '.Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    '.Items("Print").Visible = ScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'splPOTSHIPX.Visible = Not tf
        tabMain.Visible = Not tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"POTSHIPX", "POTSHIP3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_POTSHIPX()
        If ASCMAIN1.CLIENT = "NYA" Then
            Load_POTORDRX()
            sETUP_Recent()
        End If

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Sort_grdColumns(grdPOTSHIP3, "PO_SHIPMENT_LNO")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Stop
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Preparing " & Me.Text)


        Dim RPT As String = "PORWREC2"
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            RPT = "PORWRECR"
        End If

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report(RPT)
        Print_Report_End()

        '    Me.Cursor = Cursors.Default
        '    ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PO_SHIPMENT_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region


#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTSHIPX, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "PO Shipment Inquiry", "Select Entire Shipment", "Select Entire Container")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdPOTORDRS"
                '       tlb_btn = DirectCast(tlb_pop.Tools("Update Style Master (+ POs) with Changes made to Case Packs"), UltraWinToolbars.ButtonTool)
                '       tlb_btn.SharedProps.Visible = grdPOTORDRS.DisplayLayout.Bands(0).Columns("STYLE_ACTION").Hidden And Not (MENU_ITEM_OBJECT = "POFORDRI")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"
            '    Using F As New TAC.ICFSTYCX
            '        F.STYLE_CODE = ""
            '        F.Price_Caption = "Cost" & IIf(ssdDZGRD.Value = 1, "", "/Dz")
            '        F.ShowDialog()
            '        If F.STYLE_CODE <> "" Then
            '            Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
            '        End If
            '    End Using

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "PO Shipment Inquiry"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")

            Case "Select Entire Shipment"
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    Exit Sub
                End If
                grdPOTSHIPX.Selected.Rows.Clear()
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                For Each grow As UltraWinGrid.UltraGridRow In grd.ActiveRow.ParentRow.ChildBands(0).Rows
                    If Not grow.IsFilteredOut Then
                        If grow.Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO Then
                            grow.Selected = True
                        End If
                    End If
                Next

            Case "Select Entire Container"
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    Exit Sub
                End If
                grdPOTSHIPX.Selected.Rows.Clear()
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                Dim CONTAINER_NO As String = grd.ActiveRow.Cells("CONTAINER_NO").Value
                For Each grow As UltraWinGrid.UltraGridRow In grd.ActiveRow.ParentRow.ChildBands(0).Rows
                    If Not grow.IsFilteredOut Then
                        If grow.Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO And grow.Cells("CONTAINER_NO").Value = CONTAINER_NO Then
                            grow.Selected = True
                        End If
                    End If
                Next
        End Select
    End Sub


#End Region


#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

    Private Sub grdPOTSHIPX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTSHIPX.AfterRowActivate
        Setup_POTSHIP3()
    End Sub

    Private Sub grdPOTSHIPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTSHIPX.DoubleClickRow
        'If grdPOTSHIPX.ActiveRow IsNot Nothing Then
        '    Absx1.txtFor("PO_SHIPMENT_NO").Text = grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_NO").Text
        '    Click_Command("View")
        'End If

    End Sub

    Sub Setup_POTSHIP3()
        If grdPOTSHIPX.ActiveRow Is Nothing OrElse Not grdPOTSHIPX.ActiveRow.IsDataRow Then
            'grdPOTSHIP3.Visible = False
            splPOTSHIPX.Panel2Collapsed = True
        Else
            Dim PO_SHIPMENT_NO As String = grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_NO").Value
            Dim PO_SHIPMENT_LNO As String = Val(grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
            Dim CONTAINER_NO As String = grdPOTSHIPX.ActiveRow.Cells("CONTAINER_NO").Value & ""
            Dim WHSE_CODE As String = grdPOTSHIPX.ActiveRow.Cells("WHSE_CODE").Value

            Fill_Records("POTSHIP3", New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
            Sort_grdColumns(grdPOTSHIP3, "PO_ORDER_LNO")

            Fill_Records("POTSHIP7", New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
            Fill_Records("POTSHIP8", New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})

            grdPOTSHIP3.Text = "Shipment " & PO_SHIPMENT_NO & " Line " & CStr(PO_SHIPMENT_LNO) & "; Container '" & CONTAINER_NO & "' Contents"
            'grdPOTSHIP3.Visible = True
            splPOTSHIPX.Panel2Collapsed = False


            For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")

                If ASCMAIN1.CLIENT = "RGI" Then
                    If Val(rowPOTSHIP3.Item("CARTON_PACK_QTY")) <> 0 Then
                        Dim CARTON_NO As Int32 = rowPOTSHIP3.Item("PO_QTY_SHP") / rowPOTSHIP3.Item("CARTON_PACK_QTY")
                        If CARTON_NO > 0 Then
                            For i As Integer = 1 To CARTON_NO
                                rowPOTSHIP3.Item("CARTON_NOS") = rowPOTSHIP3.Item("CARTON_NOS") & CStr(i) + " | "
                            Next
                        End If
                    End If
                    Dim STYLE_CODE As String = rowPOTSHIP3.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowPOTSHIP3.Item("COLOR_CODE")
                    Dim PICK_QTY As Int64 = 9999  'as per Leo display the location with th eHighest Quantity - 04/10/2017

                    'Dim LOCATION_CODE As String = ""
                    'Dim LOCATION_ROUTE_SEQ As Int32 = 0
                    'TAC.SOCMAIN1.GET_STYLE_COLOR_LOCATIONS(WHSE_CODE, STYLE_CODE, COLOR_CODE, LOCATION_CODE, LOCATION_ROUTE_SEQ, PICK_QTY)

                    'rowPOTSHIP3.Item("STYLE_BIN") = LOCATION_CODE
                End If
            Next
        End If
    End Sub

    Sub Load_POTSHIPX()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
  
        Fill_Records("POTSHIPX")

        For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("")
            Dim PO_SHIPMENT_NO As String = rowPOTSHIPX.Item("PO_SHIPMENT_NO")
            Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIPX.Item("PO_SHIPMENT_LNO") & "")
            Dim CONTAINER_NO As String = rowPOTSHIPX.Item("CONTAINER_NO") & ""
            ASCMAIN1.sql = "Select * from POTSHIP4 where PO_SHIPMENT_NO = :PARM1 and CONTAINER_NO = :PARM2"
            Dim rowPOTSHIP4 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() {PO_SHIPMENT_NO, CONTAINER_NO})
            If rowPOTSHIP4 IsNot Nothing Then
                rowPOTSHIPX.Item("CONTAINER_TYPE_CODE") = rowPOTSHIP4.Item("CONTAINER_TYPE_CODE")
                rowPOTSHIPX.Item("CONTAINER_SEAL_NO") = rowPOTSHIP4.Item("CONTAINER_SEAL_NO")
            End If
        Next

        With grdPOTSHIPX.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("WHSE_CODE", False, True)
            .SortedColumns.Add("PO_SHIPMENT_NO", True)
        End With
        grdPOTSHIPX.Rows.ExpandAll(True)

        Setup_POTSHIP3()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_POTORDRX()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim sqlw As String = ""
        sqlw = " where POTORDR1.PO_STATUS = 'O'"
        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
            sqlw &= " and POTORDR1.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")"
        End If
        grdPOTORDRX.Text = "Open POs"

        Dim sqlDTL As String = "Select POTORDR2.PO_ORDER_NO" & vbCrLf _
            & IIf(False, ",POTORDR2.PO_DATE_SHIP_BY", "") _
            & ", SUM (POTORDR2.PO_QTY_ORD) PO_QTY_ORD" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_REC) PO_QTY_REC" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_ORD * POTORDR2.PO_COST) PO_AMT_ORD" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_SHP * POTORDR2.PO_COST) PO_AMT_SHP" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_REC * POTORDR2.PO_COST) PO_AMT_REC" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST) PO_AMT_OPN" & vbCrLf _
            & ", SUM (DECODE(POTORDR2.PO_CONF_NO,NULL,0,1)) PO_LINES_CONF" & vbCrLf _
            & ", COUNT (*) PO_LINES" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_ORD / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CTNS_ORD" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_SHP / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CTNS_SHP" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_OPN / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CTNS_OPN" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_ORD  * NVL(ICTSTYL1.CASE_CUBE,0) / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_ORD" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_SHP  * NVL(ICTSTYL1.CASE_CUBE,0) / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_SHP" & vbCrLf _
            & ", SUM (TRUNC(POTORDR2.PO_QTY_OPN  * NVL(ICTSTYL1.CASE_CUBE,0) / DECODE(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,1,NVL(ICTSTYL1.CARTON_PACK_QTY,0)) * 100) / 100) PO_CUBE_OPN" & vbCrLf _
            & ", MIN (POTORDR2.PO_DATE_SHIP_BY) PO_DATE_SHIP_BY_MIN" & vbCrLf _
            & ", MIN (POTORDR2.PO_DATE_ETA) PO_DATE_ETA_MIN" & vbCrLf _
            & ", MAX (POTORDR2.PO_DATE_SHIP_BY) PO_DATE_SHIP_BY_MAX" & vbCrLf _
            & ", MAX (POTORDR2.PO_DATE_ETA) PO_DATE_ETA_MAX" & vbCrLf _
            & " from  POTORDR2,ICTSTYL1 where POTORDR2.PO_ORDER_NO in " & vbCrLf _
            & " (Select PO_ORDER_NO from POTORDR1 " & sqlw & ")" _
            & " and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
            & " group by POTORDR2.PO_ORDER_NO" & vbCrLf _
            & IIf(False, ",POTORDR2.PO_DATE_SHIP_BY", "")
 
        ASCMAIN1.sql = "Select POTORDR1.*" & vbCrLf _
            & ", X.PO_QTY_ORD, X.PO_QTY_SHP, X.PO_QTY_REC, X.PO_QTY_OPN " & vbCrLf _
            & ", X.PO_AMT_ORD, X.PO_AMT_SHP, X.PO_AMT_REC, X.PO_AMT_OPN " & vbCrLf _
            & ", X.PO_LINES_CONF, X.PO_LINES" & vbCrLf _
            & ", X.PO_CTNS_ORD, X.PO_CTNS_SHP, X.PO_CTNS_OPN" & vbCrLf _
            & ", X.PO_CUBE_ORD, X.PO_CUBE_SHP, X.PO_CUBE_OPN " & vbCrLf _
            & ", X.PO_DATE_SHIP_BY_MIN, X.PO_DATE_ETA_MIN, X.PO_DATE_SHIP_BY_MAX, X.PO_DATE_ETA_MAX" & vbCrLf _
            & ", SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
            & " from (" & sqlDTL & ") X,POTORDR1,SOTORDR1" & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sqlw & " and SOTORDR1.ORDR_NO (+) = POTORDR1.ORDR_NO and POTORDR1.PO_ORDER_NO = X.PO_ORDER_NO")

        Fill_Records("POTORDRX", "", True, ASCMAIN1.sql)

        Sort_grdColumns(grdPOTORDRX, "PO_DATE_ETA".ToLower)

        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
             
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Recent()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Getting Styles with Status", "")
  
        Dim WHSE_CODEs As String = "'18'"

        Dim sqlx As String = "Select STYLE_CODE" _
                                 & IIf(True, ",COLOR_CODE", "") & vbCrLf _
                                 & IIf(True, ",WHSE_CODE", "") & vbCrLf _
                                 & ", SUM (WHSE_QTY_ON_HAND) QTY_ONHD" & vbCrLf _
                                 & ", SUM (WHSE_QTY_ON_ORDER) QTY_ONPO" & vbCrLf _
                                 & ", SUM (WHSE_QTY_TRAN) QTY_TRAN" & vbCrLf _
                                 & ", SUM (WHSE_QTY_OPEN) QTY_OPEN" & vbCrLf _
                                 & ", SUM (WHSE_QTY_PICK) QTY_PICK" & vbCrLf _
                                 & ", SUM (WHSE_QTY_COMM) QTY_COMM" & vbCrLf _
                                 & ", SUM (WHSE_QTY_PROD) QTY_PROD" & vbCrLf _
                                 & " from ICTSTAT2" & vbCrLf _
                                 & " where WHSE_CODE in (" & WHSE_CODEs & ")" & vbCrLf _
                                 & " group by STYLE_CODE" & vbCrLf _
                                 & IIf(True, ",COLOR_CODE", "") _
                                 & IIf(True, ",WHSE_CODE", "")

        Dim SQLW As String = ""
        SQLW = " and (NVL(QTY_ONHD,0) <> 0 or NVL(QTY_ONPO,0) <> 0 or NVL(QTY_TRAN,0) <> 0 or NVL(QTY_OPEN,0) <> 0 or NVL(QTY_PICK,0) <> 0 or NVL(QTY_COMM,0) <> 0 or NVL(QTY_PROD,0) <> 0)"

        ', SYSDATE LAST_ORDR_DATE, 'X' LAST_ORDR_NO, 'X' LAST_ORDR_CUST_CODE, 'X' LAST_ORDR_PO
        ASCMAIN1.sql = "Select ICTSTYL1.*" & vbCrLf _
            & ", X.QTY_ONHD, X.QTY_ONPO, X.QTY_TRAN, X.QTY_OPEN, X.QTY_PICK, X.QTY_COMM, X.QTY_PROD" & vbCrLf _
            & ", NVL(X.QTY_ONHD,0) + NVL(X.QTY_ONPO,0) + NVL(X.QTY_TRAN,0) - NVL(X.QTY_OPEN,0) - NVL(X.QTY_PICK,0) - NVL(X.QTY_COMM,0) + NVL(X.QTY_PROD,0) QTYAVA" & vbCrLf _
            & IIf(True, ",X.COLOR_CODE,ICTCOLR1.COLOR_DESC", "") _
            & IIf(True, ",X.WHSE_CODE,ICTWHSE1.WHSE_DESC", "") _
            & " from ICTSTYL1,(" & sqlx & ") X" & vbCrLf _
            & IIf(True, ",ICTCOLR1", "") _
            & IIf(True, ",ICTWHSE1", "") _
            & " where X.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & IIf(True, " and ICTCOLR1.COLOR_CODE (+) = X.COLOR_CODE", "") _
            & IIf(True, " and ICTWHSE1.WHSE_CODE (+) = X.WHSE_CODE", "") _
            & SQLW

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Dim SQLPO As String = "" _
                & "Select STYLE_CODE" & vbCrLf _
                & IIf(True, ", COLOR_CODE" & vbCrLf, "") _
                & IIf(True, ", WHSE_CODE" & vbCrLf, "") _
                & ", Sum (OPNQTY) OPNQTY, Sum (OPNAMT) OPNAMT" & vbCrLf _
                & ", Sum (SHPQTY) SHPQTY, Sum (SHPAMT) SHPAMT" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
                & ", SUM (POTORDR2.PO_QTY_OPN) OPNQTY" & vbCrLf _
                & ", SUM (POTORDR2.PO_QTY_OPN * (1 + NVL(ICTDUTY1.DUTY_RATE,0)/100) * POTORDR2.PO_COST) OPNAMT" & vbCrLf _
                & ", 0 SHPQTY, 0 SHPAMT" & vbCrLf _
                & " from POTORDR2,POTORDR1,ICTDUTY1,ICTSTYL1" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTDUTY1.DUTY_RATE_CODE (+) = ICTSTYL1.DUTY_RATE_CODE" & vbCrLf _
                & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE" & vbCrLf _
                & ", 0 OPNQTY, 0 OPNAMT" & vbCrLf _
                & ", SUM (POTSHIP3.PO_QTY_SHP) SHPQTY" & vbCrLf _
                & ", SUM (POTSHIP3.PO_QTY_SHP * CASE WHEN NVL(POTSHIP3.PO_COST_LANDED,0) <> 0 AND NVL(POTSHIP3.PO_COST_LANDED,0) <> NVL(POTSHIP3.PO_COST,0) THEN NVL(POTSHIP3.PO_COST_LANDED,0) ELSE (1 + NVL(ICTDUTY1.DUTY_RATE,0)/100) * NVL(POTSHIP3.PO_COST,0) END) SHPAMT" & vbCrLf _
                & " from POTORDR2,POTORDR1,POTSHIP1,POTSHIP2,POTSHIP3,ICTDUTY1,ICTSTYL1" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                & "   and ICTDUTY1.DUTY_RATE_CODE (+) = ICTSTYL1.DUTY_RATE_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE" & vbCrLf _
                & ") group by STYLE_CODE" & vbCrLf _
                & IIf(True, ", COLOR_CODE" & vbCrLf, "") _
                & IIf(True, ", WHSE_CODE" & vbCrLf, "")

            ASCMAIN1.sql = "Select A.*, B.OPNQTY, B.OPNAMT, B.SHPQTY, B.SHPAMT" & vbCrLf _
                & ", A.STYLE_COST STYLE_COST_LDP, CASE WHEN NVL(A.QTYAVA,0) > 0 THEN A.STYLE_COST * NVL(A.QTYAVA,0) ELSE 0 END STYLE_COST_CUM" & vbCrLf _
                & ", CASE WHEN NVL(B.OPNQTY,0) + NVL(B.SHPQTY,0) = 0 THEN NULL ELSE TRUNC(10000 * ((NVL(B.OPNAMT,0) + NVL(B.SHPAMT,0)) / (NVL(B.OPNQTY,0) + NVL(B.SHPQTY,0)))) / 10000 END STYLE_COST_ELC" & vbCrLf _
                & " from (" & ASCMAIN1.sql & ") A" _
                & ", (" & SQLPO & ") B" & vbCrLf _
                & " where B.STYLE_CODE (+) = A.STYLE_CODE" _
                & IIf(True, " and B.COLOR_CODE (+) = A.COLOR_CODE", "") _
                & IIf(True, " and B.WHSE_CODE (+) = A.WHSE_CODE", "")
        End If

        ICTSTYL1_Recent = ""
        If ICTSTYL1_Recent = "" Then
            ICTSTYL1_Recent = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Create Index I_" & ICTSTYL1_Recent & "_1 on " & ICTSTYL1_Recent & " (STYLE_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTSTYL1_Recent)
            ASCDATA1.ExecuteSQL("Insert into " & ICTSTYL1_Recent & " " & ASCMAIN1.sql)
        End If

        Fill_Records("ICTSTYL1_RECENT", "", True, "Select * from " & ICTSTYL1_Recent)

        With grdICTSTYL1_Recent.DisplayLayout.Bands(0)
            .Columns("LAST_ORDR_DATE").Hidden = True
            .Columns("LAST_ORDR_NO").Hidden = True
            .Columns("LAST_ORDR_CUST_CODE").Hidden = True
            .Columns("LAST_ORDR_CUST_PO").Hidden = True
        End With

        If ASCMAIN1.CLIENT = "NYA" Then
            '    Create_Summary(grdICTSTYL1_Recent, New String() {"OPNQTY", "OPNAMT", "SHPQTY", "SHPAMT"}, "Sum", grdICTSTYL1_Recent.DisplayLayout.Bands(0).Key) ' "ICTSTYL1_VIEW")
        End If

        grdICTSTYL1_Recent.Text = "Styles with Status Qtys, "
        grdICTSTYL1_Recent.Text &= " Whses:" & Replace(WHSE_CODEs, "'", "")

        Sort_grdColumns(grdICTSTYL1_Recent, "STYLE_CODE")

        For Each COLUMN_NAME As String In New String() {"QTY_ONHD", "QTY_ONPO", "QTY_TRAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_PROD", "QTY_NETA", "STYLE_COST_EXT"}
            If COLUMN_NAME = "STYLE_COST_EXT" And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
            Else
                grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
                If COLUMN_NAME = "QTY_COMM" Or COLUMN_NAME = "QTY_PROD" Then
                    grdICTSTYL1_Recent.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
                End If
            End If
        Next

        With grdICTSTYL1_Recent.DisplayLayout.Bands(0)
            .Columns("STYLE_COST_CUM").Hidden = True
            .Columns("COLOR_CODE").Hidden = True
            .Columns("COLOR_DESC").Hidden = True
            .Columns("WHSE_CODE").Hidden = True
            .Columns("WHSE_DESC").Hidden = True
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Style_grdICTSTYL1_Recent()

        With grdICTSTYL1_Recent.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"QTY_ONHD", "QTY_ONPO", "QTY_TRAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_PROD", "QTY_NETA"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    If gcol.Key = "QTY_NETA" Then
                        'gcol.CellAppearance.ForeColor = Color.Purple
                    ElseIf New String() {"QTY_ONHD", "QTY_ONPO", "QTY_TRAN"}.Contains(gcol.Key) Then
                        gcol.CellAppearance.ForeColor = System.Drawing.Color.Green
                    Else
                        gcol.CellAppearance.ForeColor = System.Drawing.Color.Red
                    End If
                    Create_Summary(grdICTSTYL1_Recent, gcol.Key)

                ElseIf New String() {"STYLE_CODE", "STYLE_STATUS", "STYLE_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    gcol.Header.Fixed = True
                ElseIf New String() {"VEND_CODE", "FACTORY_CODE", "STYLE_CART_CUBE", "CASE_CUBE", "VEND_ITEM_CODE", "STYLE_PO_QTY_MIN", "PURCH_NOTES", "REPLENISHMENT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                ElseIf New String() {""}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                ElseIf New String() {"STYLE_XMAS_DATE", "STYLE_SO_QTY_MIN", "STYLE_COST_FIRST"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Else
                        gcol.Hidden = True
                    End If
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                End If
            Next


            If ASCMAIN1.CLIENT = "VAN" Then
                .Columns("DUTY_RATE_CODE").Hidden = False
            End If

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Else
                For Each COLUMN_NAME As String In New String() {"SUB_BODY_CODE", "MASTER_BODY_CODE", "SIZE_CODE", "FASHION_PROMO", "SUB_UNIT_PACK_QTY", "STYLE_MATL_DESC", "FABRIC_CODE", "FACTORY_CODE"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
            End If

            If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                For Each COLUMN_NAME As String In New String() _
                    {"STYLE_COST", "STYLE_COST_EXT"}
                    .Columns(COLUMN_NAME).Hidden = True
                Next
                Create_Summary(grdICTSTYL1_Recent, {"STYLE_COST_CUM"})
            Else
                'For Each COLUMN_NAME As String In New String() _
                '    {"STYLE_COST_LDP", "STYLE_COST_LDP_CODE", "STYLE_COST_ELC", "STYLE_COST_CUM"}
                '    .Columns(COLUMN_NAME).Hidden = True
                'Next
            End If

        End With

    End Sub

End Class