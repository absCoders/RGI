Public Class EDF850I1

    ' CUST_STORE_NO BY GLIN

    ' Temporary Oacle Tables

    Dim EDT850T1 As String
    Dim SOTORDR1 As String
    Dim SOTORDR2 As String
    Dim SOTORDR3 As String
    Dim SOTORDR4 As String
    Dim SOTORDR9 As String

    ' Datarow

    Dim rowEDTTRPM1 As DataRow
    Dim rowEDT850T1 As DataRow
    Dim rowEDT850T2 As DataRow
    Dim rowEDTSLSP1 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim rowARTCUST1bt As DataRow
    Dim rowARTCUST2 As DataRow
    Dim rowARTCUST2_DC As DataRow
    Dim rowICTSTYL1 As DataRow
    Dim rowSOTPCLS1 As DataRow

    ' EDI Control

    Dim EDI_ERRORs As New Dictionary(Of String, Int32)
    Dim EDI_DOC_SEQ_NO_ok As Boolean = False
    Dim EDI_DOC_SEQ_NO As String
    Dim EDI_DTL_SEQ As Int32
    Dim EDI_TP_QUAL As String
    Dim EDI_TP_ID As String

    Dim EDI_MERCH_TYPE As String
    Dim EDI_SHIPPER As String
    Dim EDI_PO_TYPE As String

    Dim DROP_SHIP As Boolean
    Dim ALT_PRICING As Boolean

    Dim EDI_ACTION As String
    Dim EDI_REPLACED_VALUE As String

    Dim EDI_DOC_SEQ_NOs_no_company As New List(Of String)


    ' Order Data

    Dim ORDR_CUST_PO As String
    Dim CUST_CODE As String
    Dim CUST_STORE_NO As String
    Dim WHSE_CODE As String
    Dim CUST_BILL_TO_CUST As String
    Dim ORDR_SHIP_TO As String
    Dim TERM_CODE As String
    Dim CUST_DC_NO As String
    Dim ORDR_DEPT As String
    Dim ORDR_EDI_810 As String
    Dim ORDR_EDI_856 As String
    Dim ORDR_DATE As Date
    Dim ORDR_CANCEL_DATE As Date
    Dim ORDR_SHIP_DATE As Date
    Dim ORDR_ARRIVAL_DATE As Date
    Dim ORDR_LAST_ARRIVAL_DATE As Date
    Dim SHIP_VIA_CODE As String

    Dim PRICE_BASIS As String
    Dim PRICE_BASE_DPCT As Decimal
    Dim PRICE_LIST_CODE As String

    Dim CURR_CODE As String = "USD"
    Dim CURR_EXCH_RATE As Decimal = 1

    Dim STYLE_CODE As String
    Dim COLOR_CODE As String

    Dim ITEM_INACTIVE As String

    Dim saleable_item As Boolean

    Dim TRANSIT_BUS_DAYS As Int32
    Dim CUST_EDI_DTS_FLAG As String
    Dim FRT_TERMS As String

    Dim ITEM_OK_IF_ONE_MATCH As Boolean = True
    Dim skip_store As Boolean
    Dim Bad_Data_Cond_05 As Boolean
    Dim skip_item As Boolean
    ' Dim sdq_stores As New List(Of String)

    Dim RESOLUTIONS As New Dictionary(Of String, String)

    Dim ACTIONS_A As New List(Of String)
    Dim ACTIONS_E As New List(Of String)
    Dim ACTIONS_S As New List(Of String)
    Dim ACTIONS_R As New List(Of String)

    Dim ECOM_CODE As String = ""
    Dim ECOM_PRICE_TOLERANCE_PCT As Decimal

    Dim useClass As Boolean = False

    'EDI_SUPPLIER_NO
    'EDI_PROMOTION
    'EDI_PO_TYPE

    '01 N TRADING PARTNER SETUP REQUIRED 
    '27 A CHECK OTHER UNPROCESSED EDI ORDERS FOR SAME TP ID AND SAME CUSTOMER PO
    '28 A CHECK THIS PO FOR DUPLICATE ITEMS IN ORDER DETAIL
    '04 N CUSTOMER MASTER SETUP REQUIRED
    '35 N CUSTOMER MASTER SETUP REQUIRED
    '36 N CUSTOMER MASTER SETUP REQUIRED
    '03 N CUSTOMER MASTER SETUP REQUIRED
    '91 N CUSTOMER MASTER SETUP REQUIRED
    '09 A ACCEPT EDI DATA, GET AN EXTENSION FROM CUSTOMER, AND CHANGE CANCEL DATE
    '08 N CUSTOMER MUST RETRANSMIT WITH SHIP DATE
    '21 N CUSTOMER MASTER SETUP REQUIRED
    '11 A MAP EDI TERMS TO A VALID AR TERMS CODE
    '12 N MAP EDI TERMS TO A VALID AR TERMS CODE
    '13 N CUSTOMER MASTER SETUP REQUIRED
    '14 A CHECK ORDER HISTORY FOR THIS PO
    '15 EAN/UPC Code not found
    '16 RANGE QTY DOES NOT MATCH ORDER QTY
    '67 EDTXREF4 suplier xref not defined when one sent in
    '68 Promotional Price Mismatch
    '41 Unable to lock Reservations

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")
        Get_PARM("EDTPARM1")

        If ASCMAIN1.CLIENT = "RGI" Then
            TAC.EDCMAIN1.Fix_Bad_Styles()
        End If

        Create_Temp_Table()

        With dst
            ASCMAIN1.sql = "Select EDT850T1.*, ARTCUST1.CUST_NAME" & vbCrLf _
                & " from " & EDT850T1 & " EDT850T1,ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = EDT850T1.CUST_CODE"
            Create_TDA(.Tables.Add, "EDT850T1", "**", 0, True, "", 1)
            With .Tables("EDT850T1").Columns
                .Add("SEL")
                .Add("RESULT")
                .Add("ORDERS", GetType(System.Int32))
                .Add("AMOUNT", GetType(System.Decimal))
                .Add("WHSE_CODE")
            End With
            .Tables("EDT850T1").Columns("SEL").DefaultValue = "0"

            Dim TBL As DataTable = dst.Tables("EDT850T1").Clone
            TBL.TableName = "EDT850T1_IMPORTED"
            dst.Tables.Add(TBL)

            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select EDT850T1.*, ARTCUST1.CUST_NAME" & vbCrLf _
              & " from EDT850T1,ARTCUST1" & vbCrLf _
              & " where ARTCUST1.CUST_CODE (+) = EDT850T1.CUST_CODE" & vbCrLf _
              & "   and EDT850T1.EDI_PROCESS_IND in ('C','1')" & vbCrLf _
              & "   and EDT850T1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" & vbCrLf _
              & "   and EDT850T1.CUST_CODE = :PARM1" & vbCrLf _
              & "   and EDT850T1.EDI_PO_DATE between :PARM2 and :PARM3"
            Else
                ASCMAIN1.sql = "Select EDT850T1.*, ARTCUST1.CUST_NAME" & vbCrLf _
              & " from EDT850T1,ARTCUST1" & vbCrLf _
              & " where ARTCUST1.CUST_CODE (+) = EDT850T1.CUST_CODE" & vbCrLf _
              & "   and EDT850T1.EDI_PROCESS_IND = '1'" & vbCrLf _
              & "   and EDT850T1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" & vbCrLf _
              & "   and EDT850T1.CUST_CODE = :PARM1" & vbCrLf _
              & "   and EDT850T1.EDI_PO_DATE between :PARM2 and :PARM3"
            End If
            Create_TDA(.Tables.Add, "EDT850T1_ARCHIVED", "**", 0, False, "VDD", 1)
            With .Tables("EDT850T1_ARCHIVED").Columns
                .Add("SEL")
                .Add("RESULT")
                .Add("ORDERS", GetType(System.Int32))
                .Add("AMOUNT", GetType(System.Decimal))
                .Add("WHSE_CODE")
            End With
            .Tables("EDT850T1_ARCHIVED").Columns("SEL").DefaultValue = "0"

            'TBL = dst.Tables("EDT850T1").Clone
            'TBL.TableName = "EDT850T1_ARCHIVED"
            'dst.Tables.Add(TBL)

            Create_TDA(.Tables.Add, "EDT850T2", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "EDT850T3", "*", 1, False, "", 3)
            Create_TDA(.Tables.Add, "EDT850T4", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "EDT850T5", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "EDT850T6", "*", 1, False, "", 3)
            Create_TDA(.Tables.Add, "EDT850T7", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "EDT850T8", "*", 1, False, "", 3)
            Create_TDA(.Tables.Add, "EDT850T9", "*", 1, False, "", 2)

            Create_TDA(.Tables.Add, "EDT850TE", "*", 0, True, "", 2)
            .Tables("EDT850TE").Columns.Add("RESOLUTION")
            .Tables("EDT850TE").Columns.Add("VARIANCE")

            ASCMAIN1.sql = "Select * from " & SOTORDR1
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0)
            ASCMAIN1.sql = "Select * from " & SOTORDR2
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0)
            .Tables("SOTORDR2").Columns.Add("ORDR_AMT", GetType(System.Decimal), "ORDR_QTY * ORDR_UNIT_PRICE")

            ASCMAIN1.sql = "Select * from " & SOTORDR3
            Create_TDA(.Tables.Add, "SOTORDR3", "**", 0)
            ASCMAIN1.sql = "Select * from " & SOTORDR4
            Create_TDA(.Tables.Add, "SOTORDR4", "**", 0)
            ASCMAIN1.sql = "Select * from " & SOTORDR9
            Create_TDA(.Tables.Add, "SOTORDR9", "**", 0)

            ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = :PARM1 and CUST_ADDR_TYPE = 'MK'"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, True, "V", 3)
            ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = :PARM1 and CUST_ADDR_TYPE = 'DC'"
            Create_TDA(.Tables.Add, "ARTCUST2_DC", "**", 0, False, "V", 3)

            Create_TDA(.Tables.Add, "ARTCUST3", "*")

            Create_TDA(.Tables.Add, "SOTCSTY1", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "SOTPRIC2", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "ICTSTYL1", "*", 0, False)
            Create_TDA(.Tables.Add, "EDTUPCX1", "*", 1, False, "", 2)

            ASCMAIN1.sql = "Select EDTPPKS1.*, ICTSTYL1.CARTON_PACK_QTY CARTON_PACK_QTY_STYLE" & vbCrLf _
               & " from EDTPPKS1,ICTSTYL1" & vbCrLf _
               & " where ICTSTYL1.STYLE_CODE (+) = EDTPPKS1.STYLE_CODE"
            Create_TDA(.Tables.Add, "EDTPPKS1", "**", 0, False, "", 2)

            'ASCMAIN1.sql = "" _
            '    & "Select 'UPC' UPC_EAN, ITEM_UPC_CODE, STYLE_CODE from ICTSTYL1 where ITEM_UPC_CODE is not Null" _
            '    & " union " _
            '    & "Select 'EAN' UPC_EAN, ITEM_EAN_CODE ITEM_UPC_CODE, STYLE_CODE from ICTSTYL1 where ITEM_EAN_CODE is not Null"
            'Create_TDA(.Tables.Add, "ICTITEMX", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select UPC_CODE, STYLE_CODE, COLOR_CODE from ICTSTYC1 where UPC_CODE is not Null"
            ASCMAIN1.sql = "Select UPC_CODE, MAX (STYLE_CODE) STYLE_CODE, MAX (COLOR_CODE) COLOR_CODE from (" & ASCMAIN1.sql & ") X group by UPC_CODE"
            Create_TDA(.Tables.Add, "ICTITEMX", "**", 0, False, "", 1)


            'For Each TABLE_NAME As String In New String() {"SOTPRIC1", "SOTPRIC2"}
            '    ASCMAIN1.sql = "Select * from " & TABLE_NAME
            '    Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False)
            'Next

            With .Tables.Add("EDT850TS")
                .Columns.Add("EDI_STORE")
                .Columns.Add("EDI_SHIP_DC")
                .Columns.Add("CUST_STORE_NO")
                .Columns.Add("CUST_DC_NO")
                .PrimaryKey = New DataColumn() {.Columns("EDI_STORE")}
            End With

            ASCMAIN1.sql = "Select EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE, EDT850T2.EDI_UPC, SOTPCLS1.PRICE_BASE_DPCT" & vbCrLf _
                & ", EDT850T2.EDI_PRICE PRICE, EDT850T2.EDI_PRICE" & vbCrLf _
                & ", EDT850T2.EDI_PRICE EDI_PRICE_CURR, EDT850T2.EDI_PRICE ITEM_PRICE_CURR" & vbCrLf _
                & " from SOTORDR2, SOTPCLS1, EDT850T2"
            Create_TDA(.Tables.Add, "EDTSDQT0", "**", 0, False, "", 4)

            ASCMAIN1.sql = "Select EDT850T1.EDI_DOC_SEQ_NO, EDT850T1.EDI_STORE" & vbCrLf _
                & ", EDT850T2.EDI_DTL_SEQ, ICTSTYL1.STYLE_CODE, ICTSTYC1.COLOR_CODE" & vbCrLf _
                & ", EDT850T2.EDI_ITEM, EDT850T2.EDI_STYLE, EDT850T2.EDI_COLOR_CODE, EDT850T2.EDI_COLOR_NAME" & vbCrLf _
                & ", EDT850T2.EDI_UPC, EDT850T2.EDI_SKU, EDT850T2.EDI_SIZE_DESC" & vbCrLf _
                & ", EDT850T2.EDI_STYLE SLN_PARENT_STYLE_CODE, 0 SLN_PARENT_STYLE_QTY" & vbCrLf _
                & ", 0 SLN_PARENT_INNER_PACK_QTY, EDT850T2.EDI_PRICE, EDT850T2.EDI_DTL_SEQ EDI_SLN_SEQ" & vbCrLf _
                & ", 0 QTY, 0 RANGE_PARENT_QTY, 0 RANGE_PARENT_PRICE, EDT850T1.EDI_SHIP_DC, ICTRSTY1.RANGE_STYLE_CODE" & vbCrLf _
                & ", EDT850T2.EDI_STYLE SLN_PARENT_ITEM_DESC, 0 RANGE_STYLE_QTY_PER_PP, 0 QTY_PER_PP, ICTRSTY1.RNG_AST_FLG " & vbCrLf _
                & ", ICTSTYL1.SALES_DIVISION_CODE, ARTCUST2.CUST_CODE, ARTCUST2.CUST_ADDR_CODE CUST_STORE_NO, ARTCUST2.CUST_ADDR_CODE CUST_DC_NO " & vbCrLf _
                & " from EDT850T2, EDT850T1, ICTSTYL1, ARTCUST2,ICTSTYC1,ICTRSTY1"
            Create_TDA(.Tables.Add, "EDTSDQT1", "**", 0, False, "", 12) ' 9) ' 6)
            .Tables("EDTSDQT1").Columns("RANGE_STYLE_CODE").AllowDBNull = True
            .Tables("EDTSDQT1").Columns("CUST_DC_NO").AllowDBNull = True

            ASCMAIN1.sql = "Select SOTORDR1.CUST_CODE, EDT850T2.EDI_UPC, EDT850T2.EDI_SKU, EDT850T2.EDI_ITEM, EDT850T2.EDI_STYLE" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE, 0.01 CUST_PRICE, 0.01 ITEM_PRICE_CURR" & vbCrLf _
                & ", 0.01 PRICE_DISC, 0.01 RETAIL_PRICE, 0.01 RETAIL_PRICE_CURR" & vbCrLf _
                & ", SOTORDR1.SALES_DIVISION_CODE" & vbCrLf _
                & " from SOTORDR2, EDT850T2, SOTORDR1"
            Create_TDA(.Tables.Add, "EDTITEMX", "**", 0, False, "", 0)

            '        Call Create_Index("EDWITEMX", "I_EDWITEMX_1", "CUST_CODE,EDI_UPC,EDI_SKU,EDI_ITEM")
            '        Call Create_Index("EDWITEMX", "I_EDWITEMX_2", "CUST_CODE,STYLE_CODE")

            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select * from ECTESTY1" & vbCrLf _
                    & " where ECOM_STYLE_STATUS = 'A' " & vbCrLf _
                    & " and ECOM_CODE = :PARM1"
                Create_TDA(.Tables.Add, "ECTESTY1", "**", 0, False, "V", 2)

                ASCMAIN1.sql = "Select ECTESTY3.* from ECTESTY3" & vbCrLf _
                    & " where ECOM_CODE = :PARM1"
                Create_TDA(.Tables.Add, "ECTESTY3", "**", 0, False, "V", 3)

            End If

            Create_TDA(.Tables.Add, "ICTRSTY1", "*", 1, False)
            ASCMAIN1.sql = "Select * from ICTRSTY2 where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTRSTY2", "*", 0, False, "V", 0)

            ' "EDTUPCX1", 
            For Each TABLE_NAME As String In New String() _
                {"EDTTRPM1", "TATTERM1", "SOTSVIA1", "SOTSVIA2", "EDTTERM1", "SOTSDIV1", "EDTUPCX4", "EDTXREF3", "EDTXREF4", "ECTECOM1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, TABLE_NAME = "EDTTERM1")
                Fill_Records(TABLE_NAME)
            Next

            ' for reservations
            ASCMAIN1.sql = "Select * from SOTORDR7 where ORDR_GROUP_NO = :PARM1 " & vbCrLf _
                & "   and SOTORDR7.STYLE_CODE = :PARM2 " & vbCrLf _
                & "   and SOTORDR7.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and SOTORDR7.PICK_BATCH_NO is Null"
            Create_TDA(.Tables.Add, "SOTORDR7", "**", 0, True, "VVV", 1)

            Create_TDA(.Tables.Add, "SOTRSRV1", "*")
            Create_TDA(.Tables.Add, "SOTRSRV2", "*")

            ASCMAIN1.sql = "Select SOTRSRV2.* from SOTRSRV2,SOTRSRV1" & vbCrLf _
                & " where SOTRSRV1.CUST_CODE = :PARM1 " & vbCrLf _
                & "   and SOTRSRV2.STYLE_CODE = :PARM2 " & vbCrLf _
                & "   and SOTRSRV2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & "   and SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
                & "   and SOTRSRV2.RSRV_QTY_OPEN > 0" & vbCrLf
            Create_TDA(.Tables.Add, "SOTRSRVX", "**", 0, False, "VVV", 0)


            ' for automatically released orders
            Create_TDA(.Tables.Add, "SOTPICK0", "*")
            Create_TDA(.Tables.Add, "SOTPICK1", "*")
            Create_TDA(.Tables.Add, "SOTPICK2", "*")
            .Tables("SOTPICK2").Columns.Add("STYLE_CODE")
            .Tables("SOTPICK2").Columns.Add("COLOR_CODE")

            Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            Create_TDA(.Tables.Add, "SOTCART1", "*")
            Create_TDA(.Tables.Add, "SOTCART2", "*")
            .Tables("SOTCART2").Columns.Add("STYLE_WEIGHT", GetType(System.Decimal))
            .Tables("SOTCART2").Columns.Add("WGT", GetType(System.Decimal), "ISNULL(QTY_PACKED,0) * ISNULL(STYLE_WEIGHT,0)")

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")
            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")

            .Tables("SOTCART1").Columns.Add("QTY", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).QTY_PACKED)")
            .Tables("SOTCART1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).WGT)")

            .Tables("SOTPICK1").Columns.Add("CTNS", GetType(System.Int64), "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
            .Tables("SOTPICK1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_CALC)")

        End With

        Fill_Records("ICTITEMX")
        Fill_Records("EDTPPKS1")

        grdEDT850T2.DataSource = dst.Tables("EDT850T2")
        grdEDT850T3.DataSource = dst.Tables("EDT850T3")
        grdEDT850T4.DataSource = dst.Tables("EDT850T4")
        grdEDT850T5.DataSource = dst.Tables("EDT850T5")
        grdEDT850T6.DataSource = dst.Tables("EDT850T6")
        grdEDT850TE.DataSource = dst.Tables("EDT850TE")
        grdEDT850T1.DataSource = dst.Tables("EDT850T1")
        grdEDTTERM1.DataSource = dst.Tables("EDTTERM1")

        Create_Summary(grdEDT850T1, "CUST_CODE", "Count")
        Create_Summary(grdEDT850T1, New String() {"SEL", "ORDERS", "AMOUNT"})

        Create_Summary(grdEDT850TE, "EDI_COND_DESC", "Count")

        grdEDT850T1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdEDT850T1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        grdEDT850T1.DisplayLayout.UseFixedHeaders = True
        With grdEDT850T1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White

                If gcol.Key = "SEL" Or gcol.Key = "WHSE_CODE" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

            Next

            For Each COLUMN_NAME As String In New String() {"SEL", "RESULT", "CUST_CODE", "CUST_NAME", "EDI_DEPARTMENT", "EDI_PO_NO", "EDI_PO_DATE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            'For Each COLUMN_NAME As String In New String() {"EDI_DOC_SEQ_NO", "CUST_CODE", "ORDR_CUST_PO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            '    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            'Next
            For Each COLUMN_NAME As String In New String() {"EDI_PO_NO", "EDI_START_DATE", "EDI_END_DATE", "EDI_SHIP_DATE", "EDI_SHIP_DC", "EDI_STORE", "EDI_PO_DATE", "EDI_ARRIVAL_DATE", "EDI_LAST_ARRIVAL_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"EDI_PO_PURP", "EDI_PO_TYPE", "EDI_ORD_QTY"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
            For Each COLUMN_NAME As String In New String() {"EDI_TERMS", "EDI_TERM_TYPE", "EDI_TERM_BASIS", "EDI_TERM_RATE", "EDI_TERM_DSCDAYS", "EDI_TERM_NETDAYS", "EDI_TERM_DESC", "EDI_TERM_DOM"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
            Next
            For Each COLUMN_NAME As String In New String() {"EDI_DOC_SEQ_NO", "GEN_DOC_NO", "EDI_ISA_NO", "EDI_TP_QUAL", "EDI_TP_ID", "EDI_RECEIVED_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            Next
        End With

        Show_Filter(grdEDT850T1, True)
        grdEDT850T1.DisplayLayout.GroupByBox.Hidden = False


        With grdEDT850TE.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White

                If gcol.Key = "EDI_ACTION" Or gcol.Key = "EDI_REPLACED_VALUE" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                If gcol.Key = "VARIANCE" Then
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If
                If New String() {"EDI_REFERENCE", "VARIANCE", "EDI_RECEIVED_VALUE", "EDI_EXPECTED_VALUE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next

            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "ORDR_CUST_PO", "CUST_STORE_NO", "EDI_COND_CODE", "EDI_COND_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

        End With


        Setup_RESOLUTIONS()

        ASCMAIN1.Add_Value_List(grdEDT850TE, "EDI_ACTION", Nothing, New String() {":", "N:Do Nothing", "S:Skip Record", "E:Use Value Expected", "A:Use Value Received", "R:Use Replacement Value"})
        ' ASCMAIN1.Add_Value_List(grdEDT850T1, "ORDR_PICK_TYPE", Nothing, New String() {":", "P:Pick&Pack", "C:Full Case"})
        'ASCMAIN1.Add_Value_List(grdEDT850T1, "WO_TYPE", "Select WO_TYPE, WO_TYPE_DESC from SOTWORKT")
        ASCMAIN1.Add_Value_List(grdEDT850T1, "WHSE_CODE", "Select * from (Select WHSE_CODE, WHSE_CODE WHSE_DESC from ICTWHSE1) order by WHSE_CODE")


        dteFrom.Value = Now.Date.AddDays(-30)
        dteTo.Value = Now.Date


        If ASCMAIN1.Running_in_VS Then
            Stop
            'useClass = True
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"


            Case "Process Orders"
                Dim SEL_CNT As Int32 = dst.Tables("EDT850T1").Select("SEL='1'").Length
                If SEL_CNT = 0 Then
                    EMsg &= vbCr & "You Must Select at Least 1 Order to Process"
                End If

                If EMsg = "" Then
                    If MsgBox("This action will " & IIf(optACTION.Value = "I", "Import", "Delete") & " the " & CStr(SEL_CNT) & " Order(s) Selected." _
                              & vbCrLf & vbCrLf & "OK to Continue?", _
                              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
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
            Case "Refresh"
                Load_EDT850T1()

            Case "Process Orders"
                Mode_Settings(True)
                If optACTION.Value = "D" Then
                    Delete_Orders()
                    optACTION.Value = "I"
                Else
                    If useClass Then
                        Import_Orders_using_Class()
                    Else
                        Import_Orders()
                    End If

                End If

                Mode_Settings(False)
                'Load_EDT850T1()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                    .Items("Process Orders").Settings.Enabled = not_iScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        ' grdEDT850T1.Visible = Not tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"EDT850T1", "EDT850T2", "EDT850T3", "EDT850T4", "EDT850T5", "EDT850T6", "EDT850T7", "EDT850T8", "EDT850T9", _
                 "SOTORDR1", "SOTORDR2", "ARTCUST2", "ICTSTYL1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        tabMain.SelectedTab = tabMain.Tabs("Orders Waiting to be Imported")
        Setup_tabMain()
        Load_EDT850T1()
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Generate_Report("PORWREC2")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "USER_ID"
                sql_where = ""
            Case "CUST_CODE"
                sql_where = "CUST_CODE in (Select Distinct CUST_CODE from EDT850T1)"
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT850T1, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select Today's", "Select All", "De-Select All", "Select All for Customer", "Customer Order Inquiry")
        Load_Popup_Menu(grdEDT850T5, "BB", "Add as Store", "Add as DC")
        Load_Popup_Menu(grdEDT850T2, "BB", "Style Status Inquiry", "Add to Cross Reference")
        Load_Popup_Menu(grdEDT850T6, "BB", "Style Status Inquiry", "Add to Cross Reference")
        Load_Popup_Menu(grdEDT850TE, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Do Nothing", "Use Value Expected", "Use Value Received", "Use Replacement Value")
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
            Case "grdEDT850T1"
                tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow


                tlb_btn = DirectCast(tlb_pop.Tools("Select All for Customer"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                    tlb_btn.SharedProps.Visible = True
                    tlb_btn.SharedProps.Caption = "Select All for " & CUST_CODE
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select Today's"
                Dim sqlw As String = "EDI_RECEIVED_DATE > '" & Format(DATETIME_STAMP, "MM/dd/yy") & "'"
                For Each row As DataRow In dst.Tables("EDT850T1").Select(sqlw)
                    row.Item("SEL") = "1"
                Next

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdEDT850T1.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Do Nothing", "Use Value Expected", "Use Value Received", "Use Replacement Value"
                If grdEDT850TE.ActiveRow IsNot Nothing AndAlso Not grdEDT850TE.ActiveRow.Selected Then
                    grdEDT850TE.ActiveRow.Selected = True
                End If
                For Each grow As UltraWinGrid.UltraGridRow In grdEDT850TE.Selected.Rows
                    Dim EDI_ACTION As String = "N"
                    If e.Tool.Key = "Use Value Received" Then
                        EDI_ACTION = "A"
                    ElseIf e.Tool.Key = "Use Value Expected" Then
                        EDI_ACTION = "E"
                    ElseIf e.Tool.Key = "Use Replacement Value" Then
                        EDI_ACTION = "R"
                    End If
                    grow.Cells("EDI_ACTION").Value = EDI_ACTION
                    grow.Update()
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Select All for Customer"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                Dim sqlw As String = "CUST_CODE = '" & CUST_CODE & "'"
                For Each row As DataRow In dst.Tables("EDT850T1").Select(sqlw)
                    row.Item("SEL") = "1"
                Next

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")

            Case "Style Status Inquiry"
                Dim ITEM_UPC_CODE As String
                If grd.Name = "grdEDT850T6" Then
                    ITEM_UPC_CODE = grd.ActiveRow.Cells("EDI_SLN_UPC").Value & ""
                Else
                    ITEM_UPC_CODE = grd.ActiveRow.Cells("EDI_UPC").Value & ""
                End If

                Dim row As DataRow = dst.Tables("ICTITEMX").Rows.Find(ITEM_UPC_CODE)
                If row Is Nothing Then
                    ASCMAIN1.sql = "Select * from ICTSTYC1 where UPC_CODE = '" & ITEM_UPC_CODE & "'"
                    row = ASCDATA1.GetDataRow
                End If
                If row IsNot Nothing Then
                    Dim STYLE_CODE As String = row.Item("STYLE_CODE") & ""
                    Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                Else
                    MsgBox("Cannot Locate Style for UPC " & ITEM_UPC_CODE, MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                End If

            Case "Add as Store", "Add as DC"

                Dim CUST_CODE As String = grdEDT850T1.ActiveRow.Cells("CUST_CODE").Value & ""
                If CUST_CODE = "" Then Exit Sub
                Dim EDI_ADDR_CODE_QUAL As String = grd.ActiveRow.Cells("EDI_ADDR_CODE_QUAL").Value & ""
                Dim EDI_ADDR_CODE As String = grd.ActiveRow.Cells("EDI_ADDR_CODE").Value & ""
                Dim CUST_STORE_NO As String = Format_Store(EDI_ADDR_CODE)

                Dim GLN As String = ""

                If Len(CUST_STORE_NO) > 6 Then
                    Dim CN As String = grd.ActiveRow.Cells("EDI_CUST_NAME_ADR").Value
                    If CN.StartsWith("WAL-MART DC ") Then
                        CUST_STORE_NO = Mid(CN, 13, 5)
                        GLN = EDI_ADDR_CODE
                    End If
                End If

                If Len(CUST_STORE_NO) > 6 Then
                    Dim CUST_STORE_NO_entry As String = ASCMAIN1.Get_txt_from_User("Customer Address Code (" & CUST_STORE_NO & ") appears to be a GLN", _
                                                                 "Enter Store No to use for this Address", , 6, "000000")
                    'MsgBox("Invalid Value for Customer Address Code (" & CUST_STORE_NO & ")", _
                    '       MsgBoxStyle.OkOnly, _
                    '       "Cannot Perform Requested Action")
                    If CUST_STORE_NO_entry = "" Then
                        Exit Sub
                    Else
                        GLN = EDI_ADDR_CODE
                        CUST_STORE_NO = CUST_STORE_NO_entry.PadLeft(6, "0")
                    End If

                End If
                Dim CUST_ADDR_TYPE As String = IIf(e.Tool.Key = "Add as Store", "MK", "DC")

                If CUST_CODE = "JCPE" Then
                    CUST_STORE_NO = "000000"
                End If

                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_ADDR_TYPE, CUST_STORE_NO})
                If rowARTCUST2 IsNot Nothing Then
                    MsgBox("Address Record is alread on file for " & CUST_CODE & "-" & CUST_STORE_NO)
                ElseIf CUST_STORE_NO = "" Or Trim(CUST_STORE_NO) = "" Then
                    MsgBox("Could Not Resolve Store No for " & CUST_CODE & "-" & CUST_STORE_NO)
                Else
                    If ASCMAIN1.Logical_Lock("ARTCUST1", CUST_CODE, False, True, True, 1) Then
                        If MsgBox("OK to Add Address Record for " & CUST_CODE & "-" & CUST_STORE_NO, _
                                  MsgBoxStyle.YesNo, "Verification to " & e.Tool.Key) = MsgBoxResult.Yes Then
                            dst.Tables("ARTCUST2").Rows.Clear()
                            rowARTCUST2 = dst.Tables("ARTCUST2").NewRow
                            With rowARTCUST2
                                .Item("CUST_CODE") = CUST_CODE
                                .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
                                .Item("CUST_ADDR_CODE") = CUST_STORE_NO
                                .Item("CUST_NAME") = grd.ActiveRow.Cells("EDI_CUST_NAME_ADR").Value
                                .Item("CUST_ADDR1") = grd.ActiveRow.Cells("EDI_ADDRESS1").Value
                                .Item("CUST_ADDR2") = grd.ActiveRow.Cells("EDI_ADDRESS2").Value
                                .Item("CUST_CITY") = grd.ActiveRow.Cells("EDI_CITY").Value
                                .Item("CUST_STATE") = grd.ActiveRow.Cells("EDI_STATE").Value
                                Dim CUST_ZIP_CODE As String = grd.ActiveRow.Cells("EDI_ZIPCODE").Value & ""
                                If CUST_ZIP_CODE.Length = 9 Then CUST_ZIP_CODE = Mid(CUST_ZIP_CODE, 1, 5) & "-" & Mid(CUST_ZIP_CODE, 6, 4)
                                If CUST_ZIP_CODE.Length = 10 AndAlso Mid(CUST_ZIP_CODE, 6, 5) = "-0000" Then CUST_ZIP_CODE = Mid(CUST_ZIP_CODE, 1, 5)
                                .Item("CUST_ZIP_CODE") = CUST_ZIP_CODE
                                .Item("CUST_COUNTRY") = grd.ActiveRow.Cells("EDI_COUNTRY").Value
                                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                .Item("INIT_DATE") = DATETIME_STAMP
                                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                .Item("LAST_DATE") = DATETIME_STAMP
                                If GLN <> "" Then
                                    .Item("GLOBAL_LOCATION_NUMBER") = GLN
                                End If
                                dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2)

                                If e.Tool.Key = "Add as DC" Then
                                    rowARTCUST2.Item("CUST_DC_NO") = CUST_STORE_NO
                                    Dim rowARTCUST3 As DataRow = dst.Tables("ARTCUST3").NewRow
                                    rowARTCUST3.Item("CUST_CODE") = CUST_CODE
                                    rowARTCUST3.Item("CUST_ADDR_CODE") = CUST_STORE_NO
                                    rowARTCUST3.Item("CUST_ADDR_TYPE") = "MK"
                                    rowARTCUST3.Item("CUST_ADDR_TYPE2") = "DC"
                                    rowARTCUST3.Item("CUST_ADDR_CODE2") = CUST_STORE_NO
                                    dst.Tables("ARTCUST3").Rows.Add(rowARTCUST3)
                                    Update_Record_TDA("ARTCUST3")

                                    Dim rowARTCUST2_MK_for_DC As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
                                    If rowARTCUST2_MK_for_DC Is Nothing Then
                                        rowARTCUST2_MK_for_DC = dst.Tables("ARTCUST2").NewRow
                                        rowARTCUST2_MK_for_DC.ItemArray = rowARTCUST2.ItemArray
                                        rowARTCUST2_MK_for_DC.Item("CUST_ADDR_TYPE") = "MK"
                                        dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2_MK_for_DC)
                                    End If

                                    ' .Item("CUST_DC_IND") = "1"
                                Else
                                    '  .Item("CUST_DC_NO") = CUST_DC_NO
                                End If
                            End With

                            If GLN <> "" Then
                                ASCMAIN1.sql = "Update ARTCUST2 Set GLOBAL_LOCATION_NUMBER = NULL where GLOBAL_LOCATION_NUMBER = :PARM1"
                                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {GLN})
                            End If

                            Update_Record_TDA("ARTCUST2")
                            MsgBox("Address Record for " & CUST_CODE & "-" & CUST_STORE_NO & " has been added")
                        End If
                        ASCMAIN1.MultiTask_Release("", 0, 1)
                    End If
                End If

            Case "Add to Cross Reference"

                Dim CUST_CODE As String = grdEDT850T1.ActiveRow.Cells("CUST_CODE").Value
                Dim CUST_NAME As String = grdEDT850T1.ActiveRow.Cells("CUST_NAME").Value

                Dim UPC_TABLE As Boolean = False
                Dim EDI_SKU As String = ""
                Dim EDI_STYLE As String = ""
                Dim EDI_UPC As String = ""

                If grd.Name = "grdEDT850T6" Then
                    EDI_SKU = grdEDT850T6.ActiveRow.Cells("EDI_SLN_SKU").Value & ""
                    EDI_STYLE = grdEDT850T6.ActiveRow.Cells("EDI_SLN_STYLE").Value & ""
                    EDI_UPC = grdEDT850T6.ActiveRow.Cells("EDI_SLN_UPC").Value & ""
                Else
                    EDI_SKU = grdEDT850T2.ActiveRow.Cells("EDI_SKU").Value & ""
                    EDI_STYLE = grdEDT850T2.ActiveRow.Cells("EDI_STYLE").Value & ""
                    EDI_UPC = grdEDT850T2.ActiveRow.Cells("EDI_UPC").Value & ""
                End If

                Dim CUST_STYLE_CODE As String = EDI_SKU
                If CUST_STYLE_CODE = "" Then
                    CUST_STYLE_CODE = EDI_STYLE
                End If

                If CUST_STYLE_CODE = "" Then
                    CUST_STYLE_CODE = EDI_UPC
                    UPC_TABLE = True
                End If

                If CUST_STYLE_CODE = "" Then Exit Sub

                Dim sql_where As String = ""
                'ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE", , sql_where)
                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_COLOR", , sql_where)

                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()
                    If ASCMAIN1.CodeSelector.Selections <> 0 Then

                        Dim STYLE_CODE As String = ASCMAIN1.CodeSelector.SelectedCode
                        Dim STYLE_DESC As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STYLE_DESC") & ""
                        Dim COLOR_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("COLOR_CODE") & "" ' "AST"


                        If MsgBox("Please verify:" _
                                  & vbCrLf & vbCrLf & " Customer" & vbTab & vbTab & CUST_CODE & vbTab & CUST_NAME _
                                  & vbCrLf & " Style/SKU" & vbTab & vbTab & CUST_STYLE_CODE _
                                  & vbCrLf & " Our Style" & vbTab & vbTab & STYLE_CODE & "-" & COLOR_CODE & vbTab & STYLE_DESC _
                                  & vbCrLf & vbCrLf & "OK to Add this Cross Reference?", _
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                            If UPC_TABLE Then
                                Dim rowEDTUPCX1 As DataRow = LookUp("EDTUPCX1", New String() {CUST_CODE, CUST_STYLE_CODE})
                                If rowEDTUPCX1 IsNot Nothing Then
                                    MsgBox("Cross Reference for" _
                                           & vbCrLf & vbCrLf & " Customer " & CUST_CODE _
                                           & vbCrLf & " Customer Style/SKU " & CUST_STYLE_CODE _
                                           & vbCrLf & " to Style " & STYLE_CODE & "-" & COLOR_CODE _
                                           & vbCrLf & vbCrLf & " already exists")
                                Else
                                    ASCMAIN1.sql = "Insert into EDTUPCX1 (CUST_CODE,EDI_ITEM_CODE,STYLE_CODE,COLOR_CODE) Values (:PARM1,:PARM2,:PARM3,:PARM4)"
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New String() {CUST_CODE, CUST_STYLE_CODE, STYLE_CODE, COLOR_CODE})
                                    MsgBox("Cross Reference for" _
                                           & vbCrLf & vbCrLf & " Customer " & CUST_CODE _
                                           & vbCrLf & " Customer Style/SKU " & CUST_STYLE_CODE _
                                           & vbCrLf & " to Style " & STYLE_CODE & "-" & COLOR_CODE _
                                           & vbCrLf & vbCrLf & " has been Successfully Added")
                                End If
                            Else
                                Dim rowSOTCSTY1 As DataRow = LookUp("SOTCSTY1", New String() {CUST_CODE, CUST_STYLE_CODE})
                                If rowSOTCSTY1 IsNot Nothing Then
                                    MsgBox("Cross Reference for" _
                                           & vbCrLf & vbCrLf & " Customer " & CUST_CODE _
                                           & vbCrLf & " Customer Style/SKU " & CUST_STYLE_CODE _
                                           & vbCrLf & " to Style " & STYLE_CODE & "-" & COLOR_CODE _
                                           & vbCrLf & vbCrLf & " already exists")
                                Else
                                    ASCMAIN1.sql = "Insert into SOTCSTY1 (CUST_CODE,CUST_STYLE_CODE,STYLE_CODE,COLOR_CODE,INIT_OPER,INIT_DATE) Values (:PARM1,:PARM2,:PARM3,:PARM4,'" & ASCMAIN1.USER_ID & "',SYSDATE)"
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New String() {CUST_CODE, CUST_STYLE_CODE, STYLE_CODE, COLOR_CODE})
                                    MsgBox("Cross Reference for" _
                                           & vbCrLf & vbCrLf & " Customer " & CUST_CODE _
                                           & vbCrLf & " Customer Style/SKU " & CUST_STYLE_CODE _
                                           & vbCrLf & " to Style " & STYLE_CODE & "-" & COLOR_CODE _
                                           & vbCrLf & vbCrLf & " has been Successfully Added")
                                End If
                            End If
                        End If
                    End If
                End If

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

    Sub Load_EDT850T1()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Data")

        Create_Temp_Table()
        Fill_Records("EDT850T1")

        If ASCMAIN1.CLIENT = "RGI" Then
            EMsg = ""
            ASCMAIN1.sql = "Select * from EDT860T1 where nvl(EDI_PROCESS_IND,'0') = '0'"
            Dim dt As DataTable = ASCDATA1.GetDataTable
            If dt.Rows.Count <> 0 Then
                EMsg = vbCrLf & "Please check EDI for PO changes, New Records found"
            End If
            If Not ASCMAIN1.Running_in_VS Then
                ASCMAIN1.sql = "Select * from EDT824T1 where nvl(EDI_PROCESS_IND,'0') = '0'"
                dt = ASCDATA1.GetDataTable
                If dt.Rows.Count <> 0 Then
                    EMsg &= vbCrLf & "Please check EDI for EDI Messages for Application Advice, New Records found"
                End If
            End If
            If EMsg <> "" Then
                    MsgBox(EMsg, MsgBoxStyle.OkOnly, "EDI Documents Found")
                End If
                For Each row As DataRow In dst.Tables("EDT850T1").Select("")
                    Dim rowEDTXREF4 As DataRow = dst.Tables("EDTXREF4").Rows.Find(New String() {row.Item("EDI_TP_QUAL"), row.Item("EDI_TP_ID"), row.Item("EDI_SUPPLIER_NO")})
                    If rowEDTXREF4 IsNot Nothing Then
                        row.Item("WHSE_CODE") = rowEDTXREF4.Item("WHSE_CODE")
                    End If
                Next
            End If
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("EDT850TE"), New String() {"EDI_DOC_SEQ_NO"}).Rows
            Dim rowEDT850T1 As DataRow = dst.Tables("EDT850T1").Rows.Find(row.Item("EDI_DOC_SEQ_NO"))
            If rowEDT850T1 IsNot Nothing Then
                rowEDT850T1.Item("RESULT") = "Rejected"
            End If
        Next
        Sort_grdColumns(grdEDT850T1, "EDI_DOC_SEQ_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdEDT850T1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdEDT850T1.AfterRowActivate
        Setup_grdEDT850T1()
    End Sub

    Sub Setup_grdEDT850T1()
        If grdEDT850T1.ActiveRow Is Nothing OrElse Not grdEDT850T1.ActiveRow.IsDataRow Then
            tabEDT850T1.Visible = False
        Else
            tabEDT850T1.Visible = True
            EnforceConstraints(False)
            Dim EDI_DOC_SEQ_NO As String = grdEDT850T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value

            Fill_Records("EDT850T2", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T2, "EDI_DTL_SEQ")
            grdEDT850T2.Text = "Line Item Order Details for " & EDI_DOC_SEQ_NO

            Fill_Records("EDT850T3", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T3, "EDI_DTL_SEQ")
            grdEDT850T3.Text = "Store/Qty Order Details for " & EDI_DOC_SEQ_NO

            Fill_Records("EDT850T4", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T4, "EDI_CMT_SEQ")
            grdEDT850T4.Text = "Comments for " & EDI_DOC_SEQ_NO

            Fill_Records("EDT850T5", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T5, "EDI_ADR_SEQ")
            grdEDT850T5.Text = "Address Details for " & EDI_DOC_SEQ_NO

            Fill_Records("EDT850T6", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T6, "EDI_DTL_SEQ")
            grdEDT850T6.Text = "Assortment Details for " & EDI_DOC_SEQ_NO

            If tabEDT850T1.SelectedTab.Key = "Raw EDI Data" Then
                txtRaw.Text = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"))
            End If

            EnforceConstraints(True)
        End If
    End Sub

    Sub Create_Temp_Table()

        ASCDATA1.ExecuteSQL("Update EDT850T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_PROCESS_IND is Null")

        ASCMAIN1.sql = "Update EDT850T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
               & " where TRIM(EDI_OUR_ID) = TRIM(EDT850T1.EDI_OUR_ID) and TRIM(EDI_TP_ID) = TRIM(EDT850T1.EDI_TP_ID))" _
               & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Update EDT850T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
        '       & " where EDI_TP_QUAL = EDT850T1.EDI_TP_QUAL and EDI_TP_ID = EDT850T1.EDI_TP_ID and EDI_DOC_NO = 850)" _
        '       & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT850T1 Set CUST_CODE = (Select CUST_CODE from EDTSLSP1" _
               & " where EDI_QUAL_850 = EDT850T1.EDI_TP_QUAL and EDI_ID_850 = EDT850T1.EDI_TP_ID and NVL(EDI_CHAIN,'??') = NVL(EDT850T1.EDI_CHAIN,'??'))" _
               & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from EDT850T1 where EDI_PROCESS_IND = '0' and COMPANY_CODE is Null"
        If EDI_DOC_SEQ_NOs_no_company.Count <> 0 Then
            ASCMAIN1.sql &= " and EDI_DOC_SEQ_NO Not in ('" & Join(EDI_DOC_SEQ_NOs_no_company.ToArray, "','") & "')"
        End If
        Dim dt As DataTable = ASCDATA1.GetDataTable
        If dt.Rows.Count <> 0 Then
            For Each row As DataRow In dt.Rows
                EDI_DOC_SEQ_NOs_no_company.Add(row.Item("EDI_DOC_SEQ_NO"))
            Next
            Using frm As New ASFMSGBF
                frm.Show_grd(dt, Me, "EDI Transactions which could not be mapped to an ABSolution Company")
            End Using
        End If

        If EDT850T1 = "" Then
            ASCMAIN1.sql = "Select * from EDT850T1 where ROWNUM < 1"
            EDT850T1 = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select * from SOTORDR1 where ROWNUM < 1"
            SOTORDR1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")

            ASCMAIN1.sql = "Select * from SOTORDR2 where ROWNUM < 1"
            SOTORDR2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add Primary Key (ORDR_NO,ORDR_LNO)")

            ASCMAIN1.sql = "Select * from SOTORDR3 where ROWNUM < 1"
            SOTORDR3 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR3 & " Add Primary Key (ORDR_NO,ORDR_LNO,ORDR_SUB_LNO)")

            ASCMAIN1.sql = "Select * from SOTORDR4 where ROWNUM < 1"
            SOTORDR4 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR4 & " Add Primary Key (ORDR_NO,ORDR_CLNO)")

            ASCMAIN1.sql = "Select * from SOTORDR9 where ROWNUM < 1"
            SOTORDR9 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR9 & " Add Primary Key (ORDR_NO,RANGE_STYLE_LNO)")

        Else
            ASCMAIN1.sql = "Truncate Table " & EDT850T1
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Select * from EDT850T1 where EDI_PROCESS_IND = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            ASCDATA1.ExecuteSQL("Insert into " & EDT850T1 & " " & ASCMAIN1.sql)
        End If
    End Sub

    Sub Delete_Orders()

        BeginTrans()
        For Each rowEDT850T1 As DataRow In dst.Tables("EDT850T1").Select("SEL = '1'")
            Dim EDI_DOC_SEQ_NO As String = rowEDT850T1.Item("EDI_DOC_SEQ_NO")
            ASCMAIN1.sql = "Update EDT850T1" _
                & " Set EDI_PROCESS_IND = 'C' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next
        CommitTrans("Selected EDI Orders have been Deleted")

    End Sub


    Sub Asst_SDQT0(RANGE_STYLE_CODE As String, RANGE_STYLE_PRICE As Decimal, EDI_UPC As String)
        'Sql = "Select * from ICWAITM2 where ASST_STYLE_CODE = '" & ASST_STYLE_CODE & "'"
        'dynICWAITM2 = AccD.OpenRecordset(Sql, dbOpenDynaset)
        'Do While Not dynICWAITM2.EOF
        '    STYLE_CODE = dynICWAITM2.Item("STYLE_CODE")
        '    EDI_SKU = dynICWAITM2.Item("CUST_SKU") & ""
        '    ASST_SKU = dynICWAITM2.Item("CUST_SKU") & ""
        '    ASST_UPC = dynICWAITM2.Item("CUST_UPC") & ""
        '    PRICE = dynICWAITM2.Item("ITEM_PRICE")
        '    rowICTSTYL1.Seek("=", STYLE_CODE)
        '    If rowICTSTYL1.NoMatch Then
        '        OraD.Parameters("STYLE_CODE") = STYLE_CODE
        '        dynICTSTYL1.Refresh()
        '        If Not dynICTSTYL1.EOF Then
        '            Write_Item_to_MDB()
        '        End If
        '    End If
        '    Write_SDQT0()
        '    dynICWAITM2.MoveNext()
        'Loop

        Dim sql As String = "CUST_CODE = '" & CUST_CODE & "' and RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "'"
        For Each rowICTRSTY2 As DataRow In dst.Tables("ICTRSTY2").Select(sql, "STYLE_CODE,COLOR_CODE")

            STYLE_CODE = rowICTRSTY2.Item("STYLE_CODE")
            COLOR_CODE = rowICTRSTY2.Item("COLOR_CODE")

            'rowICTSTYL1 = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
            'If rowICTSTYL1 Is Nothing Then
            '    rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
            '    If rowICTSTYL1 IsNot Nothing Then
            '        rowICTSTYL1 = Write_Item_to_MDB(rowICTSTYL1)
            '    End If
            'End If
            'saleable_item = True

            'Dim CUST_UPC As String = rowICTRSTY2.Item("CUST_UPC") & ""
            'Dim SIZE_CODE As String = rowICTRSTY2.Item("SIZE_CODE") & ""
            'Dim CUST_SKU As String = rowICTRSTY2.Item("CUST_SKU") & ""
            'Dim STYLE_PRICE As Decimal = Val(rowICTRSTY2.Item("STYLE_PRICE") & "")
            'Dim STYLE_QTY As Decimal = Val(rowICTRSTY2.Item("STYLE_QTY") & "")

            '  Dim STYLE_PRICE As Decimal = RANGE_STYLE_PRICE
            Write_SDQT0(EDI_UPC, RANGE_STYLE_PRICE, RANGE_STYLE_PRICE, RANGE_STYLE_PRICE, RANGE_STYLE_PRICE)
        Next
    End Sub

    Sub Asst_Item_SDQ( _
        EDI_UPC As String, _
        EDI_SKU As String, _
        EDI_ITEM As String, _
        EDI_STYLE As String, _
        EDI_STORE As String, _
        EDI_SHIP_DC As String, QTY As Int32, EDI_PRICE As Decimal, _
        EDI_COLOR_CODE As String, _
        EDI_COLOR_NAME As String, _
        EDI_SIZE_DESC As String, _
        RNG_AST_FLG As String, _
        SLN_PARENT_STYLE_CODE As String, _
        SLN_PARENT_STYLE_DESC As String, _
        SLN_PARENT_STYLE_QTY As Int64, _
        SLN_PARENT_INNER_PACK_QTY As Int64, _
        EDI_SLN_SEQ As Int32, _
        RANGE_STYLE_QTY_PER_PP As Int32, _
        QTY_PER_PP As Int32, _
        WORK_QTY As Int64, _
        RANGE_STYLE_CODE As String, _
        RANGE_STYLE_PRICE As Decimal, _
        RANGE_UPC As String, _
        RANGE_SKU As String)
        'ASCMAIN1.sql = "Select * from ICTAITM2 where ASST_STYLE_CODE = '" & ASST_STYLE_CODE & "'"
        'For Each rowICTAITM2 As DataRow In ASCDATA1.GetDataTable.Select("")
        '    STYLE_CODE = rowICTAITM2.Item("STYLE_CODE")
        '    ASST_UPC = rowICTAITM2.Item("CUST_UPC") & ""
        '    ASST_SKU = rowICTAITM2.Item("CUST_SKU") & ""
        '    WORK_QTY = QTY
        '    QTY = rowICTAITM2.Item("ITEM_QTY")
        '    Write_SDQ()
        '    QTY = WORK_QTY
        'Next

        Dim sql As String = "CUST_CODE = '" & CUST_CODE & "' and RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "'"
        For Each rowICTRSTY2 As DataRow In dst.Tables("ICTRSTY2").Select(sql, "STYLE_CODE,COLOR_CODE")

            STYLE_CODE = rowICTRSTY2.Item("STYLE_CODE")
            COLOR_CODE = rowICTRSTY2.Item("COLOR_CODE")

            rowICTSTYL1 = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
            If rowICTSTYL1 Is Nothing Then
                rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    rowICTSTYL1 = Write_Item_to_MDB(rowICTSTYL1)
                End If
            End If

            Dim CUST_UPC As String = rowICTRSTY2.Item("CUST_UPC") & ""
            Dim SIZE_CODE As String = rowICTRSTY2.Item("SIZE_CODE") & ""
            Dim CUST_SKU As String = rowICTRSTY2.Item("CUST_SKU") & ""
            Dim STYLE_QTY As Int64 = Val(rowICTRSTY2.Item("STYLE_QTY") & "")

            Write_SDQ(EDI_UPC, EDI_SKU, _
                    EDI_ITEM, _
                    EDI_STYLE, _
                    EDI_STORE, _
                    EDI_SHIP_DC, QTY, EDI_PRICE, _
                    EDI_COLOR_CODE, _
                    EDI_COLOR_NAME, _
                    EDI_SIZE_DESC, _
                    RNG_AST_FLG, _
                    SLN_PARENT_STYLE_CODE, _
                    SLN_PARENT_STYLE_DESC, _
                    SLN_PARENT_STYLE_QTY, _
                    SLN_PARENT_INNER_PACK_QTY, _
                    EDI_SLN_SEQ, _
                    RANGE_STYLE_QTY_PER_PP, _
                    QTY_PER_PP, _
                    STYLE_QTY, _
                    RANGE_STYLE_CODE, _
                    RANGE_STYLE_PRICE, _
                    RANGE_UPC, _
                    RANGE_SKU)
        Next
    End Sub

    Sub Import_Orders_using_Class()
        ' do we instantiate 1 order at a time
        ' or for a list of orders - performance

        Dim EDI_DOC_SEQ_NOs As New List(Of String)
        For Each row As DataRow In dst.Tables("EDT850T1").Select("SEL='1'")
            Dim EDI_DOC_SEQ_NO As String = row.Item("EDI_DOC_SEQ_NO")
            EDI_DOC_SEQ_NOs.Add(EDI_DOC_SEQ_NO)
        Next
        Dim g As New WHC.GunEnvironment

        g.DBS_COMPANY = ASCMAIN1.DBS_COMPANY
        g.DBS_PASSWORD = ASCMAIN1.DBS_PASSWORD
        g.DBS_SERVER = ASCMAIN1.DBS_SERVER

        Dim x As New WHC.EDC850I1(EDI_DOC_SEQ_NOs, g, EDT850T1, dst.Tables("EDT850TE"))

        x.Import_Orders()

        Stop

        ' NEED EDTPARM1 SETTING THAT SAYS AUTOMATION IS ACTIVE

        ' IF AUTOMATION IS NOT ACTIVE, THEN FORM WILL:
        ' sets cust code and company
        ' process ind = '0' 


        ' SERVICE WILL:
        ' sets cust code and company
        ' process ind = 'A'

        ' PROCESS_IND = '1' IF ORDER WAS SUCCESSFULLY IMPORTED
        ' PROCESS_IND = '0' IF NOT IMPORTED

    End Sub

    Sub Import_Orders()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Importing EDI Orders")

        Fill_Records("ICTITEMX")

        ' Loop thru and Process/Cancel Selected EDI Orders

        Dim ORDERS_PROCESSED As Int32 = 0
        Dim ORDERS_IMPORTED As Int32 = 0
        Dim ORDERS_REJECTED As Int32 = 0
        Dim EDI_SUPPLIER_NO As String
        'Dim ECOM_CODE As String = ""

        Dim sqlw As String = ""
        Dim EDI_RECEIVED_VALUE As String = ""
        CUST_CODE = ""

        For Each rowEDT850T1 In dst.Tables("EDT850T1").Select("SEL = '1'", "CUST_CODE")
            ORDERS_PROCESSED += 1

            EDI_DOC_SEQ_NO = rowEDT850T1.Item("EDI_DOC_SEQ_NO")
            EDI_TP_QUAL = rowEDT850T1.Item("EDI_TP_QUAL") & ""
            EDI_TP_ID = rowEDT850T1.Item("EDI_TP_ID") & ""
            ORDR_CUST_PO = rowEDT850T1.Item("EDI_PO_NO") & ""
            ORDR_DEPT = rowEDT850T1.Item("EDI_DEPARTMENT") & ""
            ORDR_DATE = rowEDT850T1.Item("EDI_PO_DATE")
            EDI_MERCH_TYPE = rowEDT850T1.Item("EDI_MERCH_TYPE") & ""
            EDI_SHIPPER = rowEDT850T1.Item("EDI_SHIPPER") & ""

            EDI_PO_TYPE = rowEDT850T1.Item("EDI_PO_TYPE") & ""
            EDI_SUPPLIER_NO = rowEDT850T1.Item("EDI_SUPPLIER_NO") & ""

            ECOM_CODE = ""
            ECOM_PRICE_TOLERANCE_PCT = 0
            DROP_SHIP = False
            ALT_PRICING = False


            If ASCMAIN1.CLIENT = "RGI" Then
                Dim rowsECTECOM1() As DataRow = dst.Tables("ECTECOM1").Select("EDI_TP_QUAL = '" & EDI_TP_QUAL & "' and EDI_TP_ID = '" & EDI_TP_ID & "'")
                If rowsECTECOM1.Length <> 1 Then
                    ' This is an error that needs to be reported, it should never happen, until we do non-ecomm orders like Nordstrom
                    'Throw New Exception("eCommerce Code not properly defined - CALL ABS")
                Else
                    If ECOM_CODE <> rowsECTECOM1(0).Item("ECOM_CODE") & "" Then
                        Dim TMP_ECOM_CODE As String = rowsECTECOM1(0).Item("ECOM_CODE") & ""
                        ECOM_PRICE_TOLERANCE_PCT = Val(rowsECTECOM1(0).Item("ECOM_PRICE_TOLERANCE_PCT") & "")
                        Fill_Records("ECTESTY1", TMP_ECOM_CODE)
                        Fill_Records("ECTESTY3", TMP_ECOM_CODE)
                        If EDI_PO_TYPE = "DS" Or
                            (TMP_ECOM_CODE = "WAYFAIR" And (EDI_PO_TYPE = "SO" Or EDI_PO_TYPE = "PR" Or EDI_PO_TYPE = "RC")) Or
                            (TMP_ECOM_CODE = "KIRKLANDS" And EDI_PO_TYPE = "SA") Or
                            (TMP_ECOM_CODE = "HOUZZ" And EDI_PO_TYPE = "OS") Or
                            TMP_ECOM_CODE = "AMAZON" Or
                            TMP_ECOM_CODE = "HOMEDEPOT" Or
                            TMP_ECOM_CODE = "XMASCENT" Or
                            TMP_ECOM_CODE = "KOHLS" Or
                            TMP_ECOM_CODE = "QVC" Then
                            ECOM_CODE = TMP_ECOM_CODE
                            DROP_SHIP = True
                            If (TMP_ECOM_CODE = "WAYFAIR" And EDI_PO_TYPE = "SO") Or (TMP_ECOM_CODE = "HOUZZ" And EDI_PO_TYPE = "OS") Then
                                ALT_PRICING = True
                            End If
                        End If
                    End If
                End If
            End If

            ' BK is a Blanket Order - and will be imported into SOTORDRx with ORDR_HOLD = '1' 
            ' We may want to move this data into a reservation at some point, 
            '  because the process usually is that the customer will send a follow-up order

            If CUST_CODE <> rowEDT850T1.Item("CUST_CODE") & "" Then
                CUST_CODE = rowEDT850T1.Item("CUST_CODE") & ""
                Fill_Records("ARTCUST2", CUST_CODE)
                Fill_Records("SOTCSTY1", CUST_CODE)
                Fill_Records("EDTUPCX1", CUST_CODE)
                Fill_Records("SOTPRIC2", CUST_CODE)
                Dim row As DataRow = LookUp("ARTCUST1", CUST_CODE)
                PRICE_LIST_CODE = row.Item("PRICE_LIST_CODE") & ""
                If CUST_CODE <> PRICE_LIST_CODE Then
                    Fill_Records("SOTPRIC2", PRICE_LIST_CODE)
                End If

                Fill_Records("ICTRSTY1", CUST_CODE)
                Fill_Records("ICTRSTY2", CUST_CODE)
            End If

            ASCMAIN1.Progress("-", EDI_DOC_SEQ_NO)

            Dim STORE_GLOBAL_LOCATION_NUMBER As String = rowEDT850T1.Item("STORE_GLOBAL_LOCATION_NUMBER") & ""
            Dim DC_GLOBAL_LOCATION_NUMBER As String = rowEDT850T1.Item("DC_GLOBAL_LOCATION_NUMBER") & ""
            Dim GLN As String = DC_GLOBAL_LOCATION_NUMBER
            If GLN = "" And STORE_GLOBAL_LOCATION_NUMBER <> "" Then
                GLN = STORE_GLOBAL_LOCATION_NUMBER
            End If

            Dim EDI_STORE As String = rowEDT850T1.Item("EDI_STORE") & ""
            If EDI_STORE = "" Then
                If GLN <> "" Then
                    ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = :PARM1 and CUST_ADDR_TYPE = 'MK' and GLOBAL_LOCATION_NUMBER = :PARM2"
                    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() {CUST_CODE, GLN})
                    If row IsNot Nothing Then
                        EDI_STORE = row.Item("CUST_ADDR_CODE")
                    End If
                End If
            End If

            CUST_STORE_NO = ""
            If EDI_STORE <> "" Then
                If Len(EDI_STORE) > 6 Then ' probably a gln
                    ASCMAIN1.sql = "Select * from ARTCUST2 where GLOBAL_LOCATION_NUMBER = :PARM1"
                    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {EDI_STORE})
                    If row Is Nothing Then
                        CUST_STORE_NO = ""
                    Else
                        CUST_STORE_NO = row.Item("CUST_ADDR_CODE")
                    End If
                Else
                    CUST_STORE_NO = Format_Store(EDI_STORE)
                End If
            End If

            dst.Tables("EDT850TS").Rows.Clear()
            dst.Tables("EDTITEMX").Rows.Clear()
            dst.Tables("EDTSDQT1").Rows.Clear()
            dst.Tables("EDTSDQT0").Rows.Clear()

            EDI_ERRORs.Clear()
            EDI_DOC_SEQ_NO_ok = True

            Dim EDI_ERROR_NO As Int32 = 0
            For Each row As DataRow In dst.Tables("EDT850TE").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'", "EDI_ERROR_NO")
                If row.Item("EDI_ACTION") = "N" Then
                    row.Delete()
                Else
                    EDI_ERROR_NO += 1
                    row.Item("EDI_ORDER_COUNT") = 0
                    row.Item("EDI_ERROR_NO") = EDI_ERROR_NO
                    Dim EDI_ERROR As String = row.Item("EDI_COND_DESC") & vbTab & row.Item("EDI_COND_CODE") & vbTab & row.Item("EDI_RECEIVED_VALUE")
                    EDI_ERRORs.Add(EDI_ERROR, EDI_ERROR_NO)
                End If
            Next

            rowEDTTRPM1 = Get_EDTTRPM1()

            sqlw = String.Format("EDI_DOC_SEQ_NO <> '{0}' and EDI_TP_QUAL = '{1}' and EDI_TP_ID = '{2}' and EDI_PO_NO = '{3}' and ISNULL(EDI_STORE,'') = '{4}'", _
                                  EDI_DOC_SEQ_NO, EDI_TP_QUAL, EDI_TP_ID, ORDR_CUST_PO, EDI_STORE)

            If dst.Tables("EDT850T1").Select(sqlw).Length <> 0 Then
                'If ASCMAIN1.Running_in_VS Then Stop
                Bad_Data(EDI_COND_DESC:="Possible PO Duplication within Import", _
                            EDI_COND_CODE:="27", _
                            EDI_RECEIVED_VALUE:=EDI_DOC_SEQ_NO)
            End If

            If EDI_PO_TYPE = "BK" Then
                Bad_Data(EDI_COND_DESC:="Order is a Blanket Order", _
                EDI_COND_CODE:="50", _
                EDI_RECEIVED_VALUE:=EDI_DOC_SEQ_NO)
            End If

            If EDI_DOC_SEQ_NO_ok Then

                Fill_Records("EDT850T2", EDI_DOC_SEQ_NO)

                ASCMAIN1.sql = "Select EDI_ITEM, EDI_UPC, EDI_SKU, EDI_GTIN, Count (*)" _
                    & " from EDT850T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
                    & " and (EDI_ITEM IS NOT NULL OR EDI_UPC IS NOT NULL OR EDI_SKU IS NOT NULL OR EDI_GTIN IS NOT NULL)" _
                    & " group by EDI_ITEM, EDI_UPC, EDI_SKU, EDI_GTIN having Count (*) > 1"
                For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                    If row.Item("EDI_ITEM") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "Item: " & row.Item("EDI_ITEM")
                    ElseIf row.Item("EDI_UPC") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "UPC: " & row.Item("EDI_UPC")
                    ElseIf row.Item("EDI_SKU") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "SKU: " & row.Item("EDI_SKU")
                    ElseIf row.Item("EDI_GTIN") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "GTIN: " & row.Item("EDI_GTIN")
                    End If
                    Bad_Data(EDI_COND_DESC:="Item Duplication within PO", _
                            EDI_COND_CODE:="28", _
                            EDI_RECEIVED_VALUE:=EDI_RECEIVED_VALUE)
                Next

                Fill_Records("EDT850T3", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T4", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T5", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T6", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T7", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T8", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T9", EDI_DOC_SEQ_NO)
            End If

            Dim EDI_SHIP_DC As String = ""

            If EDI_DOC_SEQ_NO_ok Then
                Get_ARTCUST1()

                SHIP_VIA_CODE = ""

                If rowEDT850T1.Item("EDI_SHIP_DC") & "" <> "" Then
                    EDI_SHIP_DC = rowEDT850T1.Item("EDI_SHIP_DC") & ""
                Else
                    EDI_SHIP_DC = rowEDT850T1.Item("EDI_CENTER_CODE") & ""
                End If

                If rowEDTSLSP1.Item("EDI_CONSUMER") & "" = "1" Or ECOM_CODE <> "" Then
                    EDI_STORE = "000000"
                    CUST_STORE_NO = EDI_STORE
                    EDI_SHIP_DC = "MK"
                    ORDR_SHIP_TO = "MK"
                    ' Can only support 1 address with this indicator – throwing an exception if Not the case.
                    ' If there Then are no address record In EDT850T5 – throwing an exception,

                    'ASCMAIN1.sql = "Select * from EDT850T5 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_ADDR_TYPE = 'ST'"
                    'Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("") ' dst.Tables("EDT850T5").Select("")

                    Dim rows() As DataRow = dst.Tables("EDT850T5").Select("EDI_ADDR_TYPE = 'ST'")
                    If rows.Length = 0 Then
                        Bad_Data(EDI_COND_DESC:="No Ship To Address Records found in EDI data",
                            EDI_COND_CODE:="63",
                            EDI_RECEIVED_VALUE:="Missing EDI data")
                    ElseIf rows.Length > 1 Then
                        Bad_Data(EDI_COND_DESC:="Multiple Ship To Address Records found in EDI data",
                            EDI_COND_CODE:="64",
                            EDI_RECEIVED_VALUE:="Multiple Addresses Found")
                    End If

                    If ASCMAIN1.CLIENT = "RGI" Then
                        If EDI_SHIPPER.StartsWith("_") Then
                            If EDI_SUPPLIER_NO = "23249" Then
                                EDI_SHIPPER = "UPSN" & EDI_SHIPPER
                            Else
                                SHIP_VIA_CODE = "ECOM" ' Castlegate orders use ECOM when shipper not sent
                            End If
                        End If
                        If EDI_SHIPPER = "" And EDI_PO_TYPE = "NA" Then
                            'xfer order
                            SHIP_VIA_CODE = "UPSG"
                        End If
                        If SHIP_VIA_CODE = "" Then
                            Dim tSHIPPER = EDI_SHIPPER.ToUpper.Substring(0, Math.Min(10, EDI_SHIPPER.Length))
                            Dim rowEDTXREF3 As DataRow = dst.Tables("EDTXREF3").Rows.Find(New String() {EDI_TP_QUAL, EDI_TP_ID, tSHIPPER})
                            If rowEDTXREF3 Is Nothing Then
                                Bad_Data(EDI_COND_DESC:="Ship Via X-Ref Not Found (" & EDI_SHIPPER & ")",
                                EDI_COND_CODE:="65",
                                EDI_RECEIVED_VALUE:="Missing EDI data")
                            Else
                                SHIP_VIA_CODE = rowEDTXREF3.Item("SHIP_VIA_CODE") & ""
                                Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                                If rowSOTSVIA1 Is Nothing Then
                                    Bad_Data(EDI_COND_DESC:="Bad Ship Via Code in X-Ref for EDI Shipper (" & EDI_SHIPPER & ")",
                                    EDI_COND_CODE:="66",
                                    EDI_RECEIVED_VALUE:="Bad / Missing EDI X-Ref data")
                                ElseIf EDI_SUPPLIER_NO = "23249" And DROP_SHIP And rowSOTSVIA1.Item("CARRIER_CODE") <> "FEDEX" Then ' Dropship orders from MS should be UPS carrier
                                    Bad_Data(EDI_COND_DESC:="MountainSide Order Should ship FEDEX,  set for (" & rowSOTSVIA1.Item("CARRIER_CODE") & ")",
                                    EDI_COND_CODE:="68",
                                    EDI_RECEIVED_VALUE:="Bad Ship Via for MS")
                                End If
                            End If
                        End If
                    Else
                        Dim rowEDTXREF3 As DataRow = dst.Tables("EDTXREF3").Rows.Find(New String() {EDI_TP_QUAL, EDI_TP_ID, EDI_SHIPPER})
                        If rowEDTXREF3 Is Nothing Then
                            Bad_Data(EDI_COND_DESC:="No Shipper Defined or Ship Via X-Ref Not Found (" & EDI_SHIPPER & ")",
                            EDI_COND_CODE:="65",
                            EDI_RECEIVED_VALUE:="Missing EDI data")
                        Else
                            SHIP_VIA_CODE = rowEDTXREF3.Item("SHIP_VIA_CODE") & ""
                            Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                            If rowSOTSVIA1 Is Nothing Then
                                Bad_Data(EDI_COND_DESC:="Bad Ship Via Code in X-Ref for EDI Shipper (" & EDI_SHIPPER & ")",
                                EDI_COND_CODE:="66",
                                EDI_RECEIVED_VALUE:="Bad / Missing EDI X-Ref data")
                            End If
                        End If
                    End If
                Else
                    '  Dim DC_GLOBAL_LOCATION_NUMBER As String = rowEDT850T1.Item("DC_GLOBAL_LOCATION_NUMBER") & ""
                    ' Dim STORE_GLOBAL_LOCATION_NUMBER As String = rowEDT850T1.Item("STORE_GLOBAL_LOCATION_NUMBER") & ""
                    If DC_GLOBAL_LOCATION_NUMBER = "" And STORE_GLOBAL_LOCATION_NUMBER <> "" Then
                        DC_GLOBAL_LOCATION_NUMBER = STORE_GLOBAL_LOCATION_NUMBER
                        ' WALMART DC'S ARE MAPPED INTO STORE_GLOBAL_LOCATION_NUMBERS
                    End If
                    If EDI_SHIP_DC = "" Then
                        If DC_GLOBAL_LOCATION_NUMBER <> "" Then
                            ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = :PARM1 and CUST_ADDR_TYPE = 'DC' and GLOBAL_LOCATION_NUMBER = :PARM2"
                            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() {CUST_CODE, DC_GLOBAL_LOCATION_NUMBER})
                            If row IsNot Nothing Then
                                EDI_SHIP_DC = row.Item("CUST_ADDR_CODE")
                            End If
                        End If
                    End If
                    If EDI_SHIP_DC <> "" Then
                        ORDR_SHIP_TO = "DC"
                    End If
                End If

                If ASCMAIN1.CLIENT = "RGI" Then
                    If Not New String() {"007942915", "NORDJWN"}.Contains(EDI_TP_ID) Then ' skip nordstrom and others
                        If Not String.IsNullOrEmpty(EDI_SUPPLIER_NO) Then
                            'EDTXREF4
                            Dim rowEDTXREF4 As DataRow = dst.Tables("EDTXREF4").Rows.Find(New String() {EDI_TP_QUAL, EDI_TP_ID, EDI_SUPPLIER_NO})
                            If rowEDTXREF4 Is Nothing Then
                                Bad_Data(EDI_COND_DESC:="Supplier No X-Ref Not Found (" & EDI_SUPPLIER_NO & ")",
                                EDI_COND_CODE:="67",
                                EDI_RECEIVED_VALUE:="Missing EDI data")
                            Else
                                EDI_STORE = rowEDTXREF4.Item("CUST_STORE_NO") & ""
                                WHSE_CODE = rowEDTXREF4.Item("WHSE_CODE") & ""
                            End If
                        End If
                        If EDI_PO_TYPE = "NA" Then
                            EDI_SHIP_DC = rowEDT850T1.Item("EDI_SHIP_DC") & ""
                        End If
                    End If
                End If

                If EDI_STORE <> "" Or EDI_SHIP_DC <> "" Then
                    Get_Ship_To(EDI_STORE, EDI_SHIP_DC)
                Else
                    TRANSIT_BUS_DAYS = 5 ' SETTING A DEFAULT IN CASE WE DO NOT KNOW ANYTHING ABOUT SHIP TO AT THIS POINT
                End If

                ' Load values for Ship & Cancel Dates, and check for Reasonableness

                ORDR_CANCEL_DATE = Nothing
                ORDR_SHIP_DATE = Nothing
                ORDR_ARRIVAL_DATE = Nothing
                ORDR_LAST_ARRIVAL_DATE = Nothing

                If rowEDT850T1.Item("EDI_ARRIVAL_DATE") & "" <> "" Then
                    ORDR_ARRIVAL_DATE = rowEDT850T1.Item("EDI_ARRIVAL_DATE") & ""
                End If
                If rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE") & "" <> "" Then
                    ORDR_LAST_ARRIVAL_DATE = rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE") & ""
                End If

                Dim ORDR_CANCEL_DATE_based_on As String = ""
                If rowEDT850T1.Item("EDI_END_DATE") & "" <> "" Then
                    ORDR_CANCEL_DATE = rowEDT850T1.Item("EDI_END_DATE")
                    ORDR_CANCEL_DATE_based_on = "EDI_END_DATE"
                Else
                    If rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE") & "" <> "" Then
                        ORDR_CANCEL_DATE = rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE")
                        ORDR_CANCEL_DATE_based_on = "EDI_LAST_ARRIVAL_DATE"
                    ElseIf rowEDT850T1.Item("EDI_ARRIVAL_DATE") & "" <> "" Then
                        ORDR_CANCEL_DATE = rowEDT850T1.Item("EDI_ARRIVAL_DATE")
                        ORDR_CANCEL_DATE_based_on = "EDI_ARRIVAL_DATE"
                    End If
                End If

                If rowEDT850T1.Item("EDI_START_DATE") & "" <> "" Then
                    ORDR_SHIP_DATE = rowEDT850T1.Item("EDI_START_DATE")
                ElseIf rowEDT850T1.Item("EDI_SHIP_DATE") & "" <> "" Then
                    ORDR_SHIP_DATE = rowEDT850T1.Item("EDI_SHIP_DATE")
                ElseIf Format(ORDR_CANCEL_DATE, "yyyyMMdd") <> "00010101" Then
                    ORDR_SHIP_DATE = ASCMAIN1.DateDiff_Weekday(ORDR_CANCEL_DATE, -1 * TRANSIT_BUS_DAYS)
                    If ORDR_CANCEL_DATE_based_on = "EDI_LAST_ARRIVAL_DATE" Then
                        If Format(ORDR_ARRIVAL_DATE, "yyyyMMdd") <> "00010101" _
                            And Format(ORDR_ARRIVAL_DATE, "yyyyMMdd") < Format(ORDR_SHIP_DATE, "yyyyMMdd") Then
                            ORDR_SHIP_DATE = ORDR_ARRIVAL_DATE
                        End If
                    End If
                End If

                If Format(ORDR_CANCEL_DATE, "yyyyMMdd") = "00010101" Then
                    If Format(ORDR_SHIP_DATE, "yyyyMMdd") <> "00010101" Then
                        ORDR_CANCEL_DATE = ORDR_SHIP_DATE.AddDays(6)
                    End If
                End If

                If Format(ORDR_CANCEL_DATE, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                    EDI_REPLACED_VALUE = Bad_Data(
                        EDI_COND_DESC:="Cancel Date is past",
                        EDI_COND_CODE:="09",
                        EDI_RECEIVED_VALUE:=Format$(ORDR_CANCEL_DATE, "MM/dd/yyyy"))
                    If EDI_REPLACED_VALUE <> "" Then
                        ORDR_CANCEL_DATE = DateValue(EDI_REPLACED_VALUE)
                    End If
                    'If EDI_RECEIVED_VALUE <> "" Then
                    '    ORDR_CANCEL_DATE = DateValue(EDI_RECEIVED_VALUE)
                    'End If
                End If

                If Format(ORDR_SHIP_DATE, "yyyyMMdd") = "00010101" Then
                    EDI_REPLACED_VALUE = Bad_Data(
                                           EDI_COND_DESC:="Missing Ship Date",
                                           EDI_COND_CODE:="08",
                                           EDI_RECEIVED_VALUE:=Format$(ORDR_SHIP_DATE, "MM/dd/yyyy"))
                    ' we do not allow accepting bad edi value for this field, so follownig lines are not used
                    If EDI_REPLACED_VALUE <> "" Then
                        ORDR_SHIP_DATE = DateValue(EDI_REPLACED_VALUE)
                    End If
                    'If EDI_RECEIVED_VALUE <> "" Then
                    '    ORDR_SHIP_DATE = DateValue(EDI_RECEIVED_VALUE)
                    'End If
                Else
                    If Format(ORDR_SHIP_DATE, "yyyyMMdd") = "00010101" Then
                        ORDR_CANCEL_DATE = ORDR_SHIP_DATE.AddDays(7)
                    End If
                End If

                If CUST_CODE <> "" Then
                    Get_Terms()
                    Check_for_Possible_Order_Duplication()
                End If

                Process_EDT850T2(EDI_STORE, EDI_SHIP_DC)

                If dst.Tables("EDT850TS").Rows.Count = 0 Then
                    If EDI_STORE = "" Then
                        Bad_Data(EDI_COND_DESC:="No Store found for Order, Verify EDI Store",
                                EDI_COND_CODE:="40",
                                EDI_RECEIVED_VALUE:="Missing Store")
                    End If

                    ' NO POINT IN LOOP BELOW BECAUSE WE WILL ENTER THIS LOOP IMMEDIATELY WHEN WE CREATE ORDERS

                    'Else
                    '    For Each rowEDT850TS As DataRow In dst.Tables("EDT850TS").Select("", "EDI_STORE")
                    '        EDI_STORE = rowEDT850TS.Item("EDI_STORE")
                    '        EDI_SHIP_DC = rowEDT850TS.Item("EDI_SHIP_DC")
                    '        CUST_STORE_NO = rowEDT850TS.Item("CUST_STORE_NO") & ""
                    '        CUST_DC_NO = rowEDT850TS.Item("CUST_DC_NO") & ""
                    '        Get_Ship_To(EDI_STORE, EDI_SHIP_DC)
                    '    Next
                End If

                If ASCMAIN1.CLIENT = "RGI" Then
                    If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE, False, True, False, 2) Then

                        Bad_Data(EDI_COND_DESC:="Unable to lock Reservations Record, try again.",
                                EDI_COND_CODE:="41",
                                EDI_RECEIVED_VALUE:="Record Locked")
                    End If
                End If


                ' ****************************************************************
                ' If the Journal Passed all tests, then generate orders from SDQ's

                If EDI_DOC_SEQ_NO_ok Then

                    Dim order_count As Int32 = dst.Tables("EDT850TS").Rows.Count
                    Dim ORDER_COUNTER As Int32 = 0

                    BeginTrans()

                    Dim ORDR_NO_next As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO", order_count)
                    Dim ORDR_NO_beg = ORDR_NO_next
                    Dim ORDR_NO_end = ORDR_NO_next

                    STYLE_CODE = ""
                    'EDI_ITEM = ""

                    Dim ORDR_GROUP_NO As String = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                    Dim ORDR_LNO As Int32 = 0

                    ASCMAIN1.Progress("Generating Orders", "")

                    ORDERS_IMPORTED += 1

                    dst.Tables("SOTORDR1").Rows.Clear()
                    dst.Tables("SOTORDR2").Rows.Clear()
                    dst.Tables("SOTORDR9").Rows.Clear()
                    Dim RANGE_STYLE_CODE_seqs As New Dictionary(Of String, Integer)

                    For Each rowEDT850TS As DataRow In dst.Tables("EDT850TS").Select("", "EDI_STORE")
                        EDI_STORE = rowEDT850TS.Item("EDI_STORE")
                        EDI_SHIP_DC = rowEDT850TS.Item("EDI_SHIP_DC")
                        CUST_STORE_NO = rowEDT850TS.Item("CUST_STORE_NO") & ""
                        CUST_DC_NO = rowEDT850TS.Item("CUST_DC_NO") & ""
                        Dim SALES_DIVISION_CODE As String = "" '?
                        Get_Ship_To(EDI_STORE, EDI_SHIP_DC)

                        If CUST_STORE_NO = "" Then
                            CUST_STORE_NO = CUST_DC_NO
                        End If

                        Dim ORDR_NO As String = ORDR_NO_next
                        ORDR_NO_end = ORDR_NO_next
                        ORDR_NO_next = Format(Val(ORDR_NO_next) + 1, "0000000000")
                        ORDR_LNO = 0
                        '    EDI_DTL_SEQ = 0
                        ' CURRENT_DTL_SEQ = 0
                        '  CURRENT_ASST = ""
                        If ASCMAIN1.CLIENT = "RGI" Then
                            ORDR_GROUP_NO = ORDR_NO_beg
                        End If

                        ASCMAIN1.Progress("-", ORDR_NO)

                        ' Write Order Header

                        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow
                        With rowSOTORDR1
                            ORDER_COUNTER = ORDER_COUNTER + 1
                            .Item("ORDR_NO") = ORDR_NO
                            .Item("ORDR_DATE") = ORDR_DATE
                            .Item("CUST_CODE") = CUST_CODE
                            .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                            .Item("CUST_STORE_NO") = CUST_STORE_NO
                            ' .Item("CUST_STORE_LOCATION") = rowARTCUST2.Item("CUST_STORE_LOCATION") & ""
                            '.Item("ORDR_FOB") = ORDR_FOB
                            .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                            .Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE
                            .Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
                            .Item("ORDR_ORIG_SHIP_DATE") = ORDR_SHIP_DATE
                            .Item("ORDR_ORIG_CANCEL_DATE") = ORDR_CANCEL_DATE
                            .Item("POST_CODE") = rowARTCUST1bt.Item("POST_CODE")
                            .Item("TERM_CODE") = TERM_CODE
                            .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
                            .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE")
                            .Item("WHSE_CODE") = WHSE_CODE
                            .Item("ORDR_TYPE_CODE") = "REG"
                            .Item("SHIP_VIA_CODE") = SHIP_VIA_CODE

                            .Item("ORDR_SHIP_INSTR") = rowARTCUST1.Item("CUST_SPECIAL_INST")

                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("LAST_DATE") = DATETIME_STAMP
                            .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date
                            .Item("ORDR_DEPT") = ORDR_DEPT
                            '  .Item("ORDR_SHIP_TO") = ORDR_SHIP_TO
                            .Item("ORDR_ADDR_TYPE_ST") = ORDR_SHIP_TO

                            'If rowARTCUST1.Item("CUST_BLOCK_SALES") & "" = "1" Then
                            '.Item("ORDR_STATUS") = "X"
                            'Else
                            .Item("ORDR_STATUS") = "O"
                            .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
                            .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
                            'End If
                            '  .Item("EDI_JRNL_NO") = EDI_JRNL_NO
                            .Item("ORDR_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
                            If .Item("ORDR_PRIORITY") & "" = "" Then
                                .Item("ORDR_PRIORITY") = ROWs("SOTPARM1").Item("SO_PARM_ORDR_PRIORITY")
                            End If

                            .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                            .Item("CUST_DC_NO") = CUST_DC_NO
                            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
                            .Item("ORDR_SOURCE") = "E"
                            '.Item("PRICE_CLASS_CODE") = rowARTCUST1.Item("PRICE_CLASS_CODE") & ""
                            '.Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE") & ""

                            .Item("CURR_CODE") = CURR_CODE
                            .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE

                            '.Item("CUST_DISC_PCT") = PRICE_BASE_DPCT
                            .Item("CUST_DC_NO") = CUST_DC_NO
                            '.Item("ITEM_BRAND_CODE") = ITEM_BRAND_CODE
                            'SALES_DIVISION_CODE = "AV"
                            '.Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                            .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO

                            .Item("CUST_FACTOR_IND") = rowARTCUST1.Item("CUST_FACTOR_IND")

                            FRT_TERMS = rowARTCUST1.Item("FRT_TERMS") & ""
                            .Item("FRT_TERMS") = FRT_TERMS
                            .Item("ORDR_EDI_810") = ORDR_EDI_810
                            .Item("ORDR_EDI_856") = ORDR_EDI_856
                            If Format(ORDR_LAST_ARRIVAL_DATE, "yyyyMMdd") <> "00010101" Then .Item("ORDR_LAST_ARRIVAL_DATE") = ORDR_LAST_ARRIVAL_DATE
                            If Format(ORDR_ARRIVAL_DATE, "yyyyMMdd") <> "00010101" Then .Item("ORDR_ARRIVAL_DATE") = ORDR_ARRIVAL_DATE
                            '.Item("ORDR_BILL_SHIP_TO") = rowARTCUST1.Item("CUST_BILL_SHIP_TO") & ""
                            .Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE")

                            If EDI_PO_TYPE = "BK" Then
                                .Item("ORDR_HOLD") = "1"
                                .Item("ORDR_HOLD_REASON") = "BLANKET ORDER"
                                .Item("ORDR_SHIP_INSTR") = "BLANKET ORDER - DO NOT SHIP"
                            End If
                            If ASCMAIN1.CLIENT = "RGI" Then
                                .Item("EDI_PO_TYPE") = EDI_PO_TYPE
                                If DROP_SHIP Then
                                    .Item("ECOM_CODE") = ECOM_CODE
                                    .Item("ORDR_TYPE_CODE") = "B2C"
                                ElseIf EDI_PO_TYPE = "NA" Then
                                    .Item("ECOM_CODE") = ""
                                    .Item("ORDR_TYPE_CODE") = "XFR"
                                    .Item("WHSE_CODE") = "MS"
                                    .Item("WHSE_CODE_TO") = WHSE_CODE
                                End If
                            End If
                        End With
                        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

                        STYLE_CODE = ""
                        COLOR_CODE = ""

                        'CURRENT_ASST = ""
                        ' ASST_ITEM_LNO = 0
                        '    EDI_DTL_SEQ = 0
                        ' CURRENT_DTL_SEQ = 0

                        Dim SIZE_CODE As String = String.Empty
                        Dim ORDR_QTY As Int64 = 0
                        Dim EDI_DETL_QTY As Int64 = 0
                        Dim RANGE_STYLE_CODE As String = String.Empty
                        Dim CURRENT_RANGE As String = ""
                        Dim RANGE_STYLE_LNO As Integer = 0
                        Dim RANGE_UPC As String = String.Empty
                        Dim RANGE_SKU As String = String.Empty
                        Dim RANGE_STYLE_PRICE As Decimal = 0
                        Dim RANGE_STYLE_PRICE_CURR As Decimal = 0
                        Dim RANGE_AS_REPLACEMENT As String = String.Empty
                        Dim RANGE_QTY As Int64 = 0
                        Dim RANGE_PRICE As Decimal = 0
                        Dim RNG_AST_FLG As String = String.Empty

                        For Each rowEDTSDQT1 As DataRow In dst.Tables("EDTSDQT1").Select("EDI_STORE = '" & EDI_STORE & "'", "EDI_DTL_SEQ")
                            Dim EDI_DTL_SEQ As Int32 = Val(rowEDTSDQT1.Item("EDI_DTL_SEQ") & "")
                            Dim STYLE_CODE As String = rowEDTSDQT1.Item("STYLE_CODE") & ""
                            Dim COLOR_CODE As String = rowEDTSDQT1.Item("COLOR_CODE") & ""
                            Dim EDI_UPC As String = rowEDTSDQT1.Item("EDI_UPC") & ""
                            Dim EDI_SKU As String = rowEDTSDQT1.Item("EDI_SKU") & ""
                            Dim EDI_STYLE As String = rowEDTSDQT1.Item("EDI_STYLE") & ""
                            Dim EDI_COLOR_CODE As String = rowEDTSDQT1.Item("EDI_COLOR_CODE") & ""
                            Dim EDI_COLOR_NAME As String = rowEDTSDQT1.Item("EDI_COLOR_NAME") & ""
                            Dim EDI_ITEM As String = rowEDTSDQT1.Item("EDI_ITEM") & ""
                            If EDI_SKU = "" And EDI_ITEM <> "" Then EDI_SKU = EDI_ITEM

                            Dim EDI_SIZE As String = rowEDTSDQT1.Item("EDI_SIZE_DESC") & ""

                            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
                            If rowSOTORDR1.Item("SALES_DIVISION_CODE") & "" = "" Then rowSOTORDR1.Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")

                            Dim rowEDTSDQT0 As DataRow = dst.Tables("EDTSDQT0").Rows.Find(New String() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ, STYLE_CODE, EDI_UPC})
                            'EDI_UPC = rowEDTSDQT1.Item("EDI_UPC") & ""
                            'EDI_SKU = rowEDTSDQT1.Item("EDI_SKU") & ""

                            ' ASST_STYLE_CODE = rowEDTSDQT1.Item("ASST_STYLE_CODE")
                            Dim QTY As Int32 = Val(rowEDTSDQT1.Item("QTY") & "")
                            Dim PRICE As Decimal = Val(rowEDTSDQT0.Item("PRICE") & "")

                            'If ASST_STYLE_CODE = "" Then
                            '    ASST_ITEM_LNO = 0
                            'End If

                            Dim EDI_SLN_SEQ As Int32 = Val(rowEDTSDQT1.Item("EDI_SLN_SEQ") & "")
                            RANGE_STYLE_CODE = rowEDTSDQT1.Item("RANGE_STYLE_CODE") & ""

                            If RANGE_STYLE_CODE = "" And EDI_SLN_SEQ = 0 Then
                                RANGE_STYLE_LNO = 0
                            Else
                                ORDR_LNO += 1
                                RANGE_STYLE_LNO = ORDR_LNO
                            End If

                            Dim F As Int32 = 1
                            Dim rowEDTPPKS1 As DataRow = dst.Tables("EDTPPKS1").Rows.Find(New String() {CUST_CODE, STYLE_CODE})
                            If rowEDTPPKS1 IsNot Nothing Then
                                F = Val(rowEDTPPKS1.Item("CARTON_PACK_QTY") & "")
                                If F = 0 Then
                                    F = Val(rowEDTPPKS1.Item("CARTON_PACK_QTY_STYLE") & "")
                                End If
                                If F = 0 Then
                                    F = 1
                                End If
                            End If

                            Dim SET_QTY As Int32 = 1 ' Val(rowEDTSDQT0.Item("SET_QTY") & "")

                            If ASCMAIN1.CLIENT = "RGI" Then
                                Dim TEMP_ECOM_CODE As String = ECOM_CODE

                                If Not DROP_SHIP Then
                                    If CUST_CODE = "031013" Then
                                        TEMP_ECOM_CODE = "WAYFAIR"
                                    Else
                                        'Throw New Exception("Need to review Set Qty for this Order Type")
                                    End If
                                End If
                                If TEMP_ECOM_CODE = "" Then
                                    SET_QTY = 1
                                Else
                                    Dim rowECTESTY1 As DataRow = dst.Tables("ECTESTY1").Rows.Find(New String() {STYLE_CODE, TEMP_ECOM_CODE})
                                    If rowECTESTY1 IsNot Nothing Then
                                        SET_QTY = Val(rowECTESTY1.Item("SET_QTY") & "")
                                        If SET_QTY = 0 Then
                                            SET_QTY = 1
                                        End If
                                    Else
                                        If DROP_SHIP Then
                                            Throw New Exception("Problem with Set Qty for Style " & STYLE_CODE & " for eCommerce Partner " & TEMP_ECOM_CODE)
                                        End If
                                    End If
                                End If
                            End If

                            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
                            With rowSOTORDR2
                                .Item("ORDR_NO") = ORDR_NO
                                ORDR_LNO += 1
                                .Item("ORDR_LNO") = ORDR_LNO
                                .Item("STYLE_CODE") = STYLE_CODE
                                .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                                .Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
                                .Item("COLOR_CODE") = COLOR_CODE
                                .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                                .Item("ORDR_UNIT_PRICE") = rowEDTSDQT0.Item("PRICE") / F
                                .Item("ORDR_QTY") = QTY * F
                                .Item("ORDR_QTY_OPEN") = QTY * F
                                .Item("ORDR_QTY_ORIG") = QTY * F
                                'If rowARTCUST1.Item("CUST_BLOCK_SALES") = 1 Then
                                '.Item("ORDR_STATUS") = "X"
                                'Else
                                .Item("ORDR_STATUS") = "O"
                                'End If
                                .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                                .Item("CUST_STYLE_CODE") = EDI_STYLE

                                If EDI_COLOR_CODE <> "" Then
                                    .Item("CUST_COLOR_CODE") = EDI_COLOR_CODE
                                ElseIf EDI_COLOR_NAME <> "" Then
                                    .Item("CUST_COLOR_CODE") = EDI_COLOR_NAME
                                End If

                                ' I THINK I MIGHT WANT TO CONSIDER PLACING CARTON_PACK_QTY - WHICH IS DETERMINED IN Get_Style, INTO EDTSDQT0
                                ' GOING WITH ICTSTYL1 FOR NOW TO GET PAST NYA ISSUE WHERE WALMART CASE COUNT IS NOT BEING CALCUATED PROPERLY BECAUSE CARTON_PACK_QTY IS NOT GETTING A VALUE IN SOTORDR2
                                .Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")

                                .Item("CUST_SIZE_CODE") = EDI_SIZE
                                .Item("CUST_UPC") = EDI_UPC
                                .Item("CUST_SKU") = EDI_SKU

                                .Item("STYLE_RETAIL") = rowICTSTYL1.Item("STYLE_RETAIL")
                                '.Item("SREP_CODE") = rowARTCUST2.Item("SREP_CODE")
                                '.Item("CUST_CODE") = CUST_CODE
                                '.Item("CUST_STORE_NO") = CUST_STORE_NO
                                '.Item("WHSE_CODE") = WHSE_CODE
                                .Item("ORDR_UNIT_PRICE_CURR") = rowEDTSDQT0.Item("EDI_PRICE_CURR") / F

                                If ASCMAIN1.CLIENT = "RGI" Then
                                    .Item("SET_QTY") = SET_QTY
                                End If

                                If ASCMAIN1.CLIENT = "NYA" Then
                                    If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then

                                        If Not TAC.TACMAIN1.CanadaCustomerList.Contains(CUST_CODE) Then
                                            'If CUST_CODE <> "LOBLAW" And CUST_CODE <> "SDM" Then
                                            MsgBox("Pricing in Foreign Currency was built around LOBLAWS rules = needs to be reviewed for " & CUST_CODE)
                                            Stop
                                        End If

                                        'If PRICE_BASIS = "L' THEN" Then
                                        Dim ORDR_UNIT_PRICE_CURR As Decimal = rowEDTSDQT0.Item("EDI_PRICE_CURR")
                                        '  ORDR_UNIT_PRICE_CURR = rowEDTSDQT0.Item("PRICE")
                                        .Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE_CURR
                                        .Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_CURR * CURR_EXCH_RATE
                                        'End If
                                    End If
                                End If

                                'EDI_ITEM = ""
                                'ASST_STYLE_CODE = ""
                                If EDI_SLN_SEQ > 0 Then
                                    .Item("EDI_SLN_SEQ") = EDI_SLN_SEQ
                                End If

                                If RANGE_STYLE_CODE <> "" Then
                                    Dim rowICTRSTY1 As DataRow = dst.Tables("ICTRSTY1").Rows.Find(New String() {CUST_CODE, RANGE_STYLE_CODE})
                                    RANGE_PRICE = Val(rowICTRSTY1.Item("RANGE_PRICE") & "")
                                    RANGE_QTY = Val(rowICTRSTY1.Item("RANGE_QTY") & "")
                                    If RANGE_QTY = 0 Then
                                        RANGE_STYLE_PRICE = 0
                                    Else
                                        RANGE_STYLE_PRICE = RANGE_PRICE / RANGE_QTY
                                    End If

                                    If RANGE_STYLE_CODE_seqs.ContainsKey(RANGE_STYLE_CODE) Then
                                        RANGE_STYLE_LNO = RANGE_STYLE_CODE_seqs(RANGE_STYLE_CODE)
                                        'If RANGE_STYLE_CODE_seqs(RANGE_STYLE_CODE) <> RANGE_STYLE_LNO Then
                                        '    Throw New Exception("Cannot have Range Style on more than 1 EDI DTL SEQ")
                                        'End If
                                    Else
                                        RANGE_STYLE_CODE_seqs.Add(RANGE_STYLE_CODE, RANGE_STYLE_LNO)

                                        Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").NewRow
                                        With rowSOTORDR9
                                            .Item("ORDR_NO") = ORDR_NO
                                            .Item("RANGE_STYLE_LNO") = RANGE_STYLE_LNO ' EDI_DTL_SEQ
                                            .Item("RANGE_STYLE_CODE") = RANGE_STYLE_CODE
                                            .Item("RANGE_STYLE_QTY") = RANGE_QTY ' dynWK.Fields("SLN_PARENT_STYLE_QTY") * RANGE_STYLE_QTY_PER_PP
                                            .Item("RANGE_STYLE_PRICE") = RANGE_STYLE_PRICE ' dynWK.Fields("SLN_PARENT_STYLE_PRICE") / RANGE_STYLE_QTY_PER_PP
                                            .Item("RANGE_STYLE_PRICE_CURR") = RANGE_STYLE_PRICE ' dynWK.Fields("SLN_PARENT_STYLE_PRICE_CURR") / RANGE_STYLE_QTY_PER_PP
                                            .Item("RANGE_INNER_PACK_QTY") = 0 ' rowICTRSTY1.Item("RANGE_STYLE_QTY_PER_PP") '  dynWK.Fields("SLN_PARENT_INNER_PACK_QTY") 'RANGE_INNER_PACK_QTY
                                            .Item("RANGE_STYLE_DESC") = rowICTRSTY1.Item("RANGE_STYLE_DESC") ' dynWK.Fields("SLN_PARENT_STYLE_DESC")       ' RANGE DESC
                                            .Item("RANGE_STYLE_UOM") = "EA"    ' UNIT OF MEASURE
                                            .Item("RANGE_STYLE_PP_PRICE") = RANGE_STYLE_PRICE ' dynWK.Fields("SLN_PARENT_STYLE_PRICE")
                                            .Item("RANGE_STYLE_PP_PRICE_CURR") = RANGE_STYLE_PRICE ' dynWK.Fields("SLN_PARENT_STYLE_PRICE_CURR")
                                            .Item("RANGE_STYLE_PP_QTY") = RANGE_QTY ' dynWK.Fields("SLN_PARENT_STYLE_QTY")
                                            .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                                            .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                            .Item("RANGE_STYLE_QTY_PER_PP") = rowICTRSTY1.Item("RANGE_STYLE_QTY_PER_PP")
                                        End With
                                        dst.Tables("SOTORDR9").Rows.Add(rowSOTORDR9)
                                    End If

                                    .Item("RANGE_STYLE_CODE") = RANGE_STYLE_CODE
                                    .Item("RANGE_STYLE_LNO") = RANGE_STYLE_LNO ' EDI_DTL_SEQ
                                End If
                            End With
                            'Reservations go here - update SOTORDR2, SOTRSRV1 & 2
                            If ASCMAIN1.CLIENT = "RGI" Then Dependent_Updates(rowSOTORDR2, ORDR_NO, ORDR_GROUP_NO, WHSE_CODE)

                            dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                            'If rowSOTORDR1.Item("COLLECTION_CODE") & "" = "" Then
                            '    rowSOTORDR1.Item("COLLECTION_CODE") = rowEDTSDQT1.Item("COLLECTION_CODE")
                            '    rowSOTORDR1.Item("BRAND_CODE") = rowEDTSDQT1.Item("BRAND_CODE")
                            'End If
                        Next
                    Next

                    rowEDT850T1.Item("EDI_PROCESS_IND") = "1"
                    rowEDT850T1.Item("LAST_DATE") = DATETIME_STAMP
                    rowEDT850T1.Item("LAST_OPER") = ASCMAIN1.USER_ID

                    ASCMAIN1.sql = "Update EDT850T1" _
                        & " Set EDI_PROCESS_IND = '1' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                    ASCDATA1.ExecuteSQL()

                    Update_Record_TDA("SOTORDR1")
                    Update_Record_TDA("SOTORDR2")
                    Update_Record_TDA("SOTORDR9")

                    rowEDT850T1.Item("RESULT") = "Imported"

                    rowEDT850T1.Item("ORDERS") = order_count
                    Dim AMOUNT = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")

                    rowEDT850T1.Item("AMOUNT") = AMOUNT
                    If dst.Tables("EDT850T1_IMPORTED").Rows.Find(EDI_DOC_SEQ_NO) IsNot Nothing Then
                        dst.Tables("EDT850T1_IMPORTED").Rows.Find(EDI_DOC_SEQ_NO).Delete()
                    End If
                    dst.Tables("EDT850T1_IMPORTED").Rows.Add(rowEDT850T1.ItemArray)

                    Update_ICTSTAT2(ORDR_GROUP_NO)
                    ASCDATA1.ExecuteSQL("Begin SOPORDR0_G('" & ORDR_GROUP_NO & "'); End;")

                    If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                        ' IF WE EVER DO MULTIPLE ORDERS IN A GROUP - WE WILL NEED TO CALL THIS FOR EACH ORDER
                        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", "ORDR_NO")
                            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                            ASCDATA1.ExecuteSP("SOPORDR1_COMM", "V", New Object() {ORDR_NO}, New String() {"ORDR_NO_IN"})
                        Next
                    End If

                    ASCMAIN1.sql = "Insert into SOTORDR5 (ORDR_NO,CUST_ADDR_TYPE,CUST_ADDR_CODE,CUST_NAME,CUST_ADDR1,CUST_ADDR2 " & vbCrLf _
                        & ",CUST_CITY,CUST_STATE,CUST_ZIP_CODE,CUST_COUNTRY,CUST_CONTACT,CUST_PHONE,CUST_EXT,CUST_FAX,CUST_EMAIL,CUST_ADDR3)" & vbCrLf

                    Dim COMMENTS As New List(Of String)

                    If rowEDTSLSP1.Item("EDI_CONSUMER") & "" = "1" Then

                        Dim JCPElabel As Boolean = False

                        If (CUST_CODE = "JCPE" And ASCMAIN1.CLIENT = "NYA") _
                          And rowEDT850T1.Item("EDI_SHIP_DC") & "" <> "" _
                          And dst.Tables("SOTORDR1").Rows.Count = 1 Then
                            JCPElabel = True
                        End If

                        If JCPElabel Then
                            Dim rows() As DataRow = dst.Tables("EDT850T5").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_ADR_SEQ = 1 and EDI_ADDR_TYPE = 'ST'")
                            If rows.Length = 1 Then
                                Dim EDI_ADDRESS2 As String = rows(0).Item("EDI_ADDRESS2") & ""
                                If EDI_ADDRESS2 <> "" Then
                                    COMMENTS.Add(Mid(EDI_ADDRESS2, 1, 30))
                                    COMMENTS.Add(Mid(EDI_ADDRESS2, 31, 12))
                                    COMMENTS.Add(Mid(EDI_ADDRESS2, 44))
                                End If
                            End If

                            ASCMAIN1.sql &= "" _
                                & "Select SOTORDR1.ORDR_NO, 'ST', '000000' CUST_ADDR_CODE" & vbCrLf _
                                & ", EDT850T5.EDI_ADDRESS1, EDT850T5.EDI_ADDRESS3, NULL" & vbCrLf _
                                & ", EDT850T5.EDI_CITY, EDT850T5.EDI_STATE, EDT850T5.EDI_ZIPCODE" & vbCrLf _
                                & ", EDT850T5.EDI_COUNTRY, NULL CUST_CONTACT, NULL CUST_PHONE" & vbCrLf _
                                & ", NULL CUST_EXT,NULL CUST_FAX, NULL CUST_EMAIL, NULL" & vbCrLf _
                                & " from EDT850T5, EDT850T1, SOTORDR1" & vbCrLf _
                                & " where EDT850T5.EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
                                & "   and EDT850T5.EDI_ADDR_TYPE = 'ST'" & vbCrLf _
                                & "   and EDT850T1.EDI_DOC_SEQ_NO = EDT850T5.EDI_DOC_SEQ_NO" & vbCrLf _
                                & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                                & "   and SOTORDR1.ORDR_STATUS = 'O'"
                        Else
                            'this is address from EDT850T5 let's get Contact info
                            Dim CUST_PHONE As String = getEDT850T9()

                            ASCMAIN1.sql &= "" _
                                & "Select SOTORDR1.ORDR_NO, 'ST', '000000' CUST_ADDR_CODE" & vbCrLf _
                                & ", EDT850T5.EDI_CUST_NAME_ADR, EDT850T5.EDI_ADDRESS1, EDT850T5.EDI_ADDRESS2" & vbCrLf _
                                & ", EDT850T5.EDI_CITY, EDT850T5.EDI_STATE, substr(EDT850T5.EDI_ZIPCODE,1,10) EDI_ZIPCODE" & vbCrLf _
                                & ", EDT850T5.EDI_COUNTRY, NULL CUST_CONTACT, '" & CUST_PHONE & "' CUST_PHONE" & vbCrLf _
                                & ", NULL CUST_EXT,NULL CUST_FAX, NULL CUST_EMAIL, EDT850T5.EDI_ADDRESS3" & vbCrLf _
                                & " from EDT850T5, SOTORDR1" & vbCrLf _
                                & " where EDT850T5.EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
                                & "   and EDT850T5.EDI_ADDR_TYPE = 'ST'" & vbCrLf _
                                & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                                & "   and SOTORDR1.ORDR_STATUS = 'O'"
                        End If

                    Else
                        ASCMAIN1.sql &= "" _
                        & " Select SOTORDR1.ORDR_NO, 'ST', ARTCUST2.CUST_ADDR_CODE" & vbCrLf _
                        & ", ARTCUST2.CUST_NAME, ARTCUST2.CUST_ADDR1, ARTCUST2.CUST_ADDR2" & vbCrLf _
                        & ", ARTCUST2.CUST_CITY, ARTCUST2.CUST_STATE, ARTCUST2.CUST_ZIP_CODE" & vbCrLf _
                        & ", ARTCUST2.CUST_COUNTRY, ARTCUST2.CUST_CONTACT, ARTCUST2.CUST_PHONE" & vbCrLf _
                        & ", ARTCUST2.CUST_EXT,ARTCUST2.CUST_FAX, ARTCUST2.CUST_EMAIL, ARTCUST2.CUST_ADDR3" & vbCrLf _
                        & "" & vbCrLf _
                        & " from ARTCUST2, SOTORDR1" & vbCrLf _
                        & " where ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                        & "   and ARTCUST2.CUST_ADDR_CODE = SOTORDR1.CUST_" & IIf(ORDR_SHIP_TO = "MK", "STORE", "DC") & "_NO" & vbCrLf _
                        & "   and ARTCUST2.CUST_ADDR_TYPE = '" & ORDR_SHIP_TO & "'" & vbCrLf _
                        & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                        & "   and SOTORDR1.ORDR_STATUS = 'O'"
                    End If
                    ASCDATA1.ExecuteSQL()

                    Dim ORDR_CLNO As Integer = 0
                    If ASCMAIN1.CLIENT = "NYA" AndAlso COMMENTS.Count > 0 Then
                        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows(0)
                        Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                        For Each COMMENT As String In COMMENTS
                            ORDR_CLNO += 1
                            Dim rowSOTORDR4 As DataRow = dst.Tables("SOTORDR4").NewRow
                            rowSOTORDR4.Item("ORDR_NO") = ORDR_NO
                            rowSOTORDR4.Item("ORDR_CLNO") = ORDR_CLNO
                            rowSOTORDR4.Item("ORDR_COMMENT") = COMMENT
                            dst.Tables("SOTORDR4").Rows.Add(rowSOTORDR4)
                        Next
                    End If
                    For Each rowEDT850T4 As DataRow In dst.Tables("EDT850T4").Select("")
                        ORDR_CLNO += 1
                        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
                            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                            Dim rowSOTORDR4 As DataRow = dst.Tables("SOTORDR4").NewRow
                            rowSOTORDR4.Item("ORDR_NO") = ORDR_NO
                            rowSOTORDR4.Item("ORDR_CLNO") = ORDR_CLNO
                            rowSOTORDR4.Item("ORDR_COMMENT") = rowEDT850T4.Item("EDI_CMMNT")
                            dst.Tables("SOTORDR4").Rows.Add(rowSOTORDR4)
                        Next
                    Next
                    Update_Record_TDA("SOTORDR4")

                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1.Item("CUST_FACTOR_IND") & "" = "1" Then
                            Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                            Dim ORDR_AMT_NOW As Decimal = Val(rowSOTORDR0.Item("ORDR_AMT") & "")
                            If ORDR_AMT_NOW <> 0 Then
                                TAC.SOCMAIN1.Credit_Request(rowARTCUST1.Item("TERM_CODE"), rowSOTORDR0)
                            End If
                        End If
                    End If

                    ' THIS SECTION DOES NOT PROPERLY CORRELATE SOTORDR2 TO SOTORDR3 THE WAY IT IS CURRENTLY CODED
                    'Dim ORDR_SUB_LNO As Integer = 0
                    'For Each rowEDT850T6 As DataRow In dst.Tables("EDT850T6").Select("")
                    '    ORDR_SUB_LNO += 1
                    '    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
                    '        Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                    '        Dim rowSOTORDR3 As DataRow = dst.Tables("SOTORDR3").NewRow
                    '        rowSOTORDR3.Item("ORDR_NO") = ORDR_NO
                    '        rowSOTORDR3.Item("ORDR_LNO") = ORDR_LNO
                    '        rowSOTORDR3.Item("ORDR_SUB_LNO") = ORDR_SUB_LNO
                    '        rowSOTORDR3.Item("CUST_STYLE_CODE") = rowEDT850T6.Item("EDI_SLN_STYLE")
                    '        rowSOTORDR3.Item("CUST_COLOR_CODE") = rowEDT850T6.Item("EDI_SLN_COLOR")
                    '        'rowSOTORDR3.Item("EDI_DTL_SEQ") = rowEDT850T6.Item("EDI_SLN_COLOR")
                    '        'rowSOTORDR3.Item("ORDR_UNIT_PRICE") = rowEDT850T6.Item("EDI_SLN_COLOR")
                    '        rowSOTORDR3.Item("ORDR_QTY") = rowEDT850T6.Item("EDI_SLN_QTY")
                    '        dst.Tables("SOTORDR3").Rows.Add(rowSOTORDR3)
                    '    Next
                    'Next

                    ASCMAIN1.sql = "" _
                        & " Update SOTORDR1 " & vbCrLf _
                        & " Set CUST_STORE_NAME = (Select CUST_NAME from SOTORDR5 " & vbCrLf _
                        & "  where ORDR_NO = SOTORDR1.ORDR_NO and CUST_ADDR_TYPE = 'ST')" & vbCrLf _
                        & " where SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"

                    ASCDATA1.ExecuteSQL()

                    If ASCMAIN1.CLIENT = "RGI" Then
                        If rowEDT850T1.Item("EDI_PREMARK_IND") & "" = "BULK" Then
                            ASCMAIN1.sql = "" _
                                & "Begin " & vbCrLf _
                                & "   SOPORDR0_BULK('" & ORDR_GROUP_NO & "','" & EDI_DOC_SEQ_NO & "'); " & vbCrLf _
                                & "End;"
                            ASCDATA1.ExecuteSQL()
                        End If
                    End If

                    CommitTrans()

                    ASCDATA1.DeleteRows(dst.Tables("EDT850TE"), "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                Else
                    rowEDT850T1.Item("RESULT") = "Rejected"
                End If
            End If
            ASCMAIN1.MultiTask_Release("", 0, 2)
        Next

        ORDERS_REJECTED = ORDERS_PROCESSED - ORDERS_IMPORTED
        MsgBox(CStr(ORDERS_PROCESSED) & " Order(s) Processed, " & vbCrLf & _
               CStr(ORDERS_IMPORTED) & " Order(s) Imported, " & vbCrLf & _
               CStr(ORDERS_REJECTED) & " Order(s) Rejected", MsgBoxStyle.OkOnly, "Processing Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        If ASCMAIN1.CLIENT = "RGI" Then
            Release_ECOM_Orders()
        End If

    End Sub
    Sub Dependent_Updates(ByRef rowSOTORDR2 As DataRow, ORDR_NO As String, ORDR_GROUP_NO As String, WHSE_CODE As String)

        Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
        Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
        Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")


        Dim rowSOTRSRVX As DataRow = Fill_Record("SOTRSRVX", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
        '& " order by SOTRSRV1.ORDR_CANCEL_DATE"

        Dim Ps() As Object

        If rowSOTRSRVX IsNot Nothing Then
            rowSOTORDR2.Item("RSRV_NO") = rowSOTRSRVX.Item("RSRV_NO")
            rowSOTORDR2.Item("RSRV_LNO") = rowSOTRSRVX.Item("RSRV_LNO")
            Ps = {rowSOTRSRVX.Item("RSRV_NO"), rowSOTRSRVX.Item("RSRV_LNO")}
            Update_SOTRSRVx(rowSOTORDR2, 1, ORDR_GROUP_NO)
        End If

    End Sub

    Sub Update_SOTRSRVx(rowSOTORDR2 As DataRow, S As Integer, ORDR_GROUP_NO As String)
        Dim RSRV_NO As String = rowSOTORDR2.Item("RSRV_NO") & ""
        Dim RSRV_LNO As Int64 = Val(rowSOTORDR2.Item("RSRV_LNO") & "")

        Dim rowSOTRSRV1 As DataRow = Fill_Record("SOTRSRV1", RSRV_NO)
        Dim WHSE_CODE As String = rowSOTRSRV1.Item("WHSE_CODE")

        Dim rowSOTRSRV2 As DataRow = Fill_Record("SOTRSRV2", New String() {RSRV_NO, RSRV_LNO})
        With rowSOTRSRV2
            Dim RSRV_QTY As Int64 = .Item("RSRV_QTY")
            Dim RSRV_QTY_OPEN As Int64 = Val(.Item("RSRV_QTY_OPEN") & "")
            Dim RSRV_QTY_CANC As Int64 = Val(.Item("RSRV_QTY_CANC") & "")
            Dim RSRV_QTY_USED As Int64 = Val(.Item("RSRV_QTY_USED") & "") _
                          + S * Val(rowSOTORDR2.Item("ORDR_QTY") & "")

            '  + S * Val(rowSOTORDR2.Item("ORDR_QTY_ORIG") & "") - USING ORDR_QTY_ORIG WILL ALWAYS HAVE 0 IMPACT WHEN CHANGING THE ORDER
            Dim RSRV_QTY_OPEN_OLD As Int64 = RSRV_QTY_OPEN
            RSRV_QTY_OPEN = RSRV_QTY - RSRV_QTY_CANC - RSRV_QTY_USED
            If RSRV_QTY_OPEN < 0 Then
                RSRV_QTY_OPEN = 0
            End If
            Dim RSRV_QTY_OPEN_NEW As Int64 = RSRV_QTY_OPEN
            .Item("RSRV_QTY_USED") = RSRV_QTY_USED
            .Item("RSRV_QTY_OPEN") = RSRV_QTY_OPEN

            Dim QTY_TO_COMMIT As Int64 = RSRV_QTY_OPEN_NEW - RSRV_QTY_OPEN_OLD
            If QTY_TO_COMMIT <> 0 Then
                Dim STYLE_CODE As String = .Item("STYLE_CODE")
                Dim COLOR_CODE As String = .Item("COLOR_CODE")
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY_TO_COMMIT)
            End If
        End With

        If S = -1 Then
        Else

            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim rowSOTORDR7 As DataRow = Fill_Record("SOTORDR7", New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})

            If rowSOTORDR7 Is Nothing Then
                rowSOTORDR7 = dst.Tables("SOTORDR7").NewRow
                rowSOTORDR7.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                rowSOTORDR7.Item("STYLE_CODE") = STYLE_CODE
                rowSOTORDR7.Item("COLOR_CODE") = COLOR_CODE
                dst.Tables("SOTORDR7").Rows.Add(rowSOTORDR7)
            End If
            If rowSOTRSRV2.Item("RSRV_PRIORITY_DATE") & "" = "" Then
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = CDate(rowSOTRSRV1.Item("INIT_DATE")).Date
            Else
                rowSOTORDR7.Item("ORDR_PRIORITY_DATE") = CDate(rowSOTRSRV2.Item("RSRV_PRIORITY_DATE")).Date
            End If
            rowSOTORDR7.Item("ORDR_PRIORITY") = rowSOTRSRV2.Item("RSRV_PRIORITY")
            Update_Record_TDA("SOTORDR7", "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
        End If
        Update_Record_TDA("SOTRSRV2")

        ASCMAIN1.sql = "Select Sum (RSRV_QTY_OPEN) from SOTRSRV2 where RSRV_NO = :PARM1"
        Dim RSRV_QTY_OPEN_total As Int64 = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {RSRV_NO}))

        If RSRV_QTY_OPEN_total = 0 Then
            rowSOTRSRV1.Item("RSRV_STATUS") = "F"
        Else
            rowSOTRSRV1.Item("RSRV_STATUS") = "O"
        End If
        Update_Record_TDA("SOTRSRV1")
    End Sub

    Function Bad_Data( _
        EDI_COND_DESC As String, _
        EDI_COND_CODE As String, _
        EDI_RECEIVED_VALUE As String, _
        Optional EDI_EXPECTED_VALUE As String = "", _
        Optional EDI_REFERENCE As String = "") As String

        Dim EDI_REPLACED_VALUE As String = ""
        EDI_ACTION = ""

        Dim EDI_ERROR As String = EDI_COND_DESC & vbTab & EDI_COND_CODE & vbTab & EDI_RECEIVED_VALUE
        If EDI_ERRORs.ContainsKey(EDI_ERROR) Then
            Dim EDI_ERROR_NO As Int32 = EDI_ERRORs(EDI_ERROR)
            Dim rowEDT850TE As DataRow = dst.Tables("EDT850TE").Rows.Find(New Object() {EDI_DOC_SEQ_NO, EDI_ERROR_NO})

            EDI_ACTION = rowEDT850TE.Item("EDI_ACTION") & ""
            Select Case EDI_ACTION
                Case "A" ' Use Received Value (ie, Accept EDI Data)
                    If ACTIONS_A.Contains("*") Or ACTIONS_A.Contains(EDI_COND_CODE) Then
                        EDI_REPLACED_VALUE = EDI_RECEIVED_VALUE
                    Else
                        EDI_ACTION = "N"
                    End If

                Case "E" ' Use Expected Value
                    If ACTIONS_E.Contains("*") Or ACTIONS_E.Contains(EDI_COND_CODE) Then
                        EDI_REPLACED_VALUE = EDI_EXPECTED_VALUE
                    Else
                        EDI_ACTION = "N"
                    End If

                Case "S" ' Skip Record
                    If ACTIONS_S.Contains("*") Or ACTIONS_S.Contains(EDI_COND_CODE) Then
                        EDI_REPLACED_VALUE = "Skip"
                    Else
                        EDI_ACTION = "N"
                    End If

                Case "R" ' Use Replacement Value
                    If ACTIONS_R.Contains("*") Or ACTIONS_R.Contains(EDI_COND_CODE) Then
                        EDI_REPLACED_VALUE = rowEDT850TE.Item("EDI_REPLACED_VALUE") & ""
                    Else
                        EDI_ACTION = "N"
                    End If
            End Select

            If EDI_REPLACED_VALUE = "" Then
                rowEDT850TE.Item("EDI_ORDER_COUNT") = Val(rowEDT850TE.Item("EDI_ORDER_COUNT") & "") + 1
            End If
        Else
            Dim rowEDT850TE As DataRow = dst.Tables("EDT850TE").NewRow
            Dim EDI_ERROR_NO As Int32 = EDI_ERRORs.Count + 1
            EDI_ERRORs.Add(EDI_ERROR, EDI_ERROR_NO)
            With rowEDT850TE
                .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                .Item("EDI_ERROR_NO") = EDI_ERROR_NO
                .Item("EDI_COND_DESC") = EDI_COND_DESC
                .Item("EDI_COND_CODE") = EDI_COND_CODE
                .Item("EDI_EXPECTED_VALUE") = EDI_EXPECTED_VALUE
                .Item("EDI_RECEIVED_VALUE") = EDI_RECEIVED_VALUE
                .Item("EDI_ACTION") = "N"
                .Item("EDI_ORDER_COUNT") = 1
                .Item("EDI_TP_QUAL") = EDI_TP_QUAL
                .Item("EDI_TP_ID") = EDI_TP_ID
                .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = CUST_STORE_NO
                .Item("EDI_REFERENCE") = EDI_REFERENCE
                If RESOLUTIONS.ContainsKey(EDI_COND_CODE) Then
                    .Item("RESOLUTION") = RESOLUTIONS(EDI_COND_CODE)
                End If
                If ASCMAIN1.CLIENT = "RGI" Then
                    If EDI_COND_CODE = "17" And Val(EDI_EXPECTED_VALUE) <> 0 Then
                        .Item("VARIANCE") = Format((1 - Val(EDI_RECEIVED_VALUE) / Val(EDI_EXPECTED_VALUE)), "##.0%")
                    End If
                End If
            End With
            dst.Tables("EDT850TE").Rows.Add(rowEDT850TE)
        End If

        If EDI_REPLACED_VALUE = "" Then
            EDI_DOC_SEQ_NO_ok = False
        End If

        Return EDI_REPLACED_VALUE
    End Function

    Function Get_EDTTRPM1() As DataRow
        ORDR_EDI_810 = "0"
        ORDR_EDI_856 = "0"
        CUST_EDI_DTS_FLAG = "0"

        rowEDTTRPM1 = dst.Tables("EDTTRPM1").Rows.Find(New String() {EDI_TP_QUAL, EDI_TP_ID, "850"})
        If rowEDTTRPM1 Is Nothing Then
            Bad_Data(EDI_COND_DESC:="Unrecognized Sender ID", EDI_COND_CODE:="01", EDI_RECEIVED_VALUE:=EDI_TP_QUAL & "-" & EDI_TP_ID)
        Else
            rowEDTSLSP1 = LookUp("EDTSLSP1", CUST_CODE)
            If rowEDTSLSP1 Is Nothing Then
                Bad_Data(EDI_COND_DESC:="Unrecognized Sender ID", EDI_COND_CODE:="01", EDI_RECEIVED_VALUE:=EDI_TP_QUAL & "-" & EDI_TP_ID)
            Else
                ORDR_EDI_810 = IIf(rowEDTSLSP1.Item("EDI_ID_810") & "" = "", "0", "1")
                ORDR_EDI_856 = IIf(rowEDTSLSP1.Item("EDI_ID_856") & "" = "", "0", "1")
                CUST_EDI_DTS_FLAG = rowEDTSLSP1.Item("EDI_DTS_IND") & ""
            End If
        End If

        Return rowEDTTRPM1
    End Function

    Sub Get_ARTCUST1()

        ' WHAT ABOUT CUST_STATUS?

        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
        If rowARTCUST1 Is Nothing Then
            Bad_Data(EDI_COND_DESC:="Customer Record Missing or Not Active", EDI_COND_CODE:="02", EDI_RECEIVED_VALUE:=CUST_CODE)
        Else
            If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                Bad_Data(EDI_COND_DESC:="Customer PO is Required, but Missing", EDI_COND_CODE:="04", EDI_RECEIVED_VALUE:=ORDR_CUST_PO)
            End If
            If rowARTCUST1.Item("FRT_TERMS") & "" = "" Then
                Bad_Data(EDI_COND_DESC:="Customer Missing Freight Terms", EDI_COND_CODE:="35", EDI_RECEIVED_VALUE:="")
            End If
            If rowARTCUST1.Item("CUST_ROUTING_INST") & "" = "" Then
                Bad_Data(EDI_COND_DESC:="Customer Missing Routing Instructions", EDI_COND_CODE:="36", EDI_RECEIVED_VALUE:="")
            End If

            If rowEDT850T1.Item("WHSE_CODE") & "" <> "" Then
                WHSE_CODE = rowEDT850T1.Item("WHSE_CODE")
            Else
                If rowARTCUST1.Item("WHSE_CODE") & "" = "" Then
                    WHSE_CODE = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""
                Else
                    WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
                End If
            End If

            CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
            If CUST_BILL_TO_CUST = "" Then
                CUST_BILL_TO_CUST = CUST_CODE
            End If

            rowARTCUST1bt = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
            If rowARTCUST1bt Is Nothing Then
                Bad_Data(EDI_COND_DESC:="Bill-To Customer Missing or not Active", EDI_COND_CODE:="03", EDI_RECEIVED_VALUE:=CUST_BILL_TO_CUST)
            End If

            'rowSOTPCLS1 = dst.Tables("SOTPCLS1").Rows.Find(rowARTCUST1.Item("PRICE_CLASS_CODE") & "")
            'If rowSOTPCLS1 Is Nothing Then
            '    Bad_Data(EDI_COND_DESC:="Missing or Invalid Price Class Code for Customer " & CUST_CODE, EDI_COND_CODE:="91", EDI_RECEIVED_VALUE:=rowARTCUST1.Item("PRICE_CLASS_CODE") & "")
            'Else
            '    PRICE_BASIS = rowSOTPCLS1.Item("PRICE_BASIS") & ""
            '    PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
            'End If


            PRICE_LIST_CODE = rowARTCUST1.Item("PRICE_LIST_CODE") & ""

            CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
            If CURR_CODE = "" Then CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")

            If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                CURR_EXCH_RATE = 1
            Else
                'CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, CURR_CODE, Now.Date)
                CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me.ROWs("GLTPARM1"), CURR_CODE, Now.Date)

            End If


        End If
    End Sub

    Function Get_Ship_To(EDI_STORE As String, EDI_SHIP_DC As String) As String

        Dim skip_store As Boolean = False
        Dim Bad_Data_Cond_05 As Boolean = False

        If EDI_STORE <> "" Then
            CUST_DC_NO = ""
            If Len(EDI_STORE) < 6 And IsNumeric(EDI_STORE) Then
                CUST_STORE_NO = Format(Val(EDI_STORE), "000000")
            Else
                If Len(EDI_STORE) > 6 Then
                    ASCMAIN1.sql = "Select * from ARTCUST2 where GLOBAL_LOCATION_NUMBER = :PARM1"
                    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {EDI_STORE})
                    If row Is Nothing Then
                        CUST_STORE_NO = ""
                    Else
                        CUST_STORE_NO = row.Item("CUST_ADDR_CODE")
                        CUST_DC_NO = row.Item("CUST_DC_NO") & ""
                    End If
                Else
                    CUST_STORE_NO = EDI_STORE
                    Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_STORE_NO})
                    If rowARTCUST2 IsNot Nothing Then
                        ' SHOULD WE BE USING ARTCUST3?
                        CUST_DC_NO = rowARTCUST2.Item("CUST_DC_NO") & ""
                    End If
                End If
            End If
        Else
            If Len(EDI_SHIP_DC) > 6 Then
                ASCMAIN1.sql = "Select * from ARTCUST2 where GLOBAL_LOCATION_NUMBER = :PARM1"
                Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {EDI_SHIP_DC})
                If row Is Nothing Then
                    CUST_DC_NO = ""
                Else
                    CUST_DC_NO = row.Item("CUST_ADDR_CODE")
                End If
            Else
                If IsNumeric(EDI_SHIP_DC) Then
                    CUST_DC_NO = Format(Val(EDI_SHIP_DC), "000000")
                Else
                    CUST_DC_NO = EDI_SHIP_DC
                End If
            End If

            '    CUST_DC_NO = Format(Val(EDI_SHIP_DC), "000000")
            If CUST_DC_NO <> "" Then
                CUST_STORE_NO = CUST_DC_NO
            End If
        End If

        If EDI_SHIP_DC <> "" Then
            If CUST_DC_NO = "" Then
                If Len(EDI_SHIP_DC) > 6 Then
                    ASCMAIN1.sql = "Select * from ARTCUST2 where GLOBAL_LOCATION_NUMBER = :PARM1"
                    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {EDI_SHIP_DC})
                    If row Is Nothing Then
                        CUST_DC_NO = ""
                    Else
                        CUST_DC_NO = row.Item("CUST_ADDR_CODE")
                    End If
                Else
                    CUST_DC_NO = Format(Val(EDI_SHIP_DC), "000000")
                End If
                ' CUST_DC_NO = Format(Val(EDI_SHIP_DC), "000000")
            End If
        Else

        End If
        'CUST_STORE_NAME = ""


        rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_STORE_NO})
        'If rowARTCUST2 Is Nothing Then
        '    rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "DC", CUST_STORE_NO})
        'End If

        ' S.B USING DCGLN NOT EDI_STORE_NO BELOW

        If rowARTCUST2 Is Nothing AndAlso EDI_STORE <> "" Then
            Dim row() As DataRow = dst.Tables("ARTCUST2").Select("GLOBAL_LOCATION_NUMBER = '" & EDI_STORE & "'")
            If row.Length > 0 Then
                rowARTCUST2 = row(0)
            End If
            'Dim row() As DataRow = dst.Tables("EDT850T5").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_ADDR_TYPE = 'ST'")
            'If row.Length = 0 Then
            '    row = dst.Tables("EDT850T5").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_ADDR_TYPE = 'BY'")
            'End If
            'If row.Length > 0 Then
            '    Dim EDI_ADDR_CODE As String = row(0).Item("EDI_ADDR_CODE")
            '    row = dst.Tables("ARTCUST2").Select("GLOBAL_LOCATION_NO = '" & EDI_STORE & "'")

            'End If
        End If

        If rowARTCUST2 Is Nothing Then
            Bad_Data(EDI_COND_DESC:="Invalid Customer Store", EDI_COND_CODE:="05", EDI_RECEIVED_VALUE:=EDI_STORE)
            If EDI_ACTION = "S" Then
                skip_store = True
            Else
                Bad_Data_Cond_05 = True
            End If
        End If

        Dim rowSOTSVIA2 As DataRow

        If Not Bad_Data_Cond_05 And Not skip_store Then

            rowSOTSVIA2 = dst.Tables("SOTSVIA2").Rows.Find _
                (New String() {WHSE_CODE, rowARTCUST2.Item("CUST_STATE") & ""})

            If rowSOTSVIA2 Is Nothing Then
                TRANSIT_BUS_DAYS = 5
            Else
                TRANSIT_BUS_DAYS = Val(rowSOTSVIA2.Item("TRANSIT_BUS_DAYS") & "")
            End If

            ' THE LATTER HALF OF THE NEXT LINE NEEDS WORK BECAUSE AHA.ARTCUST2 <> NYA.ARTCUST2
            If CUST_EDI_DTS_FLAG <> "1" And CUST_DC_NO = "" AndAlso rowARTCUST2.Item("CUST_DC_NO") & "" = "" Then ' And rowARTCUST2.Item("CUST_DC_IND") & "" <> "1") Then
                ' Bad_Data(EDI_COND_DESC:="Store " & CUST_STORE_NO & " is Missing DC Code", EDI_COND_CODE:="37", EDI_RECEIVED_VALUE:="")
            End If

            CUST_STORE_NO = rowARTCUST2.Item("CUST_ADDR_CODE")
            'CUST_STORE_NAME = rowARTCUST2.Item("CUST_STORE_NAME") & ""
            ' THIS SECTION NEEDS REWORK - NEED TO SET UP A DATATABLE LINKING DC'S WITH STORES
            'If rowARTCUST2.Item("CUST_DC_IND") & "" = "1" Then
            '    CUST_DC_NO = rowARTCUST2.Item("CUST_STORE_NO") & ""
            'Else
            '    If IsNumeric(rowARTCUST2.Item("CUST_DC_NO")) Then
            '        CUST_DC_NO = Format(Val(rowARTCUST2.Item("CUST_DC_NO") & ""), "000000")
            '    Else
            '        CUST_DC_NO = rowARTCUST2.Item("CUST_DC_NO") & ""
            '    End If
            'End If

            If EDI_SHIP_DC <> "" Or CUST_EDI_DTS_FLAG <> "1" Then
                If EDI_SHIP_DC = "" Then
                    EDI_SHIP_DC = CUST_DC_NO
                End If
                If IsNumeric(EDI_SHIP_DC) And Len(EDI_SHIP_DC) < 6 Then
                    CUST_DC_NO = EDI_SHIP_DC.PadLeft(6, "0")
                Else
                    If Len(EDI_SHIP_DC) > 6 Then
                        ASCMAIN1.sql = "Select * from ARTCUST2 where GLOBAL_LOCATION_NUMBER = :PARM1"
                        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {EDI_SHIP_DC})
                        If row Is Nothing Then
                            CUST_DC_NO = ""
                        Else
                            CUST_DC_NO = row.Item("CUST_ADDR_CODE")
                        End If
                    End If
                    'CUST_DC_NO = EDI_SHIP_DC
                End If
                ORDR_SHIP_TO = "DC"
            Else
                ORDR_SHIP_TO = "MK"
            End If

            If CUST_DC_NO <> "" And ORDR_SHIP_TO = "DC" Then

                rowARTCUST2_DC = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "DC", CUST_DC_NO})
                If rowARTCUST2_DC Is Nothing Then
                    rowARTCUST2_DC = LookUp("ARTCUST2", New String() {CUST_CODE, "DC", CUST_DC_NO})
                End If
                ' DON'T DO THIS OR ELSE YOU WILL GET NO SOTORDR5 ST RECORD
                If rowARTCUST2_DC Is Nothing Then
                    Dim rowARTCUST2_DC_as_MK = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, "MK", CUST_DC_NO})

                    rowARTCUST2_DC = dst.Tables("ARTCUST2").NewRow
                    rowARTCUST2_DC.ItemArray = rowARTCUST2_DC_as_MK.ItemArray
                    rowARTCUST2_DC.Item("CUST_ADDR_TYPE") = "DC"
                    dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2_DC)
                    rowARTCUST2.Item("CUST_DC_NO") = CUST_DC_NO
                    Update_Record_TDA("ARTCUST2")

                    If dst.Tables("ARTCUST3").Rows.Find(New String() {CUST_CODE, "MK", CUST_STORE_NO, "DC"}) Is Nothing Then
                        Dim rowARTCUST3 As DataRow = dst.Tables("ARTCUST3").NewRow
                        rowARTCUST3.Item("CUST_CODE") = CUST_CODE
                        rowARTCUST3.Item("CUST_ADDR_TYPE") = "MK"
                        rowARTCUST3.Item("CUST_ADDR_CODE") = CUST_STORE_NO
                        rowARTCUST3.Item("CUST_ADDR_TYPE2") = "DC"
                        rowARTCUST3.Item("CUST_ADDR_CODE2") = CUST_STORE_NO
                        dst.Tables("ARTCUST3").Rows.Add(rowARTCUST3)
                    End If

                    Update_Record_TDA("ARTCUST3")
                End If

                If rowARTCUST2_DC Is Nothing Then
                    Bad_Data(EDI_COND_DESC:="DC not on file", EDI_COND_CODE:="07", EDI_RECEIVED_VALUE:=CUST_DC_NO)
                    CUST_DC_NO = "000000"
                Else
                    CUST_DC_NO = rowARTCUST2_DC.Item("CUST_ADDR_CODE") & ""
                    ORDR_SHIP_TO = "DC"
                    rowSOTSVIA2 = dst.Tables("SOTSVIA2").Rows.Find _
                       (New String() {WHSE_CODE, rowARTCUST2_DC.Item("CUST_STATE") & ""})

                    If rowSOTSVIA2 Is Nothing Then
                        TRANSIT_BUS_DAYS = 5
                    Else
                        TRANSIT_BUS_DAYS = Val(rowSOTSVIA2.Item("TRANSIT_BUS_DAYS") & "")
                    End If
                End If
            End If

            If CUST_DC_NO = "" And ORDR_SHIP_TO = "DC" Then
                Bad_Data(EDI_COND_DESC:="Store " & CUST_STORE_NO & " not linked to DC", EDI_COND_CODE:="33", EDI_RECEIVED_VALUE:=ORDR_SHIP_TO)
            End If
        End If

        If TRANSIT_BUS_DAYS = 0 Then
            TRANSIT_BUS_DAYS = 5
        End If

        Return CUST_STORE_NO

    End Function

    Function Get_Terms() As String

        TERM_CODE = ""
        If rowARTCUST1bt Is Nothing Then Return TERM_CODE

        If rowARTCUST1bt.Item("TERM_CODE") & "" = "" Then
            Bad_Data(EDI_COND_DESC:="Bill-To Customer Missing Terms", EDI_COND_CODE:="21", EDI_RECEIVED_VALUE:="")
        Else
            TERM_CODE = rowARTCUST1bt.Item("TERM_CODE") & ""
        End If

        Dim EDI_TERMS_key As String = "" _
            & rowEDT850T1.Item("EDI_TERMS") _
            & rowEDT850T1.Item("EDI_TERM_TYPE") _
            & rowEDT850T1.Item("EDI_TERM_BASIS") _
            & rowEDT850T1.Item("EDI_TERM_RATE") _
            & rowEDT850T1.Item("EDI_TERM_DSCDAYS") _
            & rowEDT850T1.Item("EDI_TERM_NETDAYS") _
            & rowEDT850T1.Item("EDI_TERM_DESC") _
            & rowEDT850T1.Item("EDI_TERM_DOM")

        If EDI_TERMS_key <> "" Then
            Dim EDI_TERM_CODE As String = ""

            Dim rowEDTTERM1 As DataRow
            For Each row As DataRow In dst.Tables("EDTTERM1").Select("")
                If EDI_TERMS_key = "" _
                & row.Item("EDI_TERMS") _
                & row.Item("EDI_TERM_TYPE") _
                & row.Item("EDI_TERM_BASIS") _
                & row.Item("EDI_TERM_RATE") _
                & row.Item("EDI_TERM_DSCDAYS") _
                & row.Item("EDI_TERM_NETDAYS") _
                & row.Item("EDI_TERM_DESC") _
                & row.Item("EDI_TERM_DOM") Then
                    EDI_TERM_CODE = row.Item("EDI_TERM_CODE")
                    TERM_CODE = row.Item("TERM_CODE") & ""
                    Exit For
                End If
            Next

            If TERM_CODE = "" Then
                rowEDTTERM1 = dst.Tables("EDTTERM1").NewRow
                With rowEDTTERM1
                    EDI_TERM_CODE = ASCMAIN1.Next_Control_No("EDTTERM1.EDI_TERM_CODE")
                    .Item("EDI_TERM_CODE") = EDI_TERM_CODE
                    .Item("EDI_TERMS") = rowEDT850T1.Item("EDI_TERMS")
                    .Item("EDI_TERM_TYPE") = rowEDT850T1.Item("EDI_TERM_TYPE")
                    .Item("EDI_TERM_BASIS") = rowEDT850T1.Item("EDI_TERM_BASIS")
                    .Item("EDI_TERM_RATE") = rowEDT850T1.Item("EDI_TERM_RATE")
                    .Item("EDI_TERM_DSCDAYS") = rowEDT850T1.Item("EDI_TERM_DSCDAYS")
                    .Item("EDI_TERM_NETDAYS") = rowEDT850T1.Item("EDI_TERM_NETDAYS")
                    .Item("EDI_TERM_DESC") = rowEDT850T1.Item("EDI_TERM_DESC")
                    .Item("EDI_TERM_DOM") = rowEDT850T1.Item("EDI_TERM_DOM")
                End With
                dst.Tables("EDTTERM1").Rows.Add(rowEDTTERM1)
            End If
            If TERM_CODE = "" Then
                Bad_Data(EDI_COND_DESC:="EDI Terms File", EDI_COND_CODE:="11", EDI_RECEIVED_VALUE:=EDI_TERM_CODE)
            Else
                Dim rowTATTERM1 As DataRow = dst.Tables("TATTERM1").Rows.Find(TERM_CODE)
                If rowTATTERM1 Is Nothing Then
                    Bad_Data(EDI_COND_DESC:="AR Terms File", EDI_COND_CODE:="12", EDI_RECEIVED_VALUE:=TERM_CODE)
                Else
                    If TERM_CODE <> rowARTCUST1bt.Item("TERM_CODE") & "" Then
                        TERM_CODE = Bad_Data(EDI_COND_DESC:="EDI Terms do not match Bill-To Customer", EDI_COND_CODE:="13", EDI_RECEIVED_VALUE:=TERM_CODE, EDI_EXPECTED_VALUE:=rowARTCUST1bt.Item("TERM_CODE") & "")
                    End If
                End If
            End If
        End If

        Return TERM_CODE
    End Function

    Sub Check_for_Possible_Order_Duplication(Optional CUST_STORE_NO As String = "")
        If ORDR_CUST_PO <> "" Then
            ASCMAIN1.sql = "Select MAX (ORDR_NO) from SOTORDR1 where CUST_CODE = :PARM1 and ORDR_CUST_PO = :PARM2" _
                & " and ORDR_STATUS in ('O','P','C','F')"
            Dim ORDR_NO As String
            If CUST_STORE_NO = "" Then
                ORDR_NO = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {CUST_CODE, ORDR_CUST_PO})
            Else
                ASCMAIN1.sql &= " and CUST_STORE_NO = :PARM3"
                ORDR_NO = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VVV", New String() {CUST_CODE, ORDR_CUST_PO, CUST_STORE_NO})
            End If
            If ORDR_NO <> "" Then
                Bad_Data(EDI_COND_DESC:="Possible Order Duplication", EDI_COND_CODE:="14", EDI_RECEIVED_VALUE:=ORDR_CUST_PO)
            End If
        End If
    End Sub

    Sub Process_EDT850T2(EDI_STORE As String, EDI_SHIP_DC As String)

        ' Loop thru 850T2 - each record will represent a Single Item/Style/UPC, etc
        Dim EDI_SHIP_DC_passed_in As String = EDI_SHIP_DC

        For Each rowEDT850T2 In dst.Tables("EDT850T2").Select("")
            Dim skip_item As Boolean = False

            EDI_SHIP_DC = EDI_SHIP_DC_passed_in
            If EDI_SHIP_DC = "" And rowEDT850T2.Item("EDI_LN_SHIP_DC") & "" <> "" Then
                EDI_SHIP_DC = rowEDT850T2.Item("EDI_LN_SHIP_DC")
            End If

            Dim EDI_DETL_QTY As Int64 = 0
            Dim EDI_PRICE As Decimal = 0

            Dim RANGE_STYLE_CODE As String = String.Empty
            Dim CURRENT_RANGE As String = String.Empty
            Dim RANGE_STYLE_LNO As Integer = 0
            Dim RANGE_UPC As String = String.Empty
            Dim RANGE_SKU As String = String.Empty
            Dim RANGE_STYLE_PRICE As Decimal = 0
            Dim RANGE_STYLE_PRICE_CURR As Decimal = 0
            Dim RANGE_AS_REPLACEMENT As String = String.Empty
            Dim RANGE_QTY As Int64 = 0
            Dim RANGE_PRICE As Decimal = 0
            Dim RNG_AST_FLG As String = String.Empty

            Dim T2_PO4_QTY As Int32 = 0
            Dim T6_PO4_QTY As Int32 = 0
            Dim PO4_OW As Int32 ' Dim PO4_QTY As Long

            Dim EDI_TOTAL_QTY As Int32 = rowEDT850T2.Item("EDI_TOTAL_QTY")

            Dim EDI_SLN_SEQ As Int32 = 0
            Dim EDI_SKU As String = rowEDT850T2.Item("EDI_SKU") & ""
            Dim EDI_ITEM As String = rowEDT850T2.Item("EDI_ITEM") & ""
            Dim EDI_UPC As String = rowEDT850T2.Item("EDI_UPC") & ""
            Dim EDI_EAN As String = rowEDT850T2.Item("EDI_EAN") & ""
            If EDI_UPC = "" And EDI_EAN <> "" Then EDI_UPC = EDI_EAN ' PROBABLY NEED TO ADD SUPPORT THROUGH ALL OF THE ARGUMENT LISTS
            Dim EDI_STYLE As String = rowEDT850T2.Item("EDI_STYLE") & ""
            Dim EDI_COLOR_CODE As String = rowEDT850T2.Item("EDI_COLOR_CODE") & ""
            Dim EDI_COLOR_NAME As String = rowEDT850T2.Item("EDI_COLOR_NAME") & ""
            Dim EDI_SIZE_DESC As String = rowEDT850T2.Item("EDI_SIZE_DESC") & ""

            Dim EDI_PO4_UOM As String = rowEDT850T2.Item("EDI_PO4_UOM") & ""
            Dim EDI_PRICE_UOM As String = rowEDT850T2.Item("EDI_PRICE_UOM") & ""
            Dim EDI_PO4_QTY As Integer = Val(rowEDT850T2.Item("EDI_PO4_QTY") & "")

            If ASCMAIN1.CLIENT = "RGI" Then
                EDI_STYLE = (rowEDT850T2.Item("EDI_STYLE") & "").Trim
                EDI_COLOR_CODE = (rowEDT850T2.Item("EDI_COLOR_CODE") & "").Trim
            End If

            Dim PO1_is_Range_Style As Boolean = False
            Dim rowICTRSTY1 As DataRow = Nothing
            For Each CODE_VALUE As String In New String() {EDI_STYLE, EDI_SKU, EDI_ITEM, EDI_UPC, EDI_EAN}
                If CODE_VALUE <> "" Then
                    rowICTRSTY1 = dst.Tables("ICTRSTY1").Rows.Find(New String() {CUST_CODE, CODE_VALUE})
                    If rowICTRSTY1 IsNot Nothing Then
                        PO1_is_Range_Style = True
                        RANGE_STYLE_CODE = CODE_VALUE
                        RANGE_PRICE = Val(rowICTRSTY1.Item("RANGE_PRICE") & "")
                        RANGE_QTY = Val(rowICTRSTY1.Item("RANGE_QTY") & "")
                        If RANGE_QTY = 0 Then
                            RANGE_STYLE_PRICE = 0
                        Else
                            RANGE_STYLE_PRICE = RANGE_PRICE / RANGE_QTY
                        End If

                        If RANGE_QTY <> EDI_TOTAL_QTY * EDI_PO4_QTY Then
                            Bad_Data( _
                            EDI_COND_DESC:="Range Qty (" & CStr(RANGE_QTY) & ") does not match Order (" & CStr(EDI_TOTAL_QTY) & ")", _
                            EDI_COND_CODE:="16", _
                            EDI_RECEIVED_VALUE:=CStr(EDI_TOTAL_QTY))
                        End If

                        Exit For
                    End If
                End If
            Next

            Dim QTY As Int64 = 0
            Dim WORK_QTY As Int64 = 0
            Dim PRICE As Decimal = 0
            Dim PRICE_CURR As Decimal = 0

            Dim HAVE_SDQ As Boolean = False
            Dim HAVE_SLN As Boolean = False
            Dim FROM_T2 As Boolean = False

            Dim CUST_PRICE As Decimal = 0
            Dim CUST_PRICE_CURR As Decimal = 0

            Dim SLN_PARENT_STYLE_CODE As String
            Dim SLN_PARENT_STYLE_QTY As Long
            Dim SLN_PARENT_STYLE_PRICE As Decimal
            Dim SLN_PARENT_STYLE_PRICE_CURR As Decimal
            Dim SLN_PARENT_INNER_PACK_QTY As Long

            Dim RANGE_STYLE_QTY_PER_PP As Int32 = 0
            Dim QTY_PER_PP As Int32 = 0
            Dim ASSORTMENT_PRICE_FROM_COMPONENTS As Decimal = 0

            Dim SLN_FLAG As Boolean = False  'NEW VARIABLE FOR WRITING SLN_PARENT_STYLE_CODE TO ORDR2
            Dim SLN_PARENT_STYLE_DESC As String = ""

            Dim use_case_pack As Boolean = False

            If EDI_PO4_UOM = "AS" Or EDI_PO4_UOM = "CS" Or EDI_PO4_UOM = "CA" _
            Or EDI_PRICE_UOM = "AS" Or EDI_PRICE_UOM = "CS" Or EDI_PRICE_UOM = "CA" Then
                T2_PO4_QTY = EDI_PO4_QTY
                use_case_pack = True
            Else
                If EDI_PO4_UOM = "EA" Then
                    T2_PO4_QTY = 1
                ElseIf EDI_PO4_QTY = 0 Then
                    T2_PO4_QTY = 1
                Else
                    T2_PO4_QTY = EDI_PO4_QTY
                End If
            End If

            'If CUST_CODE = "MARSHAL" And rowEDT850T2.Item("EDI_PO4_UOM") & "" = "CA" _
            '    And rowEDT850T2.Item("EDI_PO4_QTY") & "" = "" Then
            '    MsgBox("Call ABS")
            '    T2_PO4_QTY = 0
            'End If

            'If CUST_CODE = "CHARLOT" And (rowEDT850T2.Item("EDI_PO4_UOM") & "" = "CA" _
            '    Or rowEDT850T2.Item("EDI_PO4_UOM") & "" = "AS") Then
            '    MsgBox("Call ABS")
            '    T2_PO4_QTY = 1
            'End If

            EDI_DTL_SEQ = Val(rowEDT850T2.Item("EDI_DTL_SEQ") & "")

            ' needed for walmart 0000004950 where price was INCORRECTLY being taken from SLN and not from PO1
            'EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "")

            Dim rowEDT850T6s() As DataRow = dst.Tables("EDT850T6").Select("EDI_DTL_SEQ = " & CStr(EDI_DTL_SEQ))

            ' LESLIE WANTS TO BRING IN WALMARTS NOT USING THE PPK SKU - BUT i DON'T KNOW WHEN TO START THE NEW WAY
            ' TO DO THE NEW WAY, ELIMINATE THE WALMART HARD CODE
            ' If rowEDT850T6s.Length <> 0 Then

            'If rowEDT850T6s.Length <> 0 And (CUST_CODE <> "CKOUT" And CUST_CODE <> "PVH" And CUST_CODE <> "WALMART" And CUST_CODE <> "CASUALMALE" And CUST_CODE <> "MAURICES" And CUST_CODE <> "KOHLS") Then
            If rowEDT850T6s.Length <> 0 And rowEDTSLSP1.Item("EDI_IGNORE_SLN") & "" <> "1" Then

                If PO1_is_Range_Style Then
                    Throw New Exception("Range Style with SLN")
                    Stop
                End If

                HAVE_SLN = True

                If CUST_CODE = "WALMART" Then ' THIS PROBABLY NEEDS TO BE OPENED UP TO ALL
                    SLN_PARENT_STYLE_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "")
                End If

                For Each rowEDT850T6 As DataRow In rowEDT850T6s

                    With rowEDT850T6
                        If .Item("EDI_SLN_PO4_UOM") & "" = "EA" Then
                            T6_PO4_QTY = 1
                        Else
                            If .Item("EDI_PO4_INNER") & "" = "" Then
                                T6_PO4_QTY = Val(.Item("EDI_PO4_QTY") & "")
                            Else
                                T6_PO4_QTY = Val(.Item("EDI_PO4_QTY")) * Val(.Item("EDI_PO4_INNER"))
                            End If
                        End If

                        If CUST_CODE = "CHARMING" Or CUST_CODE = "LANEBRY" Or CUST_CODE = "CATHERINE" Then
                            T6_PO4_QTY = 0
                        Else
                            If T6_PO4_QTY = 0 Then T6_PO4_QTY = 1
                        End If

                        EDI_SLN_SEQ = Val(.Item("EDI_SLN_SEQ") & "")
                        EDI_UPC = .Item("EDI_SLN_UPC") & ""
                        EDI_SKU = .Item("EDI_SLN_SKU") & ""
                        EDI_ITEM = .Item("EDI_SLN_ITEM") & ""
                        EDI_STYLE = .Item("EDI_SLN_STYLE") & ""
                        EDI_COLOR_CODE = .Item("EDI_SLN_COLOR") & ""
                        EDI_COLOR_NAME = .Item("EDI_SLN_COLOR") & ""
                        EDI_SIZE_DESC = .Item("EDI_SLN_SIZE_DESC") & ""

                        If ASCMAIN1.CLIENT = "RGI" Then
                            EDI_STYLE = (.Item("EDI_SLN_STYLE") & "").Trim
                            EDI_COLOR_CODE = (.Item("EDI_SLN_COLOR") & "").Trim
                        End If
                        'Dim EDI_SIZE_DESC As String = ""
                        'If .Item("EDI_SLN_ITEM") & "" <> "" Then
                        '    EDI_SIZE_DESC = .Item("EDI_SLN_SIZE_DESC") & ""
                        'Else
                        '    EDI_SIZE_DESC = rowEDT850T2.Item("EDI_SIZE_DESC") & ""
                        'End If

                        If rowEDTSLSP1.Item("EDI_SLN_TOT_IND") & "" = "1" Then
                            EDI_DETL_QTY = .Item("EDI_SLN_QTY")
                            EDI_PRICE = Val(.Item("EDI_SLN_PRICE") & "")
                        Else
                            If CUST_CODE = "CHARMING - not necessary since we set the flag above" Then

                                EDI_DETL_QTY = T6_PO4_QTY ' .Item("EDI_SLN_QTY")
                            Else

                                ' MAYBE THIS IS CODED WRONG - MAYBE IT SHOULD BE ONE OR THE OTHER AND NOT BOTH MULTIPLIED - AND THEN WE WOULD NOT NEED SPECIAL CODING FOR CHARMING AS IN 0000002899

                                EDI_DETL_QTY = .Item("EDI_SLN_QTY") * T6_PO4_QTY
                            End If


                            If T6_PO4_QTY = 0 Then
                                EDI_PRICE = 0
                            Else
                                ' CASUALMALE PUTS THE RETAIL PRICE INTO THE EDI_SLN_PRICE FIELD, SO WE SHOULD TREAT CASUALMALE AS IF THEY DID NOT SEND PRICE IN THE SLN
                                If Val(.Item("EDI_SLN_PRICE") & "") = 0 Or CUST_CODE = "CASUALMALE" Then
                                    ' WALMART 0000004950
                                    Dim SLN_QTY As Integer = Val(dst.Tables("EDT850T6").Compute("SUM(EDI_SLN_QTY)", "EDI_DTL_SEQ = " & CStr(EDI_DTL_SEQ)) & "")
                                    EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "") / SLN_QTY

                                    If CUST_CODE = "MAURICES" Or CUST_CODE = "FAMDOLTAR" Then
                                        EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "")
                                    End If
                                Else
                                    EDI_PRICE = Val(.Item("EDI_SLN_PRICE") & "") / T6_PO4_QTY
                                End If

                            End If

                        End If

                        If CUST_CODE = "CHARMING" Or CUST_CODE = "LANEBRY" Or CUST_CODE = "CATHERINE" Then
                            EDI_DETL_QTY = EDI_DETL_QTY * EDI_TOTAL_QTY
                        End If
                        If CUST_CODE = "CHARLOTTE" Then
                            EDI_DETL_QTY = EDI_DETL_QTY * EDI_TOTAL_QTY
                        End If
                        If CUST_CODE = "CASUALMALE" Then
                            EDI_DETL_QTY = EDI_DETL_QTY * EDI_TOTAL_QTY
                        End If

                        '            If SLN_DIVIDEDBY_PO1_TIMES_SDQ = "1" Then
                        '                EDI_DETL_QTY = dynEDT850T6.Fields("EDI_SLN_QTY").Value / dynEDT850T2.Fields("EDI_TOTAL_QTY").Value
                        '            Else
                        '                EDI_DETL_QTY = dynEDT850T6.Fields("EDI_SLN_QTY").Value
                        '            End If
                        '            EDI_PRICE = Val(dynEDT850T6.Fields("EDI_SLN_PRICE").Value & "")  

                        SLN_PARENT_STYLE_CODE = rowEDT850T2.Item("EDI_STYLE") & ""
                        SLN_PARENT_STYLE_DESC = rowEDT850T2.Item("EDI_STYLE") & ""

                        If rowEDT850T2.Item("EDI_SKU") & "" <> "" Then
                            SLN_PARENT_STYLE_CODE = rowEDT850T2.Item("EDI_SKU") & ""
                        ElseIf rowEDT850T2.Item("EDI_UPC") & "" <> "" Then
                            SLN_PARENT_STYLE_CODE = rowEDT850T2.Item("EDI_UPC") & ""
                        Else
                            SLN_PARENT_STYLE_CODE = rowEDT850T2.Item("EDI_STYLE") & ""
                        End If

                        'If CUST_CODE = "BURLING" And Mid$(SLN_PARENT_STYLE_CODE, 1, 2) = "PO" Then
                        '    SLN_PARENT_STYLE_CODE = Mid$(SLN_PARENT_STYLE_CODE, 3, 7) & Format$(Val(Mid$(SLN_PARENT_STYLE_CODE, 12)), "000")
                        'End If
                        'If CUST_CODE = "JCPENNEY" And Right$(SLN_PARENT_STYLE_CODE, 5) = "*ONLY" Then
                        '    SLN_PARENT_STYLE_CODE = Left$(SLN_PARENT_STYLE_CODE, 8)
                        'End If

                        PO4_OW = T6_PO4_QTY ' PO4_QTY = T6_PO4_QTY
                        ' PO4_OverWrite(EDI_UPC, EDI_SKU, EDI_ITEM, PO4_OW, T2_PO4_QTY, T6_PO4_QTY)

                        FROM_T2 = False
                        RANGE_STYLE_QTY_PER_PP = RANGE_STYLE_QTY_PER_PP + (EDI_DETL_QTY * PO4_OW)

                        If CUST_CODE = "CHARMING" Or CUST_CODE = "LANEBRY" Or CUST_CODE = "CATHERINE" Or CUST_CODE = "CASUALMALE" Or CUST_CODE = "CHARLOTTE" Or CUST_CODE = "FAMDOLTAR" Then
                            QTY_PER_PP = PO4_OW
                        Else
                            QTY_PER_PP = EDI_DETL_QTY * PO4_OW
                        End If
                        'If CUST_CODE = "CHARLOTTE" Then
                        '    QTY_PER_PP = PO4_OW
                        'Else
                        '    QTY_PER_PP = EDI_DETL_QTY * PO4_OW
                        'End If

                        ' CHARMING EDI_DOC_SEQ_NO = '0000003010' NEEDED THE NEXT LINE TO GET THE PROPER QTY

                        If rowEDT850T6.Item("EDI_SHIP_DC") & "" <> "" Then
                            EDI_STORE = rowEDT850T6.Item("EDI_SHIP_DC")
                            EDI_SHIP_DC = rowEDT850T6.Item("EDI_SHIP_DC")
                        End If
                        Process_Items_Pre(EDI_UPC, EDI_SKU, EDI_ITEM, EDI_STYLE, EDI_DETL_QTY, EDI_STORE, EDI_SHIP_DC, EDI_PRICE, _
                                    EDI_COLOR_CODE, EDI_COLOR_NAME, _
                                    EDI_SIZE_DESC, _
                                    RNG_AST_FLG, _
                                    SLN_PARENT_STYLE_CODE, _
                                    SLN_PARENT_STYLE_DESC, _
                                    SLN_PARENT_STYLE_QTY, _
                                    SLN_PARENT_INNER_PACK_QTY, _
                                    EDI_SLN_SEQ, _
                                    RANGE_STYLE_QTY_PER_PP, _
                                    QTY_PER_PP, _
                                    WORK_QTY, _
                                    RANGE_STYLE_CODE,
                                    RANGE_STYLE_PRICE,
                                    RANGE_UPC, _
                                    RANGE_SKU, _
                                    T2_PO4_QTY, _
                                    T6_PO4_QTY)

                    End With
                Next

                Dim CALC_SLN_PRICES As String = ""
                If ASSORTMENT_PRICE_FROM_COMPONENTS = 0 Then
                    CALC_SLN_PRICES = "1"
                End If
                If ASSORTMENT_PRICE_FROM_COMPONENTS <> 0 Or CALC_SLN_PRICES <> "1" Then
                    If System.Math.Round(SLN_PARENT_STYLE_PRICE, 6) <> System.Math.Round(ASSORTMENT_PRICE_FROM_COMPONENTS, 6) Then
                        Bad_Data(EDI_COND_DESC:="Assortment Price " & Format$(SLN_PARENT_STYLE_PRICE, "####0.00") & " does not match Component Prices " & Format$(ASSORTMENT_PRICE_FROM_COMPONENTS, "####0.00") & " for " & rowEDT850T2.Item("EDI_SKU"), EDI_COND_CODE:="36", EDI_RECEIVED_VALUE:=Format$(SLN_PARENT_STYLE_PRICE, "####0.00"))
                    End If
                End If

            Else

                HAVE_SLN = False

                EDI_DETL_QTY = rowEDT850T2.Item("EDI_TOTAL_QTY")
                EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "")

                If RANGE_STYLE_CODE <> "" And CUST_CODE = "FREDMEYER" And T2_PO4_QTY <> 0 And T2_PO4_QTY <> 1 Then
                    ' this is the 1st occurrence of a range style order where the range style is a pre-pack
                    ' this code might be the way to handle this - but limiting the code to just FM until we get another one
                    EDI_PRICE = EDI_PRICE / T2_PO4_QTY
                End If

                'EDI_PRICE_CURR = dynEDT850T2.Fields("EDI_PRICE").Value
                SLN_PARENT_STYLE_CODE = ""
                SLN_PARENT_STYLE_QTY = 0
                SLN_PARENT_STYLE_PRICE = 0
                SLN_PARENT_STYLE_PRICE_CURR = 0
                SLN_PARENT_INNER_PACK_QTY = 0

                ' this next section was entirely remmed out
                '  - unremmed to handle case packs
                ' - RE-REMMED BECAUSE CASE PACK CONVERSIONS ARE HANDLED IN GET_ICTSTYL1
                'PO4_OW = T2_PO4_QTY
                'PO4_OverWrite(EDI_UPC, EDI_SKU, EDI_ITEM, PO4_OW, T2_PO4_QTY, T6_PO4_QTY)
                'If rowEDT850T2.Item("EDI_PO4_UOM") & "" = "EA" Then
                '    'If rowEDTSLSP1.Item("EDI_PO1_TOT_IND") & "" = "1" And rowEDT850T2.Item("EDI_PO4_UOM") & "" = "EA" Then
                '    EDI_DETL_QTY = rowEDT850T2.Item("EDI_TOTAL_QTY")
                '    EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "")
                'Else
                '    ' Notice bellow is checking T2_PO4_QTY, should it check PO1_TOT_IND? or T2_PO4_QTY <> 1?
                '    'If T2_PO4_QTY = 1 And (rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "CA" Or rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "CS") Then
                '    '    EDI_PRICE = rowEDT850T2.Item("EDI_PRICE") / T2_PO4_QTY
                '    'Else
                '    If T2_PO4_QTY = 0 Then
                '        EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "")
                '    Else
                '        EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "") / T2_PO4_QTY
                '    End If
                '    'End If
                '    If dst.Tables("EDT850T3").Select("EDI_DTL_SEQ = " & CStr(EDI_DTL_SEQ)).Length > 0 Then
                '        If T2_PO4_QTY <> 0 Then
                '            EDI_DETL_QTY = T2_PO4_QTY
                '        Else
                '            EDI_DETL_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "")
                '        End If
                '    Else
                '        If T2_PO4_QTY <> 0 Then
                '            EDI_DETL_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "") * T2_PO4_QTY
                '        Else
                '            EDI_DETL_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "")
                '        End If
                '    End If
                'End If

                PO4_OW = T2_PO4_QTY ' PO4_QTY = T2_PO4_QTY
                FROM_T2 = True
                Process_Items_Pre(EDI_UPC, EDI_SKU, EDI_ITEM, EDI_STYLE, EDI_DETL_QTY, EDI_STORE, EDI_SHIP_DC, EDI_PRICE, _
                                    EDI_COLOR_CODE, EDI_COLOR_NAME, _
                                    EDI_SIZE_DESC, _
                                    RNG_AST_FLG, _
                                    SLN_PARENT_STYLE_CODE, _
                                    SLN_PARENT_STYLE_DESC, _
                                    SLN_PARENT_STYLE_QTY, _
                                    SLN_PARENT_INNER_PACK_QTY, _
                                    EDI_SLN_SEQ, _
                                    RANGE_STYLE_QTY_PER_PP, _
                                    QTY_PER_PP, _
                                    WORK_QTY, _
                                    RANGE_STYLE_CODE,
                                    RANGE_STYLE_PRICE,
                                    RANGE_UPC, _
                                    RANGE_SKU, _
                                    T2_PO4_QTY, _
                                    T6_PO4_QTY)
            End If
        Next
    End Sub

    Sub Process_Items(
        EDI_UPC As String,
        EDI_SKU As String,
        EDI_ITEM As String,
        EDI_STYLE As String,
        EDI_COLOR_CODE As String,
        EDI_PRICE As Decimal,
        RANGE_STYLE_CODE As String,
        RANGE_STYLE_PRICE As Decimal,
        ByRef T2_PO4_QTY As Int32,
        ByRef EDI_DETL_QTY As Int32,
        ByRef T6_PO4_QTY As Int32,
        rowEDT850T3 As DataRow)

        ITEM_INACTIVE = False


        ' NYA FREDMEYER IS THE ONLY CUSTOMER WE HAVE HAD TO DO THIS FOR SEE EDI_DOC_SEQ_NO = '0000046887'
        Dim QTY_UOM As String = rowEDT850T2.Item("EDI_PRICE_UOM") & ""
        If QTY_UOM = "" Then QTY_UOM = rowEDT850T2.Item("EDI_PO4_UOM") & ""
        If rowEDT850T3 IsNot Nothing Then
            QTY_UOM = rowEDT850T3.Item("EDI_SDQ_UOM") & ""
        End If

        Dim PRICE As Decimal = 0
        Dim CUST_PRICE As Decimal = 0
        'Dim EDI_PRICE As Decimal = 0
        Dim EDI_PRICE_CURR As Decimal = 0
        Dim ITEM_PRICE_CURR As Decimal = 0
        Dim RETAIL_PRICE As Decimal
        Dim PRICE_DISC As Decimal

        Dim sqlw As String = String.Format("CUST_CODE = '{0}' and EDI_UPC = '{1}' and EDI_SKU = '{2}' and EDI_ITEM = '{3}' and EDI_STYLE = '{4}'", CUST_CODE, EDI_UPC, EDI_SKU, EDI_ITEM, EDI_STYLE)
        Dim rowEDTITEMXs() As DataRow = dst.Tables("EDTITEMX").Select(sqlw)
        If rowEDTITEMXs.Length <> 0 Then
            With rowEDTITEMXs(0)
                CUST_PRICE = Val(.Item("CUST_PRICE") & "")
                STYLE_CODE = .Item("STYLE_CODE")
                Dim SALES_DIVISION_CODE As String = .Item("SALES_DIVISION_CODE") & ""
                rowICTSTYL1 = dst.Tables("ICTSTYL1").Rows.Find(.Item("STYLE_CODE"))
                If rowICTSTYL1 Is Nothing Then
                    Bad_Data(EDI_COND_DESC:="Item Not in Work Table",
                              EDI_COND_CODE:="18",
                              EDI_RECEIVED_VALUE:=STYLE_CODE)
                End If
                saleable_item = True
                'If rowICTSTYL1.Item("ITEM_SNU_CODE") & "" = "S" Then
                '    saleable_item = True
                'Else
                '    saleable_item = False
                'End If
                If ASCMAIN1.CLIENT = "RGI" Then 'Color is set in Get_ICTSTYL1 need it here if style has been previously processed
                    If Not String.IsNullOrEmpty(EDI_COLOR_CODE) Then
                        COLOR_CODE = EDI_COLOR_CODE
                    End If

                    ' BUT WE HAD TO ENABLE AN ITEM ON AN ORDER MULTIPLE TIMES FOR RGI/NORDRACK BECAUSE ON A BULK ORDER, AN ITEM MAY APPEAR WITH SDQS W/OVERRIDING DC
                    ' AND THE RESULT WAS THAT THE EDI_PRICE_CURR WAS SET TO 0
                    ' SO WE ARE GOING TO AT LEAST SUPPORT THAT FIELD HERE
                    If CURR_CODE = "USD" Then
                        EDI_PRICE_CURR = EDI_PRICE
                    Else
                        Throw New Exception("No Current Support for Duplicate Item on EDI document in non-USD currency")
                    End If

                End If
            End With
        Else
            If RANGE_STYLE_CODE = "" Then
                Get_ICTSTYL1(EDI_UPC, EDI_SKU, EDI_ITEM, EDI_STYLE, EDI_COLOR_CODE,
                             PRICE, CUST_PRICE, EDI_PRICE, EDI_PRICE_CURR, ITEM_PRICE_CURR, RETAIL_PRICE, PRICE_DISC,
                             T2_PO4_QTY, EDI_DETL_QTY, T6_PO4_QTY, rowEDT850T3)
            Else

            End If
            If STYLE_CODE = "" And RANGE_STYLE_CODE = "" Then
                ' MODECRAFT EDI DOC SEQ 0000009745 NEEDED THE NEXT LINE REMMED OUT B/C THE STYLE WAS DETERMINED IN THE GET_ICTSTYL1 CALL A FEW LINES DOWN
                EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="EAN/UPC Code not found",
                                            EDI_COND_CODE:="15",
                                            EDI_RECEIVED_VALUE:=EDI_UPC, EDI_REFERENCE:=Mid$(":", 1, System.Math.Sign(Len(EDI_UPC))) & EDI_SKU & Mid$(":", 1, System.Math.Sign(Len(EDI_UPC & EDI_UPC))) & EDI_ITEM)
                If EDI_ACTION = "S" Then
                    ITEM_INACTIVE = True
                    Return
                Else
                    EDI_UPC = EDI_REPLACED_VALUE
                    Get_ICTSTYL1(EDI_UPC, EDI_SKU, EDI_ITEM, EDI_STYLE, EDI_COLOR_CODE,
                         PRICE, CUST_PRICE, EDI_PRICE, EDI_PRICE_CURR, ITEM_PRICE_CURR, RETAIL_PRICE, PRICE_DISC,
                         T2_PO4_QTY, EDI_DETL_QTY, T6_PO4_QTY, rowEDT850T3)
                End If
            Else
                'If ASST_STYLE_CODE = "" Then
                '    If rowICTSTYL1.Item("ITEM_STATUS") & "" <> "A" Then
                '        For i = 1 To UBound(INACTIVE_ITEMS)
                '            If INACTIVE_ITEMS(i) = STYLE_CODE Then
                '                ITEM_INACTIVE = True
                '            End If
                '        Next i
                '        If ITEM_INACTIVE = False Then
                '            MsgBox("Item" & STYLE_CODE & " will be ignored", vbOKOnly, "Inactive Item")
                '            i = UBound(INACTIVE_ITEMS)
                '            ReDim Preserve INACTIVE_ITEMS(i + 1)
                '            INACTIVE_ITEMS(i + 1) = STYLE_CODE
                '            ITEM_INACTIVE = True
                '            Return
                '        Else
                '            Return
                '        End If
                '    End If
                'End If
            End If
        End If

        Dim rowEDTSDQT0 As DataRow = dst.Tables("EDTSDQT0").Rows.Find(New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ, STYLE_CODE, EDI_UPC})
        If rowEDTSDQT0 Is Nothing Then
            PRICE = CUST_PRICE
            If RANGE_STYLE_CODE <> "" Then
                PRICE = RANGE_STYLE_PRICE
            End If
            If EDI_PRICE < 0.01 Or (PRICE = 0 And EDI_PRICE = 0.01) Then
                EDI_PRICE = PRICE
                EDI_PRICE_CURR = ITEM_PRICE_CURR
            End If
            If (ASCMAIN1.CLIENT = "RGI" And (EDI_PO_TYPE = "NA" Or DROP_SHIP = True)) Or (ASCMAIN1.CLIENT = "NYA" AndAlso CUST_CODE = "LOBLAW") Then
                'Xfer order no pricing check - also DropShip for RGI has separate E-comm logic
            Else

                If EDI_PRICE <> PRICE And System.Math.Abs(EDI_PRICE - PRICE) > 0.01 Then ' And PRICE <> 0 Then
                    EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Item " & STYLE_CODE & "" & " Price s/b " & Format(PRICE, ".00"),
                                 EDI_COND_CODE:="17",
                                 EDI_RECEIVED_VALUE:=Format(EDI_PRICE, ".00"),
                                 EDI_EXPECTED_VALUE:=Format(PRICE, ".00"),
                                 EDI_REFERENCE:=EDI_SKU)
                    If EDI_ACTION = "S" Then
                        skip_item = True
                        Return
                    End If
                    If EDI_REPLACED_VALUE <> "" Then
                        If PRICE <> Val(EDI_REPLACED_VALUE) Then
                            PRICE = Val(EDI_REPLACED_VALUE)
                            RETAIL_PRICE = PRICE / ((100 - PRICE_DISC) / 100)
                            ITEM_PRICE_CURR = PRICE / CURR_EXCH_RATE
                        End If
                    End If
                Else
                    If saleable_item And PRICE = 0 Then
                        EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Item " & STYLE_CODE & " is Saleable with a Price of $0",
                               EDI_COND_CODE:="38",
                               EDI_RECEIVED_VALUE:=Format(PRICE, ".00"), EDI_REFERENCE:=EDI_UPC)
                        If EDI_ACTION = "S" Then
                            skip_item = True
                            Return
                        End If
                    End If

                    PRICE = EDI_PRICE
                    ITEM_PRICE_CURR = EDI_PRICE_CURR
                End If

            End If
            If ASCMAIN1.CLIENT = "NYA" Then
                If TAC.TACMAIN1.CanadaCustomerList.Contains(CUST_CODE) Then ' CUST_CODE = "LOBLAW" Or CUST_CODE = "SDM" Then

                    'Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
                    'Dim C As Integer = 1
                    'If rowICTSTYL1 IsNot Nothing Then
                    '    C = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
                    'End If

                    Dim rowSOTPRIC2 As DataRow = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE, STYLE_CODE})
                    If rowSOTPRIC2 Is Nothing Then
                        ' NO PRICE RECORD - note - there is a already a place where we issue error 39 if the customer has a price class using basis code L (price list)
                        ' but for cad - doing this check no matter what

                        PRICE = -1
                        CUST_PRICE = -1
                        Bad_Data(EDI_COND_DESC:="No Price List record for Style " & STYLE_CODE, _
                                 EDI_COND_CODE:="39", _
                                 EDI_RECEIVED_VALUE:=STYLE_CODE)

                    ElseIf System.Math.Abs(Val(rowSOTPRIC2.Item("STYLE_PRICE") & "") - EDI_PRICE_CURR) > 0.01 Then
                        'ElseIf Val(rowSOTPRIC2.Item("STYLE_PRICE") & "") <> EDI_PRICE_CURR Then
                        ' INVALID PRICE RECORD

                        PRICE = Val(rowSOTPRIC2.Item("STYLE_PRICE") & "")
                        CUST_PRICE = Val(rowSOTPRIC2.Item("STYLE_PRICE") & "")

                        EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Style " & STYLE_CODE & "" & " Price s/b " & Format(PRICE, ".00"),
                                                     EDI_COND_CODE:="17",
                                                     EDI_RECEIVED_VALUE:=Format(EDI_PRICE_CURR, ".00"),
                                                     EDI_EXPECTED_VALUE:=Format(PRICE, ".00"),
                                                     EDI_REFERENCE:=EDI_SKU)
                        If EDI_ACTION = "S" Then
                            skip_item = True
                            Return
                        End If
                        If EDI_REPLACED_VALUE <> "" Then
                            If PRICE <> Val(EDI_REPLACED_VALUE) Then
                                PRICE = Val(EDI_REPLACED_VALUE)
                                '      RETAIL_PRICE = PRICE / ((100 - PRICE_DISC) / 100)
                                '      ITEM_PRICE_CURR = PRICE / CURR_EXCH_RATE
                            End If
                        End If

                    End If
                End If
            End If
            If ASCMAIN1.CLIENT = "RGI" Then
                Dim TEMP_ECOM_CODE As String = ECOM_CODE
                If CUST_CODE = "031013" And EDI_PO_TYPE = "NA" Then
                    'XFER orders have no ECOM_CODE but we still need to verify price record, not price
                    TEMP_ECOM_CODE = "WAYFAIR"
                End If
                If TEMP_ECOM_CODE <> "" Then
                    Dim rowECTESTY1 As DataRow = dst.Tables("ECTESTY1").Rows.Find(New String() {STYLE_CODE, TEMP_ECOM_CODE})

                    If rowECTESTY1 Is Nothing OrElse rowECTESTY1.Item("ECOM_STYLE_STATUS") & "" <> "A" Then
                        EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="No Active Price List record for Style " & STYLE_CODE,
                                 EDI_COND_CODE:="49",
                                 EDI_RECEIVED_VALUE:=Format(EDI_PRICE, ".00"),
                                 EDI_REFERENCE:=EDI_SKU)
                        If EDI_ACTION = "S" Then
                            skip_item = True
                            Return
                        End If

                    Else 'If EDI_PO_TYPE = "DS" Then
                        'check price for ECOM orders 
                        Dim ORDR_DATE_string As String = Format(ORDR_DATE, "MM/dd/yyyy")
                        Dim PriceType As String = ""

                        Dim sqlw3 As String = "STYLE_CODE = '" & STYLE_CODE & "' and ECOM_CODE = '" & ECOM_CODE & "'" _
                            & " and PROMO_START_DATE <= '" & ORDR_DATE_string & "' and PROMO_END_DATE >= '" & ORDR_DATE_string & "'"
                        Dim rowECTESTY3s() As DataRow = dst.Tables("ECTESTY3").Select(sqlw3)

                        Dim ECOM_PRICE As Decimal = 0

                        ' if length>1 then error message that there are multiple promos that apply - should not have been allowed
                        If ALT_PRICING Then
                            ECOM_PRICE = Val(rowECTESTY1.Item("ALT_UNIT_PRICE") & "")
                            PriceType = " Alternate"
                        ElseIf rowECTESTY3s.Length = 1 Then
                            ECOM_PRICE = Val(rowECTESTY3s(0).Item("PROMO_UNIT_PRICE") & "")
                            PriceType = " Promo"
                        Else
                            ECOM_PRICE = Val(rowECTESTY1.Item("ECOM_UNIT_PRICE") & "")
                            PriceType = ""
                        End If

                        Dim SET_QTY As Integer = Val(rowECTESTY1.Item("SET_QTY") & "")
                        If SET_QTY = 0 Then SET_QTY = 1
                        ' If SET_QTY <> 1 Then Stop

                        If SET_QTY <> 1 Then
                            EDI_DETL_QTY = EDI_DETL_QTY * SET_QTY
                            ' put something in sotordr2
                        End If

                        If DROP_SHIP Then ' Order type = 'OS' for alt pricing fields, also show alt pricing err msg
                            If Math.Abs((1 - EDI_PRICE / (ECOM_PRICE * SET_QTY)) * 100) <= ECOM_PRICE_TOLERANCE_PCT Or System.Math.Abs(EDI_PRICE - PRICE) <= 0.01 Then
                                'Price is within tolerance, use EDI_PRICE
                                PRICE = EDI_PRICE / SET_QTY
                                CUST_PRICE = EDI_PRICE / SET_QTY
                                EDI_PRICE_CURR = EDI_PRICE / SET_QTY
                            Else
                                EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Style " & STYLE_CODE & PriceType & " Price s/b " & Format(ECOM_PRICE * SET_QTY, ".00"),
                                                             EDI_COND_CODE:="17",
                                                             EDI_RECEIVED_VALUE:=Format(EDI_PRICE, ".00"),
                                                             EDI_EXPECTED_VALUE:=Format(ECOM_PRICE * SET_QTY, ".00"),
                                                             EDI_REFERENCE:=EDI_SKU)
                                If EDI_ACTION = "S" Then
                                    skip_item = True
                                    Return
                                Else
                                    If EDI_REPLACED_VALUE <> "" Then
                                        PRICE = Val(EDI_REPLACED_VALUE) / SET_QTY
                                        CUST_PRICE = EDI_PRICE / SET_QTY
                                        EDI_PRICE_CURR = Val(EDI_REPLACED_VALUE) / SET_QTY
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If RANGE_STYLE_CODE = "" Then
                Write_SDQT0(EDI_UPC, PRICE, EDI_PRICE, EDI_PRICE_CURR, EDI_PRICE)
            Else
                Asst_SDQT0(RANGE_STYLE_CODE, RANGE_STYLE_PRICE, EDI_UPC)
            End If
        Else
            If Val(rowEDTSDQT0.Item("EDI_PRICE") & "") <> EDI_PRICE Then
                EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Same Item (" & EDI_UPC & ") Diff Lines, Diff Prices",
                            EDI_COND_CODE:="10",
                            EDI_RECEIVED_VALUE:=rowEDTSDQT0.Item("EDI_PRICE") & " vs " & EDI_PRICE)
            End If
        End If

    End Sub

    Sub Process_Items_Pre( _
        EDI_UPC As String, _
        EDU_SKU As String, _
        EDI_ITEM As String, _
        EDI_STYLE As String, _
        EDI_DETL_QTY As Int32, _
        EDI_STORE As String, _
        EDI_SHIP_DC As String, _
        EDI_PRICE As Decimal, _
        EDI_COLOR_CODE As String, _
        EDI_COLOR_NAME As String, _
        EDI_SIZE_DESC As String, _
        RNG_AST_FLG As String, _
        SLN_PARENT_STYLE_CODE As String, _
        SLN_PARENT_STYLE_DESC As String, _
        SLN_PARENT_STYLE_QTY As Int64, _
        SLN_PARENT_INNER_PACK_QTY As Int64, _
        EDI_SLN_SEQ As Int32, _
        RANGE_STYLE_QTY_PER_PP As Int32, _
        QTY_PER_PP As Int32, _
        WORK_QTY As Int64, _
        RANGE_STYLE_CODE As String,
        RANGE_STYLE_PRICE As Decimal,
        RANGE_UPC As String, _
        RANGE_SKU As String, _
        ByRef T2_PO4_QTY As Int32, _
        ByRef T6_PO4_QTY As Int32)

        Dim QTY As Int32

        Dim found_sdq As Boolean = False
        ' Note that we do not even look at EDT850T2.EDI_UPC & EDI_SKU
        For Each rowEDT850T3 As DataRow In dst.Tables("EDT850T3").Select("EDI_DTL_SEQ = " & CStr(EDI_DTL_SEQ))
            With rowEDT850T3
                For sdqi As Int32 = 1 To 10
                    If .Item("EDI_STORE_" & Format(sdqi, "00")) & "" <> "" Then
                        EDI_STORE = Format_Store(.Item("EDI_STORE_" & Format(sdqi, "00")) & "")
                        skip_store = False
                        Bad_Data_Cond_05 = False

                        ' THIS VARIABLE SHOULD BE FORM GLOBAL AND SET IN IMPORT ORDERS
                        'If rowEDT850T1.Item("EDI_SHIP_DC") & "" <> "" Then
                        '    EDI_SHIP_DC = rowEDT850T1.Item("EDI_SHIP_DC") & ""
                        'Else
                        '    EDI_SHIP_DC = rowEDT850T1.Item("EDI_CENTER_CODE") & ""
                        'End If

                        'If rowEDTSLSP1.Item("MIX_DIV_ORDERS") & "" = "1" Then
                        '    Get_Cust()
                        '    If CUST_CODE = "" Then
                        '        Return
                        '    End If
                        '    Get_Terms()
                        'Else
                        '    If (CUST_CODES.Count > 1 Or MILITARY_FLAG = "1") And next_sdq = 1 And dynEDW850T3.Item("EDI_DTL_SEQ") = 1 And dynEDW850T3.Item("EDI_SDQ_SEQ") = 1 Then
                        '        Get_Cust()
                        '        If CUST_CODE = "" Then
                        '            Return
                        '        End If
                        '        Get_Terms()
                        '    End If
                        'End If
                        Get_Ship_To(EDI_STORE, EDI_SHIP_DC)

                        If dst.Tables("EDT850TS").Rows.Find(EDI_STORE) Is Nothing Then
                            ' Check for Possible Order Duplication
                            If Not skip_store And Not Bad_Data_Cond_05 Then
                                dst.Tables("EDT850TS").Rows.Add(New String() {EDI_STORE, EDI_SHIP_DC, CUST_STORE_NO, CUST_DC_NO})
                                Check_for_Possible_Order_Duplication(EDI_STORE)
                            End If
                        End If

                        If sdqi = 1 And Not found_sdq Then
                            found_sdq = True
                            Process_Items(EDI_UPC, EDU_SKU, EDI_ITEM, EDI_STYLE, EDI_COLOR_CODE, EDI_PRICE,
                                RANGE_STYLE_CODE, RANGE_STYLE_PRICE,
                                T2_PO4_QTY, EDI_DETL_QTY, T6_PO4_QTY, rowEDT850T3)
                            If skip_item Then
                                Exit For
                            End If
                            If ITEM_INACTIVE = True Then
                                Exit For
                            End If
                        End If

                        If rowEDTSLSP1.Item("EDI_SDQ_TOT_IND") & "" <> "1" Then
                            QTY = .Item("EDI_QTY_" & Format(sdqi, "00"))
                        Else
                            QTY = .Item("EDI_QTY_" & Format(sdqi, "00")) * EDI_DETL_QTY
                        End If
                        If RANGE_STYLE_CODE = "" Then
                            Write_SDQ(EDI_UPC, EDU_SKU, EDI_ITEM, EDI_STYLE, EDI_STORE, EDI_SHIP_DC, QTY, EDI_PRICE, _
                                    EDI_COLOR_CODE, EDI_COLOR_NAME, _
                                    EDI_SIZE_DESC, _
                                    RNG_AST_FLG, _
                                    SLN_PARENT_STYLE_CODE, _
                                    SLN_PARENT_STYLE_DESC, _
                                    SLN_PARENT_STYLE_QTY, _
                                    SLN_PARENT_INNER_PACK_QTY, _
                                    EDI_SLN_SEQ, _
                                    RANGE_STYLE_QTY_PER_PP, _
                                    QTY_PER_PP, _
                                    WORK_QTY, _
                                    RANGE_STYLE_CODE, _
                                    RANGE_STYLE_PRICE, _
                                    RANGE_UPC, _
                                    RANGE_SKU)

                        Else
                            Asst_Item_SDQ(EDI_UPC, EDU_SKU, EDI_ITEM, EDI_STYLE, EDI_STORE, EDI_SHIP_DC, QTY, EDI_PRICE, _
                                    EDI_COLOR_CODE, EDI_COLOR_NAME, _
                                    EDI_SIZE_DESC, _
                                    RNG_AST_FLG, _
                                    SLN_PARENT_STYLE_CODE, _
                                    SLN_PARENT_STYLE_DESC, _
                                    SLN_PARENT_STYLE_QTY, _
                                    SLN_PARENT_INNER_PACK_QTY, _
                                    EDI_SLN_SEQ, _
                                    RANGE_STYLE_QTY_PER_PP, _
                                    QTY_PER_PP, _
                                    WORK_QTY, _
                                    RANGE_STYLE_CODE, _
                                    RANGE_STYLE_PRICE, _
                                    RANGE_UPC, _
                                    RANGE_SKU)
                        End If
                    End If
                Next
            End With
        Next

        If found_sdq = False Then
            'If CUST_CODE = "" Then
            '    Return
            'End If
            If dst.Tables("EDT850TS").Rows.Find(EDI_STORE) Is Nothing Then
                dst.Tables("EDT850TS").Rows.Add(New String() {EDI_STORE, EDI_SHIP_DC, CUST_STORE_NO, CUST_DC_NO})
            End If
            Process_Items(EDI_UPC, EDU_SKU, EDI_ITEM, EDI_STYLE, EDI_COLOR_CODE, EDI_PRICE,
                         RANGE_STYLE_CODE, RANGE_STYLE_PRICE,
                         T2_PO4_QTY, EDI_DETL_QTY, T6_PO4_QTY, Nothing)
            If skip_item Then
                Return
            End If
            If ITEM_INACTIVE = True Then
                Return
            End If
            QTY = EDI_DETL_QTY
            If STYLE_CODE <> "" Or RANGE_STYLE_CODE <> "" Then
                If RANGE_STYLE_CODE = "" Then
                    Write_SDQ(EDI_UPC, EDU_SKU, EDI_ITEM, EDI_STYLE, EDI_STORE, EDI_SHIP_DC, QTY, EDI_PRICE, _
                                    EDI_COLOR_CODE, EDI_COLOR_NAME, _
                                    EDI_SIZE_DESC, _
                                    RNG_AST_FLG, _
                                    SLN_PARENT_STYLE_CODE, _
                                    SLN_PARENT_STYLE_DESC, _
                                    SLN_PARENT_STYLE_QTY, _
                                    SLN_PARENT_INNER_PACK_QTY, _
                                    EDI_SLN_SEQ, _
                                    RANGE_STYLE_QTY_PER_PP, _
                                    QTY_PER_PP, _
                                    WORK_QTY, _
                                    RANGE_STYLE_CODE, _
                                    RANGE_STYLE_PRICE, _
                                    RANGE_UPC, _
                                    RANGE_SKU)
                Else
                    Asst_Item_SDQ(EDI_UPC, EDU_SKU, EDI_ITEM, EDI_STYLE, EDI_STORE, EDI_SHIP_DC, QTY, EDI_PRICE, _
                                    EDI_COLOR_CODE, EDI_COLOR_NAME, _
                                    EDI_SIZE_DESC, _
                                    RNG_AST_FLG, _
                                    SLN_PARENT_STYLE_CODE, _
                                    SLN_PARENT_STYLE_DESC, _
                                    SLN_PARENT_STYLE_QTY, _
                                    SLN_PARENT_INNER_PACK_QTY, _
                                    EDI_SLN_SEQ, _
                                    RANGE_STYLE_QTY_PER_PP, _
                                    QTY_PER_PP, _
                                    WORK_QTY, _
                                    RANGE_STYLE_CODE, _
                                    RANGE_STYLE_PRICE, _
                                    RANGE_UPC, _
                                    RANGE_SKU)
                End If
            End If
        End If

    End Sub

    Sub Write_SDQT0(EDI_UPC As String, PRICE As Decimal, EDI_PRICE As Decimal, EDI_PRICE_CURR As Decimal, ITEM_PRICE_CURR As Decimal)
        If STYLE_CODE <> "" Then
            Dim rowEDTSDQT0 As DataRow = dst.Tables("EDTSDQT0").NewRow
            With rowEDTSDQT0
                .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                .Item("EDI_UPC") = EDI_UPC
                .Item("PRICE") = PRICE
                .Item("EDI_PRICE") = EDI_PRICE
                .Item("EDI_PRICE_CURR") = EDI_PRICE_CURR
                .Item("ITEM_PRICE_CURR") = ITEM_PRICE_CURR
            End With
            dst.Tables("EDTSDQT0").Rows.Add(rowEDTSDQT0)
        End If
    End Sub

    Sub Get_ICTSTYL1(
        EDI_UPC As String,
        EDI_SKU As String,
        EDI_ITEM As String,
        EDI_STYLE As String,
        EDI_COLOR_CODE As String,
        ByRef PRICE As Decimal,
        ByRef CUST_PRICE As Decimal,
        ByRef EDI_PRICE As Decimal,
        ByRef EDI_PRICE_CURR As Decimal,
        ByRef ITEM_PRICE_CURR As Decimal,
        ByRef RETAIL_PRICE As Decimal,
        ByRef PRICE_DISC As Decimal,
        ByRef T2_PO4_QTY As Int32,
        ByRef EDI_DETL_QTY As Int32,
        ByRef T6_PO4_QTY As Int32, rowEDT850T3 As DataRow)

        STYLE_CODE = ""
        COLOR_CODE = ""

        Dim ALL_EDI_ITEM_ok As Boolean = False

        Dim EDI_UPC_ok As Boolean = (EDI_UPC = "")
        Dim EDI_SKU_ok As Boolean = (EDI_SKU = "")
        Dim EDI_ITEM_ok As Boolean = (EDI_ITEM = "")
        Dim EDI_STYLE_ok As Boolean = (EDI_STYLE = "")

        Dim EDI_UPC_ITEM As String = IIf(EDI_UPC_ok, "none", "")
        Dim EDI_SKU_ITEM As String = IIf(EDI_SKU_ok, "none", "")
        Dim EDI_ITM_ITEM As String = IIf(EDI_ITEM_ok, "none", "")
        Dim EDI_STY_ITEM As String = IIf(EDI_STYLE_ok, "none", "")

        CUST_PRICE = 0

        ' Locate Item via UPC, if supplied

        ' DO NOT USE THE EDI_UPC AS IF IT WERE A CUSTOMER ITEM CODE
        'If Not EDI_UPC_ok Then
        '    Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, EDI_UPC})
        '    If rowSOTCSTY1 IsNot Nothing Then
        '        If STYLE_CODE = "" Or rowSOTCSTY1.Item("FORCE_SUB") & "" = "1" Then
        '            STYLE_CODE = rowSOTCSTY1.Item("STYLE_CODE") & ""
        '        End If
        '        EDI_UPC_ITEM = rowSOTCSTY1.Item("STYLE_CODE") & ""
        '        'CUST_PRICE = rowSOTCSTY1.Item("CUST_PRICE") & ""
        '        EDI_UPC_ok = True
        '        ALL_EDI_ITEM_ok = True
        '    End If
        'End If

        If Not EDI_UPC_ok And EDI_MERCH_TYPE <> "" Then
            Dim rowEDTUPCX4 As DataRow = dst.Tables("EDTUPCX4").Rows.Find(New String() {CUST_CODE, EDI_MERCH_TYPE, EDI_UPC})
            If rowEDTUPCX4 IsNot Nothing Then
                Dim row As DataRow = LookUp("ICTSTYL1", rowEDTUPCX4.Item("STYLE_CODE"))
                If row IsNot Nothing Then
                    STYLE_CODE = rowEDTUPCX4.Item("STYLE_CODE")
                    COLOR_CODE = rowEDTUPCX4.Item("COLOR_CODE")
                    EDI_UPC_ITEM = STYLE_CODE
                    EDI_UPC_ok = True
                    ALL_EDI_ITEM_ok = True
                End If
            End If
        End If

        If Not EDI_UPC_ok Then
            Dim rowEDTUPCX1 As DataRow = dst.Tables("EDTUPCX1").Rows.Find(New String() {CUST_CODE, EDI_UPC})
            If rowEDTUPCX1 IsNot Nothing Then
                Dim row As DataRow = LookUp("ICTSTYL1", rowEDTUPCX1.Item("STYLE_CODE"))
                If row IsNot Nothing Then
                    STYLE_CODE = rowEDTUPCX1.Item("STYLE_CODE")
                    COLOR_CODE = rowEDTUPCX1.Item("COLOR_CODE")
                    EDI_UPC_ITEM = STYLE_CODE
                    EDI_UPC_ok = True
                    ALL_EDI_ITEM_ok = True
                End If
            End If
        End If

        If Not EDI_UPC_ok Then
            Dim rowICTITEMX As DataRow = dst.Tables("ICTITEMX").Rows.Find(EDI_UPC)
            If rowICTITEMX IsNot Nothing Then
                STYLE_CODE = rowICTITEMX.Item("STYLE_CODE")
                COLOR_CODE = rowICTITEMX.Item("COLOR_CODE")
                EDI_UPC_ITEM = STYLE_CODE
                EDI_UPC_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If

        ' Locate Item via Customer Item Code, if supplied

        If Not EDI_SKU_ok Then
            Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, EDI_SKU})
            If rowSOTCSTY1 IsNot Nothing Then
                If STYLE_CODE = "" Then ' Or rowSOTCSTY1.Item("FORCE_SUB") & "" = "1" Then
                    STYLE_CODE = rowSOTCSTY1.Item("STYLE_CODE") & ""
                    COLOR_CODE = rowSOTCSTY1.Item("COLOR_CODE") & ""
                End If
                EDI_SKU_ITEM = rowSOTCSTY1.Item("STYLE_CODE") & ""
                'CUST_PRICE = rowSOTCSTY1.Item("CUST_PRICE") & ""
                EDI_SKU_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If

        If Not EDI_ITEM_ok Then
            Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, EDI_ITEM})
            If rowSOTCSTY1 IsNot Nothing Then
                If STYLE_CODE = "" Then ' Or rowSOTCSTY1.Item("FORCE_SUB") & "" = "1" Then
                    STYLE_CODE = rowSOTCSTY1.Item("STYLE_CODE") & ""
                    COLOR_CODE = rowSOTCSTY1.Item("COLOR_CODE") & ""
                End If
                EDI_ITM_ITEM = rowSOTCSTY1.Item("STYLE_CODE") & ""
                'CUST_PRICE = rowSOTCSTY1.Item("CUST_PRICE") & ""
                EDI_ITEM_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If

        If Not EDI_STYLE_ok Then
            Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, EDI_STYLE})
            If rowSOTCSTY1 IsNot Nothing Then
                If STYLE_CODE = "" Then '  Or rowSOTCSTY1.Item("FORCE_SUB") & "" = "1" Then
                    STYLE_CODE = rowSOTCSTY1.Item("STYLE_CODE") & ""
                    COLOR_CODE = rowSOTCSTY1.Item("COLOR_CODE") & ""
                End If
                EDI_STY_ITEM = rowSOTCSTY1.Item("STYLE_CODE") & ""
                'CUST_PRICE = rowSOTCSTY1.Item("CUST_PRICE") & ""
                EDI_STYLE_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If

        ' Locate Item using STYLE_CODE
        If STYLE_CODE = "" Then
            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(New String() {EDI_ITEM})
            If rowICTSTYL1 Is Nothing Then
                rowICTSTYL1 = LookUp("ICTSTYL1", EDI_ITEM)
            End If
            If rowICTSTYL1 IsNot Nothing Then
                STYLE_CODE = EDI_ITEM
                '    Stop ' COLOR_CODE = "?"
                COLOR_CODE = "AST"
                EDI_ITM_ITEM = EDI_ITEM
                EDI_ITEM_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If

        ' Stop ' ASSUMES THAT GLOBAL STYLE_CODE IS IN USE
        Check_Item(EDI_UPC_ok, EDI_UPC, EDI_UPC_ITEM, ALL_EDI_ITEM_ok)
        Check_Item(EDI_SKU_ok, EDI_SKU, EDI_SKU_ITEM, ALL_EDI_ITEM_ok)
        Check_Item(EDI_ITEM_ok, EDI_ITEM, EDI_ITM_ITEM, ALL_EDI_ITEM_ok)
        Check_Item(EDI_STYLE_ok, EDI_STYLE, EDI_STY_ITEM, ALL_EDI_ITEM_ok)

        If ASCMAIN1.CLIENT = "RGI" Then
            If Not String.IsNullOrEmpty(EDI_COLOR_CODE) And (EDI_UPC_ok = False Or String.IsNullOrEmpty(EDI_UPC)) Then
                If COLOR_CODE = "AST" Or String.IsNullOrEmpty(COLOR_CODE) Then
                    COLOR_CODE = EDI_COLOR_CODE
                End If

            End If
        End If

        '' check if a asst item
        'tblICWAITM1.Seek("=", CUST_CODE_edi, EDI_ITEM)
        'If Not tblICWAITM1.NoMatch Then
        '    ASST_STYLE_CODE = tblICWAITM1.Item("ASST_STYLE_CODE")
        '    EDI_ITEM_ok = True
        '    ALL_EDI_ITEM_ok = True
        '    EDI_ITM_ITEM = tblICWAITM1.Item("ASST_STYLE_CODE")
        'End If
        'tblICWAITM1.Seek("=", CUST_CODE_edi, EDI_UPC)
        'If Not tblICWAITM1.NoMatch Then
        '    ASST_STYLE_CODE = tblICWAITM1.Item("ASST_STYLE_CODE")
        '    EDI_UPC_ok = True
        '    ALL_EDI_ITEM_ok = True
        '    EDI_UPC_ITEM = tblICWAITM1.Item("ASST_STYLE_CODE")

        'End If
        'tblICWAITM1.Seek("=", CUST_CODE_edi, EDI_SKU)
        'If Not tblICWAITM1.NoMatch Then
        '    ASST_STYLE_CODE = tblICWAITM1.Item("ASST_STYLE_CODE")
        '    EDI_SKU_ok = True
        '    ALL_EDI_ITEM_ok = True
        '    EDI_SKU_ITEM = tblICWAITM1.Item("ASST_STYLE_CODE")
        'End If

        If Not EDI_UPC_ok And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) And STYLE_CODE <> "" Then
            Bad_Data(EDI_COND_DESC:="EDI UPC " & EDI_UPC & " missing for Style " & STYLE_CODE, EDI_COND_CODE:="23", EDI_RECEIVED_VALUE:=EDI_UPC)
        End If
        If Not EDI_SKU_ok And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) And STYLE_CODE <> "" Then
            Bad_Data(EDI_COND_DESC:="EDI SKU " & EDI_SKU & " missing for Style " & STYLE_CODE, EDI_COND_CODE:="24", EDI_RECEIVED_VALUE:=EDI_SKU)
        End If
        If Not EDI_ITEM_ok And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI Item " & EDI_ITEM & " missing for Style " & STYLE_CODE, EDI_COND_CODE:="25", EDI_RECEIVED_VALUE:=EDI_ITEM)
        End If
        If Not EDI_STYLE_ok And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI Style " & EDI_STYLE & " missing for Style " & STYLE_CODE, EDI_COND_CODE:="25", EDI_RECEIVED_VALUE:=EDI_STYLE)
        End If

        If EDI_UPC_ITEM <> "none" And EDI_UPC_ITEM <> STYLE_CODE And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI UPC " & EDI_UPC & " does not match Style " & STYLE_CODE, EDI_COND_CODE:="30", EDI_RECEIVED_VALUE:=EDI_UPC_ITEM)
        End If
        If EDI_SKU_ITEM <> "none" And EDI_SKU_ITEM <> STYLE_CODE And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI SKU " & EDI_SKU & " does not match Style " & STYLE_CODE, EDI_COND_CODE:="31", EDI_RECEIVED_VALUE:=EDI_SKU_ITEM)
        End If
        If EDI_ITM_ITEM <> "none" And EDI_ITM_ITEM <> STYLE_CODE And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI Item " & EDI_ITEM & " does not match Style " & STYLE_CODE, EDI_COND_CODE:="32", EDI_RECEIVED_VALUE:=EDI_ITM_ITEM)
        End If
        If EDI_STY_ITEM <> "none" And EDI_STY_ITEM <> STYLE_CODE And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI Style " & EDI_STYLE & " does not match Style " & STYLE_CODE, EDI_COND_CODE:="32", EDI_RECEIVED_VALUE:=EDI_STY_ITEM)
        End If

        If STYLE_CODE <> "" Then
            rowICTSTYL1 = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
            If rowICTSTYL1 Is Nothing Then
                rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    rowICTSTYL1 = Write_Item_to_MDB(rowICTSTYL1)
                End If
            End If
            saleable_item = True
            'If rowICTSTYL1.Item("ITEM_SNU_CODE") & "" = "S" Then
            '    saleable_item = True
            'Else
            '    saleable_item = False
            'End If

            Dim rowSOTSDIV1 As DataRow = dst.Tables("SOTSDIV1").Rows.Find(rowICTSTYL1.Item("SALES_DIVISION_CODE") & "")
            If rowSOTSDIV1 Is Nothing Then
                Bad_Data(EDI_COND_DESC:="Missing/Bad Sales Division for Item " & STYLE_CODE, EDI_COND_CODE:="34", EDI_RECEIVED_VALUE:=rowICTSTYL1.Item("SALES_DIVISION_CODE") & "")
            Else
                'If rowSOTSDIV1.Item("STATUS") & "" <> "A" Then
                '    Bad_Data(EDI_COND_DESC:="Inactive Sales Division for Item " & STYLE_CODE, EDI_COND_CODE:="19", EDI_RECEIVED_VALUE:=rowICTSTYL1.Item("SALES_DIVISION_CODE") & "")
                '    If EDI_ACTION = "S" Then
                '        Stop
                '        ITEM_INACTIVE = True
                '        Return
                '    End If
                'End If
            End If


            EDI_PRICE_CURR = 0
            ' EDI_PRICE = 0
            'Dim T2_PO4_QTY As Int32 = 0
            'Dim EDI_DETL_QTY As Int32 = 0
            'Dim T6_PO4_QTY As Int32 = 0

            If rowEDTSLSP1.Item("EDI_RETAIL_PRICE_FLAG") & "" = "1" Then
                EDI_PRICE_CURR = EDI_PRICE - (EDI_PRICE * (PRICE_BASE_DPCT / 100))
                EDI_PRICE = EDI_PRICE * CURR_EXCH_RATE - (EDI_PRICE * CURR_EXCH_RATE * (PRICE_BASE_DPCT / 100))
            Else
                EDI_PRICE_CURR = EDI_PRICE
                EDI_PRICE = EDI_PRICE * CURR_EXCH_RATE
            End If

            Dim CARTON_PACK_QTY As Integer = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
            If CARTON_PACK_QTY = 0 Then CARTON_PACK_QTY = 1

            Dim INNER_PACK_QTY As Integer = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")

            Dim QF As Decimal = 1

            If ASCMAIN1.CLIENT = "NYA" Then
                If CUST_CODE = "LOBLAW" Then
                    QF = CARTON_PACK_QTY / INNER_PACK_QTY
                    ' QF = INNER_PACK_QTY
                    QF = CARTON_PACK_QTY / T2_PO4_QTY
                End If
            End If

            If rowEDTSLSP1.Item("EDI_PO1_TOT_IND") & "" <> "1" Then
                If (T2_PO4_QTY <> 0 And T2_PO4_QTY <> 1) And T2_PO4_QTY * QF <> CARTON_PACK_QTY Then

                    ' NEXT LINE CODED FOR FREDMEYER - MAY NEED TO BE CODED FOR ALL CUSTOMERS
                    If rowEDT850T3 IsNot Nothing AndAlso rowEDT850T3.Item("EDI_SDQ_UOM") = "EA" AndAlso CUST_CODE = "FREDMEYER" Then
                        ' QTY IS IN EACHES, NO CASE PACK CONVERSIONS REQUIRED
                    Else

                        EDI_REPLACED_VALUE = Bad_Data(
                            EDI_COND_DESC:="PO4 (" & CStr(T2_PO4_QTY) & ") <> Case Qty (" & CStr(CARTON_PACK_QTY) & ") " & STYLE_CODE,
                                 EDI_COND_CODE:="55",
                                 EDI_RECEIVED_VALUE:=CStr(T2_PO4_QTY))
                        If EDI_ACTION = "R" Or EDI_ACTION = "A" Then
                            If Val(EDI_REPLACED_VALUE) <> 0 Then
                                CARTON_PACK_QTY = Val(EDI_REPLACED_VALUE)
                            End If

                            EDI_DETL_QTY = EDI_DETL_QTY * CARTON_PACK_QTY
                            EDI_PRICE = EDI_PRICE / CARTON_PACK_QTY
                            EDI_PRICE_CURR = EDI_PRICE_CURR / CARTON_PACK_QTY

                        End If
                        ' REMMING NEXT 4 LINES OUT BECAUSE WE SHOULD ONLY DO THIS IF WE ARE TAKING ACTION
                        'EDI_DETL_QTY = EDI_DETL_QTY * CARTON_PACK_QTY
                        'EDI_PRICE = EDI_PRICE / CARTON_PACK_QTY
                        ''EDI_PRICE_CURR = EDI_PRICE_CURR * T2_PO4_QTY / CARTON_PACK_QTY
                        'EDI_PRICE_CURR = EDI_PRICE_CURR / CARTON_PACK_QTY

                    End If
                End If

                If (T6_PO4_QTY <> 0 And T6_PO4_QTY <> 1) And T6_PO4_QTY <> CARTON_PACK_QTY Then
                    Bad_Data(EDI_COND_DESC:="PO6 (" & CStr(T6_PO4_QTY) & ") <> Case Qty (" & CStr(CARTON_PACK_QTY) & ") " & STYLE_CODE,
                             EDI_COND_CODE:="56",
                             EDI_RECEIVED_VALUE:=CStr(T6_PO4_QTY))
                    'If Good_Data_Value <> "" Then
                    '    ' PROCEED
                    'End If
                    EDI_DETL_QTY = EDI_DETL_QTY / CARTON_PACK_QTY
                    EDI_PRICE = EDI_PRICE * CARTON_PACK_QTY
                    EDI_PRICE_CURR = EDI_PRICE_CURR * CARTON_PACK_QTY
                End If
            End If

            If T2_PO4_QTY <> 0 Then
                If CUST_CODE = "SEARSALLO" Or CUST_CODE = "CVS" Then
                    EDI_DETL_QTY = EDI_DETL_QTY * T2_PO4_QTY
                    EDI_PRICE = EDI_PRICE / T2_PO4_QTY
                    EDI_PRICE_CURR = EDI_PRICE_CURR / T2_PO4_QTY
                End If
                If CUST_CODE = "SDM" Then
                    EDI_DETL_QTY = EDI_DETL_QTY * T2_PO4_QTY
                    EDI_PRICE = EDI_PRICE / T2_PO4_QTY
                    EDI_PRICE_CURR = EDI_PRICE_CURR / T2_PO4_QTY
                End If
                If CUST_CODE = "LOBLAW" Then
                    EDI_DETL_QTY = EDI_DETL_QTY * T2_PO4_QTY * QF
                    EDI_PRICE = EDI_PRICE / (T2_PO4_QTY * QF)
                    EDI_PRICE_CURR = EDI_PRICE_CURR / (T2_PO4_QTY * QF)
                End If
            End If

            'MAYBE WE SHOULD OPEN THIS UP TO ALL - IT WAS A BLOCK THAT WAS REMMED OUT AND i NEEDED IT FOR THE CUSTOMERS BELOW
            ' THIS CODE ADJUSTS THE CASE PACK QTY - A CUSTOMER WILL SEND IN 1 AND WE KNOW IT IS 6
            If CUST_CODE = "CASUALMALE" Or CUST_CODE = "CHARMING" Or CUST_CODE = "LANEBRY" Or CUST_CODE = "CATHERINE" Or CUST_CODE = "CHARLOTTE" Then
                If (T2_PO4_QTY = 0 Or T2_PO4_QTY = 1) And (CARTON_PACK_QTY <> 0 And CARTON_PACK_QTY <> 1) And
                    (rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "CA" Or
                     rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "CS" Or
                     (rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "" And
                      (rowEDT850T2.Item("EDI_PO4_UOM") & "" = "CA" _
                       Or rowEDT850T2.Item("EDI_PO4_UOM") & "" = "CS"))) Then
                    EDI_PRICE_CURR = EDI_PRICE / CARTON_PACK_QTY
                    EDI_PRICE = EDI_PRICE * CURR_EXCH_RATE / CARTON_PACK_QTY
                    EDI_DETL_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "") * CARTON_PACK_QTY
                End If
            End If


            'If T2_PO4_QTY = 0 Then
            '    T2_PO4_QTY = 1
            '    '            T2_PO4_QTY = dynICTSTYL1.Item("CARTON_PACK_QTY")
            'End If


            'OraD.Parameters("CODE") = CURR_CODE
            'OraD.Parameters("STYLE_CODE") = STYLE_CODE
            'dynICTRETL1.Refresh()
            'If Not dynICTRETL1.EOF Then
            '    RETAIL_PRICE = Val(dynICTRETL1.Item("ITEM_RETAIL_PRICE"))
            'Else
            '    RETAIL_PRICE = Val(dynICTSTYL1.Item("ITEM_RETAIL_PRICE") & "")
            'End If

            PRICE = EDI_PRICE
            CUST_PRICE = EDI_PRICE
            RETAIL_PRICE = Val(rowICTSTYL1.Item("STYLE_RETAIL") & "")
            ITEM_PRICE_CURR = 0
            PRICE_DISC = 0

            Select Case PRICE_BASIS
                Case "R"
                    PRICE = (RETAIL_PRICE - (RETAIL_PRICE * (PRICE_BASE_DPCT / 100))) * CURR_EXCH_RATE
                    ITEM_PRICE_CURR = RETAIL_PRICE - (RETAIL_PRICE * (PRICE_BASE_DPCT / 100))

                    PRICE = System.Math.Round(PRICE, 2)
                    ITEM_PRICE_CURR = System.Math.Round(ITEM_PRICE_CURR, 2)

                    PRICE_DISC = PRICE_BASE_DPCT
                    CUST_PRICE = PRICE

                Case "L"
                    Dim rowSOTPRIC2 As DataRow = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE, STYLE_CODE})
                    If rowSOTPRIC2 IsNot Nothing Then
                        PRICE = rowSOTPRIC2.Item("ITEM_PRICE")
                        CUST_PRICE = PRICE
                        ITEM_PRICE_CURR = rowSOTPRIC2.Item("ITEM_PRICE")
                        PRICE_DISC = 0
                    Else
                        Bad_Data(EDI_COND_DESC:="No Std Price for Item " & STYLE_CODE,
                             EDI_COND_CODE:="39",
                             EDI_RECEIVED_VALUE:=STYLE_CODE)
                    End If
                    'price = lookup from price file

                Case "E"
                    PRICE = rowEDT850T2.Item("EDI_PRICE") * CURR_EXCH_RATE / T2_PO4_QTY
                    CUST_PRICE = PRICE
                    ITEM_PRICE_CURR = rowEDT850T2.Item("EDI_PRICE") / T2_PO4_QTY
                    PRICE_DISC = 0

            End Select

            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
            If rowICTSTYC1 Is Nothing Then
                Bad_Data(EDI_COND_DESC:="Bad Style/Color for " & STYLE_CODE & "/" & COLOR_CODE,
                                   EDI_COND_CODE:="29",
                                   EDI_RECEIVED_VALUE:=STYLE_CODE & "/" & COLOR_CODE)
            End If

            Dim rowEDTITEMX As DataRow = dst.Tables("EDTITEMX").NewRow
            With rowEDTITEMX
                .Item("CUST_CODE") = CUST_CODE
                .Item("EDI_UPC") = EDI_UPC
                .Item("EDI_SKU") = EDI_SKU
                .Item("EDI_ITEM") = EDI_ITEM
                .Item("EDI_STYLE") = EDI_STYLE
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("CUST_PRICE") = CUST_PRICE
                .Item("ITEM_PRICE_CURR") = ITEM_PRICE_CURR
                .Item("PRICE_DISC") = PRICE_DISC
                .Item("RETAIL_PRICE") = RETAIL_PRICE
                .Item("RETAIL_PRICE_CURR") = RETAIL_PRICE / CURR_EXCH_RATE
                '.Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
            End With
            dst.Tables("EDTITEMX").Rows.Add(rowEDTITEMX)
        End If

        If COLOR_CODE.Length > 4 And ASCMAIN1.Running_in_VS Then
            Stop
            COLOR_CODE = ""
        End If

    End Sub

    Function getEDT850T9() As String
        Dim PHONE_NO As String = ""
        Dim rCnt As Integer = 0

        Dim rowSOTSHIP5 As DataRow = dst.Tables("EDT850T5").Select("EDI_ADDR_TYPE = 'ST'").FirstOrDefault
        For Each rContact As DataRow In dst.Tables("EDT850T9").Select("", "EDI_PER_SEQ")
            Dim holdPhone As String = ""
            'highest priority last to override previous
            If rContact("CONTACT_COMM_NO_QUAL_3") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_3")
            If rContact("CONTACT_COMM_NO_QUAL_2") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_2")
            If rContact("CONTACT_COMM_NO_QUAL_1") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_1")
            'Next highest priority first, we leave as soon as first hit, if we need additional info change code below
            If holdPhone <> "" Then
                'Found a Phone, Count it
                rCnt += 1
                If rowSOTSHIP5("EDI_CUST_NAME_ADR") & "" <> "" And rowSOTSHIP5("EDI_CUST_NAME_ADR") & "" = rContact("CONTACT_NAME") & "" Then
                    PHONE_NO = holdPhone
                    Exit For
                End If
                If rContact("CONTACT_NAME") & "" = "ST" Or rContact("CONTACT_FUNC_CODE") & "" = "ST" Or rContact("CONTACT_FUNC_CODE") & "" = "SH" Then
                    PHONE_NO = holdPhone
                    Exit For
                End If
            End If
        Next
        'Only a single rec with a Phone, Use it
        If rCnt = 1 And PHONE_NO = "" Then
            For Each rContact As DataRow In dst.Tables("EDT850T9").Select("", "EDI_PER_SEQ")
                Dim holdPhone As String = ""
                'highest priority last to override previous
                If rContact("CONTACT_COMM_NO_QUAL_3") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_3")
                If rContact("CONTACT_COMM_NO_QUAL_2") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_2")
                If rContact("CONTACT_COMM_NO_QUAL_1") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_1")
                'Found our Single record
                If holdPhone <> "" Then
                    PHONE_NO = holdPhone
                    Exit For
                End If
            Next
        End If

        'exhaust all other options before matching by line no, it has to be separete loop
        If rowSOTSHIP5("EDI_ADR_SEQ") & "" <> "" And PHONE_NO = "" Then
            For Each rContact As DataRow In dst.Tables("EDT850T9").Select("EDI_PER_SEQ = '" & rowSOTSHIP5("EDI_ADR_SEQ") & "'")
                Dim holdPhone As String = ""
                'highest priority last to override previous
                If rContact("CONTACT_COMM_NO_QUAL_3") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_3")
                If rContact("CONTACT_COMM_NO_QUAL_2") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_2")
                If rContact("CONTACT_COMM_NO_QUAL_1") & "" = "TE" Then holdPhone = rContact("CONTACT_COMM_NO_1")
                If rowSOTSHIP5("EDI_ADR_SEQ") & "" = rContact("EDI_PER_SEQ") & "" Then
                    PHONE_NO = holdPhone
                End If
            Next
        End If
        'Cleanup Phone no - only send digits, clear common format chars
        PHONE_NO = PHONE_NO.Replace("-", "").Replace("(", "").Replace(")", "").Replace(".", "")
        Dim phonetemp() As String = PHONE_NO.Split(" ")
        If phonetemp.Length > 1 And phonetemp(0).Length > 9 Then
            PHONE_NO = phonetemp(0)
        End If
        If PHONE_NO.Length > 20 Then
            PHONE_NO = PHONE_NO.Substring(0, 20)
        End If

        Return PHONE_NO
    End Function

    Sub Check_Item(ByRef EDI_XXX_ok, ByRef EDI_XXX, ByRef EDI_XXX_ITEM, ByRef ALL_EDI_ITEM_ok)

        Dim rowICTSTYL1 As DataRow
        If EDI_XXX_ok = False Then
            rowICTSTYL1 = dst.Tables("ICTSTYL1").Rows.Find(EDI_XXX)
            If rowICTSTYL1 IsNot Nothing Then
                If STYLE_CODE = "" Then
                    STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & ""
                    COLOR_CODE = "AST" ' "?"
                End If
                EDI_XXX_ITEM = rowICTSTYL1.Item("STYLE_CODE") & ""
                EDI_XXX_ok = True
                ALL_EDI_ITEM_ok = True
            Else
                rowICTSTYL1 = LookUp("ICTSTYL1", EDI_XXX)
                If rowICTSTYL1 IsNot Nothing Then
                    rowICTSTYL1 = Write_Item_to_MDB(rowICTSTYL1)
                    If STYLE_CODE = "" Then
                        STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & ""
                        COLOR_CODE = "AST" ' "?"
                    End If
                    EDI_XXX_ITEM = STYLE_CODE
                    EDI_XXX_ok = True
                    ALL_EDI_ITEM_ok = True
                End If
            End If
        End If
    End Sub

    Sub PO4_OverWrite( _
        EDI_UPC As String, _
        EDI_SKU As String, _
        EDI_ITEM As String, _
        PO4_OW As Int32, _
        T2_PO4_QTY As Int32, _
        T6_PO4_QTY As Int32)

        Dim rowSOTCSTY1 As DataRow = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, EDI_UPC})
        If rowSOTCSTY1 IsNot Nothing Then
            rowSOTCSTY1 = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, EDI_ITEM})
        ElseIf rowSOTCSTY1 IsNot Nothing Then
            rowSOTCSTY1 = dst.Tables("SOTCSTY1").Rows.Find(New String() {CUST_CODE, EDI_SKU})
        End If

        If rowSOTCSTY1 IsNot Nothing Then
            If PO4_OW = Val(rowSOTCSTY1.Item("CUST_PO4_QTY") & "") + 0 Then
                T2_PO4_QTY = Val(rowSOTCSTY1.Item("ITEM_PO4_QTY") & "")
                T6_PO4_QTY = Val(rowSOTCSTY1.Item("ITEM_PO4_QTY") & "")
            End If
        End If
    End Sub

    Function Write_Item_to_MDB(ROW As DataRow) As DataRow
        Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").NewRow
        rowICTSTYL1.ItemArray = ROW.ItemArray

        If Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "") = 0 Then
            rowICTSTYL1.Item("CARTON_PACK_QTY") = 1
        End If
        dst.Tables("ICTSTYL1").Rows.Add(rowICTSTYL1)

        Return rowICTSTYL1
    End Function

    Sub Write_SDQ( _
        EDI_UPC As String, _
        EDI_SKU As String, _
        EDI_ITEM As String, _
        EDI_STYLE As String, _
        EDI_STORE As String, _
        EDI_SHIP_DC As String, QTY As Int32, EDI_PRICE As Decimal, _
        EDI_COLOR_CODE As String, _
        EDI_COLOR_NAME As String, _
        EDI_SIZE_DESC As String, _
        RNG_AST_FLG As String, _
        SLN_PARENT_STYLE_CODE As String, _
        SLN_PARENT_STYLE_DESC As String, _
        SLN_PARENT_STYLE_QTY As Int64, _
        SLN_PARENT_INNER_PACK_QTY As Int64, _
        EDI_SLN_SEQ As Int32, _
        RANGE_STYLE_QTY_PER_PP As Int32, _
        QTY_PER_PP As Int32, _
        WORK_QTY As Int64, _
        RANGE_STYLE_CODE As String, _
        RANGE_STYLE_PRICE As Decimal, _
        RANGE_UPC As String, _
        RANGE_SKU As String)

        If STYLE_CODE = "" Then
            Return
        End If
        If skip_store Or Bad_Data_Cond_05 Then
            Return
        End If

        '  If EDI_STORE = "000002" And EDI_DTL_SEQ = 1 Then Stop

        Dim rowEDTSDQT1 As DataRow = dst.Tables("EDTSDQT1").Rows.Find _
            (New Object() {EDI_DOC_SEQ_NO, EDI_STORE, EDI_DTL_SEQ, STYLE_CODE, COLOR_CODE, _
                           EDI_ITEM, EDI_STYLE, EDI_COLOR_CODE, EDI_COLOR_NAME, EDI_UPC, EDI_SKU, EDI_SIZE_DESC})
        'Dim rowEDTSDQT1 As DataRow = dst.Tables("EDTSDQT1").Rows.Find _
        '(New Object() {EDI_DOC_SEQ_NO, EDI_STORE, EDI_DTL_SEQ, STYLE_CODE, COLOR_CODE})

        ' If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" And RANGE_STYLE_QTY_PER_PP > 1 Then MsgBox("RANGE_STYLE_QTY_PP = " & CStr(RANGE_STYLE_QTY_PER_PP))

        Dim QTY_SDQ As Int64 = 0 ' QTY * IIf(SLN_PARENT_STYLE_QTY = 0, 1, SLN_PARENT_STYLE_QTY) * IIf(QTY_PER_PP = 0, 1, QTY_PER_PP) '       IIf(RANGE_STYLE_QTY_PER_PP = 0, 1, RANGE_STYLE_QTY_PER_PP)
        If SLN_PARENT_STYLE_QTY <> 0 Then
            QTY_SDQ = QTY * IIf(SLN_PARENT_STYLE_QTY = 0, 1, SLN_PARENT_STYLE_QTY)
        Else
            QTY_SDQ = QTY * IIf(QTY_PER_PP = 0, 1, QTY_PER_PP)
        End If

        If RANGE_STYLE_CODE <> "" Then
            QTY_SDQ = WORK_QTY
        End If

        If rowEDTSDQT1 IsNot Nothing Then
            rowEDTSDQT1.Item("QTY") = Val(rowEDTSDQT1.Item("QTY") & "") + QTY_SDQ
        Else
            rowEDTSDQT1 = dst.Tables("EDTSDQT1").NewRow
            With rowEDTSDQT1
                .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                .Item("EDI_STORE") = EDI_STORE
                .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("EDI_ITEM") = EDI_ITEM
                .Item("EDI_STYLE") = EDI_STYLE
                .Item("EDI_COLOR_CODE") = EDI_COLOR_CODE
                .Item("EDI_COLOR_NAME") = EDI_COLOR_NAME
                .Item("EDI_UPC") = EDI_UPC
                .Item("EDI_SKU") = EDI_SKU
                .Item("EDI_SIZE_DESC") = EDI_SIZE_DESC

                .Item("RANGE_STYLE_CODE") = RANGE_STYLE_CODE
                .Item("RNG_AST_FLG") = RNG_AST_FLG
                'If found_sdq = False And EDI_SLN_SEQ <> 0 Then
                '  QTY = QTY * IIf(SLN_PARENT_STYLE_QTY = 0, 1, SLN_PARENT_STYLE_QTY)
                'End If
                .Item("QTY") = QTY_SDQ
                .Item("EDI_SHIP_DC") = EDI_SHIP_DC


                .Item("SLN_PARENT_STYLE_CODE") = SLN_PARENT_STYLE_CODE
                '.Item("SLN_PARENT_STYLE_DESC") = SLN_PARENT_STYLE_DESC
                .Item("SLN_PARENT_STYLE_QTY") = SLN_PARENT_STYLE_QTY
                .Item("SLN_PARENT_INNER_PACK_QTY") = SLN_PARENT_INNER_PACK_QTY
                .Item("EDI_SLN_SEQ") = EDI_SLN_SEQ
                .Item("RANGE_STYLE_QTY_PER_PP") = RANGE_STYLE_QTY_PER_PP
                .Item("QTY_PER_PP") = QTY_PER_PP ' this is not being cleared or set on non Range styles - Rick

                If RANGE_STYLE_CODE <> "" Then
                    .Item("RANGE_PARENT_QTY") = WORK_QTY
                    If RANGE_UPC <> "" Then
                        .Item("EDI_UPC") = RANGE_UPC
                    End If
                    If RANGE_SKU <> "" Then
                        .Item("EDI_SKU") = RANGE_SKU
                    End If
                End If

                .Item("EDI_PRICE") = EDI_PRICE
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = CUST_STORE_NO
                .Item("CUST_DC_NO") = CUST_DC_NO
            End With
            dst.Tables("EDTSDQT1").Rows.Add(rowEDTSDQT1)
        End If
    End Sub

    Sub Update_ICTSTAT2(ORDR_GROUP_NO As String)
        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, Sum (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & "  from SOTORDR1,SOTORDR2" & vbCrLf _
            & "  where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "    and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            & "  group by SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   ICPSTAT2(R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,0,0,0,R1.ORDR_QTY,0,0);" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()
    End Sub

    Private Sub grdEDT850TE_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdEDT850TE.AfterRowActivate

    End Sub

    Function Format_Store(EDI_STORE As String)
        If Len(EDI_STORE) = 13 Then Return EDI_STORE ' looks like a GLN

        Dim CUST_STORE_NO As String = EDI_STORE
        If IsNumeric(EDI_STORE) Then
            CUST_STORE_NO = Format(Val(EDI_STORE), "000000")
        End If
        Return CUST_STORE_NO
    End Function

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()

        If tabMain.SelectedTab.Key = "Orders Waiting to be Imported" Then
            grdEDT850T1.DataSource = dst.Tables("EDT850T1")
            grdEDT850T1.Parent = splEDT850T1.Panel1
            With grdEDT850T1.DisplayLayout.Bands(0)
                .Columns("SEL").Hidden = False
                .Columns("ORDERS").Hidden = True
                .Columns("AMOUNT").Hidden = True
            End With

        ElseIf tabMain.SelectedTab.Key = "Imported Orders" Then

            'Dim ios As System.IO.Stream
            Dim iosf As String = ASCMAIN1.Folders("Temp") & "ISI" & XNO & ".xml"
            grdEDT850T1.DisplayLayout.Save(iosf)
            grdEDT850T1.DataSource = Nothing
            grdEDT850T1.DataSource = dst.Tables("EDT850T1_IMPORTED")
            grdEDT850T1.DisplayLayout.Load(iosf)
            ASCMAIN1.grdInitializeLayout(grdEDT850T1)

            'grdEDT850T1.DataSource = dst.Tables("EDT850T1_IMPORTED")
            grdEDT850T1.Parent = tabMain.Tabs("Imported Orders").TabPage
            With grdEDT850T1.DisplayLayout.Bands(0)
                .Columns("SEL").Hidden = True
                .Columns("ORDERS").Hidden = False
                .Columns("AMOUNT").Hidden = False
            End With

        ElseIf tabMain.SelectedTab.Key = "Archived EDI" Then
            dst.Tables("EDT850T1_ARCHIVED").Rows.Clear()

            'Dim ios As System.IO.Stream
            Dim iosf As String = ASCMAIN1.Folders("Temp") & "ISI" & XNO & ".xml"
            grdEDT850T1.DisplayLayout.Save(iosf)
            grdEDT850T1.DataSource = Nothing
            grdEDT850T1.DataSource = dst.Tables("EDT850T1_ARCHIVED")
            grdEDT850T1.DisplayLayout.Load(iosf)
            ASCMAIN1.grdInitializeLayout(grdEDT850T1)

            ' grdEDT850T1.DataSource = dst.Tables("EDT850T1_ARCHIVED")
            grdEDT850T1.Parent = tabMain.Tabs("Archived EDI").TabPage
            With grdEDT850T1.DisplayLayout.Bands(0)
                .Columns("SEL").Hidden = False
                .Columns("ORDERS").Hidden = True
                .Columns("AMOUNT").Hidden = True
            End With
        End If

        UltraExplorerBar1.Groups("Screen Control").Visible = (tabMain.SelectedTab.Key = "Orders Waiting to be Imported")
        UltraExplorerBar1.Groups("Action").Visible = (tabMain.SelectedTab.Key = "Orders Waiting to be Imported")
        UltraExplorerBar1.Groups("EDIs Already Imported").Visible = (tabMain.SelectedTab.Key = "Archived EDI")
    End Sub

    Private Sub grdEDT850TE_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDT850TE.InitializeRow
        'If e.Row.Cells("EDI_ACTION").Value & "" = "A" Then
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Orange
        'ElseIf e.Row.Cells("EDI_ACTION").Value & "" = "E" Then
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Blue
        'ElseIf e.Row.Cells("EDI_ACTION").Value & "" = "R" Then
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
        'ElseIf e.Row.Cells("EDI_ACTION").Value & "" = "S" Then
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Pink
        'Else
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Empty
        'End If

        Dim EDI_ACTION As String = e.Row.Cells("EDI_ACTION").Value & ""
        Dim EDI_COND_CODE As String = e.Row.Cells("EDI_COND_CODE").Value & ""

        e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Empty

        If EDI_ACTION = "A" Then
            If Not ACTIONS_A.Contains("*") And Not ACTIONS_A.Contains(EDI_COND_CODE) Then
                e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
            End If
        ElseIf EDI_ACTION = "E" Then
            If Not ACTIONS_E.Contains("*") And Not ACTIONS_E.Contains(EDI_COND_CODE) Then
                e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
            End If
        ElseIf EDI_ACTION = "S" Then
            If Not ACTIONS_S.Contains("*") And Not ACTIONS_S.Contains(EDI_COND_CODE) Then
                e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
            End If
        ElseIf EDI_ACTION = "R" Then
            If Not ACTIONS_R.Contains("*") And Not ACTIONS_R.Contains(EDI_COND_CODE) Then
                e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
            End If
        End If

        If e.Row.Cells("RESOLUTION").Value & "" <> "" Then
            e.Row.Cells("EDI_COND_DESC").ToolTipText = e.Row.Cells("RESOLUTION").Value
        End If
    End Sub

    Private Sub grdEDT850T1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDT850T1.InitializeRow
        If e.Row.Cells("RESULT").Value & "" = "Rejected" Then
            e.Row.Cells("RESULT").Appearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("RESULT").Value & "" = "Imported" Then
            e.Row.Cells("RESULT").Appearance.ForeColor = Drawing.Color.Blue
        Else
            e.Row.Cells("RESULT").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Setup_RESOLUTIONS()
        RESOLUTIONS.Add("01", "Trading Partner Setup Required")
        RESOLUTIONS.Add("27", "Check other unprocessed EDI Orders for Same TP ID and Same Customer PO")
        RESOLUTIONS.Add("28", "Chec this PO for Duplicate Items in Order Detail")
        RESOLUTIONS.Add("04", "Customer Master Setup Required")
        RESOLUTIONS.Add("35", "Customer Master Setup Required")
        RESOLUTIONS.Add("36", "Customer Master Setup Required")
        RESOLUTIONS.Add("03", "Customer Master Setup Required")
        RESOLUTIONS.Add("91", "Customer Master Setup Required")
        RESOLUTIONS.Add("09", "Accept EDI Data, Get an Extension from Customer, and Change Cancel Date in Sales Order Entry")
        RESOLUTIONS.Add("08", "Customer must re-transmit with a Ship Date")
        RESOLUTIONS.Add("21", "Customer Master Setup Required")
        RESOLUTIONS.Add("11", "Map EDI Terms to a Valid AR Terms Code")
        RESOLUTIONS.Add("12", "Map EDI Terms to a Valid AR Terms Code")
        RESOLUTIONS.Add("13", "Customer Master Setup Required")
        RESOLUTIONS.Add("14", "Check Customer Order History for this PO")
        RESOLUTIONS.Add("34", "Item Master Setup Required")

        RESOLUTIONS.Add("63", "EDI Mapping issue - Call ABS")
        RESOLUTIONS.Add("64", "EDI Mapping issue - Call ABS")
        RESOLUTIONS.Add("65", "EDI Mapping issue - Call ABS")
        RESOLUTIONS.Add("66", "EDI Mapping issue - Call ABS")

        ACTIONS_A.Add("27")
        ACTIONS_A.Add("28")
        ACTIONS_A.Add("09")
        ACTIONS_A.Add("40")
        ACTIONS_A.Add("17")
        ACTIONS_A.Add("38")
        ACTIONS_A.Add("14")
        ACTIONS_A.Add("04")
        ACTIONS_A.Add("06")
        ACTIONS_A.Add("33")
        ACTIONS_A.Add("13")
        ACTIONS_A.Add("55")
        ACTIONS_A.Add("56")
        ACTIONS_A.Add("39")
        ACTIONS_A.Add("50")
        ACTIONS_A.Add("68")

        ACTIONS_S.Add("15")
        ACTIONS_S.Add("17")
        ACTIONS_S.Add("38")
        ACTIONS_S.Add("05")
        ACTIONS_S.Add("23")
        ACTIONS_S.Add("24")
        ACTIONS_S.Add("25")
        ACTIONS_S.Add("30")
        ACTIONS_S.Add("31")
        ACTIONS_S.Add("32")


        ACTIONS_R.Add("15")
        If ASCMAIN1.Running_in_VS Then
            ACTIONS_R.Add("*")
        End If

        ACTIONS_E.Add("*")
        'If ASCMAIN1.CLIENT = "NYA" Then
        '    ACTIONS_E.Add("55") ' NEEDED FOR LOBLAW
        'End If

    End Sub

    Private Sub tabEDT850T1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabEDT850T1.SelectedTabChanged
        If tabEDT850T1.SelectedTab.Key = "Raw EDI Data" Then
            If grdEDT850T1.ActiveRow IsNot Nothing AndAlso grdEDT850T1.ActiveRow.IsDataRow Then
                Dim EDI_DOC_SEQ_NO As String = grdEDT850T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value
                txtRaw.Text = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"))
            End If
        End If
    End Sub

    Private Sub cmdFetchEDIs_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetchEDIs.Click
        If txtCUST_CODE.Text = "" Then
            MsgBox("You Must First Select a Customer and a Date Range", MsgBoxStyle.OkOnly, "Cannot Fetch")
            Exit Sub
        End If

        Load_EDT850T1_ARCHIVED()

    End Sub

    Private Sub cmdReQueue_Click(sender As System.Object, e As System.EventArgs) Handles cmdReQueue.Click
        If dst.Tables("EDT850T1_ARCHIVED").Select("SEL = '1'").Length = 0 Then
            MsgBox("Nothing Selected", MsgBoxStyle.OkOnly, "Cannot Re-Queue")
            Exit Sub
        End If

        Dim I As Integer = 0
        For Each rowEDT850T1_ARCHIVED As DataRow In dst.Tables("EDT850T1_ARCHIVED").Select("SEL = '1'")
            Dim EDI_DOC_SEQ_NO As String = rowEDT850T1_ARCHIVED.Item("EDI_DOC_SEQ_NO")
            ASCDATA1.ExecuteSQL("Update EDT850T1 Set EDI_PROCESS_IND = '0' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
            I += 1
        Next

        MsgBox(CStr(I) & " EDI Documents have been Re-Queued", MsgBoxStyle.OkOnly, "Verification")

        tabMain.SelectedTab = tabMain.Tabs("Orders Waiting to be Imported")
        Load_EDT850T1()

    End Sub

    Sub Load_EDT850T1_ARCHIVED()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Data")

        Fill_Records("EDT850T1_ARCHIVED", New Object() {txtCUST_CODE.Text, dteFrom.Value, dteTo.Value})

        For Each row As DataRow In dst.Tables("EDT850T1_ARCHIVED").Select("")
            row.Item("SEL") = "0"
        Next
        Sort_grdColumns(grdEDT850T1, "EDI_DOC_SEQ_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdEDT850TE_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDT850TE.InitializeLayout

    End Sub

    Private Sub grdEDT850TE_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT850TE.DoubleClickRow
        For Each grow As UltraWinGrid.UltraGridRow In grdEDT850T1.Rows
            If grow.Cells("EDI_DOC_SEQ_NO").Value = e.Row.Cells("EDI_DOC_SEQ_NO").Value Then
                grdEDT850T1.ActiveRow = grow
                grow.Selected = True
                Exit For
            End If
        Next
        tabMain.SelectedTab = tabMain.Tabs("Orders Waiting to be Imported")
    End Sub

    Private Sub grdEDT850T1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles grdEDT850T1.KeyDown
        If e.KeyCode = Keys.Delete Then
            If grdEDT850T1.ActiveCell IsNot Nothing AndAlso grdEDT850T1.ActiveCell.Column.Key = "WHSE_CODE" AndAlso grdEDT850T1.ActiveCell.Value & "" <> "" Then
                grdEDT850T1.ActiveCell.Value = DBNull.Value
            End If
        End If
    End Sub

    Sub Release_ECOM_Orders()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Releasing E-Commerce Orders")

        ' Create Pick Tickets, Shipment BOL, and Cartons
        '   do this for all ecommerce orders imported using EDI

        Dim PICK_RELEASED As Date = DATETIME_STAMP

        ASCMAIN1.sql = "Select ORDR_NO, CUST_CODE, ORDR_GROUP_NO from SOTORDR1" & vbCrLf _
            & " where ORDR_STATUS = 'O' and ECOM_CODE is Not Null and WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where LP_CODE is Not Null)"
        ' PROBABLY SHOULD RELEASE ALL ECOM ORDERS WITH A 3PL WHSE
        ' ASCMAIN1.sql = "Select ORDR_NO, CUST_CODE, ORDR_GROUP_NO from SOTORDR1 where ORDR_STATUS = 'O' and ECOM_CODE is Not Null and WHSE_CODE = 'MS'"
        ' NEED A MORE DEFINITIVE WAY TO IDENTIFY ECOMMERCE ORDERS SHIPPED BY ECOMMERCE PARTNER
        Dim SOTORDR1_ECOM As String = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Select SOTORDR1.*, DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST,'DC',SOTORDR1.CUST_DC_NO,SOTORDR1.CUST_STORE_NO) SHIP_TO" & vbCrLf _
            & " from SOTORDR1, " & SOTORDR1_ECOM & " SOTORDR1_ECOM where SOTORDR1.ORDR_NO = SOTORDR1_ECOM.ORDR_NO"
        Fill_Records("SOTORDR1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
            & " from SOTORDR2, " & SOTORDR1_ECOM & " SOTORDR1_ECOM where SOTORDR2.ORDR_NO = SOTORDR1_ECOM.ORDR_NO"
        Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select ICTSTYL1.*" & vbCrLf _
            & " from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from SOTORDR2, " & SOTORDR1_ECOM & " SOTORDR1_ECOM where SOTORDR2.ORDR_NO = SOTORDR1_ECOM.ORDR_NO)"
        Fill_Records("ICTSTYL1", "", True, ASCMAIN1.sql)

        If dst.Tables("SOTORDR1").Rows.Count = 0 Then Exit Sub

        BeginTrans()

        Dim PICK_BATCH_NO As String = ""

        If ASCMAIN1.CLIENT = "VAN" Then
            PICK_BATCH_NO = ASCMAIN1.Next_Control_No("PICK_BATCH_NO")
        Else
            PICK_BATCH_NO = ASCMAIN1.Next_Control_No("SOTPICK0.PICK_BATCH_NO")
        End If

        Dim rowARTCUST1 As DataRow = Nothing
        Dim CUST_CODE As String = ""
        For Each rowSOTORDR1_rel As DataRow In dst.Tables("SOTORDR1").Select("", "CUST_CODE, ORDR_GROUP_NO, ORDR_ADDR_TYPE_ST, SHIP_TO, ORDR_NO")
            If CUST_CODE <> rowSOTORDR1_rel.Item("CUST_CODE") Then
                CUST_CODE = rowSOTORDR1_rel.Item("CUST_CODE")
                rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
            End If

            Dim ORDR_NO As String = rowSOTORDR1_rel.Item("ORDR_NO")
            Dim PICK_SEQ_NO As Integer = Val(rowSOTORDR1_rel.Item("ORDR_PICK_SEQ") & "") + 1
            rowSOTORDR1_rel.Item("ORDR_PICK_SEQ") = PICK_SEQ_NO
            rowSOTORDR1_rel.Item("ORDR_STATUS") = "P"

            Dim PICK_NO As String = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO") ' NG FOR VAN
            Dim SHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO") ' NG FOR VAN

            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
            With rowSOTPICK1
                .Item("PICK_NO") = PICK_NO
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_PICK_SEQ") = PICK_SEQ_NO
                .Item("PICK_STATUS") = "P"
                .Item("PICK_BATCH_NO") = PICK_BATCH_NO
                .Item("PICK_RELEASED") = PICK_RELEASED
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("CCPA_NO_STATUS") = "0"
            End With
            dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)

            Dim CART_NO As String = ASCMAIN1.Next_Control_No("SOTCART1.CART_NO") ' NG FOR VAN

            Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
            rowSOTCART1.Item("CART_NO") = CART_NO
            rowSOTCART1.Item("PICK_NO") = PICK_NO
            rowSOTCART1.Item("CART_TOTAL_UNITS") = 0
            rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = 0
            dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

            Dim CART_LNO_seq As Integer = 0

            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_QTY_OPEN <> 0"
            For Each rowSOTORDR2_rel As DataRow In dst.Tables("SOTORDR2").Select(sqlw, "ORDR_LNO")
                Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
                With rowSOTPICK2
                    .Item("PICK_NO") = PICK_NO
                    .Item("PICK_LNO") = rowSOTORDR2_rel.Item("ORDR_LNO")
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = rowSOTORDR2_rel.Item("ORDR_LNO")

                    rowICTSTYL1 = dst.Tables("ICTSTYL1").Rows.Find(rowSOTORDR2_rel.Item("STYLE_CODE"))

                    .Item("STYLE_CODE") = rowSOTORDR2_rel.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = rowSOTORDR2_rel.Item("COLOR_CODE")
                    .Item("PICK_UNIT_PRICE") = rowSOTORDR2_rel.Item("ORDR_UNIT_PRICE")

                    Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") & "")

                    rowSOTORDR2_rel.Item("ORDR_QTY_PICK") = Val(rowSOTORDR2_rel.Item("ORDR_QTY_PICK") & "") + ORDR_QTY_OPEN
                    rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") = 0
                    .Item("PICK_QTY") = ORDR_QTY_OPEN

                    'rowSOTORDR2_rel.Item("ORDR_QTY_ALLO_CUR") = 0


                    CART_LNO_seq = CART_LNO_seq + 1
                    Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                    With rowSOTCART2
                        .Item("CART_NO") = CART_NO
                        .Item("CART_LNO") = CART_LNO_seq
                        .Item("ORDR_NO") = ORDR_NO
                        .Item("ORDR_LNO") = rowSOTORDR2_rel.Item("ORDR_LNO")
                        .Item("STYLE_CODE") = rowSOTORDR2_rel.Item("STYLE_CODE")
                        .Item("COLOR_CODE") = rowSOTORDR2_rel.Item("COLOR_CODE")
                        .Item("QTY_PACKED") = ORDR_QTY_OPEN
                        .Item("STYLE_WEIGHT") = rowICTSTYL1.Item("STYLE_WEIGHT")
                    End With
                    dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)


                End With
                dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
            Next

            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
            With rowSOTSHIP1
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO

                For Each COLUMN_NAME As String In New String() _
                    {"SHIP_VIA_CODE", "ORDR_GROUP_NO", "TERM_CODE", "SREP_CODE", "FRT_TERMS", "WHSE_CODE", "ORDR_DEPT"}
                    .Item(COLUMN_NAME) = rowSOTORDR1_rel.Item(COLUMN_NAME)
                Next

                .Item("SHIP_ADDR_TYPE") = rowSOTORDR1_rel.Item("ORDR_ADDR_TYPE_ST")
                .Item("SHIP_ADDR_CODE") = rowSOTORDR1_rel.Item("SHIP_TO")
                .Item("SHIP_STATUS") = "P"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LP_STATUS") = "1" ' ECOM PARTNER IS ACTING AS A 3PL - SO MANUAL CONFIRMATION IS NOT PERMITTED
                .Item("ORDR_PICK_TYPE") = "P" ' PICK AND PACK
                .Item("SHIP_CART_REQD") = rowARTCUST1.Item("CUST_CART_REQD")

                .Item("CUST_FACTOR_TRANS_IND") = rowARTCUST1.Item("CUST_FACTOR_IND")
                .Item("SHIP_856_IND") = "0" ' SINCE ECOM PARTNER IS SHIPPING, WE DO NOT SEND AN 856
                .Item("SHIP_810_IND") = "1" ' SINCE ECOM PARTNER IS SHIPPING, WE WILL INVOICE THE PARTNER USING THE CONTRACTED PRICE

            End With
            dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)
        Next

        'dst.Tables("SOTCART2").Columns.Add("WGT", GetType(System.Decimal), "ISNULL(QTY_PACKED,0) * ISNULL(STYLE_WEIGHT,0)")
        'dst.Tables("SOTCART1").Columns.Add("QTY", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).QTY_PACKED)")
        'dst.Tables("SOTCART1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).WGT)")
        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
            rowSOTCART1.Item("CART_TOTAL_UNITS") = rowSOTCART1.Item("QTY")
            rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = rowSOTCART1.Item("WGT")
        Next

        'dst.Tables("SOTPICK1").Columns.Add("CTNS", GetType(System.Int64), "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
        'dst.Tables("SOTPICK1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_CALC)")
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = rowSOTPICK1.Item("CTNS")
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = rowSOTPICK1.Item("WGT")
        Next

        Dim rowSOTPICK0 As DataRow = dst.Tables("SOTPICK0").NewRow
        With rowSOTPICK0
            .Item("PICK_BATCH_NO") = PICK_BATCH_NO
            .Item("PICK_SHPS") = dst.Tables("SOTSHIP1").Rows.Count
            .Item("PICK_CTNS") = dst.Tables("SOTCART1").Rows.Count
            .Item("PICK_PKTS") = dst.Tables("SOTPICK1").Rows.Count
            .Item("PICK_BATCH_STATUS") = "O"
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("PICK_SHIP_REL_DATE") = Now.Date
            .Item("PICK_FORCED") = "0"
        End With
        dst.Tables("SOTPICK0").Rows.Add(rowSOTPICK0)


        Update_Record_TDA("SOTORDR1")
        Update_Record_TDA("SOTORDR2")
        Update_Record_TDA("SOTPICK0")
        Update_Record_TDA("SOTPICK1")
        Update_Record_TDA("SOTPICK2")
        Update_Record_TDA("SOTSHIP1")
        Update_Record_TDA("SOTCART1")
        Update_Record_TDA("SOTCART2")

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is " & vbCrLf _
            & " Select SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", Sum (ORDR_QTY_PICK) PICK_QTY" & vbCrLf _
            & " from SOTORDR2,SOTORDR1," & SOTORDR1_ECOM & " SOTORDR1_ECOM " & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR1_ECOM.ORDR_NO" & vbCrLf _
            & " group by SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   ICPSTAT2(R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,0,0,0,-R1.PICK_QTY,R1.PICK_QTY,0);" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select ORDR_GROUP_NO from " & SOTORDR1_ECOM & ";" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   SOPORDR0_G(R1.ORDR_GROUP_NO); " & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        CommitTrans()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
