Imports Infragistics.UltraChart.Resources

Imports Infragistics.Win

Public Class SOFCORD1

    ' charts
    ' invoice history grid
    ' tracking links
    ' sales division code and radio buttons for status
    ' ORIG FORM USED SALES_DIVISION_CODE TO CULL DATA IN DETAILS, LIKE SOTPICK2, SOTORDR2
    'right click on sotpick1 to show invoice or edi
    ' ASSUMING THAT ORDR_EXTD_COST IS ACCURATE

    ' IF FIND BY THEN SET THE ACTIVE ROW TO THAT WHATEVER

    '       SQL = "Select SUM (SHIP_CNT_CARTONS) SHIP_CNT_CARTONS, "
    '   SQL = SQL & " SUM (SHIP_TOTAL_WGT) SHIP_TOTAL_WGT "
    '   SQL = SQL & " from SOTSHIP1 where ORDR_GROUP_NO = :CODE"
    '   dynSOTSHIPX = OraD.CreateDynaset(SQL, 8&)

    'InStr(UserSecs, "X2") = 0 
    '   grdSOTORDRS.Columns("GP_PCT").Visible = False
    'whr SOFCTRAK

#Region "Declarations"
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim SOTORDR0 As String
    Dim sqlSOTORDR0 As String
    Dim sqlSOTPICK2 As String
    Dim sqlSOTORDRS As String
    Dim sqlSOTRSRVS As String
    Dim ORDR_GROUP_NO As String
    Dim sqlSOTORDRP As String
    Dim sqlSOTORDR1 As String
    Dim SOTORDR0_ALL As String

    Dim sqlSOTORDRT As String
    Dim PO_CARTON_ORIG As String = ""
    Dim PO_CARTON_COMBINED As String = ""

    Dim appRed As New Infragistics.Win.Appearance
    Dim YYMM_exp As String = Format(Now, "yyMM")

    Dim COLUMN_NAMEs_All As New List(Of String)
    Dim COLUMN_NAMEs_Short As New List(Of String)

    Dim cols3PL() As String

    Dim SOTORDRS_BASE As String
    Dim sqlSOTORDRS_BASE As String

    Dim SOFCORD1_LAYOUT_ORIG As String = "SOFCORD1_LAYOUT_ORIG.xml"
    Dim SOFCORD1_LAYOUT_SHORT As String = ""
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("EDTPARM1")

        grpVerifyBL.Visible = (ASCMAIN1.CLIENT = "VAN" And (ASCMAIN1.USER_ID = "naseema" Or ASCMAIN1.USER_ID = "wendy" Or ASCMAIN1.USER_ID = "wjz"))
        appRed.ForeColor = Drawing.Color.Red

        With dst
            sqlSOTORDR0 = "Select 'O' ORDR_TYPE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR0.CUST_DC_NO, SOTORDR0.ORDR_DEPT, EDT850T1.EDI_MERCH_TYPE, SOTORDR0.SALES_DIVISION_CODE, SOTORDR0.ORDR_DATE" & vbCrLf _
                & ", SOTORDR0.ORDR_SHIP_DATE,SOTORDR0. ORDR_CANCEL_DATE, SOTORDR0.ORDR_ORIG_SHIP_DATE, SOTORDR0.ORDR_ORIG_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR0.WHSE_CODE, SOTORDR0.SREP_CODE" & vbCrLf _
                & ", SOTORDR0.ORDR_TYPE_CODE, SOTORDR0.ORDR_SOURCE, SOTORDR0.EDI_DOC_SEQ_NO" & vbCrLf _
                & ", SOTORDR0.ORDR_AMT, SOTORDR0.ORDR_AMT_OPEN, SOTORDR0.ORDR_AMT_PICK, SOTORDR0.ORDR_AMT_SHIP, SOTORDR0.ORDR_AMT_CANC" & vbCrLf _
                & ", SOTORDR0.ORDR_QTY, SOTORDR0.ORDR_QTY_OPEN, SOTORDR0.ORDR_QTY_PICK, SOTORDR0.ORDR_QTY_SHIP, SOTORDR0.ORDR_QTY_CANC" & vbCrLf _
                & ", SOTORDR0.ORDR_CNT, SOTORDR0.ORDR_CNT_OPEN, SOTORDR0.ORDR_CNT_PICK" & vbCrLf _
                & ", SOTORDR0.ORDR_DATE_RECD, SOTORDR0.ORDR_PRIORITY, SOTORDR0.ORDR_ARRIVAL_DATE, SOTORDR0.ORDR_LAST_ARRIVAL_DATE" & vbCrLf _
                & ", SOTORDR0.ORDR_NO_MIN, SOTORDR0.ORDR_NO_MAX, SOTORDR0.ORDR_RELEASE_AVAIL_MIN, SOTORDR0.ORDR_RELEASE_AVAIL_MAX" & vbCrLf _
                & ", SOTORDRG.ORDR_REL_SHORT, SOTORDRG.ORDR_REL_SHORT_OPER, SOTORDRG.ORDR_REL_ACTION_DATE, SOTORDRG.ORDR_REL_ACTION_OPER" & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "VAN", ", EDT850T1.EDI_CONS_NO", ", '0' EDI_CONS_NO") & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "VAN", ", SOTPCKP2.PACK_NO", ", ' ' PACK_NO") & vbCrLf _
                & " from SOTORDR0,EDT850T1,SOTORDRG" & IIf(ASCMAIN1.CLIENT = "VAN", ",SOTPCKP2", "")
            ASCMAIN1.sql = sqlSOTORDR0 & " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTORDR0.CUST_CODE = ''"
            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql &= " and SOTPCKP2.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTPCKP2.PACK_GROUP_STATUS (+) = 'A'"
            End If
            ASCMAIN1.sql = "Select X.*, SOTORDR1.TERM_CODE, SOTORDR1.LAST_DATE, SOTORDR1.LAST_OPER, SOTORDR1.ORDR_SHIP_INSTR, SOTORDR1.ORDR_MESSAGE, ARTCCPA1.CUST_CREDIT_CARD_EXP_DATE, ARTCCPA1.CUST_CREDIT_CARD_LAST4, ARTCUST1.CUST_NAME, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY" & vbCrLf _
                & " from (" & ASCMAIN1.sql & ") X, ARTCUST1,SOTORDR1,ARTCCPA1" _
                & " where ARTCUST1.CUST_CODE = X.CUST_CODE and SOTORDR1.ORDR_NO = X.ORDR_NO_MIN and ARTCCPA1.CCPA_NO (+) = SOTORDR1.CCPA_NO"
            SOTORDR0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add WAVE_NO VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add EDI_LOAD_ID VARCHAR2(20)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_AMT_ALLO_CUR NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_AMT_ALLO_FUT NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_AMT_ALLO_CXL NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_TYPE, ORDR_GROUP_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add EDI_PO_TYPE VARCHAR2(2)")
            ASCMAIN1.sql = "Select * from " & SOTORDR0
            'Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "V", 2)
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "", 2)
            For Each A As String In New String() {"CUR", "FUT", "CXL"}
                .Tables("SOTORDR0").Columns.Add("PCT_ALLO_" & A, GetType(System.Decimal), "IIF(ORDR_AMT=0,0,100*ORDR_AMT_ALLO_" & A & "/ORDR_AMT)")
            Next
            For i As Int64 = 1 To 10
                Dim COLN As String = Format(i, "00")
                .Tables("SOTORDR0").Columns.Add($"COMMENT{COLN}", GetType(System.String))
            Next

            Create_TDA(.Tables.Add, "SOTORDRG", "*")

            Dim TBL As DataTable = .Tables("SOTORDR0").Clone
            TBL.TableName = "SOTCORDG"
            .Tables.Add(TBL)

            ASCMAIN1.sql = "Select ORDR_NO, ORDR_CUST_PO, ORDR_DATE" & vbCrLf _
                & ", ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_ORIG_SHIP_DATE, ORDR_ORIG_CANCEL_DATE" & vbCrLf _
                & ", CUST_STORE_NO, SALES_DIVISION_CODE, ORDR_SOURCE, ORDR_DEPT, ORDR_ADDR_TYPE_ST" & vbCrLf _
                & ", ORDR_STATUS, CUST_STORE_NAME, SREP_CODE, ORDR_PRIORITY, ORDR_HOLD" & vbCrLf _
                & ", ORDR_REL_HOLD_CODES, CUST_DC_NO, ORDR_PRE_ALLOC, WHSE_CODE, EDI_DOC_SEQ_NO, REASON_CODE" & vbCrLf _
                & ", 'O' ORDR_TYPE, ORDR_GROUP_NO, X.ORDR_QTY, X.ORDR_QTY_OPEN, X.ORDR_QTY_PICK, X.ORDR_QTY_SHIP, X.ORDR_QTY_CANC, X.ORDR_AMT, X.ORDR_AMT_OPEN, X.ORDR_AMT_PICK, X.ORDR_AMT_SHIP, X.ORDR_AMT_CANC" & vbCrLf _
                & ", ORDR_MESSAGE, ORDR_INV_COMMENT, ORDR_SHIP_INSTR, EDI_PO_TYPE" & vbCrLf _
                & " from SOTORDR1, (Select ORDR_NO ORDR_NO_DTL, Sum (ORDR_QTY) ORDR_QTY, Sum (ORDR_QTY_OPEN) ORDR_QTY_OPEN, Sum (ORDR_QTY_PICK) ORDR_QTY_PICK, Sum (ORDR_QTY_SHIP) ORDR_QTY_SHIP, Sum (ORDR_QTY_CANC) ORDR_QTY_CANC, Sum (ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT, Sum (ORDR_QTY_OPEN * ORDR_UNIT_PRICE) ORDR_AMT_OPEN, Sum (ORDR_QTY_PICK * ORDR_UNIT_PRICE) ORDR_AMT_PICK, Sum (ORDR_QTY_SHIP * ORDR_UNIT_PRICE) ORDR_AMT_SHIP, Sum (ORDR_QTY_CANC * ORDR_UNIT_PRICE) ORDR_AMT_CANC from SOTORDR2 group by ORDR_NO) X where X.ORDR_NO_DTL = SOTORDR1.ORDR_NO"
            sqlSOTORDR1 = ASCMAIN1.sql
            ASCMAIN1.sql &= vbCrLf & " and ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V", 1)


            Create_TDA(.Tables.Add, "SOTORDR4", "*", 1)


            ASCMAIN1.sql = sqlSOTORDR1
            ASCMAIN1.sql &= vbCrLf & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTORDR1_ALL", "**", 0, False, "", 1)

            sqlSOTORDRS = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC, SOTORDR2.RANGE_STYLE_CODE" & vbCrLf _
                & ", ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY" & vbCrLf _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_ALLO) ORDR_QTY_ALLO" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_EXTD_COST) ORDR_CGS" & vbCrLf _
                & ", MAX (SOTORDR2.ORDR_RELEASE_AVAIL) ORDR_RELEASE_AVAIL" & vbCrLf _
                & " from SOTORDR2,ICTCOLR1,SOTORDR1,ICTSTYL1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC, SOTORDR2.RANGE_STYLE_CODE" & vbCrLf _
                & ", ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY" & vbCrLf _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU"
            ASCMAIN1.sql = Replace(sqlSOTORDRS, " group by ", " and ROWNUM < 1 group by ")
            Create_TDA(.Tables.Add, "SOTORDRS", "**", 0, False, "V", 0)
            With .Tables("SOTORDRS").Columns
                .Add("WIP_IND")
                .Add("ORDR_UNIT_PRICE", GetType(System.Decimal), "IIF(ISNULL(ORDR_QTY,0)=0,0,ISNULL(ORDR_AMT,0) / ISNULL(ORDR_QTY,0))")
                .Add("ORDR_GP", GetType(System.Decimal), "ISNULL(ORDR_AMT,0)-ISNULL(ORDR_CGS,0)")
                .Add("ORDR_GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ORDR_AMT,0)=0,0,100*ORDR_GP/ISNULL(ORDR_AMT,0))")
                '.Add("QTY_XFR_3PL", GetType(System.Int32))
                '.Add("QTY_XIT_3PL", GetType(System.Int32))
                '.Add("QTY_XIT_3PL_ETA", GetType(System.DateTime))
                .Add("QTY_ONH_3PL", GetType(System.Int32))
                .Add("QTY_PCK_3PL", GetType(System.Int32))
                .Add("QTY_OPN_3PL", GetType(System.Int32))
                .Add("QTY_AVA_3PL", GetType(System.Int32))
            End With


            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", WHSE_QTY_ON_HAND QTY_ONH_3PL" & vbCrLf _
                & ", WHSE_QTY_PICK QTY_PCK_3PL" & vbCrLf _
                & ", WHSE_QTY_OPEN QTY_OPN_3PL" & vbCrLf _
                & " from ICTSTAT2" & vbCrLf _
                & " where WHSE_CODE = 'US' and (STYLE_CODE, COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & " from SOTORDR2,SOTORDR1 where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO and SOTORDR1.ORDR_GROUP_NO = :PARM1)"
            Create_TDA(.Tables.Add, "SOTORDRS_3PL", "**", 0, False, "V", 2)
            With .Tables("SOTORDRS_3PL").Columns
                .Add("QTY_AVA_3PL", GetType(System.Int32), "ISNULL(QTY_ONH_3PL,0) - ISNULL(QTY_PCK_3PL,0)- ISNULL(QTY_OPN_3PL,0)")
            End With


            ASCMAIN1.sql = Get_SQL_SOTORDRS_BASE(True)
            SOTORDRS_BASE = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = Get_SQL_SOTORDRS_ORDERS()
            Create_TDA(.Tables.Add, "SOTORDRS_ORDERS", "**", 0, False, "", 1)
            With .Tables("SOTORDRS_ORDERS").Columns
                Dim XXQTY As String = "(ISNULL(CUR_QTY_ONHD,0)+ISNULL(FUT_QTY_ONPO,0)+ISNULL(FUT_QTY_ONPO,0)+ISNULL(UNAV_QTY,0))"
                Dim XXAMT As String = Replace(XXQTY, "QTY", "AMT")
                .Add("UNAV_QTY_PCT", GetType(System.Decimal), $"IIF({XXQTY}=0,0,100*ISNULL(UNAV_QTY,0)/{XXQTY})")
                .Add("UNAV_AMT_PCT", GetType(System.Decimal), $"IIF({XXAMT}=0,0,100*ISNULL(UNAV_QTY,0)/{XXAMT})")
            End With

            ASCMAIN1.sql = Get_SQL_SOTORDRS_SC()
            Create_TDA(.Tables.Add, "SOTORDRS_SC", "**", 0, False, "", 2)
            With .Tables("SOTORDRS_SC").Columns
                Dim XXQTY As String = "(ISNULL(CUR_QTY_ONHD,0)+ISNULL(FUT_QTY_ONPO,0)+ISNULL(FUT_QTY_ONPO,0)+ISNULL(UNAV_QTY,0))"
                Dim XXAMT As String = Replace(XXQTY, "QTY", "AMT")
                .Add("UNAV_QTY_PCT", GetType(System.Decimal), $"IIF({XXQTY}=0,0,100*ISNULL(UNAV_QTY,0)/{XXQTY})")
                .Add("UNAV_AMT_PCT", GetType(System.Decimal), $"IIF({XXAMT}=0,0,100*ISNULL(UNAV_QTY,0)/{XXAMT})")
                .Add("SEL", GetType(System.String))
            End With

            sqlSOTRSRVS = "Select SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE" & vbCrLf _
                 & ", ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, NULL RANGE_STYLE_CODE" & vbCrLf _
                 & ", NULL CUST_STYLE_CODE, NULL CUST_COLOR_CODE, NULL CUST_SIZE_CODE, NULL CUST_UPC, NULL CUST_SKU" & vbCrLf _
                 & ", SUM (SOTRSRV2.RSRV_QTY) ORDR_QTY" & vbCrLf _
                 & ", SUM (SOTRSRV2.RSRV_QTY * ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
                 & ", SUM (SOTRSRV2.RSRV_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                 & ", SUM (SOTRSRV2.RSRV_QTY_ALLO) ORDR_QTY_ALLO" & vbCrLf _
                 & ", SUM (0) ORDR_QTY_PICK" & vbCrLf _
                 & ", SUM (SOTRSRV2.RSRV_QTY_USED) ORDR_QTY_SHIP" & vbCrLf _
                 & ", SUM (SOTRSRV2.RSRV_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                 & ", SUM (0) ORDR_CGS" & vbCrLf _
                 & ", NULL ORDR_RELEASE_AVAIL" & vbCrLf _
                 & " from SOTRSRV2,ICTCOLR1,SOTRSRV1,ICTSTYL1 " & vbCrLf _
                 & " where ICTCOLR1.COLOR_CODE (+) = SOTRSRV2.COLOR_CODE" & vbCrLf _
                 & "   and SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO" & vbCrLf _
                 & "   and ICTSTYL1.STYLE_CODE = SOTRSRV2.STYLE_CODE" & vbCrLf _
                 & " group by SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE" & vbCrLf _
                 & ", ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf

            sqlSOTORDRP = "Select SOTORDR9.RANGE_STYLE_CODE" & vbCrLf _
                & ", SOTORDR1.CUST_STORE_NO" & vbCrLf _
                & ", SOTORDR9.RANGE_STYLE_DESC, SOTORDR9.RANGE_STYLE_PP_PRICE" & vbCrLf _
                & ", SOTORDR9.RANGE_STYLE_PRICE" & vbCrLf _
                & ", SUM (SOTORDR9.RANGE_STYLE_PP_QTY) RANGE_STYLE_PP_QTY" & vbCrLf _
                & ", SUM (SOTORDR9.RANGE_STYLE_QTY) RANGE_STYLE_QTY" & vbCrLf _
                & " from SOTORDR9, SOTORDR1" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTORDR9.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = :PARM1" & vbCrLf _
                & "   and RANGE_STYLE_PP = '1'" & vbCrLf _
                & " group by SOTORDR9.RANGE_STYLE_CODE" & vbCrLf _
                & ", SOTORDR1.CUST_STORE_NO" & vbCrLf _
                & ", SOTORDR9.RANGE_STYLE_DESC, SOTORDR9.RANGE_STYLE_PP_PRICE" & vbCrLf _
                & ", SOTORDR9.RANGE_STYLE_PRICE"
            ASCMAIN1.sql = sqlSOTORDRP
            Create_TDA(.Tables.Add, "SOTORDRP", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK1.PICK_NO, SOTORDR1.CUST_STORE_NO, SOTPICK1.ORDR_NO" & vbCrLf _
                & ", SOTPICK1.PICK_STATUS, SOTPICK1.PICK_RELEASED, SOTPICK1.PICK_FREIGHT" & vbCrLf _
                & ", SOTPICK1.PICK_PICKER, SOTPICK1.PICK_NO_REV" & vbCrLf _
                & ", SOTPICK1.PICK_PRINTED, SOTPICK1.PICK_PACKED, SOTPICK1.PICK_SHIPPED" & vbCrLf _
                & ", SOTPICK1.PICK_BATCH_NO, SOTPICK1.SHIP_BOL_NO, SOTPICK1.INV_NO" & vbCrLf _
                & ", SOTPICK1.PICK_CNT_CARTONS, SOTPICK1.PICK_TOTAL_WGT" & vbCrLf _
                & ", SOTPICK1.INIT_OPER, SOTPICK1.LAST_OPER, SOTPICK1.INIT_DATE, SOTPICK1.LAST_DATE" & vbCrLf _
                & ", SOTPICK0.PICK_FORCED" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTPICK0 " & vbCrLf _
                & " where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTPICK0.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTPICK1.SHIP_BOL_NO, EDT850T2.EDI_COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" & vbCrLf _
                & " from SOTPICK1,SOTPICK2,SOTORDR2,ICTCOLR1,EDT850T2" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and EDT850T2.EDI_DOC_SEQ_NO (+) = SOTORDR2.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and EDT850T2.EDI_DTL_SEQ (+) = SOTORDR2.EDI_DTL_SEQ" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "   and SOTPICK2.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")

            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO,SOTSHIP1.SHIP_DATE_SHIPPED,SOTSHIP1.SHIP_VIA_CODE,SOTSHIP1.SHIP_REF" & vbCrLf _
                & ",SOTSHIP1.SHIP_TOTAL_WGT,SOTSHIP1.SHIP_CNT_CARTONS,SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
                & ",SOTSHIP1.SHIP_PICK_PRINTED,SOTSHIP1.PICK_BATCH_NO,SOTSHIP1.SHIP_STATUS,SOTSHIP1.LP_STATUS" & vbCrLf _
                & ",SOTSHIP1.BILL_OF_LADING_NO,SOTSHIP1.FRT_TERMS,SOTSHIP1.SHIP_PULL_BY_STYLE" & vbCrLf _
                & ",SOTSHIP1.SHIP_856_BATCH_NO,SOTSHIP1.SHIP_810_BATCH_NO,SOTSHIP1.WHSE_CODE,SOTSHIP1.INV_DATE,SOTSHIP1.SHIP_MANIFEST_NO" & vbCrLf _
                & ",SOTSHIP1.SHIP_BOL_NO_REV,SOTSHIP1.SHIP_NOTES,SOTSHIP1.SHIPPED_ACTUAL,SOTSHIP1.SHIP_SEAL_NO" & vbCrLf _
                & ",SOTSHIP1.SHIP_BOL_NO_ORIG,SOTSHIP1.SHIP_BOL_NO_SPLIT,SOTSHIP1.BOL_PRINTED,SOTSHIP1.SHIP_SPEC_INST" & vbCrLf _
                & ",SOTSHIP1.MASTER_SHIP_BOL_NO,SOTSHIP1.SHIP_940_BATCH_NO,SOTSHIP1.SHIP_753_IND,SOTSHIP1.SHIP_DATE_PACKED" & vbCrLf _
                & ",SOTSHIP1.INIT_DATE,SOTSHIP1.INIT_OPER,SOTSHIP1.SHIP_LOAD_NO,SOTSHIP1.SHIP_APPT_NO,SOTSHIP1.SHIP_WAVE_STATUS" & vbCrLf _
                & " from SOTSHIP1" & vbCrLf _
                & " where ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, False, "V", 1)


            ASCMAIN1.sql = "Select SOTPICK1.SHIP_BOL_NO" _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
                & ", EDT850T2.EDI_COLOR_CODE" _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" _
                & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" _
                & ", SUM (SOTPICK2.PICK_QTY * SOTPICK2.PICK_UNIT_PRICE) PICK_AMT" _
                & ", SUM (SOTPICK2.PICK_QTY_CONF) PICK_QTY_CONF" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC) PICK_QTY_CANC" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK) PICK_QTY_BACK" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK_REL) PICK_QTY_BACK_REL" _
                & " from SOTPICK1,SOTPICK2,SOTORDR2,ICTCOLR1,EDT850T2" _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
                & "   and EDT850T2.EDI_DOC_SEQ_NO (+) = SOTORDR2.EDI_DOC_SEQ_NO" _
                & "   and EDT850T2.EDI_DTL_SEQ (+) = SOTORDR2.EDI_DTL_SEQ" _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1" _
                & " group by SOTPICK1.SHIP_BOL_NO" _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
                & ", EDT850T2.EDI_COLOR_CODE" _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU"
            Create_TDA(.Tables.Add, "SOTSHIP2", "**", 0, False, "V", 0)
            .Tables("SOTSHIP2").Columns.Add("PICK_UNIT_PRICE", GetType(System.Decimal), "IIF(PICK_QTY=0,0,PICK_AMT/PICK_QTY)")
            With .Tables("SOTSHIP2")
                .Columns("PICK_QTY").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CONF").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CANC").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_BACK").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CANC_REL").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_BACK_REL").DataType = GetType(System.Int64)
            End With

            Create_Relation("SOTSHIP1", "SOTSHIP2", "SHIP_BOL_NO")

            With .Tables.Add("SOTCORDR")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("STYLE_DESC")
                .Columns.Add("CUST_STYLE_CODE")
                .Columns.Add("CUST_COLOR_CODE")
                .Columns.Add("CUST_SIZE_CODE")
                .Columns.Add("CUST_UPC")
                .Columns.Add("CUST_SKU")
                .Columns.Add("ORDR", GetType(System.Int64))
                .Columns.Add("OPEN", GetType(System.Int64))
                .Columns.Add("PICK", GetType(System.Int64))
                .Columns.Add("SHIP", GetType(System.Int64))
                .Columns.Add("CANC", GetType(System.Int64))
                .Columns.Add("ALLO", GetType(System.Int64))
                .Columns.Add("ORDR_AMT", GetType(System.Decimal))


                .Columns("ORDR").DefaultValue = 0
                .Columns("OPEN").DefaultValue = 0
                .Columns("PICK").DefaultValue = 0
                .Columns("SHIP").DefaultValue = 0
                .Columns("CANC").DefaultValue = 0
                .Columns("ALLO").DefaultValue = 0
                .Columns("ORDR_AMT").DefaultValue = 0

                .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With

            With .Tables.Add("SOTORDRX")
                .Columns.Add("ORDR_NO")
                .Columns.Add("CUST_STORE_NO")
            End With

            With .Tables.Add("SOTORDRM")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("QTY")
            End With

            With .Tables.Add("SOTCORDX")
                .Columns.Add("SORT_SEQ")
                .Columns.Add("CODE_VALUE")
                Dim TOT As String = ""
                Dim YTD As String = ""
                For I As Integer = 1 To 12
                    Dim TC As String = "V" & Format(I, "00")
                    .Columns.Add(TC, GetType(System.Decimal))
                    TOT &= "+ISNULL(" & TC & ",0)"
                    If I > 12 - Val(Mid(ASCMAIN1.CYP, 5, 2)) Then
                        YTD &= "+ISNULL(" & TC & ",0)"
                    End If
                Next
                .Columns.Add("TOT", GetType(System.Decimal), TOT)
                .Columns.Add("YTD", GetType(System.Decimal), YTD)
                .Columns.Add("TOTPCT", GetType(System.Decimal), "TOT / 1")
                .Columns.Add("YTDPCT", GetType(System.Decimal), "YTD / 1")
                .PrimaryKey = New DataColumn() { .Columns("SORT_SEQ")}
            End With

            Dim T As DataTable = .Tables("SOTCORDX").Clone
            T.TableName = "SOTCORDD"
            T.PrimaryKey = New DataColumn() {T.Columns("SORT_SEQ"), T.Columns("CODE_VALUE")}
            .Tables.Add(T)

            ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
                & " from SOTINVH1 " & vbCrLf _
                & " where SOTINVH1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "V", 2)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            ASCMAIN1.sql = "SELECT SOTINVH2.INV_NO, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" _
                & ", SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_QTY_SHIP" _
                & ", SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO, SOTINVH1.CUST_STORE_NO, SOTINVH1.WHSE_CODE" _
                & " from SOTINVH2,SOTINVH1,SOTORDR1" _
                & " where SOTINVH2.CUST_CODE = :PARM1" _
                & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & "   and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" _
                & "   and SOTINVH2.STYLE_CODE = :PARM2 AND SOTINVH2.ORDR_YYYYPP_UPDATED = :PARM3"
            Create_TDA(.Tables.Add, "SOTCORDY", "**", 0, False, "VVV", 0)
            .Tables("SOTCORDY").Columns.Add("AMT", GetType(System.Decimal), "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")

            ASCMAIN1.sql = "Select SOTCART1.*,SOTPICK1.SHIP_BOL_NO,SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE,SOTORDR1.CUST_STORE_NO,SOTPICK1.PICK_STATUS, SOTORDR1.CUST_DC_NO" _
                & " from SOTCART1,SOTPICK1,SOTSHIP1,SOTORDR1" _
                & " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO" _
                & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
               & "   and SOTSHIP1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

            If Not .Tables("SOTCART1").Columns.Contains("PALLET_NO") Then
                .Tables("SOTCART1").Columns.Add("PALLET_NO")
            End If
            .Tables("SOTCART1").Columns.Add("STYLES", GetType(System.Int32))
            .Tables("SOTCART1").Columns.Add("QTY_PACKED", GetType(System.Int32))
            .Tables("SOTCART1").Columns.Add("SHIP_TRAILER_NO")
            .Tables("SOTCART1").Columns.Add("PALLET_INIT_DATE", GetType(System.DateTime))
            .Tables("SOTCART1").Columns.Add("PALLET_INIT_OPER")
            .Tables("SOTCART1").Columns.Add("QTY_RFID", GetType(System.Int32))

            'SOTCARTP is a copy of Cart1, but w/o the detail relation to Cart2
            ASCMAIN1.sql = "Select SOTCART1.*,SOTPICK1.SHIP_BOL_NO,SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE,SOTORDR1.CUST_STORE_NO,SOTPICK1.PICK_STATUS, SOTORDR1.CUST_DC_NO" _
                & " from SOTCART1,SOTPICK1,SOTSHIP1,SOTORDR1" _
                & " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO" _
                & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
                & "   and SOTSHIP1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCARTP", "**", 0, False, "V", 1)

            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql = "Select SOTCARM1.CART_NO,SOTCARM1.CART_FREIGHT,SOTCARM1.CART_PACKER,SOTCARM1.CART_PACKED,SOTCARM1.CART_SHIPPED,SOTCARM1.PICK_NO,SOTCARM1.CART_TOTAL_UNITS,SOTCARM1.CART_TOTAL_WGT_ACTUAL," _
                & " SOTCARM1.CART_TOTAL_WGT_CALC, SOTCARM1.CART_TRACKING_NO, SOTCARM1.CART_SEQ, SOTCARM1.CART_MEMO, SOTCARM1.CART_TYPE, SOTCARM1.PACKAGING_TYPE," _
                & " SOTCARM1.PKG_CODE, SOTCARM1.PKG_L, SOTCARM1.PKG_W, SOTCARM1.PKG_H, SOTCARM1.CART_TOTAL_UNITS_REL, SOTCARM1.PALLET_NO," _
                & " SOTPICK1.SHIP_BOL_NO, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTORDR1.CUST_STORE_NO, SOTPICK1.PICK_STATUS, SOTORDR1.CUST_DC_NO," _
                & " SOTORDR1.ORDR_CUST_PO, SUM(SOTCARM2.QTY_PACKED) UNITS_PO, SUM(SOTCARM2.QTY_PACKED * SOTORDR2.ORDR_UNIT_PRICE) CART_VALUE_PO" _
                & " From SOTCARM1, SOTPICK1, SOTSHIP1, SOTORDR1, SOTORDR2, SOTCARM2" _
                & " Where SOTPICK1.PICK_NO_CONS = SOTCARM1.PICK_NO" _
                & " And SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
                & " And SOTORDR2.ORDR_NO = SOTCARM2.ORDR_NO" _
                & " And SOTCARM2.CART_NO = SOTCARM1.CART_NO" _
                & " And SOTORDR2.ORDR_LNO = SOTCARM2.ORDR_LNO" _
                & " And SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
                & " And SOTORDR1.ORDR_NO = SOTCARM2.ORDR_NO" _
                & " And (SOTSHIP1.ORDR_GROUP_NO = :PARM1 OR SOTSHIP1.ORDR_GROUP_NO = :PARM2)" _
                & " GROUP by SOTCARM1.CART_NO, SOTCARM1.CART_FREIGHT, SOTCARM1.CART_PACKER, SOTCARM1.CART_PACKED, SOTCARM1.CART_SHIPPED, SOTCARM1.PICK_NO, SOTCARM1.CART_TOTAL_UNITS, SOTCARM1.CART_TOTAL_WGT_ACTUAL," _
                & " SOTCARM1.CART_TOTAL_WGT_CALC, SOTCARM1.CART_TRACKING_NO, SOTCARM1.CART_SEQ, SOTCARM1.CART_MEMO, SOTCARM1.CART_TYPE, SOTCARM1.PACKAGING_TYPE," _
                & " SOTCARM1.PKG_CODE, SOTCARM1.PKG_L, SOTCARM1.PKG_W, SOTCARM1.PKG_H, SOTCARM1.CART_TOTAL_UNITS_REL, SOTCARM1.PALLET_NO," _
                & " SOTPICK1.SHIP_BOL_NO, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTORDR1.CUST_STORE_NO, SOTPICK1.PICK_STATUS, SOTORDR1.CUST_DC_NO," _
                & " SOTORDR1.ORDR_CUST_PO"
                Create_TDA(.Tables.Add, "SOTCARTX", "**", 0, False, "VV", 0)
            End If

            If Not .Tables("SOTCART1").Columns.Contains("PALLET_NO") Then
                .Tables("SOTCARTP").Columns.Add("PALLET_NO")
            End If
            .Tables("SOTCARTP").Columns.Add("STYLES", GetType(System.Int32))
            .Tables("SOTCARTP").Columns.Add("QTY_PACKED", GetType(System.Int32))
            .Tables("SOTCARTP").Columns.Add("SHIP_TRAILER_NO")
            .Tables("SOTCARTP").Columns.Add("PALLET_INIT_DATE", GetType(System.DateTime))
            .Tables("SOTCARTP").Columns.Add("PALLET_INIT_OPER")
            .Tables("SOTCARTP").Columns.Add("TRACKING_NO")
            .Tables("SOTCARTP").Columns.Add("CARTON_WEIGHT", GetType(System.Double))
            .Tables("SOTCARTP").Columns.Add("CARTON_VALUE", GetType(System.Double))
            .Tables("SOTCARTP").Columns.Add("PALLET_VALUE", GetType(System.Double))
            .Tables("SOTCARTP").Columns.Add("SCAN_TIME")
            .Tables("SOTCARTP").Columns.Add("WALMART_REC")
            .Tables("SOTCARTP").Columns.Add("SHIP_LOAD_NO")
            .Tables("SOTCARTP").Columns.Add("BILL_OF_LADING_NO")
            .Tables("SOTCARTP").Columns.Add("MASTER_SHIP_BOL_NO")
            .Tables("SOTCARTP").Columns.Add("MULTI_PO")

            ASCMAIN1.sql = "Select SOTCART2.*, SOTPICK1.ORDR_NO PICK_ORDR_NO" _
                & " from SOTCART2,SOTCART1,SOTPICK1,SOTSHIP1" _
                & " where SOTCART1.CART_NO = SOTCART2.CART_NO" _
                & "   and SOTPICK1.PICK_NO = SOTCART1.PICK_NO" _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
                & "   and SOTSHIP1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "V", 2)
            .Tables("SOTCART2").Columns.Add("QTY_RFID", GetType(System.Int32))

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")

            .Tables("SOTCART1").Columns("STYLES").Expression = "COUNT(CHILD.CART_LNO)"
            .Tables("SOTCART1").Columns("QTY_PACKED").Expression = "SUM(CHILD.QTY_PACKED)"
            .Tables("SOTCART1").Columns("QTY_RFID").Expression = "SUM(CHILD.QTY_RFID)"

            sqlSOTORDRT = "Select SOTORDRS.*" & vbCrLf _
           & ", ICTCOLR1.COLOR_DESC" & vbCrLf _
           & ", ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY" & vbCrLf _
           & " from SOTORDR0,SOTORDRS,ICTSTYL1,ICTCOLR1" & vbCrLf _
           & " where SOTORDRS.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
           & "   and ICTSTYL1.STYLE_CODE = SOTORDRS.STYLE_CODE" & vbCrLf _
           & "   and ICTCOLR1.COLOR_CODE = SOTORDRS.COLOR_CODE" & vbCrLf _
           & "   and SOTORDR0.ORDR_GROUP_NO = :PARM1"
            ASCMAIN1.sql = sqlSOTORDRT
            Create_TDA(.Tables.Add, "SOTORDRT", "**", 0, False, "V", 3)
            With .Tables("SOTORDRT").Columns
                .Add("ORDR_QTY_ALLO_NOW", GetType(System.Int32))
                .Add("ORDR_QTY_BACK_NOW", GetType(System.Int32))
                .Add("ORDR_QTY_CANC_NOW", GetType(System.Int32))
            End With

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                Dim SQL As New System.Text.StringBuilder With {.Length = 0}
                SQL.AppendLine("Select")
                SQL.AppendLine("X.STYLE_CODE,")
                SQL.AppendLine("X.COLOR_CODE,")
                SQL.AppendLine("ICTSTYL1.STYLE_DESC,")
                SQL.AppendLine("ICTCOLR1.COLOR_DESC,")
                SQL.AppendLine("SUM(X.BEG) BEG, SUM(X.SHP) SHP, SUM(X.RTN) RTN, SUM(X.REC) REC,")
                SQL.AppendLine("SUM(X.ADJ) ADJ, SUM(X.XFR) XFR, SUM(X.PHY) PHY, SUM(X.ON_HAND) ON_HAND,")
                SQL.AppendLine("SUM(X.ON_ORDER) ON_ORDER, SUM(X.TRAN) TRAN, SUM(X.OPEN) OPEN,")
                SQL.AppendLine("SUM(X.PICK) PICK, SUM(X.ALLO) ALLO, SUM(X.COMM) COMM, SUM(X.PROD) PROD,")
                SQL.AppendLine("MAX(UPC_CODE) UPC_CODE, MAX(STYLE_COLOR_STATUS) STYLE_COLOR_STATUS from ICTCOLR1, ICTSTYL1,")
                SQL.AppendLine("(")
                SQL.AppendLine("        Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE,")
                SQL.AppendLine("        SUM(0) BEG,")
                SQL.AppendLine("        SUM (0) SHP,")
                SQL.AppendLine("        SUM (0) RTN,")
                SQL.AppendLine("        SUM (0) REC,")
                SQL.AppendLine("        SUM (0) ADJ,")
                SQL.AppendLine("        SUM (0) XFR,")
                SQL.AppendLine("        SUM (0) PHY,")
                SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_ON_HAND) ON_HAND,")
                SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_ON_ORDER) ON_ORDER,")
                SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_TRAN) TRAN,")
                SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_OPEN) OPEN, SUM(ICTSTAT2.WHSE_QTY_PICK) PICK,")
                SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_ALLO) ALLO,")
                SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_COMM) COMM, SUM(ICTSTAT2.WHSE_QTY_PROD) PROD,")
                SQL.AppendLine("        NULL UPC_CODE, NULL STYLE_COLOR_STATUS from ICTSTAT2")
                SQL.AppendLine("        group by ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE")
                SQL.AppendLine(") X")
                SQL.AppendLine("where ICTCOLR1.COLOR_CODE (+) = X.COLOR_CODE")
                SQL.AppendLine("and ICTSTYL1.STYLE_CODE (+) = X.STYLE_CODE")
                SQL.AppendLine("group by X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC")
                ASCMAIN1.sql = SQL.ToString
                Create_TDA(.Tables.Add, "ICTSTATA", "**", 0, False, "V", 2)
                With .Tables("ICTSTATA").Columns
                    .Add("OTS_INV", GetType(System.Int64), "ISNULL(ON_HAND,0) - ISNULL(PICK,0)")
                    .Add("OTS_WIP", GetType(System.Int64), "ISNULL(OTS_INV,0) + ISNULL(TRAN,0) + ISNULL(ON_ORDER,0)")
                    .Add("NET_POS", GetType(System.Int64), "ISNULL(OTS_WIP,0) - ISNULL(OPEN,0) - ISNULL(COMM,0) - ISNULL(PROD,0)")
                    .Add("THIS_PO", GetType(System.Int64))
                End With

                'With .Tables.Add("ICTSTATA")
                '    For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "COLOR_CODE", "WHSE_CODE", "STYLE_DESC", "COLOR_DESC", "WHSE_DESC"}
                '        If TABLE_NAME = "ICTSTATA" And (COLUMN_NAME = "WHSE_CODE" Or COLUMN_NAME = "WHSE_DESC") Then
                '        ElseIf TABLE_NAME = "ICTSTATW" And (COLUMN_NAME = "STYLE_DESC" Or COLUMN_NAME = "COLOR_DESC") Then
                '        Else
                '            .Columns.Add(COLUMN_NAME)
                '        End If
                '    Next
                '    For Each COLUMN_NAME As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "PHY",
                '                                "ON_HAND", "ON_ORDER", "TRAN", "OPEN", "PICK", "ALLO", "COMM", "PROD"}
                '        .Columns.Add(COLUMN_NAME, GetType(System.Int64))
                '    Next
                '    .Columns.Add("UPC_CODE")
                '    .Columns.Add("STYLE_COLOR_STATUS")
                '    .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
                '    .Columns.Add("THEME_DESC")
                '    .Columns.Add("OTS_INV", GetType(System.Int64), "ISNULL(ON_HAND,0) - ISNULL(PICK,0)")
                '    .Columns.Add("OTS_WIP", GetType(System.Int64), "ISNULL(OTS_INV,0) + ISNULL(TRAN,0) + ISNULL(ON_ORDER,0)")
                '    .Columns.Add("NET_POS", GetType(System.Int64), "ISNULL(OTS_WIP,0) - ISNULL(OPEN,0) - ISNULL(COMM,0) - ISNULL(PROD,0)")
                'End With
            End If

            ASCMAIN1.sql = "Select *" & vbCrLf _
                & " from SOTORDR4 " & vbCrLf _
                & " where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add("SOTORDRC"), "SOTORDR4", "**", 0, False, "V", 2)
        End With

        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdSOTORDRT.DataSource = dst.Tables("SOTORDRT")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdSOTORDRS.DataSource = dst.Tables("SOTORDRS")
        grdSOTORDRS_ORDERS.DataSource = dst.Tables("SOTORDRS_ORDERS")
        grdSOTORDRS_SC.DataSource = dst.Tables("SOTORDRS_SC")
        grdSOTCORDX.DataSource = dst.Tables("SOTCORDX")
        grdSOTCORDD.DataSource = dst.Tables("SOTCORDD")
        grdSOTCORDY.DataSource = dst.Tables("SOTCORDY")
        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTORDRX.DataSource = grdSOTORDR1.DataSource
        grdSOTORDRM.DataSource = dst.Tables("SOTORDRM")
        grdSOTORDRP.DataSource = dst.Tables("SOTORDRP")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTCARTP.DataSource = dst.Tables("SOTCARTP")
        grdSOTORDR4.DataSource = dst.Tables("SOTORDR4")
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            grdICTSTATA.DataSource = dst.Tables("ICTSTATA")
            With grdICTSTATA.DisplayLayout.Bands(0)
                .Columns("STYLE_CODE").Header.Fixed = True
                .Columns("COLOR_CODE").Header.Fixed = True
            End With
            grdSOTCARTX.DataSource = dst.Tables("SOTCARTX")
        End If

        Bind_Controls(splComments.Panel1, "SOTORDR1")

        With grdSOTORDRT.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                If gcol.Key.StartsWith("ORDR_QTY") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Width = 55
                    gcol.Format = "#,##0"
                ElseIf gcol.Key.StartsWith("ORDR_AMT") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Width = 70
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If

                If gcol.Key = "ORDR_QTY_ALLO_NOW" Or gcol.Key = "ORDR_QTY_BACK_NOW" Or gcol.Key = "ORDR_QTY_CANC_NOW" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        Create_Summary(grdSOTORDRS_ORDERS, "CUST_CODE", "Count")
        Create_Summary(grdSOTORDRS_ORDERS, New String() {"CUR_QTY_ONHD", "FUT_QTY_ONPO", "FUT_QTY_TRAN", "UNAV_QTY"})
        Create_Summary(grdSOTORDRS_ORDERS, New String() {"CUR_AMT_ONHD", "FUT_AMT_ONPO", "FUT_AMT_TRAN", "UNAV_AMT"})
        Create_Summary(grdSOTORDRS_SC, "STYLE_CODE", "Count")
        Create_Summary(grdSOTORDRS_SC, New String() {"CUR_QTY_ONHD", "FUT_QTY_ONPO", "FUT_QTY_TRAN", "UNAV_QTY"})
        Create_Summary(grdSOTORDRS_SC, New String() {"CUR_AMT_ONHD", "FUT_AMT_ONPO", "FUT_AMT_TRAN", "UNAV_AMT"})

        Create_Summary(grdSOTORDRT, "STYLE_CODE", "Count")
        Create_Summary(grdSOTORDRT, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_ALLO_CUR", "ORDR_QTY_ALLO_FUT", "ORDR_QTY_ALLO_CXL", "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO_CUR", "ORDR_AMT_ALLO_FUT", "ORDR_AMT_ALLO_CXL"})
        Create_Summary(grdSOTORDRT, New String() {"ORDR_QTY_ALLO_NOW", "ORDR_QTY_BACK_NOW", "ORDR_QTY_CANC_NOW"})

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Create_Summary(grdICTSTATA, "COLOR_CODE", "Count")
            Create_Summary(grdICTSTATA, New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "PHY",
                                                  "ON_HAND", "ON_ORDER", "TRAN", "OPEN", "PICK", "ALLO", "COMM", "PROD", "OTS_INV", "OTS_WIP", "NET_POS"})
        End If

        grdSOTORDR1.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDR1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "CUST_STORE_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"EDI_PO_TYPE"}.Contains(gcol.Key) Then
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If
            Next

            If ASCMAIN1.DBS_COMPANY = "VAN" Then
                For Each SFX As String In New String() {""}
                    For Each TYP As String In New String() {"QTY", "AMT"}
                        Dim C As String = "ORDR_" & TYP & SFX
                        .Columns(C).Format = grdSOTORDR0.DisplayLayout.Bands(0).Columns(C).Format
                        .Columns(C).Width = grdSOTORDR0.DisplayLayout.Bands(0).Columns(C).Width
                        With .Columns(C).Header
                            .Caption = grdSOTORDR0.DisplayLayout.Bands(0).Columns(C).Header.Caption
                            .Appearance = grdSOTORDR0.DisplayLayout.Bands(0).Columns(C).Header.Appearance
                        End With
                        .Columns(C).Hidden = False
                        Create_Summary(grdSOTORDR1, C)
                    Next
                Next
                Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")

            End If


        End With

        'cols3PL = {"QTY_XFR_3PL", "QTY_XIT_3PL", "QTY_XIT_3PL_ETA", "QTY_ONH_3PL", "QTY_PCK_3PL", "QTY_OPN_3PL", "QTY_AVA_3PL"}
        cols3PL = {"QTY_ONH_3PL", "QTY_PCK_3PL", "QTY_OPN_3PL", "QTY_AVA_3PL"}

        With grdSOTORDRS.DisplayLayout.Bands(0)
            .Override.AllowUpdate = DefaultableBoolean.True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit

                Dim COLUMN_NAME = gcol.Key
                If cols3PL.Contains(COLUMN_NAME) Or COLUMN_NAME = "CARTON_PACK_QTY" Or COLUMN_NAME = "INNER_PACK_QTY" Then
                    If ASCMAIN1.CLIENT = "RGI" Then
                        .Columns(COLUMN_NAME).Hidden = False
                        If cols3PL.Contains(COLUMN_NAME) Then
                            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        End If
                    Else
                        .Columns(COLUMN_NAME).Hidden = True
                    End If
                End If
            Next

            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE"}
                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    If COLUMN_NAME = "STYLE_CODE" Then .Columns(COLUMN_NAME).Header.Fixed = True
                Else
                    .Columns(COLUMN_NAME).Header.Fixed = True
                End If


            Next
            .Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        End With


        For Each grd As UltraWinGrid.UltraGrid In {grdSOTORDRS_ORDERS, grdSOTORDRS_SC}
            With grd.DisplayLayout.Bands(0)
                ' .Override.AllowUpdate = DefaultableBoolean.True
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    If {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC", "ORDR_NO", "ORDR_CUST_PO", "CUST_CODE", "CUST_NAME", "ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(gcol.Key) Then
                        gcol.Header.Fixed = True
                    Else
                        If gcol.Key = "ONHD_3PL" Or gcol.Key = "TRAN_3PL" Then
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                        ElseIf gcol.Key.Contains("AMT") Then
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        ElseIf gcol.Key.Contains("QTY") Then
                            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                        End If
                    End If

                Next

                For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}


                    If cols3PL.Contains(COLUMN_NAME) Or COLUMN_NAME = "CARTON_PACK_QTY" Or COLUMN_NAME = "INNER_PACK_QTY" Then
                        If ASCMAIN1.CLIENT = "RGI" Then
                            .Columns(COLUMN_NAME).Hidden = False
                        Else
                            .Columns(COLUMN_NAME).Hidden = True
                        End If
                    End If
                Next
                .Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
            End With
        Next


        'grdSOTORDR0.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            .Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                'If gcol.Key = "PO_QTY" Then
                '    gcol.CellAppearance.BackColor = Drawing.Color.LightYellow
                '    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                '    '  gcol.CellAppearance.BackColor = Drawing.Color.Beige
                '    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                'End If

                If gcol.Key.StartsWith("ORDR_AMT_ALLO_") Or gcol.Key.StartsWith("PCT_ALLO_") Then
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If
                ', "ORDR_ARRIVAL_DATE", "ORDR_LAST_ARRIVAL_DATE"
                If New String() {"ORDR_DATE_RECD", "ORDR_PRIORITY",
                                 "ORDR_RELEASE_AVAIL_MIN", "ORDR_RELEASE_AVAIL_MAX", "ORDR_REL_SHORT", "ORDR_REL_SHORT_OPER",
                                 "ORDR_REL_ACTION_DATE", "ORDR_REL_ACTION_OPER", "TERM_CODE", "LAST_DATE", "LAST_OPER", "ORDR_SHIP_INSTR", "ORDR_MESSAGE", "EDI_PO_TYPE"}.Contains(gcol.Key) Then
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If

                If New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf gcol.Key.StartsWith("ORDR_AMT_ALLO_") Or gcol.Key.StartsWith("PCT_ALLO_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT", "WHSE_CODE", "EDI_MERCH_TYPE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "SALES_DIVISION_CODE", "SREP_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"CUST_CITY", "CUST_STATE", "CUST_COUNTRY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                Else
                    gcol.Header.Appearance.BackColor = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            Next

            .Columns("EDI_CONS_NO").Hidden = Not (ASCMAIN1.CLIENT = "VAN")
        End With


        If ASCMAIN1.CLIENT = "VAN" Then
            With grdSOTCART1.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If gcol.Key.StartsWith("PALLET_INIT_DATE") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                        gcol.Header.Caption = "Pallet Date"
                        gcol.Width = 100
                        ' gcol.Format = "#,##0"
                    ElseIf gcol.Key.StartsWith("PALLET_INIT_OPER") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                        gcol.Header.Caption = "Pallet User"
                        gcol.Width = 100
                    ElseIf gcol.Key.StartsWith("SHIP_TRAILER_NO") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                        gcol.Header.Caption = "Trailer No"
                        gcol.Width = 100
                    ElseIf gcol.Key.StartsWith("PALLET_NO") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                        gcol.Header.Caption = "Pallet No"
                        gcol.Width = 100
                    ElseIf gcol.Key.StartsWith("CUST_DC_NO") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                        gcol.Width = 70
                    ElseIf gcol.Key.StartsWith("QTY_PACKED") Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                    End If
                Next
            End With
            With grdSOTCARTP.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    Select Case gcol.Key
                        Case "PALLET_INIT_DATE"
                            gcol.Header.Caption = "Pallet Date"
                            gcol.Width = 100
                            gcol.Hidden = True
                        Case "PALLET_INIT_OPER"
                            gcol.Header.Caption = "Pallet User"
                            gcol.Width = 100
                            gcol.Hidden = True
                        Case "TRACKING_NO"
                            gcol.Header.Caption = "Tracking #"
                            gcol.Width = 200
                            gcol.Hidden = True
                        Case "CARTON_WEIGHT"
                            gcol.Header.Caption = "Cart Weight"
                            gcol.Width = 100
                            gcol.Hidden = True
                        Case "CARTON_VALUE"
                            gcol.Header.Caption = "Cart Value"
                            gcol.Width = 100
                            gcol.Hidden = True
                        Case "PALLET_VALUE"
                            gcol.Header.Caption = "Pallet Value"
                            gcol.Width = 100
                            gcol.Hidden = True
                        Case "SCAN_TIME"
                            gcol.Header.Caption = "Scan Time"
                            gcol.Width = 100
                            gcol.Hidden = True
                        Case "WALMART_REC"
                            gcol.Header.Caption = "Walmart Rec"
                            gcol.Width = 100
                            gcol.Hidden = True
                        Case "SHIP_LOAD_NO"
                            gcol.Header.Caption = "Load No"
                            gcol.Width = 100
                            gcol.Hidden = True
                        Case "BILL_OF_LADING_NO"
                            gcol.Header.Caption = "BOL No"
                            gcol.Width = 200
                            gcol.Hidden = True
                        Case "MASTER_SHIP_BOL_NO"
                            gcol.Header.Caption = "Master BOL"
                            gcol.Width = 125
                            gcol.Hidden = True
                        Case "MULTI_PO"
                            gcol.Header.Caption = "Multi PO"
                            gcol.Width = 200
                            gcol.Hidden = True
                    End Select
                    'gcol.Format = "#,##0"
                Next
            End With
        End If

        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}, , , "#,##0")

        Create_Summary(grdSOTORDRS, "STYLE_CODE", "Count")
        Create_Summary(grdSOTORDRS, New String() {"ORDR_QTY", "ORDR_AMT", "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        If ASCMAIN1.CLIENT = "VAN" Then
            Create_Summary(grdSOTCART1, New String() {"CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL", "CART_TOTAL_WGT_CALC"}, , , "###,##0")
            Create_Summary(grdSOTCART1, New String() {"QTY_PACKED", "STYLES"}, , , "#,##0")

            Create_Summary(grdSOTCARTP, New String() {"CART_NO"}, "Count")
            Create_Summary(grdSOTCARTP, "PALLET_NO", "Custom")
            Create_Summary(grdSOTCARTP, "SHIP_TRAILER_NO", "Custom", , "###,##0")
            Create_Summary(grdSOTCARTP, New String() {"CART_TOTAL_UNITS", "CART_TOTAL_UNITS_REL"}, , , "###,##0")
            Create_Summary(grdSOTCARTP, New String() {"QTY_PACKED", "STYLES"}, , , "#,##0")
            Create_Summary(grdSOTCARTP, New String() {"CARTON_WEIGHT", "CARTON_VALUE"}, , , "#,##0.00")

            Create_Summary(grdSOTCARTX, New String() {"ORDR_CUST_PO", "CART_NO"}, "Count")

            Create_Summary(grdSOTCARTX, New String() {"UNITS_PO"}, , , "##,##0")
            Create_Summary(grdSOTCARTX, New String() {"CART_VALUE_PO"}, , , "##,##0.00")

        End If

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTCORDD, grdSOTCORDX}
            grd.DisplayLayout.UseFixedHeaders = True
            With grd.DisplayLayout.Bands(0)
                For I As Integer = 1 To 12
                    .Columns("V" & Format(I, "00")).Format = "#,##0"
                Next
                .Columns("TOT").Format = "#,##0"
                .Columns("YTD").Format = "#,##0"
                .Columns("TOTPCT").Format = "#,##0.0"
                .Columns("YTDPCT").Format = "#,##0.0"
                For Each COLUMN_NAME As String In New String() {"CODE_VALUE", "TOT", "YTD", "TOTPCT", "YTDPCT"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
            End With
        Next
        Create_Summary(grdSOTCORDD, "CODE_VALUE", "Count")
        Create_Summary(grdSOTCORDD, New String() {"TOT", "TOTPCT", "YTD", "YTDPCT", "V01", "V02", "V03", "V04", "V05", "V06", "V07", "V08", "V09", "V10", "V11", "V12"})

        Create_Summary(grdSOTCORDY, "INV_NO", "Count")
        Create_Summary(grdSOTCORDY, New String() {"ORDR_QTY_SHIP", "AMT"})


        Show_Filter(grdSOTORDR0, True)
        grdSOTORDR0.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdSOTCART1, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTCARTP, "PICK_STATUS")

        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "SHIP_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "LP_STATUS", Nothing, New String() {":", "0:Pending Xmit", "1:Transmitted", "2:Shipped", "3:Confirmed", "V:Xmit/De-Rel", "D:Deleted"})
        '  Set_cmbYP("RYP_TO", ASCMAIN1.CYP, -36, 12, 0)
        cmb12Months.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cmb12Months.SelectedRow = cmb12Months.Rows(0)

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            ASCMAIN1.Add_Value_List(grdSOTORDR0, "EDI_PO_TYPE")
            ASCMAIN1.Add_Value_List(grdSOTORDR1, "EDI_PO_TYPE")
        End If

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            chkReservations.Checked = True
            chkReservations_FindCustomerBy.Checked = True
        End If

        grdSOTORDR0.DisplayLayout.Bands(0).Columns("ORDR_TYPE_CODE").Hidden = Not (ASCMAIN1.CLIENT = "RGI")
        grdSOTORDR0.DisplayLayout.Bands(0).Columns("ORDR_SOURCE").Hidden = Not (ASCMAIN1.CLIENT = "RGI")

        grdSOTORDR0.DisplayLayout.Bands(0).Columns("WAVE_NO").Hidden = Not (ASCMAIN1.CLIENT = "VAN")
        grdSOTORDR0.DisplayLayout.Bands(0).Columns("EDI_LOAD_ID").Hidden = Not (ASCMAIN1.CLIENT = "VAN")


        grdSOTORDR0.DisplayLayout.Bands(0).Columns("TERM_CODE").Hidden = Not (ASCMAIN1.CLIENT = "RGI")


        'If ASCMAIN1.CLIENT = "VAN" Then
        '    grdSOTCART1.DisplayLayout.Bands(0).Columns("PALLET_NO").Hidden = True
        'End If

        grdSOTCART1.DisplayLayout.Bands(0).Columns("PALLET_NO").Hidden = True
        grdSOTCART1.DisplayLayout.Bands(0).Columns("SHIP_TRAILER_NO").Hidden = True
        grdSOTCART1.DisplayLayout.Bands(0).Columns("PALLET_INIT_DATE").Hidden = True
        grdSOTCART1.DisplayLayout.Bands(0).Columns("PALLET_INIT_OPER").Hidden = True
        grdSOTCART1.DisplayLayout.Bands(0).Columns("CUST_DC_NO").Hidden = True

        grdSOTCARTP.DisplayLayout.Bands(0).Columns("PALLET_NO").Hidden = False
        grdSOTCARTP.DisplayLayout.Bands(0).Columns("SHIP_TRAILER_NO").Hidden = False
        grdSOTCARTP.DisplayLayout.Bands(0).Columns("PALLET_INIT_DATE").Hidden = False
        grdSOTCARTP.DisplayLayout.Bands(0).Columns("PALLET_INIT_OPER").Hidden = False
        grdSOTCARTP.DisplayLayout.Bands(0).Columns("CUST_DC_NO").Hidden = False
        '  grdSOTCARTP.DisplayLayout.Bands(0).Columns("MULTI_PO").Hidden = False



        If ASCMAIN1.CLIENT = "VAN" Then
            Show_Filter(grdSOTCARTP, True)
            grdSOTCARTP.DisplayLayout.GroupByBox.Hidden = False

            Show_Filter(grdSOTCARTX, True)
            grdSOTCARTX.DisplayLayout.GroupByBox.Hidden = False
        End If

        chkActionDate.Visible = (ASCMAIN1.CLIENT = "RGI")
        dteActionDate.Visible = (ASCMAIN1.CLIENT = "RGI")
        chkIfReceivedSince.Visible = (ASCMAIN1.CLIENT = "RGI")
        dteIfReceivedSince.Visible = (ASCMAIN1.CLIENT = "RGI")

        dteActionDate.DateTime = Now.Date.AddDays(7)
        dteIfReceivedSince.DateTime = Now.Date.AddDays(-3)

        chkShortView.Visible = (ASCMAIN1.CLIENT = "RGI")

        SplitContainer1.Panel2Collapsed = (ASCMAIN1.CLIENT = "RGI")

        MakeTransparent(chkShortView)
        MakeTransparent(chkSynchronize)

        If ASCMAIN1.CLIENT = "RGI" Then
            SOFCORD1_LAYOUT_SHORT = ASCMAIN1.Folders("SharedRoot") & "Templates\" & "SOFCORD1_LAYOUT_SHORT.xml"
            grdSOTORDR0.DisplayLayout.SaveAsXml(SOFCORD1_LAYOUT_ORIG, UltraWinGrid.PropertyCategories.All)
        End If

        tabDetails.Tabs("Comments").Visible = (ASCMAIN1.CLIENT = "RGI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text

                        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                            If CUST_CODE <> "LOBLAW" Then
                                EMsg &= vbCr & "Invalid Customer " & CUST_CODE
                            End If
                        End If

                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
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

            Case "Select"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Select").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

            .Groups("Find Customer By").Visible = Not ScreenMode
            .Groups("Show Orders").Visible = ScreenMode
            .Groups("Styles").Visible = False
            .Groups("12 Month History").Visible = False
            .Groups("Multi PO Pallet").Visible = False


            If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                .Groups("Find Customer By").Visible = False
            End If

        End With

        grpSOTORDR0.Visible = Not ScreenMode
        ' splSOTORDR0s.Visible = Not ScreenMode
        tab0.Visible = Not ScreenMode

        chkDetails.Visible = Not ScreenMode And (ASCMAIN1.CLIENT = "RGI")

        tabMain.Visible = ScreenMode

        grdSOTORDR0.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        spl12Months.Visible = False
        chtSATCSLS1_X.Visible = False

        grdSOTORDR0.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = ScreenMode

        If ScreenMode Then
            grdSOTORDR0.Parent = splSOTORDR0.Panel1
            Setup_tabMain()
            Setup_Summary()

            splComments.Parent = tabDetails.Tabs("Comments").TabPage

        Else
            Clear_Record()
            grdSOTORDR0.Parent = splSOTORDR0s.Panel1 ' grpSOTORDR0
            splComments.Parent = UltraTabControl1.Tabs("Comments").TabPage

            toggle_Show_Details()

            For Each COLUMN_NAME As String In New String() _
                            {"STYLE_CODE", "COLOR_CODE", "CUST_UPC", "RANGE_STYLE_CODE",
                             "CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_SKU"}
                Absx1.chkFor("SHOW_" & COLUMN_NAME).Checked = (COLUMN_NAME = "STYLE_CODE" Or COLUMN_NAME = "COLOR_CODE")
            Next
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            tabStyles.Tabs("Status").Visible = True
        Else
            tabStyles.Tabs("Status").Visible = False
        End If

        btnSaveShort.Visible = ASCMAIN1.Running_in_VS And (ASCMAIN1.CLIENT = "RGI")
    End Sub

    Sub Clear_Record()

        Absx1.txtFor("CUST_CODE").Text = ""
        Dim CUST_CODE_prev As String = CUST_CODE
        CUST_CODE = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDRS", "SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTSHIP2",
             "SOTCORDX", "SOTCORDY", "SOTCORDD", "SOTORDRM", "SOTORDRP", "SOTCART1", "SOTCART2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        optOrders.Value = "OP"
        optMULTIPOP.Value = "I"

        tabMain.SelectedTab = tabMain.Tabs("Orders")
        tabMonth.Tabs("Details").Visible = False

        grdSOTORDR0.Tag = ""

        If Not Me.IsClosing Then
            'If CUST_CODE_prev = "" Then
            Load_SOTORDR0("")
            'End If
        End If

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        grdSOTORDR0.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        Load_SOTORDR0("", CUST_CODE)
        Setup_SOTORDR0()

        tabDetails.Tabs("Pallets").Visible = ASCMAIN1.CLIENT = "VAN"

        EnforceConstraints(True)

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSP("WJZOP", "VV", New String() {Me.Name, ASCMAIN1.USER_ID}, New String() {"FORM_NAME", "USER_ID"})
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Select"

                Dim CUST_CODE As String = Split(key, ":")(0)
                Dim ORDR_GROUP_NO As String = Split(key & ":", ":")(1)
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Click_Command("Select")
                If ORDR_GROUP_NO <> "" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                        If grow.Cells("ORDR_GROUP_NO").Value = ORDR_GROUP_NO Then
                            grdSOTORDR0.ActiveRow = grow
                            grdSOTORDR0.ActiveRowScrollRegion.FirstRow = grow
                            grow.Selected = True
                        End If

                    Next
                End If
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ARTCUST1"
            E.COLUMN_NAME = "CUST_CODE"
            E.CODE_VALUE = Absx1.txtFor("CUST_CODE").Text
            E.DESC_VALUE = "Customer"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDR0, "SSSSBBSSBBBBBBBBBBBBBBBS", "Show Filter", "Show GroupBox", "Show Pins", "Short View",
                        "Store Configuration Report", "Customer Order Summary", "Show Original Ship/Cancel", "Show Orders with Changed Ship/Cancel",
                        "Sales Order Entry", "Sales Order Inquiry", "Show Raw EDI", "Export Sales Order Details", "Convert CTF to Reservation", "Wave Inquiry",
                        "Create Billing Batch", "Create Master Carton Label", "Set Manual Release", "Clear Manual Release", "Summary by DC", "Carton Pack Configuration", "Customer Order Status", "Rebuild Order Summary", "Customer Order Inquiry", "Show Comments")
        Load_Popup_Menu(grdSOTORDR1, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Sales Order Entry", "Show Raw EDI")
        Load_Popup_Menu(grdSOTORDRS, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
        Load_Popup_Menu(grdSOTPICK1, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Invoice", "Pro-Forma Invoice", "EDI Data", "Show EDI Invoice")
        Load_Popup_Menu(grdSOTSHIP1, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Invoice", "Pro-Forma Invoice", "Show EDI ASN", "Show Shipments For Style")
        Load_Popup_Menu(grdSOTCORDY, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTORDRX, "BB", "Sales Order Inquiry", "Show Raw EDI")
        Load_Popup_Menu(grdSOTCART1, "SS", "Show Filter", "Show Details")
        Load_Popup_Menu(grdSOTCARTP, "SSSB", "Show Filter", "Show GroupBox", "Show Cart Stats", "Export With Stats")

        Load_Popup_Menu(grdSOTORDRS_ORDERS, "SSB", "Show Filter", "Show GroupBox", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTORDRS_SC, "SSB", "Show Filter", "Show GroupBox", "Style Status Inquiry")

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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '   e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTORDR0"
                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_btn = DirectCast(tlb_pop.Tools("Store Configuration Report"), UltraWinToolbars.ButtonTool)
                    Dim ORDR_TYPE As String = ""
                    Dim ORDR_CUST_PO As String = ""
                    Dim WAVE_NO As String = ""
                    Dim CUST_CODE_ORDR_GROUP As String = ""
                    Dim EDI_CONS_NO As String = ""
                    If grd.ActiveRow IsNot Nothing And grd.ActiveRow.IsDataRow Then
                        ORDR_TYPE = grd.ActiveRow.Cells("ORDR_TYPE").Value & ""
                        ORDR_CUST_PO = grd.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
                        WAVE_NO = grd.ActiveRow.Cells("WAVE_NO").Value & ""
                        CUST_CODE_ORDR_GROUP = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                        If ASCMAIN1.CLIENT = "VAN" Then
                            EDI_CONS_NO = grd.ActiveRow.Cells("EDI_CONS_NO").Value & ""
                        End If
                    End If
                    tlb_btn.SharedProps.Visible = (ORDR_TYPE = "O")
                    tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Summary"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode

                    tlb_btn = DirectCast(tlb_pop.Tools("Convert CTF to Reservation"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ORDR_TYPE = "O" And ORDR_CUST_PO.ToUpper Like "CTF*") And ScreenMode

                    tlb_btn = DirectCast(tlb_pop.Tools("Wave Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ORDR_TYPE = "O" And WAVE_NO <> "") And ScreenMode

                    tlb_btn = DirectCast(tlb_pop.Tools("Create Billing Batch"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ((ASCMAIN1.USER_ID = "naseema" Or ASCMAIN1.USER_ID = "wendy" Or ASCMAIN1.USER_ID = "avani") Or ASCMAIN1.Running_in_VS) And ScreenMode

                    tlb_btn = DirectCast(tlb_pop.Tools("Create Master Carton Label"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ((ASCMAIN1.CLIENT = "VAN") Or ASCMAIN1.Running_in_VS) And ScreenMode

                    tlb_btn = DirectCast(tlb_pop.Tools("Rebuild Order Summary"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ((ASCMAIN1.USER_ID = "naseema" Or ASCMAIN1.USER_ID = "wendy" Or ASCMAIN1.USER_ID = "avani") Or ASCMAIN1.Running_in_VS) And ASCMAIN1.CLIENT = "VAN" And ScreenMode
                    ' Or ASCMAIN1.Running_in_VS

                    tlb_btn = DirectCast(tlb_pop.Tools("Set Manual Release"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "RGI")
                    tlb_btn = DirectCast(tlb_pop.Tools("Clear Manual Release"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "RGI")

                    tlb_sbt = DirectCast(tlb_pop.Tools("Short View"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.SharedProps.Visible = (ASCMAIN1.CLIENT = "RGI")

                    tlb_sbt = DirectCast(tlb_pop.Tools("Show Comments"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.SharedProps.Visible = (ASCMAIN1.CLIENT = "RGI")

                    tlb_btn = DirectCast(tlb_pop.Tools("Summary by DC"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "VAN") And (CUST_CODE_ORDR_GROUP = "WALMART" Or CUST_CODE_ORDR_GROUP = "WALMARTCOM") And EDI_CONS_NO <> ""

                    tlb_btn = DirectCast(tlb_pop.Tools("Carton Pack Configuration"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "VAN") And (CUST_CODE_ORDR_GROUP = "WALMART" Or CUST_CODE_ORDR_GROUP = "WALMARTCOM") And ScreenMode

                    tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Status"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") And ScreenMode

                Case "grdSOTPICK1"
                    Dim PICK_STATUS As String = ""
                    If grd.ActiveRow.Band.Key = "SOTPICK1" Then
                        PICK_STATUS = grd.ActiveRow.Cells("PICK_STATUS").Value & ""
                    Else
                        PICK_STATUS = grd.ActiveRow.ParentRow.Cells("PICK_STATUS").Value & ""
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Sales Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (PICK_STATUS = "F")
                    tlb_btn = DirectCast(tlb_pop.Tools("Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (PICK_STATUS = "P")
                    tlb_btn = DirectCast(tlb_pop.Tools("Show EDI Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (PICK_STATUS = "F") And ASCMAIN1.CLIENT = "VAN"

                Case "grdSOTSHIP1"
                    Dim SHIP_STATUS As String = grd.ActiveRow.Cells("SHIP_STATUS").Value & ""
                    tlb_btn = DirectCast(tlb_pop.Tools("Sales Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (SHIP_STATUS = "F")
                    tlb_btn = DirectCast(tlb_pop.Tools("Pro-Forma Invoice"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (SHIP_STATUS = "P")
                    tlb_btn = DirectCast(tlb_pop.Tools("Show EDI ASN"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (SHIP_STATUS = "F") And ASCMAIN1.CLIENT = "VAN"
                    'Show Shipments For Style
                    tlb_btn = DirectCast(tlb_pop.Tools("Show Shipments For Style"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (SHIP_STATUS = "F") And ASCMAIN1.CLIENT = "VAN" And (Absx1.txtFor("CUST_CODE").Text = "WALMART" Or Absx1.txtFor("CUST_CODE").Text = "WALMARTCOM")


                Case "grdSOTCART1"
                    tlb_sbt = DirectCast(tlb_pop.Tools("Show Details"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.Checked = Not grdSOTCART1.DisplayLayout.Bands(0).Columns("PALLET_NO").Hidden
                    tlb_sbt.SharedProps.Visible = ASCMAIN1.CLIENT = "VAN"

                Case "grdSOTCARTP"
                    tlb_sbt = DirectCast(tlb_pop.Tools("Show Cart Stats"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.SharedProps.Visible = ASCMAIN1.CLIENT = "VAN" And (Absx1.txtFor("CUST_CODE").Text = "WALMART" Or Absx1.txtFor("CUST_CODE").Text = "WALMARTCOM")

                    tlb_btn = DirectCast(tlb_pop.Tools("Export With Stats"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ASCMAIN1.CLIENT = "VAN" And (Absx1.txtFor("CUST_CODE").Text = "WALMART" Or Absx1.txtFor("CUST_CODE").Text = "WALMARTCOM") And DirectCast(tlb_pop.Tools("Show Cart Stats"), UltraWinToolbars.StateButtonTool).Checked

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd IsNot Nothing AndAlso grd.Name = "grdSOTCART1" AndAlso e.Tool.Key = "Show Filter" Then
            grdSOTCART1.DisplayLayout.RowScrollRegions(0).Scroll(UltraWinGrid.RowScrollAction.Top)
        End If


        Select Case e.Tool.Key
            Case "Show Orders with Changed Ship/Cancel"
                Toggle_ChgShipCancel()
            Case "Show Original Ship/Cancel"
                Toggle_ShowShipCancel()
            Case "Short View"
                Toggle_OrderGridView()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"

                Dim ORDR_NO As String = ""
                If ASCMAIN1.CLIENT = "RGI" Then
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                Else
                    Dim ORDR_GROUP_NO = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                    ORDR_NO = ASCDATA1.GetDataValue("Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")

                    'ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value
                End If
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Sales Order Entry"

                If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                    Exit Sub
                End If

                Dim ORDR_NO As String = ""
                If grd.Name = "grdSOTORDR0" Then
                    Dim ORDR_GROUP_NO = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                    ORDR_NO = ASCDATA1.GetDataValue("Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                Else
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value
                End If

                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDR1")
                End If

            Case "Style Status Inquiry"
                If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                    Exit Sub
                End If

                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Show Raw EDI"

                'If grdSOTORDR1.ActiveRow IsNot Nothing Then
                '    Dim EDI_DOC_SEQ_NO As String = grdSOTORDR1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                'End If
                '  Display_Raw(grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value & "")

                If grd.Name = "grdSOTORDR1" And grdSOTORDR1.ActiveRow IsNot Nothing Then
                    Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                    Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"))
                    Using frm As New ASFTEXT1
                        frm.t = RAW_EDI
                        frm.Text = "Raw EDI for " & CUST_CODE & " PO No " & grdSOTORDR1.ActiveRow.Cells("ORDR_CUST_PO").Value
                        frm.ShowDialog()
                    End Using
                ElseIf grd.Name = "grdSOTORDR0" And grdSOTORDR0.ActiveRow IsNot Nothing Then
                    Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                    Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"))
                    Using frm As New ASFTEXT1
                        frm.t = RAW_EDI
                        frm.Text = "Raw EDI for " & CUST_CODE & " PO No " & grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value
                        frm.ShowDialog()
                    End Using
                End If

            Case "Show EDI ASN"
                If grd.ActiveRow IsNot Nothing Then
                    Dim EDI_DOCUMENT_NAME As String = grd.ActiveRow.Cells("BILL_OF_LADING_NO").Value & ""
                    Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI("", ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), "856", EDI_DOCUMENT_NAME)
                    Using frm As New ASFTEXT1
                        frm.t = RAW_EDI
                        frm.Text = "Raw EDI for " & CUST_CODE & " PO No " & grdSOTORDR1.ActiveRow.Cells("ORDR_CUST_PO").Value
                        frm.ShowDialog()
                    End Using
                End If

            Case "Show EDI Invoice"
                If grd.ActiveRow IsNot Nothing Then
                    Dim EDI_DOCUMENT_NAME As String = grd.ActiveRow.Cells("INV_NO").Value & ""
                    Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI("", ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), "810", EDI_DOCUMENT_NAME)
                    Using frm As New ASFTEXT1
                        frm.t = RAW_EDI
                        frm.Text = "Raw EDI for " & CUST_CODE & " PO No " & grdSOTORDR1.ActiveRow.Cells("ORDR_CUST_PO").Value
                        frm.ShowDialog()
                    End Using
                End If


            Case "Create Billing Batch"
                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                Dim ORDR_GROUP_NOs_to_batch As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Text
                    ORDR_GROUP_NOs_to_batch.Add(ORDR_GROUP_NO)
                Next

                Using FRM As New SOFBINV1
                    FRM.ORDR_GROUP_NOs = ORDR_GROUP_NOs_to_batch
                    FRM.ShowDialog()

                    If FRM.BATCH_NO <> "" Then
                        MsgBox("Batch No " & FRM.BATCH_NO & " has been created", MsgBoxStyle.OkOnly, "Verification")
                    End If
                End Using

                grdSOTORDR0.Selected.Rows.Clear()

            Case "Create Master Carton Label"
                If grd.Selected.Rows.Count > 1 Then
                    MsgBox("You may only select one Order Group when doing Master Carton Labels", MsgBoxStyle.OkOnly, "")

                    Exit Sub
                End If
                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                Dim ORDR_GROUP_NOs_for_MCL As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Text
                    ORDR_GROUP_NOs_for_MCL.Add(ORDR_GROUP_NO)
                Next

                Using FRM As New SOFMLAB1
                    FRM.ORDR_GROUP_NOs = ORDR_GROUP_NOs_for_MCL
                    FRM.CUST_CODE = CUST_CODE
                    FRM.CUST_NAME = Absx1.txtFor("CUST_NAME").Text
                    FRM.ORDR_GROUP_NO = ORDR_GROUP_NO

                    FRM.ShowDialog()

                    If FRM.printed Then
                        MsgBox("Master Carton Labels for Order Groups Selected have been created", MsgBoxStyle.OkOnly, "Verification")
                    End If
                End Using

                grdSOTORDR0.Selected.Rows.Clear()


            Case "Set Manual Release", "Clear Manual Release"
                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                Dim ORDR_GROUP_NOs_to_manually_Release As New List(Of String)
                Dim ORDR_GROUP_NOs_to_manually_Release_but_cannot As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Text
                    Dim ORDR_TYPE_CODE As String = grow.Cells("ORDR_TYPE_CODE").Text
                    Dim rowSOTORDR0 As DataRow = dst.Tables("SOTORDR0").Rows.Find(New String() {"O", ORDR_GROUP_NO})
                    If rowSOTORDR0 Is Nothing Then
                        ORDR_GROUP_NOs_to_manually_Release_but_cannot.Add(ORDR_GROUP_NO)
                        grow.Selected = False
                    Else
                        If ORDR_TYPE_CODE <> "REG" Then
                            ORDR_GROUP_NOs_to_manually_Release_but_cannot.Add(ORDR_GROUP_NO)
                            grow.Selected = False
                        Else
                            Dim ORDR_REL_SHORT As String = grow.Cells("ORDR_REL_SHORT").Value & ""
                            If e.Tool.Key.StartsWith("Set") And ORDR_REL_SHORT <> "1" Then
                                ORDR_GROUP_NOs_to_manually_Release.Add(ORDR_GROUP_NO)
                            ElseIf e.Tool.Key.StartsWith("Clear") And ORDR_REL_SHORT = "1" Then
                                ORDR_GROUP_NOs_to_manually_Release.Add(ORDR_GROUP_NO)
                            Else
                                ORDR_GROUP_NOs_to_manually_Release_but_cannot.Add(ORDR_GROUP_NO)
                                grow.Selected = False
                            End If
                        End If
                    End If
                Next

                If ORDR_GROUP_NOs_to_manually_Release.Count = 0 Then
                    MsgBox("No eligible records selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Else
                    For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs_to_manually_Release
                        Dim rowSOTORDR0 As DataRow = dst.Tables("SOTORDR0").Rows.Find(New String() {"O", ORDR_GROUP_NO})
                        If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO, , , , 1) Then
                            Exit Sub
                        End If
                        Dim ORDR_NO_MIN As String = rowSOTORDR0.Item("ORDR_NO_MIN")
                        If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO_MIN, , , , 1) Then
                            Exit Sub
                        End If
                        If Not ASCMAIN1.Logical_Open("R", "SOROREL1", , , , 1) Then
                            Exit Sub
                        End If
                    Next

                    If MsgBox("OK To " & e.Tool.Key & " For the " & CStr(ORDR_GROUP_NOs_to_manually_Release.Count) & " Order Groups Selected?" _
                              & IIf(ORDR_GROUP_NOs_to_manually_Release_but_cannot.Count > 0, vbCrLf & "Note: the following Order Groups are not Eligible for this action" & vbCrLf & Join(ORDR_GROUP_NOs_to_manually_Release_but_cannot.ToArray, ","), ""),
                              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                        'Using FRM As New SOFBINV1
                        '    FRM.ORDR_GROUP_NOs = ORDR_GROUP_NOs_to_batch
                        '    FRM.ShowDialog()

                        '    If FRM.BATCH_NO <> "" Then
                        '        MsgBox("Batch No " & FRM.BATCH_NO & " has been created", MsgBoxStyle.OkOnly, "Verification")
                        '    End If
                        'End Using

                        dst.Tables("SOTORDRG").Rows.Clear()
                        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs_to_manually_Release
                            Dim rowSOTORDRG As DataRow = Fill_Record("SOTORDRG", ORDR_GROUP_NO, True, False)
                            Dim rowSOTORDR0 As DataRow = dst.Tables("SOTORDR0").Rows.Find(New String() {"O", ORDR_GROUP_NO})
                            If rowSOTORDR0 IsNot Nothing Then

                                If e.Tool.Key.StartsWith("Set") Then
                                    rowSOTORDRG.Item("ORDR_REL_SHORT") = "1"
                                    rowSOTORDRG.Item("ORDR_REL_SHORT_OPER") = ASCMAIN1.USER_ID
                                    rowSOTORDRG.Item("ORDR_REL_SHORT_DATE") = DATETIME_STAMP
                                    rowSOTORDRG.Item("ORDR_REL_SHORT_MIN") = Val(rowSOTORDR0.Item("ORDR_AMT_ALLO_CUR") & "")
                                Else
                                    rowSOTORDRG.Item("ORDR_REL_SHORT") = "0"
                                End If
                                rowSOTORDR0.Item("ORDR_REL_SHORT") = rowSOTORDRG.Item("ORDR_REL_SHORT")
                                rowSOTORDR0.Item("ORDR_REL_SHORT_OPER") = rowSOTORDRG.Item("ORDR_REL_SHORT_OPER")
                                '  rowSOTORDR0.Item("ORDR_REL_SHORT_DATE") = rowSOTORDRG.Item("ORDR_REL_SHORT_DATE")
                            End If
                        Next
                        Update_Record_TDA("SOTORDRG")

                        MsgBox("Done")
                        grdSOTORDR0.Selected.Rows.Clear()

                    End If
                    ASCMAIN1.MultiTask_Release(, , 1)

                End If

            Case "Store Configuration Report"
                'Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                'Store_Configuration_Report(ORDR_GROUP_NO)

                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If
                Dim ORDR_GROUP_NOs As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                    Dim ORDR_TYPE As String = grow.Cells("ORDR_TYPE").Value
                    If ORDR_TYPE = "O" Then
                        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                    End If
                Next

                If ORDR_GROUP_NOs.Count <> 0 Then
                    Store_Configuration_Report(ORDR_GROUP_NOs)
                Else
                    MsgBox("No Order Groups Selected", MsgBoxStyle.OkOnly, "Cannot Proceed")
                End If

            Case "Customer Order Summary - OLD"

                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                Dim ORDR_TYPE As String = grd.ActiveRow.Cells("ORDR_TYPE").Value
                Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDR_CUST_PO").Value
                Dim CUST_DC_NO As String = grd.ActiveRow.Cells("CUST_DC_NO").Value & ""

                Dim ORDR_DATE As Date = grd.ActiveRow.Cells("ORDR_DATE").Value
                Dim ORDR_SHIP_DATE As Date = grd.ActiveRow.Cells("ORDR_SHIP_DATE").Value
                Dim ORDR_CANCEL_DATE As Date = grd.ActiveRow.Cells("ORDR_CANCEL_DATE").Value
                Dim CUST_STORE_NO As String = ""
                If chkShowSelectedOrder.Checked Then
                    If grdSOTORDRX.ActiveRow IsNot Nothing Then
                        CUST_STORE_NO = grdSOTORDRX.ActiveRow.Cells("CUST_STORE_NO").Value & ""
                    End If
                End If


                dst.Tables("SOTCORDR").Rows.Clear()
                For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("", "STYLE_CODE,COLOR_CODE")
                    Dim rowSOTCORDR As DataRow = dst.Tables("SOTCORDR").NewRow
                    rowSOTCORDR.Item("STYLE_CODE") = rowSOTORDRS.Item("STYLE_CODE")
                    rowSOTCORDR.Item("COLOR_CODE") = rowSOTORDRS.Item("COLOR_CODE")
                    rowSOTCORDR.Item("STYLE_DESC") = rowSOTORDRS.Item("STYLE_DESC")
                    rowSOTCORDR.Item("CUST_STYLE_CODE") = rowSOTORDRS.Item("CUST_STYLE_CODE")
                    rowSOTCORDR.Item("CUST_COLOR_CODE") = rowSOTORDRS.Item("CUST_COLOR_CODE")
                    rowSOTCORDR.Item("CUST_SIZE_CODE") = rowSOTORDRS.Item("CUST_SIZE_CODE")
                    rowSOTCORDR.Item("CUST_UPC") = rowSOTORDRS.Item("CUST_UPC")
                    rowSOTCORDR.Item("CUST_SKU") = rowSOTORDRS.Item("CUST_SKU")
                    rowSOTCORDR.Item("ORDR") = rowSOTORDRS.Item("ORDR_QTY")
                    rowSOTCORDR.Item("OPEN") = rowSOTORDRS.Item("ORDR_QTY_OPEN")
                    rowSOTCORDR.Item("PICK") = rowSOTORDRS.Item("ORDR_QTY_PICK")
                    rowSOTCORDR.Item("SHIP") = rowSOTORDRS.Item("ORDR_QTY_SHIP")
                    rowSOTCORDR.Item("CANC") = rowSOTORDRS.Item("ORDR_QTY_CANC")
                    rowSOTCORDR.Item("ALLO") = rowSOTORDRS.Item("ORDR_QTY_ALLO")
                    rowSOTCORDR.Item("ORDR_AMT") = rowSOTORDRS.Item("ORDR_AMT")
                    dst.Tables("SOTCORDR").Rows.Add(rowSOTCORDR)
                Next

                Print_Report_Begin()
                Dim SUBT As String = CUST_CODE _
                                     & ", PO " & ORDR_CUST_PO & ", Order Date " & Format(ORDR_DATE, "MM/dd/yy") _
                                     & ", Ship " & Format(ORDR_SHIP_DATE, "MM/dd/yy") _
                                     & ", Cancel " & Format(ORDR_CANCEL_DATE, "MM/dd/yy") _
                                     & IIf(ORDR_TYPE = "O" And chkShowSelectedOrder.Checked, ", Store No " & CUST_STORE_NO, "") _
                                     & IIf(ORDR_TYPE = "O", ", Order Group No " & ORDR_GROUP_NO, ", Reservation " & ORDR_GROUP_NO)
                Generate_Report("SORCORDR", "Customer Order Summary", SUBT, , , , False)
                Print_Report_End()

            Case "Customer Order Summary"

                dst.Tables("SOTCORDR").Rows.Clear()
                dst.Tables("SOTCORDG").Rows.Clear()

                Dim TBL As DataTable = dst.Tables("SOTORDRS").Copy

                'Dim TBL As DataTable = Nothing
                'If grd.Selected.Rows.Count = 1 Or chkShowSelectedOrder.Checked Then
                '    TBL = dst.Tables("SOTORDRS").Copy
                'End If

                If grd.Selected.Rows.Count = 0 Then grd.ActiveRow.Selected = True
                If chkShowSelectedOrder.Checked Then
                    grd.Selected.Rows.Clear()
                    grd.ActiveRow.Selected = True
                End If

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                    Dim ORDR_TYPE As String = grow.Cells("ORDR_TYPE").Value
                    Dim ORDR_CUST_PO As String = grow.Cells("ORDR_CUST_PO").Value & ""

                    If grd.Selected.Rows.Count = 1 Or chkShowSelectedOrder.Checked Then
                    Else
                        Dim sql As String = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' group by ")
                        sql = "Select X.*, SOTORDRS.WIP_IND from (" & sql & ") X, SOTORDRS where SOTORDRS.ORDR_GROUP_NO (+) = '" & ORDR_GROUP_NO & "' and SOTORDRS.STYLE_CODE (+) = X.STYLE_CODE and SOTORDRS.COLOR_CODE (+) = X.COLOR_CODE"
                        Fill_Records("SOTORDRS", "", True, sql)
                    End If
                    Summarize_Group(ORDR_TYPE, ORDR_GROUP_NO)
                Next

                If grd.Selected.Rows.Count = 1 Or chkShowSelectedOrder.Checked Then
                    dst.Tables("SOTORDRS").Rows.Clear()
                    dst.Tables("SOTORDRS").Merge(TBL)
                End If

                Print_Report_Begin()
                Dim SUBT As String
                If grd.Selected.Rows.Count = 1 Or chkShowSelectedOrder.Checked Then
                    With grd.Selected.Rows(0)
                        Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value
                        Dim ORDR_TYPE As String = grd.ActiveRow.Cells("ORDR_TYPE").Value
                        Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
                        Dim CUST_DC_NO As String = grd.ActiveRow.Cells("CUST_DC_NO").Value & ""

                        Dim ORDR_DATE As Date = grd.ActiveRow.Cells("ORDR_DATE").Value
                        Dim ORDR_SHIP_DATE As Date = grd.ActiveRow.Cells("ORDR_SHIP_DATE").Value
                        Dim ORDR_CANCEL_DATE As Date = grd.ActiveRow.Cells("ORDR_CANCEL_DATE").Value

                        Dim CUST_STORE_NO As String = ""
                        If chkShowSelectedOrder.Checked Then
                            If grdSOTORDRX.ActiveRow IsNot Nothing Then
                                CUST_STORE_NO = grdSOTORDRX.ActiveRow.Cells("CUST_STORE_NO").Value & ""
                            End If
                        End If

                        SUBT = CUST_CODE _
                            & ", PO " & ORDR_CUST_PO & " " & Format(ORDR_DATE, "MM/dd/yy") _
                            & ", Ship " & Format(ORDR_SHIP_DATE, "MM/dd/yy") _
                            & ", Cancel " & Format(ORDR_CANCEL_DATE, "MM/dd/yy") _
                            & IIf(ORDR_TYPE = "O" And chkShowSelectedOrder.Checked, ", Store No " & CUST_STORE_NO, "") _
                            & IIf(ORDR_TYPE = "O", ", Order Group No " & ORDR_GROUP_NO, ", Reservation " & ORDR_GROUP_NO)
                    End With
                Else
                    SUBT = CUST_CODE & ", " & grd.Selected.Rows.Count & " Selected POs"
                End If

                Generate_Report("SORCORDR", "Customer Order Summary", SUBT, , , , False)
                Print_Report_End()


            Case "Sales Invoice"
                If grd.Name = "grdSOTSHIP1" Then
                    Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                    TAC.SOCMAIN1.Create_Invoice(Me, SHIP_BOL_NO, False, False, "SHIP_BOL_NO")
                Else
                    Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Value
                    TAC.SOCMAIN1.Create_Invoice(Me, PICK_NO, False, False, "PICK_NO")
                End If
                'Create_Invoice(SHIP_BOL_NO, PICK_NO)

            Case "Pro-Forma Invoice"
                If grd.Name = "grdSOTSHIP1" Then
                    Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                    TAC.SOCMAIN1.Create_Invoice(Me, SHIP_BOL_NO, False, True, "SHIP_BOL_NO")
                Else
                    Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Value
                    TAC.SOCMAIN1.Create_Invoice(Me, PICK_NO, False, True, "PICK_NO")
                End If
                'Create_Invoice(SHIP_BOL_NO, PICK_NO, False, True)

            Case "Export Sales Order Details"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
                ASCMAIN1.sql = "Select ORDR_NO, ORDR_LNO, STYLE_CODE, COLOR_CODE, EDI_DTL_SEQ, ORDR_QTY" & vbCrLf _
                    & ",EDI_DOC_SEQ_NO, CUST_UPC,CUST_SKU, ORDR_QTY NEW_QTY, CUST_UPC NEW_UPC, CUST_SKU NEW_SKU" & vbCrLf _
                    & " from SOTORDR2 where ORDR_NO in (" & vbCrLf _
                    & " Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')"
                Dim TBL As DataTable = ASCDATA1.GetDataTable
                Using F As New ASFMSGBF
                    F.Show_grd(TBL, Me, "Style Details for Orders related to " & CUST_CODE & " PO " & ORDR_CUST_PO)
                End Using

            Case "Show Shipments For Style"
                Dim PeriodDate As DateTime = grd.ActiveRow.Cells("SHIP_PICK_PRINTED").Value
                Dim PeriodDate_Str As String = Format(PeriodDate, "dd-MMM-yyyy")
                Dim CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                Dim Style As String = ""
                Using F As New ASFMSGBF
                    Style = F.Get_txt_from_User("Enter Style Prefix", "WALMART RECAP")
                End Using
                ASCMAIN1.sql = $"select o0.ORDR_CUST_PO, sum(i2.ORDR_QTY_SHIP) Shipped from SOTORDR0 O0
                                join SOTINVH1 I1 on (O0.ORDR_CUST_PO = I1.ORDR_CUST_PO and O0.CUST_CODE = I1.CUST_CODE)
                                join SOTINVH2 I2 on (I1.INV_TYPE = I2.INV_TYPE and I1.INV_NO = I2.INV_NO)
                                where o0.CUST_CODE = '{CUST_CODE}'  and I1.INV_NO_REV is null and I1.INV_NO_REV_BY is null
                                and O0.ORDR_DATE between TRUNC(TO_DATE('{PeriodDate_Str}', 'DD-MON-YYYY'), 'MM') and ADD_MONTHS(TRUNC(TO_DATE('{PeriodDate_Str}', 'DD-MON-YYYY'), 'MM'), 1)
                                and I2.STYLE_CODE like '{Style}%'
                                group by o0.ORDR_CUST_PO"
                Dim TBL As DataTable = ASCDATA1.GetDataTable()
                Using F As New ASFMSGBF
                    F.Show_grd(TBL, Me, "Walmart summary for " & Style)
                End Using

            Case "Convert CTF to Reservation"
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
                Dim ORDR_CNT As Integer = Val(grd.ActiveRow.Cells("ORDR_CNT").Value & "")
                If ORDR_CNT = 1 Then

                    ASCMAIN1.sql = "Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                    Dim ORDR_NO As String = ASCDATA1.GetDataValue

                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub

                    If MsgBox("Are you sure that you want to convert " & ORDR_CUST_PO _
                              & vbCrLf & " Order Group " & ORDR_GROUP_NO _
                              & vbCrLf & " to a Reservation?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        ASCDATA1.ExecuteSP("SOPRSRV1_CTF", "V", New Object() {ORDR_NO}, New String() {"ORDR_NO_X"})
                        chkReservations.Checked = True
                        Load_SOTORDR0("", CUST_CODE)
                        Setup_SOTORDR0()
                    End If

                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Wave Inquiry"
                Dim WAVE_NO As String = grd.ActiveRow.Cells("WAVE_NO").Value
                If WAVE_NO <> "" Then
                    Context_Launch("View", WAVE_NO, e.Tool.Key, "WHFWAVEI")
                End If

            Case "Carton Pack Configuration"
                If grdSOTORDR0.Selected.Rows.Count = 0 Then
                    If grdSOTORDR0.ActiveRow IsNot Nothing AndAlso grdSOTORDR0.ActiveRow.IsDataRow Then
                        grdSOTORDR0.ActiveRow.Selected = True
                    End If
                End If

                If grdSOTORDR0.Selected.Rows.Count = 0 Then
                    MsgBox("You must select all of the order groups for which you want to prepare a carton pack configuration")
                Else

                    ' NEED TO MULTI-TASK LOCK CUSTOMER

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

                    Dim ORDR_GROUP_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                        Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                        Dim CUST_CODE_selected_row As String = grow.Cells("CUST_CODE").Value
                        If CUST_CODE <> CUST_CODE_selected_row Then
                            MsgBox("Cannot mix Customers (" & CUST_CODE & "," & CUST_CODE_selected_row & ") when performing Pack Configuration", MsgBoxStyle.OkOnly, "Cannot Proceed")
                            Exit Sub
                        End If

                        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                        Dim ORDR_CNT_PICK As Int32 = Val(rowSOTORDR0.Item("ORDR_CNT_PICK") & "")
                        Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTORDR0.Item("ORDR_QTY_SHIP") & "")
                        Dim ORDR_CNT_OPEN As Int32 = Val(rowSOTORDR0.Item("ORDR_CNT_OPEN") & "")
                        If ORDR_CNT_PICK <> 0 Or ORDR_QTY_SHIP <> 0 Or ORDR_CNT_OPEN = 0 Then
                            MsgBox("Cannot pack orders that have been released or shipped or are not open (See Order Group " & ORDR_GROUP_NO & ")", MsgBoxStyle.OkOnly, "Cannot Proceed")
                            Exit Sub
                        End If
                        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                    Next

                    ASCMAIN1.sql = "Select * from SOTPCKP2" & vbCrLf _
                        & " where NVL(PACK_GROUP_STATUS,'?') <> 'D'" & vbCrLf _
                        & "   and ORDR_GROUP_NO in ('" & Join(ORDR_GROUP_NOs.ToArray, "','") & "')"
                    Dim tblSOTPCKP2_existing As DataTable = ASCDATA1.GetDataTable
                    If tblSOTPCKP2_existing.Rows.Count <> 0 Then
                        If MsgBox("Carton Pack Data Exists for 1 or more of the Order Groups" _
                                  & vbCrLf & "  that you have just selected." _
                                  & vbCrLf & vbCrLf & "These Carton Pack Definitions will be Deleted and Replaced" _
                                  & vbCrLf & "  by the new Pack Configurations that you Create." _
                                  & vbCrLf & vbCrLf & "OK to Continue?", MsgBoxStyle.YesNo,
                                  "Carton Packs have already been created") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    Using F As New TAC.SOFCARTP
                        F.ORDR_GROUP_NOs = ORDR_GROUP_NOs
                        F.frm = Me
                        F.CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        F.tblSOTPCKP2_existing = tblSOTPCKP2_existing
                        F.ShowDialog()
                        'If F.STYLE_CODE <> "" Then
                        '    Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
                        'End If
                    End Using

                    grdSOTORDR0.Selected.Rows.Clear()


                End If
            Case "Customer Order Status"
                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                Dim ORDR_GROUP_NOs_to_batch As New List(Of String)
                '   For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Text
                    ORDR_GROUP_NOs_to_batch.Add(ORDR_GROUP_NO)
                Next

                Using FRM As New SOFCORS1
                    FRM.ORDR_GROUP_NOs = ORDR_GROUP_NOs_to_batch
                    FRM.CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    FRM.ShowDialog()
                End Using

               ' grdSOTORDR0.Selected.Rows.Clear()


            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")
                End If


            Case "Summary by DC"

                Dim EDI_CONS_NO As String = grd.ActiveRow.Cells("EDI_CONS_NO").Value & ""
                Dim ORDR_CNT_PICK As Integer = Val(grd.ActiveRow.Cells("ORDR_CNT_PICK").Value & "")
                Dim ORDR_CNT_OPEN As Integer = Val(grd.ActiveRow.Cells("ORDR_CNT_OPEN").Value & "")
                Dim qtyC As String = ""
                If ORDR_CNT_PICK <> 0 Then
                    qtyC = "ORDR_QTY_PICK"
                ElseIf ORDR_CNT_OPEN <> 0 Then
                    qtyC = "ORDR_QTY_OPEN"
                Else
                    qtyC = "ORDR_QTY_SHIP"
                End If

                ASCMAIN1.sql = "" _
                    & "Select CUST_DC_NO, MIN (PO29) PO29, MIN (PO30) PO30" & vbCrLf _
                    & ", SUM (STORES) STORES, SUM (ORD29) ORD29, SUM (ORD30) ORD30 " & vbCrLf _
                    & ", SUM (UNITS29) UNITS29, SUM (UNITS30) UNITS30" & vbCrLf _
                    & " from (" & vbCrLf _
                    & "Select SOTORDR0.CUST_DC_NO, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                    & ", MIN (DECODE(SOTORDR0.ORDR_DEPT,'00029',SOTORDR0.ORDR_CUST_PO,NULL)) PO29" & vbCrLf _
                    & ", MIN (DECODE(SOTORDR0.ORDR_DEPT,'00030',SOTORDR0.ORDR_CUST_PO,NULL)) PO30" & vbCrLf _
                    & ", COUNT (DISTINCT CASE WHEN NVL(SOTORDR1.ORDR_STATUS,'?') <> 'C' THEN SOTORDR1.CUST_STORE_NO END) STORES" & vbCrLf _
                    & ", SUM (DECODE(SOTORDR0.ORDR_DEPT,'00029',1,0)) ORD29" & vbCrLf _
                    & ", SUM (DECODE(SOTORDR0.ORDR_DEPT,'00030',1,0)) ORD30" & vbCrLf _
                    & ", 0 UNITS29" & vbCrLf _
                    & ", 0 UNITS30" & vbCrLf _
                    & " from SOTORDR0,SOTORDR1,EDT850T1 " & vbCrLf _
                    & " where EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO" & vbCrLf _
                    & "   and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                    & "   and EDT850T1.EDI_CONS_NO = '" & EDI_CONS_NO & "'" & vbCrLf _
                    & " group by SOTORDR0.CUST_DC_NO, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                    & " UNION " & vbCrLf _
                    & "Select SOTORDR0.CUST_DC_NO, NULL CUST_STORE_NO" & vbCrLf _
                    & ", MIN (DECODE(SOTORDR0.ORDR_DEPT,'00029',SOTORDR0.ORDR_CUST_PO,NULL)) PO29" & vbCrLf _
                    & ", MIN (DECODE(SOTORDR0.ORDR_DEPT,'00030',SOTORDR0.ORDR_CUST_PO,NULL)) PO30" & vbCrLf _
                    & ", 0 STORES" & vbCrLf _
                    & ", 0 ORD29" & vbCrLf _
                    & ", 0 ORD30" & vbCrLf _
                    & ", SUM (DECODE(SOTORDR0.ORDR_DEPT,'00029'," & qtyC & ",0)) UNITS29" & vbCrLf _
                    & ", SUM (DECODE(SOTORDR0.ORDR_DEPT,'00030'," & qtyC & ",0)) UNITS30" & vbCrLf _
                    & " from SOTORDR0,EDT850T1 " & vbCrLf _
                    & " where EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO" & vbCrLf _
                    & "   and EDT850T1.EDI_CONS_NO = '" & EDI_CONS_NO & "'" & vbCrLf _
                    & " group by SOTORDR0.CUST_DC_NO" & vbCrLf _
                    & ")  group by CUST_DC_NO" & vbCrLf _
                    & " order by CUST_DC_NO"

                Dim DT As DataTable = ASCDATA1.GetDataTable
                Dim xls_filename As String = ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & "_XLS") & ".XLS"
                Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

                Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
                Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1")
                'range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)
                ' format columns in xls first for best results
                Dim RX As Int32 = 0
                Dim CX As Int32 = 0


                With oSheet.Cells(Excel_Cell(RX + 0 + 1, CX + 0 + 1) & ":" & Excel_Cell(RX + 0 + 1, CX + DT.Columns.Count - 1 + 1))
                    .Font.Color = SpreadsheetGear.Colors.White
                    .Font.Bold = True
                    .Interior.Color = SpreadsheetGear.Colors.Blue
                End With

                For c As Integer = 0 To DT.Columns.Count - 1

                    Dim COLUMN_NAME As String = DT.Columns(c).ColumnName
                    If COLUMN_NAME = "STORES" Or COLUMN_NAME = "ORD29" Or COLUMN_NAME = "ORD30" Or COLUMN_NAME.StartsWith("UNITS") Then

                        With oSheet.Cells(Excel_Cell(RX + 0 + 1, CX + c + 1)).EntireColumn
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                            .NumberFormat = "#,##0"
                            If COLUMN_NAME.StartsWith("UNITS") Then
                                .ColumnWidth = .ColumnWidth * 2
                            End If

                        End With

                        oSheet.Cells(Excel_Cell(RX + 1, CX + c + 1)).Formula = "=SUBTOTAL(9," & Excel_Cell(RX + 1 + 1 + 1, CX + c + 1) & ":" & Excel_Cell(RX + 1 + DT.Rows.Count, CX + c + 1) & ")"
                    Else
                        With oSheet.Cells(Excel_Cell(RX + 0 + 1, CX + c + 1)).EntireColumn
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                            .NumberFormat = "@"
                            If COLUMN_NAME = "PO29" Or COLUMN_NAME = "PO30" Then
                                .ColumnWidth = .ColumnWidth * 2
                            End If

                        End With
                    End If
                Next

                range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

                With oSheet.Cells(Excel_Cell(RX + 1, CX + 1)).EntireRow
                    .Insert(SpreadsheetGear.InsertShiftDirection.Down)
                End With

                For c As Integer = 0 To DT.Columns.Count - 1
                    Dim COLUMN_NAME As String = DT.Columns(c).ColumnName
                    If COLUMN_NAME = "STORES" Or COLUMN_NAME = "ORD29" Or COLUMN_NAME = "ORD30" Or COLUMN_NAME.StartsWith("UNITS") Then
                        oSheet.Cells(Excel_Cell(RX + 1, CX + c + 1)).Formula = "=SUBTOTAL(9," & Excel_Cell(RX + 1 + 1 + 1, CX + c + 1) & ":" & Excel_Cell(RX + 1 + 1 + DT.Rows.Count, CX + c + 1) & ")"
                    End If
                Next

                oWB.SaveAs(ASCMAIN1.Folders("Temp") & xls_filename, SpreadsheetGear.FileFormat.Excel8)
                oWB.Close()
                range = Nothing
                oSheet = Nothing
                oWB = Nothing
                Dim p As Process = Process.Start(ASCMAIN1.Folders("Temp") & xls_filename)

            Case "Show Details"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                '   SplitContainer1.Panel2Collapsed = Not tlb_sbt.Checked
                grdSOTCART1.DisplayLayout.Bands(0).Columns("PALLET_NO").Hidden = Not tlb_sbt.Checked
                grdSOTCART1.DisplayLayout.Bands(0).Columns("SHIP_TRAILER_NO").Hidden = Not tlb_sbt.Checked
                grdSOTCART1.DisplayLayout.Bands(0).Columns("PALLET_INIT_DATE").Hidden = Not tlb_sbt.Checked
                grdSOTCART1.DisplayLayout.Bands(0).Columns("PALLET_INIT_OPER").Hidden = Not tlb_sbt.Checked
                grdSOTCART1.DisplayLayout.Bands(0).Columns("CUST_DC_NO").Hidden = Not tlb_sbt.Checked

            Case "Rebuild Order Summary"
                If grd.ActiveRow IsNot Nothing Then
                    Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                    ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
                    ASCDATA1.ExecuteSQL()
                End If

            Case "Show Cart Stats"
                Dim hideCols As Boolean = True
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    hideCols = False
                    ASCMAIN1.Progress("Now loading Carton Details")
                    UpdateCartStats()
                    ASCMAIN1.Progress("", "")
                End If
                For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTCARTP.DisplayLayout.Bands(0).Columns
                    Select Case gcol.Key
                        '"WALMART_REC"
                        Case "TRACKING_NO", "CARTON_WEIGHT", "CARTON_VALUE", "PALLET_VALUE", "SCAN_TIME", "SHIP_LOAD_NO", "BILL_OF_LADING_NO", "MASTER_SHIP_BOL_NO", "MULTI_PO"
                            gcol.Hidden = hideCols
                    End Select
                    'gcol.Format = "#,##0"
                Next

            Case "Show Comments"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Dim frmASFMSGBF As New ASFMSGBF
                    Dim SepChar As String = frmASFMSGBF.Get_txtblock_from_User("Char to Seperate Date Or Blank For No Date", "Change Notes", ":", False, 5)

                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Fetching Comments", "")
                    Application.DoEvents()

                    For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select()
                        For c As Int64 = 1 To 10
                            Dim COLN As String = Format(c, "00")
                            rowSOTORDR0.Item($"COMMENT{COLN}") = ""
                        Next

                        Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO").ToString & String.Empty
                        'dst.Tables("SOTORDRC").Clear()
                        Dim sql As New Text.StringBuilder With {.Length = 0}
                        sql.AppendLine("SELECT ORDR_NO")
                        sql.AppendLine("FROM SOTORDR1")
                        sql.AppendLine($"WHERE ORDR_GROUP_NO = '{ORDR_GROUP_NO}'")
                        Dim tblORDR_NO As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)

                        Dim maxc As Int64 = 0
                        For Each rowORDR_NO As DataRow In tblORDR_NO.Rows
                            If maxc > 10 Then Exit For
                            Fill_Records("SOTORDRC", rowORDR_NO.Item("ORDR_NO").ToString & String.Empty)
                            For Each rowSOTORDRC As DataRow In dst.Tables("SOTORDRC").Select("", "ORDR_CLNO DESC")
                                maxc += 1
                                If maxc > 10 Then Exit For
                                Dim ORDR_COMMENT As String = rowSOTORDRC.Item("ORDR_COMMENT").ToString & String.Empty
                                Dim INIT_D As String = ""
                                If IsDate(rowSOTORDRC.Item("INIT_DATE").ToString & String.Empty) Then
                                    INIT_D = Format(CDate(rowSOTORDRC.Item("INIT_DATE").ToString & String.Empty), "MM/dd/yy")
                                End If
                                Dim DT_CMT As String = ""
                                If INIT_D.Length > 0 Then
                                    If SepChar.Length > 0 Then
                                        DT_CMT = $"{INIT_D}{SepChar} {ORDR_COMMENT}"
                                    Else
                                        DT_CMT = ORDR_COMMENT
                                    End If
                                Else
                                    DT_CMT = ORDR_COMMENT
                                End If
                                Dim COLN As String = Format(maxc, "00")
                                rowSOTORDR0.Item($"COMMENT{COLN}") = DT_CMT
                            Next
                        Next
                    Next
                    For i As Int64 = 1 To 10
                        Dim COLN As String = Format(i, "00")
                        With grdSOTORDR0.DisplayLayout.Bands(0).Columns
                            .Item($"COMMENT{COLN}").Hidden = False
                            .Item($"COMMENT{COLN}").Width = 350
                        End With
                    Next
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                    Application.DoEvents()
                Else
                    For i As Int64 = 1 To 10
                        Dim COLN As String = Format(i, "00")
                        With grdSOTORDR0.DisplayLayout.Bands(0).Columns
                            .Item($"COMMENT{COLN}").Hidden = True
                        End With
                    Next
                End If
            Case "Export With Stats"

                Dim SHIP_LOAD_NO As String = ""
                Dim CUST_DC_NO As String = ""
                Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
                Dim FILE_NAME As String = ""

                If dst.Tables.Item("SOTCARTP").Rows.Count > 0 Then
                    SHIP_LOAD_NO = dst.Tables.Item("SOTCARTP").Rows(0).Item("SHIP_LOAD_NO").ToString & String.Empty
                    CUST_DC_NO = dst.Tables.Item("SOTCARTP").Rows(0).Item("CUST_DC_NO").ToString & String.Empty
                End If

                Dim DL As String = ASCMAIN1.Folders("Temp") & "DISPLAY.TXT"
                If IO.File.Exists(DL) Then
                    IO.File.Delete(DL)
                End If
                grdSOTCARTP.DisplayLayout.Save(DL)
                PrepareGridForExport(grdSOTCARTP)

                Dim WB As Infragistics.Documents.Excel.Workbook = Export_to_Excel(grdSOTCARTP)

                grdSOTCARTP.DisplayLayout.Load(DL)

                Dim TITLE As String = ""

                Dim LastRow As Int64 = 0
                Dim FirstRow As Int64 = 0
                For i As Int64 = 1 To 10000
                    If WB.Worksheets(0).GetCell("A" & i).GetText = "DC No" Or WB.Worksheets(0).GetCell("A" & i).GetText = "Pallet" Then
                        FirstRow = i + 1
                    End If
                    If WB.Worksheets(0).GetCell("A" & i).GetText = "Totals" And FirstRow > 0 Then
                        'WB.Worksheets(0).Rows.Remove(i - 1, 2)
                        LastRow = i
                        Exit For
                    End If
                Next
                If LastRow = 0 Or FirstRow = 0 Then
                    MsgBox("Can Not Find First or Last Row To Create Summary", vbCritical, "Problems")
                Else
                    Dim LAST_PALLET As String = ""
                    Dim LAST_PALLET_TRAIL As String = ""
                    Dim LAST_PALLET_VAL As Double = 0
                    Dim LAST_PALLET_CARTS As Int64 = 0

                    Dim sql As New System.Text.StringBuilder With {.Length = 0}
                    sql.AppendLine("SELECT")
                    sql.AppendLine("PALLET_NO,")
                    sql.AppendLine("SHIP_TRAILER_NO,")
                    sql.AppendLine("1000000.00 AS PALLET_VALUE,")
                    sql.AppendLine("10000 AS CART_COUNT,")
                    sql.AppendLine("10000 AS LAST_PAL_LINE")
                    sql.AppendLine("FROM WHTPALT1")
                    sql.AppendLine("WHERE ROWNUM < 0")
                    Dim tblSUM As DataTable = ASCDATA1.GetDataTable(sql.ToString())
                    TITLE = $"Carton Scanning Report For PO# {ORDR_CUST_PO}, Customer DC {CUST_DC_NO}, Load# {SHIP_LOAD_NO}"
                    If PO_CARTON_COMBINED <> "" Then
                        TITLE = $"Carton Scanning Report For DSDC Multi PO# {ORDR_CUST_PO}  &  {PO_CARTON_COMBINED}  , Customer DC {CUST_DC_NO}, Load# {SHIP_LOAD_NO}"

                    End If
                    WB.Worksheets(0).Rows(0).Cells(0).Value = ""
                    WB.Worksheets(0).Rows(0).Cells(1).Value = ""
                    WB.Worksheets(0).Rows(1).Cells(0).Value = TITLE
                    FILE_NAME = TITLE.Replace("#", "").Replace(",", "") & ".xls"

                    For rw As Int64 = FirstRow To LastRow
                        Dim PLT As String = WB.Worksheets(0).GetCell("A" & rw).GetText()
                        If LAST_PALLET = "" Then
                            LAST_PALLET = PLT
                        End If
                        If LAST_PALLET = PLT Then
                            LAST_PALLET_TRAIL = WB.Worksheets(0).GetCell("C" & rw).GetText()
                            LAST_PALLET_VAL = Val(WB.Worksheets(0).GetCell("J" & rw).GetText().Replace(",", ""))
                            LAST_PALLET_CARTS = LAST_PALLET_CARTS + 1
                        Else
                            Dim rowSUM As DataRow = tblSUM.NewRow
                            rowSUM.Item("PALLET_NO") = LAST_PALLET
                            rowSUM.Item("SHIP_TRAILER_NO") = LAST_PALLET_TRAIL
                            rowSUM.Item("PALLET_VALUE") = LAST_PALLET_VAL
                            rowSUM.Item("CART_COUNT") = LAST_PALLET_CARTS
                            rowSUM.Item("LAST_PAL_LINE") = rw
                            tblSUM.Rows.Add(rowSUM)
                            LAST_PALLET = PLT
                            LAST_PALLET_TRAIL = WB.Worksheets(0).GetCell("D" & rw).GetText()
                            LAST_PALLET_VAL = Val(WB.Worksheets(0).GetCell("J" & rw).GetText().Replace(",", ""))
                            LAST_PALLET_CARTS = 1
                        End If
                    Next

                    Dim PALLET_COUNT As Int64 = tblSUM.Rows.Count
                    Dim SumRow As Int64 = LastRow + 4
                    Dim CART_TOT As Int64 = 0
                    Dim PALLET_VAL As Decimal = 0
                    If PALLET_COUNT > 0 Then

                        WB.Worksheets(0).Rows(SumRow).SetCellValue(3, "Pallet")
                        WB.Worksheets(0).Rows(SumRow).Cells(4).Value = "Trailer No"
                        WB.Worksheets(0).Rows(SumRow).Cells(5).Value = "Pallet Value"
                        WB.Worksheets(0).Rows(SumRow).Cells(6).Value = "# Cartons"
                        WB.Worksheets(0).Rows(SumRow).Cells(3).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
                        WB.Worksheets(0).Rows(SumRow).Cells(4).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
                        WB.Worksheets(0).Rows(SumRow).Cells(5).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
                        WB.Worksheets(0).Rows(SumRow).Cells(6).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
                        SumRow += 1

                        For Each rowSUM As DataRow In tblSUM.Select()
                            WB.Worksheets(0).Rows(SumRow).Cells(3).Value = rowSUM.Item("PALLET_NO").ToString & String.Empty
                            WB.Worksheets(0).Rows(SumRow).Cells(4).Value = rowSUM.Item("SHIP_TRAILER_NO").ToString & String.Empty
                            WB.Worksheets(0).Rows(SumRow).Cells(5).Value = Val(rowSUM.Item("PALLET_VALUE").ToString & String.Empty)
                            WB.Worksheets(0).Rows(SumRow).Cells(5).CellFormat.FormatString = "###,###,##0.00"
                            WB.Worksheets(0).Rows(SumRow).Cells(6).Value = Val(rowSUM.Item("CART_COUNT").ToString & String.Empty)
                            CART_TOT = CART_TOT + Val((rowSUM.Item("CART_COUNT").ToString & String.Empty).Replace(",", ""))
                            PALLET_VAL = PALLET_VAL + Val(rowSUM.Item("PALLET_VALUE").ToString & String.Empty)
                            SumRow += 1
                        Next
                        WB.Worksheets(0).Rows(SumRow).Cells(3).Value = "Grand Total"
                        WB.Worksheets(0).Rows(SumRow).Cells(3).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
                        WB.Worksheets(0).Rows(SumRow).Cells(5).Value = PALLET_VAL
                        WB.Worksheets(0).Columns(5).Width = 4000
                        WB.Worksheets(0).Rows(SumRow).Cells(5).CellFormat.FormatString = "###,###,##0.00"
                        WB.Worksheets(0).Rows(SumRow).Cells(5).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
                        WB.Worksheets(0).Rows(SumRow).Cells(6).Value = CART_TOT
                        WB.Worksheets(0).Rows(SumRow).Cells(6).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
                    End If
                    SumRow += 3
                    Dim Verbiage As String = $"All {CART_TOT} cartons were scanned to {PALLET_COUNT} pallets (see attached Scanning Report).   We ship Collect, & the Cases Shipped equal Cases Invoiced. (See attached signed BOL)"
                    WB.Worksheets(0).Rows(SumRow).SetCellValue(0, Verbiage)

                    'Now Add Pallet Lines
                    WB.Worksheets(0).Rows.Remove(LastRow - 1, 2)
                    Dim rowsAdded As Int64 = 0
                    For Each rowSUM As DataRow In tblSUM.Select()
                        Dim AddRow As Int64 = Val(rowSUM.Item("LAST_PAL_LINE").ToString & String.Empty)
                        WB.Worksheets(0).Rows.Insert(AddRow - 1 + rowsAdded, 2)
                        WB.Worksheets(0).Rows(AddRow - 1 + rowsAdded).Cells(0).Value = Val(rowSUM.Item("CART_COUNT").ToString & String.Empty)
                        WB.Worksheets(0).Rows(AddRow - 1 + rowsAdded).Cells(8).Value = Val(rowSUM.Item("PALLET_VALUE").ToString & String.Empty)
                        For i As Int64 = 0 To 12
                            WB.Worksheets(0).Rows(AddRow - 1 + rowsAdded).Cells(i).CellFormat.TopBorderStyle = Infragistics.Documents.Excel.CellBorderLineStyle.Medium
                            WB.Worksheets(0).Rows(AddRow - 1 + rowsAdded).Cells(i).CellFormat.BottomBorderStyle = Infragistics.Documents.Excel.CellBorderLineStyle.Double
                            WB.Worksheets(0).Rows(AddRow - 1 + rowsAdded).Cells(i).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
                        Next
                        rowsAdded += 2
                    Next

                    'Freeze Top Column
                    WB.Worksheets(0).DisplayOptions.PanesAreFrozen = True
                    WB.Worksheets(0).DisplayOptions.FrozenPaneSettings.FrozenRows = FirstRow - 1

                End If
                Try
                    If IO.File.Exists(ASCMAIN1.Folders("Temp") & FILE_NAME) Then
                        IO.File.Delete(ASCMAIN1.Folders("Temp") & FILE_NAME)
                    End If

                    WB.Save(ASCMAIN1.Folders("Temp") & FILE_NAME)
                    WB = Nothing
                    Dim p As Process = Process.Start(ASCMAIN1.Folders("Temp") & FILE_NAME)
                Catch ex As Exception

                End Try

        End Select
    End Sub

    Private Sub PrepareGridForExport(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid)
        For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
            gcol.Hidden = True
            gcol.Header.SetVisiblePosition(1000, False)
        Next
        grd.DisplayLayout.Bands(0).Columns.Item("PALLET_NO").Header.VisiblePosition = 1
        grd.DisplayLayout.Bands(0).Columns.Item("PALLET_NO").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("CART_NO").Header.VisiblePosition = 2
        grd.DisplayLayout.Bands(0).Columns.Item("CART_NO").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("SHIP_TRAILER_NO").Header.VisiblePosition = 3
        grd.DisplayLayout.Bands(0).Columns.Item("SHIP_TRAILER_NO").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("SHIP_BOL_NO").Header.VisiblePosition = 4
        grd.DisplayLayout.Bands(0).Columns.Item("SHIP_BOL_NO").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("CUST_STORE_NO").Header.VisiblePosition = 5
        grd.DisplayLayout.Bands(0).Columns.Item("CUST_STORE_NO").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("CART_TOTAL_UNITS").Header.Caption = "Units Shipped"
        grd.DisplayLayout.Bands(0).Columns.Item("CART_TOTAL_UNITS").Header.VisiblePosition = 6
        grd.DisplayLayout.Bands(0).Columns.Item("CART_TOTAL_UNITS").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("PKG_CODE").Header.Caption = "Carton ID#"
        grd.DisplayLayout.Bands(0).Columns.Item("PKG_CODE").Header.VisiblePosition = 7
        grd.DisplayLayout.Bands(0).Columns.Item("PKG_CODE").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("CARTON_WEIGHT").Header.VisiblePosition = 8
        grd.DisplayLayout.Bands(0).Columns.Item("CARTON_WEIGHT").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("CARTON_VALUE").Header.VisiblePosition = 9
        grd.DisplayLayout.Bands(0).Columns.Item("CARTON_VALUE").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("PALLET_VALUE").Header.VisiblePosition = 10
        grd.DisplayLayout.Bands(0).Columns.Item("PALLET_VALUE").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("PALLET_INIT_OPER").Header.VisiblePosition = 11
        grd.DisplayLayout.Bands(0).Columns.Item("PALLET_INIT_OPER").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("PALLET_INIT_DATE").Header.VisiblePosition = 12
        grd.DisplayLayout.Bands(0).Columns.Item("PALLET_INIT_DATE").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("SCAN_TIME").Header.VisiblePosition = 13
        grd.DisplayLayout.Bands(0).Columns.Item("SCAN_TIME").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("MULTI_PO").Header.VisiblePosition = 14
        grd.DisplayLayout.Bands(0).Columns.Item("MULTI_PO").Hidden = False
        grd.DisplayLayout.Bands(0).Columns.Item("CART_TRACKING_NO").Header.VisiblePosition = 15
        grd.DisplayLayout.Bands(0).Columns.Item("CART_TRACKING_NO").Hidden = False


        'For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
        '    Select Case gcol.Key
        '        Case "CUST_DC_NO", "PICK_NO", "CART_TOTAL_UNITS_REL", "TRACKING_NO", "SHIP_LOAD_NO"
        '            gcol.Hidden = True
        '        Case "PALLET_NO"
        '            gcol.Header.VisiblePosition = 1
        '        Case "CART_NO"
        '            gcol.Header.VisiblePosition = 2
        '        Case "SHIP_TRAILER_NO"
        '            gcol.Header.VisiblePosition = 3
        '        Case "SHIP_BOL_NO"
        '            gcol.Header.VisiblePosition = 4
        '        Case "CUST_STORE_NO"
        '            gcol.Header.VisiblePosition = 5
        '        Case "CART_TOTAL_UNITS"
        '            gcol.Header.Caption = "Units Shipped"
        '            gcol.Header.VisiblePosition = 6
        '        Case "PKG_CODE"
        '            gcol.Header.Caption = "Carton ID#"
        '            gcol.Header.VisiblePosition = 7
        '        Case "CARTON_WEIGHT"
        '            gcol.Header.VisiblePosition = 8
        '        Case "CARTON_VALUE"
        '            gcol.Header.VisiblePosition = 9
        '        Case "PALLET_VALUE"
        '            gcol.Header.VisiblePosition = 10
        '        Case "PALLET_INIT_OPER"
        '            gcol.Header.VisiblePosition = 11
        '        Case "PALLET_INIT_DATE"
        '            gcol.Header.VisiblePosition = 12
        '        Case "SCAN_TIME"
        '            gcol.Header.VisiblePosition = 13
        '    End Select
        'Next
    End Sub

#End Region

    Sub Summarize_Group(ORDR_TYPE As String, ORDR_GROUP_NO As String)

        Dim rowSOTORDR0 As DataRow = dst.Tables("SOTORDR0").Rows.Find(New Object() {ORDR_TYPE, ORDR_GROUP_NO})
        dst.Tables("SOTCORDG").Rows.Add(rowSOTORDR0.ItemArray)

        For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("", "STYLE_CODE,COLOR_CODE")
            Dim STYLE_CODE As String = rowSOTORDRS.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDRS.Item("COLOR_CODE")
            Dim rowSOTCORDR As DataRow
            rowSOTCORDR = dst.Tables("SOTCORDR").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If rowSOTCORDR Is Nothing Then
                rowSOTCORDR = dst.Tables("SOTCORDR").NewRow
                rowSOTCORDR.Item("STYLE_CODE") = rowSOTORDRS.Item("STYLE_CODE")
                rowSOTCORDR.Item("STYLE_DESC") = rowSOTORDRS.Item("STYLE_DESC")
                rowSOTCORDR.Item("COLOR_CODE") = rowSOTORDRS.Item("COLOR_CODE")
                dst.Tables("SOTCORDR").Rows.Add(rowSOTCORDR)
            End If

            rowSOTCORDR.Item("ORDR") += Val(rowSOTORDRS.Item("ORDR_QTY") & "")
            rowSOTCORDR.Item("OPEN") += Val(rowSOTORDRS.Item("ORDR_QTY_OPEN") & "")
            rowSOTCORDR.Item("PICK") += Val(rowSOTORDRS.Item("ORDR_QTY_PICK") & "")
            rowSOTCORDR.Item("SHIP") += Val(rowSOTORDRS.Item("ORDR_QTY_SHIP") & "")
            rowSOTCORDR.Item("CANC") += Val(rowSOTORDRS.Item("ORDR_QTY_CANC") & "")
            rowSOTCORDR.Item("ORDR_AMT") += Val(rowSOTORDRS.Item("ORDR_AMT") & "")
            'dst.Tables("SOTCORDR").Rows.Add(rowSOTCORDR)
        Next

    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select")
                End If

            Case "SREP_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    grdSOTORDR0.Tag = "SREP_CODE"
                    Load_SOTORDR0(Absx1.txtFor("SREP_CODE").Text)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "CUST_CODE"
            '    If Not ScreenMode Then
            '        Load_SOTORDRX()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("Select")
            Case "SREP_CODE"
                grdSOTORDR0.Tag = "SREP_CODE"
                Load_SOTORDR0(Absx1.txtFor("SREP_CODE").Text)
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_SOTORDR0(Optional PARM1 As String = "", Optional CUST_CODE As String = "")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Order Summary", "")

        Dim sqlReservations As String = " UNION " & vbCrLf _
                & "Select 'R' ORDR_TYPE, SOTRSRV1.RSRV_NO ORDR_GROUP_NO, SOTRSRV1.CUST_CODE, SOTRSRV1.ORDR_CUST_PO" & vbCrLf _
                & ", NULL CUST_DC_NO, SOTRSRV1.ORDR_DEPT, NULL EDI_MERCH_TYPE, SOTRSRV1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTRSRV1.INIT_DATE ORDR_DATE" & vbCrLf _
                & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTRSRV1.ORDR_ORIG_SHIP_DATE, SOTRSRV1.ORDR_ORIG_CANCEL_DATE" & vbCrLf _
                & ", SOTRSRV1.WHSE_CODE, SOTRSRV1.SREP_CODE" & vbCrLf _
                & ", 'RSV' ORDR_TYPE_CODE, 'K' ORDR_SOURCE, NULL EDI_DOC_SEQ_NO" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY_OPEN * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN" & vbCrLf _
                & ", SUM (0) ORDR_AMT_PICK, SUM (0) ORDR_AMT_SHIP" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY_CANC * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_CANC" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY) ORDR_QTY, SUM (SOTRSRV2.RSRV_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (0) ORDR_QTY_PICK, SUM (0) ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", SUM (1) ORDR_CNT, SUM (1) ORDR_CNT_OPEN, SUM (0) ORDR_CNT_PICK" & vbCrLf _
                & ", NULL ORDR_DATE_RECD, SOTRSRV1.RSRV_PRIORITY ORDR_PRIORITY, NULL ORDR_ARRIVAL_DATE, NULL ORDR_LAST_ARRIVAL_DATE" & vbCrLf _
                & ", NULL ORDR_NO_MIN, NULL ORDR_NO_MAX, NULL ORDR_RELEASE_AVAIL_MIN, NULL ORDR_RELEASE_AVAIL_MAX" & vbCrLf _
                & ", NULL ORDR_REL_SHORT, NULL ORDR_REL_SHORT_OPER, NULL ORDR_REL_ACTION_DATE, NULL ORDR_REL_ACTION_OPER" & vbCrLf _
                & ", '0' EDI_CONS_NO" & vbCrLf _
                & ", ' ' PACK_NO" & vbCrLf _
                & " from SOTRSRV1, SOTRSRV2" & vbCrLf _
                & " where SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & "   and SOTRSRV2.RSRV_QTY_OPEN <> 0" & vbCrLf _
                & " group by SOTRSRV1.RSRV_NO, SOTRSRV1.CUST_CODE, SOTRSRV1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTRSRV1.ORDR_DEPT, SOTRSRV1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTRSRV1.INIT_DATE" & vbCrLf _
                & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTRSRV1.WHSE_CODE, SOTRSRV1.SREP_CODE" & vbCrLf _
                & ", SOTRSRV1.ORDR_ORIG_SHIP_DATE, SOTRSRV1.ORDR_ORIG_CANCEL_DATE" & vbCrLf _
                & ", SOTRSRV1.WHSE_CODE, SOTRSRV1.SREP_CODE, SOTRSRV1.RSRV_PRIORITY"


        If CUST_CODE <> "" Then ' ScreenMode Then
            ASCMAIN1.sql = sqlSOTORDR0
            Dim sqlw As String = " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf
            If ASCMAIN1.CLIENT = "VAN" Then
                sqlw &= " and SOTPCKP2.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTPCKP2.PACK_GROUP_STATUS (+) = 'A'"
            End If

            If cmbSALES_DIVISION_CODE.Value & "" <> "" Then
                sqlw &= " and SOTORDR0.SALES_DIVISION_CODE = '" & cmbSALES_DIVISION_CODE.Value & "'" & vbCrLf
            End If

            If optOrders.Value = "O" Then
                sqlw &= " AND SOTORDR0.ORDR_QTY_OPEN <> 0" & vbCrLf
            ElseIf optOrders.Value = "P" Then
                sqlw &= " AND SOTORDR0.ORDR_QTY_PICK <> 0" & vbCrLf
            ElseIf optOrders.Value = "S" Then
                sqlw &= " AND SOTORDR0.ORDR_QTY_SHIP <> 0" & vbCrLf
            ElseIf optOrders.Value = "C" Then
                sqlw &= " AND SOTORDR0.ORDR_QTY_CANC <> 0" & vbCrLf
            ElseIf optOrders.Value = "OP" Then
                sqlw &= " AND (NVL(SOTORDR0.ORDR_QTY_OPEN,0) <> 0 OR NVL(SOTORDR0.ORDR_QTY_PICK,0) <> 0)" & vbCrLf
            End If

            grdSOTORDR0.Text = "Orders for " & CUST_CODE & "; Status: " & optOrders.Text

            ASCMAIN1.sql &= sqlw

            If (optOrders.Value = "A" Or optOrders.Value = "O" Or optOrders.Value = "OP" Or optOrders.Value = "C") And chkReservations.Checked Then
                ASCMAIN1.sql &= Replace(sqlReservations, " group by ", "   and SOTRSRV1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf & " group by ")
            End If

        Else


            If Not chkReservations_FindCustomerBy.Checked Then
                sqlReservations = ""
            End If


            Dim SQLW As String = ""
            If chkActionDate.Checked And dteActionDate.Value & "" <> "" Then
                SQLW &= " AND (SOTORDRG.ORDR_REL_ACTION_DATE is Null or SOTORDRG.ORDR_REL_ACTION_DATE < '" & Format(dteActionDate.Value, "dd-MMM-yyyy") & "')" & vbCrLf
            End If

            If chkIfReceivedSince.Checked And dteIfReceivedSince.Value & "" <> "" Then
                SQLW &= " AND SOTORDR0.ORDR_GROUP_NO in (Select ORDR_GROUP_NO from SOTORDR1,SOTORDR2 where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO and SOTORDR2.ORDR_STATUS = 'O' and (SOTORDR2.STYLE_CODE,SOTORDR2.COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE from ICTIREC1,ICTIREC2 where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO and ICTIREC1.RECEIPT_DATE > '" & Format(dteIfReceivedSince.Value, "dd-MMM-yyyy") & "'))" & vbCrLf
            End If
            If ASCMAIN1.CLIENT = "VAN" Then
                SQLW &= " and SOTPCKP2.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTPCKP2.PACK_GROUP_STATUS (+) = 'A'"
            End If

            ASCMAIN1.sql = sqlSOTORDR0
            PARM1 = Replace(Replace(PARM1, ";", ""), "'", "")

            Dim sqlORDR_STATUS As String = ""

            Select Case grdSOTORDR0.Tag & ""
                Case ""
                    grdSOTORDR0.Text = "Orders which are either Open or In Pick"
                    ASCMAIN1.sql &= " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and (SOTORDR0.ORDR_CNT_OPEN <> 0 or SOTORDR0.ORDR_CNT_PICK <> 0)"
                    ASCMAIN1.sql &= SQLW
                    ASCMAIN1.sql &= sqlReservations

                Case "SREP_CODE"
                    grdSOTORDR0.Text = "Open Orders for Sales Rep " & PARM1
                    ASCMAIN1.sql &= " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTORDR0.ORDR_GROUP_NO in " _
                        & " (Select Distinct ORDR_GROUP_NO from SOTORDR1 " _
                        & " where ORDR_STATUS >= 'O' and ORDR_STATUS <= 'P'" _
                        & "   and SREP_CODE = '" & PARM1 & "')"
                    ASCMAIN1.sql &= SQLW
                    ASCMAIN1.sql &= Replace(sqlReservations, " group by ", " and SOTRSRV1.SREP_CODE = '" & PARM1 & "'" & vbCrLf & " group by ")

                Case "ORDR_CUST_PO"
                    grdSOTORDR0.Text = "All Customer Orders using Customer PO " & PARM1
                    ASCMAIN1.sql &= " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTORDR0.ORDR_GROUP_NO in " _
                        & " (Select Distinct ORDR_GROUP_NO from SOTORDR1 where ORDR_CUST_PO = '" & PARM1 & "')"
                    ASCMAIN1.sql &= SQLW
            End Select



        End If

        Dim sqlSOTORDRS As String = "" _
            & "Select ORDR_GROUP_NO" & vbCrLf _
            & ", SUM (ORDR_AMT_ALLO_CUR) ORDR_AMT_ALLO_CUR" & vbCrLf _
            & ", SUM (ORDR_AMT_ALLO_FUT) ORDR_AMT_ALLO_FUT" & vbCrLf _
            & ", SUM (ORDR_AMT_ALLO_CXL) ORDR_AMT_ALLO_CXL" & vbCrLf _
            & " from SOTORDRS" & vbCrLf _
            & " group by ORDR_GROUP_NO"


        ASCMAIN1.sql = "Select X.*, SOTORDR1.TERM_CODE, SOTORDR1.LAST_DATE, SOTORDR1.LAST_OPER, SOTORDR1.ORDR_SHIP_INSTR, SOTORDR1.ORDR_MESSAGE" & vbCrLf _
            & ", ARTCCPA1.CUST_CREDIT_CARD_EXP_DATE, ARTCCPA1.CUST_CREDIT_CARD_LAST4" & vbCrLf _
            & ", ARTCUST1.CUST_NAME, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY" & vbCrLf _
            & ", NULL WAVE_NO, NULL EDI_LOAD_ID" & vbCrLf _
            & ", SOTORDRS.ORDR_AMT_ALLO_CUR, SOTORDRS.ORDR_AMT_ALLO_FUT, SOTORDRS.ORDR_AMT_ALLO_CXL, SOTORDR1.EDI_PO_TYPE" & vbCrLf _
            & " from (" & ASCMAIN1.sql & ") X,ARTCUST1,SOTORDR1, ARTCCPA1, (" & sqlSOTORDRS & ") SOTORDRS" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO (+) = X.ORDR_NO_MIN" & vbCrLf _
            & "   and ARTCCPA1.CCPA_NO (+) = SOTORDR1.CCPA_NO" & vbCrLf _
            & "   and SOTORDRS.ORDR_GROUP_NO (+) = X.ORDR_GROUP_NO"
        'Fill_Records("SOTORDR0", "", , ASCMAIN1.sql)


        If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
            ASCMAIN1.sql &= " and X.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")"
        End If

        ASCDATA1.ExecuteSQL("Delete from " & SOTORDR0)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR0 & " " & ASCMAIN1.sql)

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select ORDR_GROUP_NO, MIN (WAVE_NO) WAVE_NO, MIN (EDI_LOAD_ID) EDI_LOAD_ID" & vbCrLf _
                & "   from SOTSHIP1 where ORDR_GROUP_NO in " & vbCrLf _
                & "    (Select ORDR_GROUP_NO from " & SOTORDR0 & " where ORDR_TYPE = 'O')" & vbCrLf _
                & "   group by ORDR_GROUP_NO;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDR0 & " Set WAVE_NO = R1.WAVE_NO, EDI_LOAD_ID = R1.EDI_LOAD_ID" & vbCrLf _
                & "    where ORDR_TYPE = 'O' and ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select ORDR_GROUP_NO, MIN(EDI_PO_TYPE) AS EDI_PO_TYPE" & vbCrLf _
                & "   from SOTORDR1 where ORDR_GROUP_NO in " & vbCrLf _
                & "    (Select ORDR_GROUP_NO from " & SOTORDR0 & ")" & vbCrLf _
                & "   group by ORDR_GROUP_NO;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDR0 & " Set EDI_PO_TYPE = R1.EDI_PO_TYPE" & vbCrLf _
                & "    where ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        Fill_Records("SOTORDR0")

        Setup_SOTORDR0()
        Toggle_ChgShipCancel()

        If optOrders.Value = "S" Or optOrders.Value = "C" Or optOrders.Value = "A" Then
            Sort_grdColumns(grdSOTORDR0, "ORDR_GROUP_NO".ToLower)
        Else
            Sort_grdColumns(grdSOTORDR0, "ORDR_CANCEL_DATE")
        End If

        grdSOTORDR0.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub

    Sub Toggle_ChgShipCancel()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Orders with Changed Ship/Cancel"), UltraWinToolbars.StateButtonTool)

        Dim dvw As DataView = dst.Tables("SOTORDR0").DefaultView

        If tlb_sbt.Checked Then
            dvw.RowFilter = "ORDR_ORIG_SHIP_DATE <> ORDR_SHIP_DATE or ORDR_ORIG_CANCEL_DATE <> ORDR_CANCEL_DATE"
            grdSOTORDR0.Text &= " (Orders with Changes to Ship or Cancel Dates)"
        Else
            dvw.RowFilter = ""
            grdSOTORDR0.Text = Replace(grdSOTORDR0.Text, " (Orders with Changes to Ship or Cancel Dates)", "")
        End If

        Toggle_ShowShipCancel()
    End Sub

    Sub Toggle_OrderGridView()
        If ASCMAIN1.CLIENT = "RGI" Then
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Short View"), UltraWinToolbars.StateButtonTool)

            grdSOTORDR0.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            'grdSOTORDR0.DisplayLayout.SaveAsXml("nl.xml", UltraWinGrid.PropertyCategories.All)
            'With grdSOTORDR0.DisplayLayout.Bands(0)
            '    For Each C As String In New String() _
            '        {"CUST_COUNTRY", "ORDR_DEPT", "EDI_MERCH_TYPE", "CUST_DC_NO", "ORDR_DATE", "ORDR_TYPE_CODE",
            '         "ORDR_SOURCE", "ORDR_AMT_CANC", "ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK", "ORDR_QTY_CANC",
            '         "ORDR_DATE_RECD", "ORDR_PRIORITY", "LAST_DATE", "LAST_OPER"}
            '        .Columns(C).Hidden = tlb_sbt.Checked
            '    Next
            'End With
            'grdSOTORDR0.DisplayLayout.SaveAsXml("nl_short.xml", UltraWinGrid.PropertyCategories.All)


            If tlb_sbt.Checked Then
                grdSOTORDR0.DisplayLayout.LoadFromXml(SOFCORD1_LAYOUT_SHORT)
            Else
                grdSOTORDR0.DisplayLayout.LoadFromXml(SOFCORD1_LAYOUT_ORIG)
            End If
        End If
    End Sub

    Sub Toggle_ShowShipCancel()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Original Ship/Cancel"), UltraWinToolbars.StateButtonTool)

        With grdSOTORDR0.DisplayLayout.Bands(0)
            .Columns("ORDR_ORIG_SHIP_DATE").Hidden = Not tlb_sbt.Checked
            .Columns("ORDR_ORIG_CANCEL_DATE").Hidden = Not tlb_sbt.Checked
        End With
    End Sub

    Sub Setup_SOTORDR0()

        If grdSOTORDR0.ActiveRow Is Nothing OrElse Not grdSOTORDR0.ActiveRow.IsDataRow Then
            tabDetails.Visible = False
        Else
            tabDetails.Visible = True
            ORDR_GROUP_NO = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value
            Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
            Dim ORDR_TYPE As String = grdSOTORDR0.ActiveRow.Cells("ORDR_TYPE").Value
            'ASCMAIN1.Progress("Now Setting up Details")
            If ORDR_TYPE = "R" Then
                chkShowSelectedOrder.Checked = False
                chkShowSelectedOrder.Enabled = False
            Else
                chkShowSelectedOrder.Enabled = True
            End If
            EnforceConstraints(False)
            dst.Tables("SOTPICK2").Rows.Clear()
            dst.Tables("SOTSHIP2").Rows.Clear()
            If ORDR_TYPE = "O" Then
                Fill_Records("SOTORDR1", ORDR_GROUP_NO)
                If tabDetails.SelectedTab.Key = "All Orders" Then
                Else
                    Sort_grdColumns(grdSOTORDR1, "ORDR_NO")
                    grdSOTORDR1.Text = "Sales Orders for Order Group " & ORDR_GROUP_NO
                End If

                Fill_Records("SOTPICK1", ORDR_GROUP_NO)
                Sort_grdColumns(grdSOTPICK1, "PICK_NO")
                grdSOTPICK1.Text = "Pick Tickets for Order Group " & ORDR_GROUP_NO

                Fill_Records("SOTSHIP1", ORDR_GROUP_NO)
                Sort_grdColumns(grdSOTSHIP1, "SHIP_BOL_NO")
                grdSOTSHIP1.Text = "Shipments for Order Group " & ORDR_GROUP_NO
            Else
                dst.Tables("SOTORDR1").Rows.Clear()
                dst.Tables("SOTPICK1").Rows.Clear()
                dst.Tables("SOTSHIP1").Rows.Clear()
            End If
            EnforceConstraints(True)
            Load_SOTORDRS()
            ' ASCMAIN1.Progress("")
        End If

    End Sub

    Private Sub grdSOTORDR0_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR0.AfterRowActivate

        If ScreenMode Then

            Setup_SOTORDR0()

            If grdSOTORDR0.ActiveRow Is Nothing OrElse Not grdSOTORDR0.ActiveRow.IsDataRow Then
                grdSOTORDR4.Visible = False
            Else
                If ASCMAIN1.CLIENT = "RGI" Then
                    Dim ORDR_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_NO_MIN").Value
                    Fill_Records("SOTORDR4", ORDR_NO)
                    grdSOTORDR4.Text = $"Internal Comments for Order {ORDR_NO}"
                End If
            End If
        Else
            If ASCMAIN1.CLIENT = "RGI" Then
                Setup_grdSOTORDRT()
            End If

        End If
    End Sub


    Sub Setup_grdSOTORDRT()
        If grdSOTORDR0.ActiveRow Is Nothing OrElse Not grdSOTORDR0.ActiveRow.IsDataRow Then
            grdSOTORDRT.Visible = False
        Else
            Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value
            Dim ORDR_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_NO_MIN").Value
            Fill_Records("SOTORDRT", ORDR_GROUP_NO)

            Sort_grdColumns(grdSOTORDRT, "STYLE_CODE,COLOR_CODE")
            grdSOTORDRT.Text = "Order Details for Order Group " & ORDR_GROUP_NO
            grdSOTORDRT.Visible = True

            Fill_Records("SOTORDR1", ORDR_GROUP_NO)
            Fill_Records("SOTORDR4", ORDR_NO)
            grdSOTORDR4.Text = $"Internal Comments for Order {ORDR_NO}"
        End If
    End Sub

    Private Sub grdSOTORDR0_AfterRowLayoutItemResized(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterRowLayoutItemResizedEventArgs) Handles grdSOTORDR0.AfterRowLayoutItemResized

    End Sub

    Private Sub grdSOTORDR0_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDR0.DoubleClickRow
        If e.Row.IsDataRow Then
            If Not ScreenMode Then
                Dim ORDR_GROUP_NO As String = e.Row.Cells("ORDR_GROUP_NO").Value & ""
                Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
                Click_Command("Select")
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                    If grow.Cells("ORDR_GROUP_NO").Value = ORDR_GROUP_NO Then
                        grdSOTORDR0.ActiveRow = grow
                        grdSOTORDR0.DisplayLayout.RowScrollRegions(0).FirstRow = grow
                    End If
                Next

            End If
        End If
    End Sub

    Private Sub txtFindBy_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtFindBy.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then

            Dim FIND_BY As String = txtFindBy.Text
            FIND_BY = Replace(Replace(FIND_BY, ";", ""), "'", "")

            If optFindBy.Value <> "C" Then
                If Len(FIND_BY) > 10 Then
                    FIND_BY = Mid(FIND_BY, 1, 10)
                Else
                    FIND_BY = FIND_BY.PadLeft(10, "0")
                End If
            End If

            Select Case optFindBy.Value
                Case "C"
                    ASCMAIN1.sql = "Select Distinct CUST_CODE from SOTORDR1 where ORDR_CUST_PO = :PARM1"
                Case "O"
                    ASCMAIN1.sql = "Select CUST_CODE from SOTORDR1 where ORDR_NO = :PARM1"
                Case "I"
                    ASCMAIN1.sql = "Select CUST_CODE from SOTINVH1 where INV_TYPE = 'I' AND INV_NO = :PARM1"
                Case "P"
                    ASCMAIN1.sql = "Select SOTORDR1.CUST_CODE from SOTPICK1,SOTORDR1 where SOTPICK1.PICK_NO = :PARM1 and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
            End Select

            If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then
                ASCMAIN1.sql &= " and SOTORDR1.WHSE_CODE IN (" & TAC.TACMAIN1.NyaCanadaWhseQueryString & ")"
            End If

            Dim rows() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {FIND_BY}).Select()
            If rows.Length = 0 Then
                MsgBox("No Customer(s) found with " & optFindBy.Text & " " & FIND_BY,
                       MsgBoxStyle.OkOnly, "Could Not Locate a Matching Customer")
            ElseIf rows.Length = 1 Then
                txtFindBy.Text = ""
                Absx1.txtFor("CUST_CODE").Text = rows(0).Item(0)
                Click_Command("Select")

                If optFindBy.Value = "C" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                        If grow.IsDataRow Then
                            If grow.Cells("ORDR_CUST_PO").Value & "" = FIND_BY Then
                                grdSOTORDR0.ActiveRow = grow
                                grdSOTORDR0.ActiveRowScrollRegion.FirstRow = grow
                                grow.Selected = True
                                Exit For
                            End If
                        End If
                    Next
                End If
            Else
                grdSOTORDR0.Tag = "ORDR_CUST_PO"
                Load_SOTORDR0(FIND_BY)
            End If
        End If
    End Sub

    Private Sub cmbSALES_DIVISION_CODE_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles cmbSALES_DIVISION_CODE.KeyDown
        If e.KeyCode = Keys.Delete Then
            cmbSALES_DIVISION_CODE.Value = ""
        End If
    End Sub

    Private Sub grdSOTORDR0_ImeModeChanged(sender As Object, e As System.EventArgs) Handles grdSOTORDR0.ImeModeChanged

    End Sub

    Private Sub grdSOTORDR0_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR0.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("ORDR_TYPE").Value & "" = "R" Then
                e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Reservation"
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Red
            ElseIf Val(e.Row.Cells("ORDR_CNT_OPEN").Value & "") <> 0 Then
                e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Some or All Orders are still Open"
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Green
            ElseIf Val(e.Row.Cells("ORDR_CNT_PICK").Value & "") <> 0 Then
                e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Some or All Orders are In Pick"
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Blue
            ElseIf Val(e.Row.Cells("ORDR_QTY_SHIP").Value & "") = 0 Then
                e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Cancelled Order"
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Orange
            End If

            If ASCMAIN1.CLIENT = "VAN" Then
                If e.Row.Cells("EDI_CONS_NO").Value & "" <> "" Then
                    e.Row.Cells("EDI_CONS_NO").ToolTipText = "Multi-PO"
                    e.Row.Cells("EDI_CONS_NO").Appearance.BackColor = Drawing.Color.LightBlue
                End If
            End If


            If e.Row.Cells("ORDR_SHIP_DATE").Value & "" <> "" And e.Row.Cells("ORDR_ORIG_SHIP_DATE").Value & "" <> "" Then
                If Format(e.Row.Cells("ORDR_SHIP_DATE").Value, "yyyyMMdd") <> Format(e.Row.Cells("ORDR_ORIG_SHIP_DATE").Value, "yyyyMMdd") Then
                    e.Row.Cells("ORDR_SHIP_DATE").ToolTipText = "Ship Date has been Changed"
                    e.Row.Cells("ORDR_SHIP_DATE").Appearance.ForeColor = Drawing.Color.Blue
                Else
                    e.Row.Cells("ORDR_SHIP_DATE").ToolTipText = ""
                    e.Row.Cells("ORDR_SHIP_DATE").Appearance.ForeColor = Drawing.Color.Empty
                End If
            End If

            If e.Row.Cells("ORDR_CANCEL_DATE").Value & "" <> "" And e.Row.Cells("ORDR_ORIG_CANCEL_DATE").Value & "" <> "" Then
                If Format(e.Row.Cells("ORDR_CANCEL_DATE").Value, "yyyyMMdd") <> Format(e.Row.Cells("ORDR_ORIG_CANCEL_DATE").Value, "yyyyMMdd") Then
                    e.Row.Cells("ORDR_CANCEL_DATE").ToolTipText = "Cancel Date has been Changed"
                    e.Row.Cells("ORDR_CANCEL_DATE").Appearance.ForeColor = Drawing.Color.Blue
                Else
                    e.Row.Cells("ORDR_CANCEL_DATE").ToolTipText = "Ship Date has been Changed"
                    e.Row.Cells("ORDR_CANCEL_DATE").Appearance.ForeColor = Drawing.Color.Empty
                End If
            End If

            If e.Row.Cells("ORDR_CANCEL_DATE").Value & "" <> "" And e.Row.Cells("ORDR_RELEASE_AVAIL_MIN").Value & "" <> "" Then
                If Format(e.Row.Cells("ORDR_CANCEL_DATE").Value, "yyyyMMdd") < Format(e.Row.Cells("ORDR_RELEASE_AVAIL_MIN").Value, "yyyyMMdd") Then
                    e.Row.Cells("ORDR_RELEASE_AVAIL_MIN").ToolTipText = "Cancel Date is prior to 1st Availability Date"
                    e.Row.Cells("ORDR_RELEASE_AVAIL_MIN").Appearance = appRed
                Else
                    e.Row.Cells("ORDR_RELEASE_AVAIL_MIN").ToolTipText = ""
                    e.Row.Cells("ORDR_RELEASE_AVAIL_MIN").Appearance.ForeColor = Drawing.Color.Empty
                End If
            End If
            If e.Row.Cells("ORDR_CANCEL_DATE").Value & "" <> "" And e.Row.Cells("ORDR_RELEASE_AVAIL_MAX").Value & "" <> "" Then
                If Format(e.Row.Cells("ORDR_CANCEL_DATE").Value, "yyyyMMdd") < Format(e.Row.Cells("ORDR_RELEASE_AVAIL_MAX").Value, "yyyyMMdd") Then
                    e.Row.Cells("ORDR_RELEASE_AVAIL_MAX").ToolTipText = "Cancel Date is prior to last Availability Date"
                    e.Row.Cells("ORDR_RELEASE_AVAIL_MAX").Appearance = appRed
                Else
                    e.Row.Cells("ORDR_RELEASE_AVAIL_MAX").ToolTipText = ""
                    e.Row.Cells("ORDR_RELEASE_AVAIL_MAX").Appearance.ForeColor = Drawing.Color.Empty
                End If
            End If

            Dim MMYY As String = e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Value & ""
            If MMYY <> "" Then
                If Mid(MMYY, 3, 2) & Mid(MMYY, 1, 2) <= YYMM_exp Then
                    e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Appearance = appRed
                Else
                    e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Appearance.ForeColor = Drawing.Color.Empty
                End If
            End If


        End If
    End Sub


    Private Sub Generate_History()
        dst.Tables("SOTCORDD").Rows.Clear()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Collecting History")
        Dim RYP As String = cmb12Months.Value
        RYP = Mid(RYP, 1, 4) & Mid(RYP, 6, 2)
        Dim YPs(12) As String
        For i As Integer = 1 To 12
            YPs(i) = ASCMAIN1.Period_Calc(RYP, i - 12)
            Dim LEGEND = ASCMAIN1.Get_Legend(YPs(i))
            grdSOTCORDX.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Header.Caption = Mid(LEGEND, 10, 6)
            grdSOTCORDD.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Header.Caption = Mid(LEGEND, 10, 6)
            grdSOTCORDD.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Tag = YPs(i)
        Next i

        ' this ought to be processed by a datareader with 1 pass thru all of styles

        For Each SORT_SEQ As String In New String() {"1", "2", "4", "8"}
            Dim exp As String = ""
            Dim sqlexpw As String = ""
            If SORT_SEQ = "1" Then
                exp = "DECODE(SOTINVH2.INV_TYPE,'I',(SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE),0)"
                sqlexpw = " and SOTINVH2.INV_TYPE = 'I' and SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE <> 0"
            ElseIf SORT_SEQ = "2" Then
                exp = "DECODE(SOTINVH2.INV_TYPE,'C',(SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE),0)"
                sqlexpw = " and SOTINVH2.INV_TYPE = 'C' and SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE <> 0"
            ElseIf SORT_SEQ = "4" Then
                exp = "ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST"
                sqlexpw = " and SOTINVH2.ORDR_QTY_SHIP  * SOTINVH2.ORDR_UNIT_COST <> 0"
            ElseIf SORT_SEQ = "8" Then
                exp = "SOTINVH2.ORDR_QTY_SHIP"
                sqlexpw = " and SOTINVH2.ORDR_QTY_SHIP <> 0"
            End If
            ASCMAIN1.sql = "Select '" & SORT_SEQ & "' SORT_SEQ, STYLE_CODE CODE_VALUE" & vbCrLf
            For i As Integer = 1 To 12
                ASCMAIN1.sql &= ", SUM (DECODE (ORDR_YYYYPP_UPDATED,'" & YPs(i) & "'," & exp & ",0)) V" & Format(i, "00") & vbCrLf
            Next i
            ASCMAIN1.sql &= " from SOTINVH2" & vbCrLf
            ASCMAIN1.sql &= " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf
            ASCMAIN1.sql &= "   and ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(RYP, -11) & "' AND ORDR_YYYYPP_UPDATED <= '" & RYP & "'" & vbCrLf
            ASCMAIN1.sql &= sqlexpw & vbCrLf
            ASCMAIN1.sql &= " group by STYLE_CODE" & vbCrLf
            If SORT_SEQ <> "4" Or ASCMAIN1.USER_SECURITY_CODEs.Contains("X2") Then
                dst.Tables("SOTCORDD").Merge(ASCDATA1.GetDataTable)
                'Fill_Records("SOTCORDD", "", False, ASCMAIN1.sql)
            End If
        Next

        ASCMAIN1.sql = "Select '6' SORT_SEQ, ARTPYMT5.REASON_CODE CODE_VALUE" & vbCrLf
        For i As Integer = 1 To 12
            ASCMAIN1.sql &= " , SUM (DECODE (ARTPYMT1.OPS_YYYYPP,'" & YPs(i) & "',ARTPYMT5.GL_DIST_AMT,0)) V" & Format(i, "00") & vbCrLf
        Next i
        ASCMAIN1.sql &= "" _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT5" & vbCrLf _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
        & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
        & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
        & "   and Decode(ARTPYMT5.CUST_CODE_SO,Null,ARTPYMT2.CUST_CODE,ARTPYMT5.CUST_CODE_SO) = '" & CUST_CODE & "'" & vbCrLf _
        & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0'" & vbCrLf _
        & " group by ARTPYMT5.REASON_CODE" & vbCrLf
        dst.Tables("SOTCORDD").Merge(ASCDATA1.GetDataTable)
        ' Fill_Records("SOTORDDD", "", False, ASCMAIN1.sql)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Set_History_Summary()

        spl12Months.Visible = True
    End Sub

    Sub Set_History_Summary()
        dst.Tables("SOTCORDX").Rows.Clear()

        Dim V(9, 12) As Decimal
        For i As Integer = 1 To 12
            Dim PP As String = Format(i, "00")
            V(1, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '1'") & "")
            V(2, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '2'") & "")
            V(3, i) = V(1, i) + V(2, i)
            V(4, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '4'") & "")
            V(5, i) = V(3, i) - V(4, i)
            V(6, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '6'") & "")
            V(7, i) = V(5, i) - V(6, i)
            V(8, i) = Val(dst.Tables("SOTCORDD").Compute("SUM(V" & PP & ")", "SORT_SEQ = '8'") & "")
            If V(8, i) <> 0 Then V(9, i) = V(3, i) / V(8, i)
        Next i

        Dim S(9) As String
        S(1) = "Gross Sales"
        S(2) = "Credits"
        S(3) = "Net Sales"
        S(4) = "CGS"
        S(5) = "GP"
        S(6) = "Deductions"
        S(7) = "Net Profit"
        S(8) = "Units"
        S(9) = "Price"

        Dim rowSOTCORDX As DataRow

        For j As Integer = 1 To 8
            If ASCMAIN1.USER_SECURITY_CODEs.Contains("X2") Or j = 1 Or j = 2 Or j = 3 Or j = 6 Then
                rowSOTCORDX = dst.Tables("SOTCORDX").NewRow
                rowSOTCORDX.Item("SORT_SEQ") = Format(j, "0")
                rowSOTCORDX.Item("CODE_VALUE") = S(j)
                For i As Integer = 1 To 12
                    Dim PP As String = Format(i, "00")
                    rowSOTCORDX.Item("V" & PP) = V(j, i)
                Next i
                dst.Tables("SOTCORDX").Rows.Add(rowSOTCORDX)
            End If
        Next j

        rowSOTCORDX = dst.Tables("SOTCORDX").Rows.Find("1")
        Dim TOT_SALES = Val(rowSOTCORDX.Item("TOT") & "")
        Dim YTD_SALES = Val(rowSOTCORDX.Item("YTD") & "")
        dst.Tables("SOTCORDX").Columns("TOTPCT").Expression = IIf(TOT_SALES = 0, "0", "100 * TOT / " & CStr(TOT_SALES))
        dst.Tables("SOTCORDX").Columns("YTDPCT").Expression = IIf(YTD_SALES = 0, "0", "100 * YTD / " & CStr(YTD_SALES))
        CreateGraph_SATCSLS1_X()
    End Sub


    Private Sub grdSOTPICK1_BeforeRowExpanded(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPICK1.BeforeRowExpanded
        Dim PICK_NO As String = e.Row.Cells("PICK_NO").Value
        Fill_Records("SOTPICK2", PICK_NO)
        Sort_grdColumns(grdSOTPICK1, "PICK_LNO", False, 1)
    End Sub

    Private Sub grdSOTPICK1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICK1.InitializeRow
        If e.Row.Band.Key = "SOTPICK1" Then
            If e.Row.Cells("PICK_FORCED").Value & "" = "1" Then
                e.Row.Cells("PICK_NO").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("PICK_NO").ToolTipText = "Force Picked"
            End If
            If e.Row.Cells("PICK_NO_REV").Value & "" <> "" Then
                e.Row.Cells("PICK_NO").Appearance.ForeColor = Drawing.Color.Orange
                e.Row.Cells("PICK_NO").ToolTipText = "Reversed"
            End If
            If e.Row.Cells("PICK_STATUS").Value & "" = "D" Then
                e.Row.Cells("PICK_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("PICK_STATUS").ToolTipText = "De-Released"
            End If
        End If
    End Sub

    Private Sub grdSOTSHIP1_BeforeRowExpanded(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSHIP1.BeforeRowExpanded
        Dim SHIP_BOL_NO As String = e.Row.Cells("SHIP_BOL_NO").Value
        Fill_Records("SOTSHIP2", SHIP_BOL_NO)
        Sort_grdColumns(grdSOTSHIP1, "STYLE_CODE,COLOR_CODE", False, 1)
    End Sub

    Private Sub grdSOTSHIP1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIP1.InitializeRow

        If e.Row.IsDataRow Then
            If e.Row.Band.Key = "SOTSHIP1" Then
                If e.Row.Cells("SHIP_BOL_NO_REV").Value & "" <> "" Then
                    e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Red
                End If

                If e.Row.Cells("SHIP_STATUS").Value = "D" Then
                    e.Row.Cells("SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Deleted"
                ElseIf e.Row.Cells("SHIP_STATUS").Value = "F" Then
                    e.Row.Cells("SHIP_STATUS").Appearance.BackColor = Drawing.Color.LightGreen
                    e.Row.ToolTipText = "Shipped"
                Else
                    If e.Row.Cells("SHIP_WAVE_STATUS").Value & "" = "1" Then
                        e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Blue
                        e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Waved"
                    Else
                        e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Empty
                        e.Row.Cells("SHIP_BOL_NO").ToolTipText = ""
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub grdSOTCORDX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCORDX.AfterRowActivate
        Setup_SOTCORDD()
        If tabMonth.SelectedTab.Key = "Details" Then tabMonth.SelectedTab = tabMonth.Tabs("Summary")
    End Sub

    Sub Setup_SOTCORDD()
        If grdSOTCORDX.ActiveRow Is Nothing Then
            tabMonth.Visible = False
        Else
            Dim TOT_SALES = Val(grdSOTCORDX.ActiveRow.Cells("TOT").Value & "")
            Dim YTD_SALES = Val(grdSOTCORDX.ActiveRow.Cells("YTD").Value & "")
            dst.Tables("SOTCORDD").Columns("TOTPCT").Expression = IIf(TOT_SALES = 0, "0", "100 * TOT / " & CStr(TOT_SALES))
            dst.Tables("SOTCORDD").Columns("YTDPCT").Expression = IIf(YTD_SALES = 0, "0", "100 * YTD / " & CStr(YTD_SALES))

            Dim dvw As DataView = DirectCast(grdSOTCORDD.DataSource, DataTable).DefaultView
            Dim SORT_SEQ As String = grdSOTCORDX.ActiveRow.Cells("SORT_SEQ").Value
            dvw.RowFilter = "SORT_SEQ = '" & SORT_SEQ & "'"
            Sort_grdColumns(grdSOTCORDD, "CODE_VALUE")
            tabMonth.Visible = True
        End If
    End Sub

    Private Sub cmdGenerateHistory_Click(sender As System.Object, e As System.EventArgs) Handles cmdGenerateHistory.Click
        Generate_History()
    End Sub

    Private Sub chkShowSelectedOrder_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowSelectedOrder.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_SOTORDRS()
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Fill_ICTSTATA()
        End If

        If tabStyles.SelectedTab.Key = "Summary" Then tabStyles.SelectedTab = tabStyles.Tabs("Detail")
        tabStyles.Tabs("Summary").Visible = Not chkShowSelectedOrder.Checked
    End Sub

    Private Sub grdSOTORDRX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDRX.AfterRowActivate
        If grdSOTORDRX.ActiveRow Is Nothing OrElse grdSOTORDRX.ActiveRow.IsFilterRow Or Not grdSOTORDRX.ActiveRow.IsDataRow Then Exit Sub
        If grdSOTORDR0.ActiveRow.IsFilterRow Then Exit Sub
        Setup_SOTORDRS()
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Fill_ICTSTATA()
        End If
    End Sub

    Sub Setup_Summary_SOTORDRM(ORDR_TYPE As String, ORDR_GROUP_NO As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Summary by Store")

        Dim COLUMN_NAME As String = optQTY.Value

        ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
        ASCMAIN1.sql &= ",Sum (" & COLUMN_NAME & ") QTY"
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), New String() {"CUST_STORE_NO"}).Select("", "CUST_STORE_NO")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            ASCMAIN1.sql &= ", Sum (Decode(SOTORDR1.CUST_STORE_NO,'" & CUST_STORE_NO & "'," & COLUMN_NAME & ",0)) QTY_" & CUST_STORE_NO
        Next
        'For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("", "CUST_STORE_NO")
        '    Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO")
        '    ASCMAIN1.sql &= ", Sum (Decode(SOTORDR1.CUST_STORE_NO,'" & CUST_STORE_NO & "'," & COLUMN_NAME & ",0)) QTY_" & CUST_STORE_NO
        'Next
        ASCMAIN1.sql &= " from SOTORDR1,SOTORDR2" _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
            & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
        grdSOTORDRM.DataSource = Nothing
        grdSOTORDRM.DisplayLayout.Bands(0).Summaries.Clear()
        grdSOTORDRM.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        dst.Tables.Remove("SOTORDRM")
        Dim t As DataTable = ASCDATA1.GetDataTable
        t.TableName = "SOTORDRM"
        dst.Tables.Add(t)
        grdSOTORDRM.DataSource = t
        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORDRM.DisplayLayout.Bands(0).Columns
            If gcol.Key = "STYLE_CODE" Then
                gcol.Width = 90
                gcol.Header.Caption = "Style"
                Create_Summary(grdSOTORDRM, "STYLE_CODE", "Count")
            ElseIf gcol.Key = "COLOR_CODE" Then
                gcol.Width = 40
                gcol.Header.Caption = "Color"
            ElseIf gcol.Key = "QTY" Then
                gcol.Width = 70
                gcol.Header.Caption = "Total"
                gcol.Format = "#,##0"
                Create_Summary(grdSOTORDRM, "QTY")
            Else
                gcol.Width = 70
                gcol.Header.Caption = Mid(gcol.Key, 5)
                gcol.Format = "#,##0"
                Create_Summary(grdSOTORDRM, gcol.Key)
            End If
        Next

        grdSOTORDRM.Text = "Order Group " & ORDR_GROUP_NO & ", Style Summary by Store, " & optQTY.Text

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Summary()
        If SELECTION_NO = 0 Then Exit Sub
        If tabDetails.SelectedTab.Key = "Styles" AndAlso tabStyles.SelectedTab.Key = "Summary" Then
            Dim ORDR_TYPE As String = grdSOTORDR0.ActiveRow.Cells("ORDR_TYPE").Value
            Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value
            If Not chkShowSelectedOrder.Checked And ORDR_TYPE = "O" Then
                Setup_Summary_SOTORDRM(ORDR_TYPE, ORDR_GROUP_NO)
                grdSOTORDRM.Visible = True
            Else
                grdSOTORDRM.Visible = False
            End If
        End If
    End Sub

    Private Sub cmdRefreshStyles_Click(sender As System.Object, e As System.EventArgs) Handles cmdRefreshStyles.Click
        Load_SOTORDRS()
    End Sub

    Sub Load_SOTORDRS()
        If Not chkShowSelectedOrder.Checked And grdSOTORDR0.ActiveRow Is Nothing Or Not grdSOTORDR0.ActiveRow.IsDataRow Then
            tabDetails.Visible = False
        ElseIf chkShowSelectedOrder.Checked And grdSOTORDRX.ActiveRow Is Nothing Then
            tabDetails.Visible = False
        Else
            Setup_SOTORDRS()
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                Fill_ICTSTATA()
            End If

            Dim ORDR_TYPE As String = grdSOTORDR0.ActiveRow.Cells("ORDR_TYPE").Value
            Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
            Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value

            With grdSOTORDRS.DisplayLayout.Bands(0)
                If ORDR_TYPE = "O" Then
                    .Columns("ORDR_QTY_SHIP").Header.Caption = "#Ship"
                    ' .Columns("ORDR_AMT_SHIP").Header.Caption = "$Ship"
                Else
                    .Columns("ORDR_QTY_SHIP").Header.Caption = "#Used"
                    ' .Columns("ORDR_AMT_SHIP").Header.Caption = "$Used"
                End If
            End With

            Setup_Summary()

            Dim SqlP As String = sqlSOTORDRP
            If Not chkShowSelectedOrder.Checked Then
                SqlP = Replace(SqlP, "SOTORDR1.CUST_STORE_NO", "NULL CUST_STORE_NO", , 1)
                SqlP = Replace(SqlP, "SOTORDR1.CUST_STORE_NO", "NULL", , 1)
            End If
            If chkShowSelectedOrder.Checked Then
                Dim ORDR_NO As String = grdSOTORDRX.ActiveRow.Cells("ORDR_NO").Value
                SqlP = Replace(SqlP, "SOTORDR1.ORDR_NO = :PARM1", "SOTORDR1.ORDR_NO = '" & ORDR_NO & "'")
                grdSOTORDRP.Text = "Pre-Packs on Order " & ORDR_NO
            Else
                SqlP = Replace(SqlP, "ORDR_NO = :PARM1", "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                grdSOTORDRP.Text = "Pre-Packs on All Orders in Group " & ORDR_GROUP_NO & "; PO " & ORDR_CUST_PO
            End If
            Fill_Records("SOTORDRP", "", True, SqlP)
        End If
        If tabDetails.SelectedTab.Key = "Cartons" Or tabDetails.SelectedTab.Key = "Pallets" Then
            Setup_Cartons()
        End If

    End Sub

    Sub Setup_SOTORDRS()

        Dim ORDR_TYPE As String = grdSOTORDR0.ActiveRow.Cells("ORDR_TYPE").Value & ""
        Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
        Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""

        ASCMAIN1.Progress("Now Getting Style Details")

        Dim sql As String = ""

        If ORDR_TYPE = "O" Then
            If Not chkShowSelectedOrder.Checked Then
                sql = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' group by ")
                sql = "Select X.*, SOTORDRS.WIP_IND from (" & sql & ") X, SOTORDRS where SOTORDRS.ORDR_GROUP_NO (+) = '" & ORDR_GROUP_NO & "' and SOTORDRS.STYLE_CODE (+) = X.STYLE_CODE and SOTORDRS.COLOR_CODE (+) = X.COLOR_CODE"
                grdSOTORDRS.Text = "Style Summary for Order Group " & ORDR_GROUP_NO & ", Customer PO " & ORDR_CUST_PO
            Else
                Dim ORDR_NO As String = ""
                Dim CUST_STORE_NO As String = ""

                If ASCMAIN1.CLIENT = "RGI" Then
                    ORDR_NO = ORDR_GROUP_NO ' grdSOTORDRX.ActiveRow.Cells("ORDR_NO").Value
                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                    CUST_STORE_NO = rowSOTORDR1.Item("CUST_STORE_NO")
                Else
                    ORDR_NO = grdSOTORDRX.ActiveRow.Cells("ORDR_NO").Value
                    CUST_STORE_NO = grdSOTORDRX.ActiveRow.Cells("CUST_STORE_NO").Value
                End If

                sql = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_NO = '" & ORDR_NO & "' group by ")
                sql = "Select X.*, SOTORDRS.WIP_IND from (" & sql & ") X, SOTORDRS where SOTORDRS.ORDR_GROUP_NO (+) = '" & ORDR_GROUP_NO & "' and SOTORDRS.STYLE_CODE (+) = X.STYLE_CODE and SOTORDRS.COLOR_CODE (+) = X.COLOR_CODE"
                grdSOTORDRS.Text = "Style Details for Order No " & ORDR_NO & ", Customer PO " & ORDR_CUST_PO & ", Store No " & CUST_STORE_NO
            End If
        Else
            sql = Replace(sqlSOTRSRVS, " group by ", " and SOTRSRV1.RSRV_NO = '" & ORDR_GROUP_NO & "' group by ")
            grdSOTORDRS.Text = "Style Summary for Reservation " & ORDR_GROUP_NO & ", Customer PO " & ORDR_CUST_PO
        End If

        For Each COLUMN_NAME As String In New String() _
                {"STYLE_CODE", "COLOR_CODE", "CUST_UPC", "RANGE_STYLE_CODE",
                 "CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_SKU"}

            grdSOTORDRS.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not Absx1.chkFor("SHOW_" & COLUMN_NAME).Checked

            If ORDR_TYPE = "O" Then
                If Not Absx1.chkFor("SHOW_" & COLUMN_NAME).Checked Then
                    sql = Replace(sql, "SOTORDR2." & COLUMN_NAME, "NULL " & COLUMN_NAME, , 1)
                    sql = Replace(sql, "SOTORDR2." & COLUMN_NAME, "NULL")
                    If COLUMN_NAME = "STYLE_CODE" Then
                        sql = Replace(sql, "SOTORDR2.STYLE_DESC", "NULL " & "STYLE_DESC", , 1)
                        sql = Replace(sql, "SOTORDR2.STYLE_DESC", "NULL")
                    End If
                    If COLUMN_NAME = "COLOR_CODE" Then
                        sql = Replace(sql, "ICTCOLR1.COLOR_DESC", "NULL " & "COLOR_DESC", , 1)
                        sql = Replace(sql, "ICTCOLR1.COLOR_DESC", "NULL")
                    End If
                End If
            Else
                If (COLUMN_NAME <> "STYLE_CODE" And COLUMN_NAME <> "COLOR_CODE") Then

                End If
            End If
        Next

        sql = Replace(sql, "ICTCOLR1.COLOR_CODE (+) = NULL", "ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE")
        'If ORDR_TYPE = "R" Then
        '    sql = Replace(sql, "SOTORDR2.QTY_PICK", "0")
        '    sql = Replace(sql, "SOTORDR2.QTY_SHIP", "SOTORDR2.QTY_USED")
        '    sql = Replace(Replace(sql, "SOTORDR1", "SOTRSRV1"), "SOTORDR2", "SOTRSRV2")
        'End If

        Fill_Records("SOTORDRS", "", True, sql)
        Sort_grdColumns(grdSOTORDRS, "STYLE_CODE, COLOR_CODE, RANGE_STYLE_CODE, CUST_STYLE_CODE, CUST_COLOR_CODE, CUST_SKU")

        If ASCMAIN1.CLIENT = "RGI" Then
            Fill_Records("SOTORDRS_3PL", ORDR_GROUP_NO)
            For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                Dim STYLE_CODE As String = rowSOTORDRS.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTORDRS.Item("COLOR_CODE")
                Dim rowSOTORDRS_3PL As DataRow = dst.Tables("SOTORDRS_3PL").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                If rowSOTORDRS_3PL IsNot Nothing Then
                    rowSOTORDRS.Item("QTY_ONH_3PL") = rowSOTORDRS_3PL.Item("QTY_ONH_3PL")
                    rowSOTORDRS.Item("QTY_PCK_3PL") = rowSOTORDRS_3PL.Item("QTY_PCK_3PL")
                    rowSOTORDRS.Item("QTY_OPN_3PL") = rowSOTORDRS_3PL.Item("QTY_OPN_3PL")
                    rowSOTORDRS.Item("QTY_AVA_3PL") = rowSOTORDRS_3PL.Item("QTY_AVA_3PL")
                End If
            Next
        End If

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Dim extra_decimals As Boolean = False
            For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("", "")
                Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTORDRS.Item("ORDR_UNIT_PRICE") & "")
                If ORDR_UNIT_PRICE <> Val(Format(ORDR_UNIT_PRICE, "#.00")) Then
                    extra_decimals = True
                    Exit For
                End If
            Next
            If extra_decimals Then
                grdSOTORDRS.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Format = "#,##0.0000"
            Else
                grdSOTORDRS.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Format = "#,##0.00"
            End If
        End If

        ASCMAIN1.Progress("")

        tabDetails.Visible = True

    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        If e.Tab.Key = "Orders" Then
        Else
            If Not spl12Months.Visible Then
                Generate_History()
            End If
        End If

        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        With UltraExplorerBar1
            .Groups("12 Month History").Visible = (tabMain.SelectedTab.Key = "12 Mos")
            .Groups("Show Orders").Visible = (tabMain.SelectedTab.Key = "Orders")
            .Groups("Styles").Visible = (tabMain.SelectedTab.Key = "Orders") And (tabDetails.SelectedTab.Key = "Styles")
        End With

    End Sub

    Private Sub tabDetails_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        If tabDetails.SelectedTab Is Nothing Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up " & tabDetails.SelectedTab.Key)

        With grdSOTORDR1
            If tabDetails.SelectedTab.Key = "Sales Orders / Reservations" Then
                .Parent = tabDetails.SelectedTab.TabPage
                .DataSource = dst.Tables("SOTORDR1")
                .Text = "Sales Orders / Reservations"
                .DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Hidden = True
                For Each SFX As String In New String() {"", "_OPEN", "_PICK", "_SHIP", "_CANC"}
                    For Each TYP As String In New String() {"QTY", "AMT"}
                        Dim C As String = "ORDR_" & TYP & SFX
                        .DisplayLayout.Bands(0).Columns(C).Hidden = True
                        .DisplayLayout.Bands(0).Columns(C).Hidden = True
                        If ASCMAIN1.CLIENT = "VAN" And (C = "ORDR_QTY" Or C = "ORDR_AMT") Then
                            .DisplayLayout.Bands(0).Columns(C).Hidden = False
                        End If

                    Next
                Next
                .DisplayLayout.Bands(0).Summaries.Clear()
                If ASCMAIN1.CLIENT = "VAN" Then
                    Create_Summary(grdSOTORDR1, New String() {"ORDR_AMT", "ORDR_QTY"}, , , "#,##0")
                    Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")

                End If


                Setup_SOTORDR0()

            ElseIf tabDetails.SelectedTab.Key = "All Orders" Then
                .Parent = tabDetails.SelectedTab.TabPage
                .DataSource = dst.Tables("SOTORDR1_ALL")
                Dim selected_only As Boolean = False
                If grdSOTORDR0.Selected.Rows.Count < 2 Then
                    .Text = "All Store Orders for All " & CStr(grdSOTORDR0.Rows.Count) & " Groups"
                Else
                    selected_only = True
                    .Text = "All Store Orders for " & CStr(grdSOTORDR0.Selected.Rows.Count) & " Selected Groups"
                End If

                .DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Hidden = False
                .DisplayLayout.Bands(0).Summaries.Clear()
                For Each SFX As String In New String() {"", "_OPEN", "_PICK", "_SHIP", "_CANC"}
                    For Each TYP As String In New String() {"QTY", "AMT"}
                        Dim C As String = "ORDR_" & TYP & SFX
                        .DisplayLayout.Bands(0).Columns(C).Format = grdSOTORDR0.DisplayLayout.Bands(0).Columns(C).Format
                        .DisplayLayout.Bands(0).Columns(C).Width = grdSOTORDR0.DisplayLayout.Bands(0).Columns(C).Width
                        With .DisplayLayout.Bands(0).Columns(C).Header
                            .Caption = grdSOTORDR0.DisplayLayout.Bands(0).Columns(C).Header.Caption
                            .Appearance = grdSOTORDR0.DisplayLayout.Bands(0).Columns(C).Header.Appearance
                        End With
                        .DisplayLayout.Bands(0).Columns(C).Hidden = False
                        Create_Summary(grdSOTORDR1, C)
                    Next
                Next
                Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")

                If SOTORDR0_ALL = "" Then
                    ASCMAIN1.sql = "Select ORDR_GROUP_NO from SOTORDR0 where ROWNUM < 1"
                    SOTORDR0_ALL = ASCMAIN1.Temp_Table
                    Create_TDA(dst.Tables.Add("SOTORDR0_ALL"), SOTORDR0_ALL, "*")
                End If
                dst.Tables("SOTORDR0_ALL").Rows.Clear()

                'Dim ORDR_GROUP_NOs As String = ""
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                    If Not selected_only Or grow.Selected Then
                        If grow.Cells("ORDR_TYPE").Value = "O" Then
                            Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                            'ORDR_GROUP_NOs &= ",'" & ORDR_GROUP_NO & "'"
                            dst.Tables("SOTORDR0_ALL").Rows.Add(New String() {ORDR_GROUP_NO})
                        End If
                    End If
                Next
                Update_Record_TDA("SOTORDR0_ALL")
                'If ORDR_GROUP_NOs = "" Then ORDR_GROUP_NOs = ",''"
                'ASCMAIN1.sql = sqlSOTORDR1 & " and ORDR_GROUP_NO in (" & Mid(ORDR_GROUP_NOs, 2) & ")"
                ASCMAIN1.sql = sqlSOTORDR1 & " and ORDR_GROUP_NO in (Select ORDR_GROUP_NO from " & SOTORDR0_ALL & ")"
                Fill_Records("SOTORDR1_ALL", "", True, ASCMAIN1.sql)
                Sort_grdColumns(grdSOTORDR1, "ORDR_GROUP_NO,ORDR_NO")
            End If
        End With


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


        Setup_tabMain()
        Setup_Summary()

        If tabDetails.SelectedTab.Key = "Cartons" Or tabDetails.SelectedTab.Key = "Pallets" Then
            If CUST_CODE = "WALMART" Then ' walmart is the onlyone with RFID at this point
                For band As Integer = 0 To 1
                    With grdSOTCART1.DisplayLayout.Bands(band)
                        For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                            Select Case gcol.Key
                                Case "QTY_RFID"
                                    gcol.Header.Caption = "RFID"
                                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                                    gcol.Width = 60
                                    gcol.Hidden = False
                            End Select
                        Next
                    End With
                Next
            End If
            Setup_Cartons()
        End If

    End Sub

    Sub Setup_Cartons()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now loading Carton Details")
        PO_CARTON_ORIG = ""
        PO_CARTON_COMBINED = ""
        optMULTIPOP.Value = "I"

        EnforceConstraints(False)
        Dim IsMultiPO = grdSOTORDR0.ActiveRow.Cells("EDI_CONS_NO").Value & ""
        Dim ORDR_GROUP_NO2 As String = ORDR_GROUP_NO

        If IsMultiPO <> "" And ASCMAIN1.CLIENT = "VAN" Then

            UltraExplorerBar1.Groups("Multi PO Pallet").Visible = True

            ASCMAIN1.sql = "select SOTORDR0.ORDR_GROUP_NO from SOTSHIP1, SOTSHIP1 CONS, SOTORDR0
                                    where SOTSHIP1.ORDR_GROUP_NO = :PARM1
                                    and SOTSHIP1.SHIP_BOL_NO_CONS = CONS.SHIP_BOL_NO_CONS
                                    and SOTSHIP1.SHIP_BOL_NO <> CONS.SHIP_BOL_NO
                                    and SOTORDR0.ORDR_GROUP_NO = cons.ORDR_GROUP_NO"
            Dim MULTI_PO_ORDR_GROUP_NO As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {ORDR_GROUP_NO})
            ORDR_GROUP_NO2 = MULTI_PO_ORDR_GROUP_NO
            ASCMAIN1.sql = "Select SOTCARM1.*,SOTPICK1.SHIP_BOL_NO,SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE,SOTORDR1.CUST_STORE_NO,SOTPICK1.PICK_STATUS, SOTORDR1.CUST_DC_NO
                            from SOTCARM1,SOTPICK1,SOTSHIP1,SOTORDR1
                            where SOTPICK1.PICK_NO_CONS = SOTCARM1.PICK_NO
                               and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                               and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                               and SOTSHIP1.SHIP_BOL_NO_REV is null
                               and SOTSHIP1.ORDR_GROUP_NO = :PARM1
                            and SOTORDR1.ORDR_cUST_PO IN (SELECT DISTINCT SOTORDR1.ORDR_CUST_PO FROM SOTCARM2,SOTORDR1 WHERE CART_NO = SOTCARM1.CART_NO AND SOTORDR1.ORDR_NO = SOTCARM2.ORDR_NO)"


            Fill_Records("SOTCART1", ORDR_GROUP_NO, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTCARM1.*,SOTPICK1.SHIP_BOL_NO,SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE,SOTORDR1.CUST_STORE_NO,SOTPICK1.PICK_STATUS, SOTORDR1.CUST_DC_NO
                            from SOTCARM1,SOTPICK1,SOTSHIP1,SOTORDR1
                            where SOTPICK1.PICK_NO_CONS = SOTCARM1.PICK_NO
                              and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                              and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                              and SOTSHIP1.SHIP_BOL_NO_REV is null
                              and SOTSHIP1.ORDR_GROUP_NO = :PARM1
            			      and SOTORDR1.ORDR_cUST_PO IN (SELECT DISTINCT SOTORDR1.ORDR_CUST_PO FROM SOTCARM2,SOTORDR1 WHERE CART_NO = SOTCARM1.CART_NO AND SOTORDR1.ORDR_NO = SOTCARM2.ORDR_NO)"

            Fill_Records("SOTCARTP", ORDR_GROUP_NO, True, ASCMAIN1.sql)
            Fill_Records("SOTCARTX", New String() {ORDR_GROUP_NO, ORDR_GROUP_NO2})
            ASCMAIN1.sql = "Select SOTCARM2.*, SOTPICK1.ORDR_NO PICK_ORDR_NO
                            from SOTCARM2,SOTCARM1,SOTPICK1,SOTSHIP1,sotordr1
                            where SOTCARM1.CART_NO = SOTCARM2.CART_NO
                              AND SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                              and SOTPICK1.PICK_NO_CONS = SOTCARM1.PICK_NO
                              and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                              and SOTSHIP1.SHIP_BOL_NO_REV is null
                              and SOTSHIP1.ORDR_GROUP_NO = :PARM1
                             and SOTORDR1.ORDR_cUST_PO IN (SELECT DISTINCT SOTORDR1.ORDR_CUST_PO FROM SOTCARM2,SOTORDR1 WHERE CART_NO = SOTCARM1.CART_NO AND SOTORDR1.ORDR_NO = SOTCARM2.ORDR_NO)"

            Fill_Records("SOTCART2", ORDR_GROUP_NO, True, ASCMAIN1.sql)
        Else
            UltraExplorerBar1.Groups("Multi PO Pallet").Visible = False
            Fill_Records("SOTCART1", ORDR_GROUP_NO)
            Fill_Records("SOTCARTP", ORDR_GROUP_NO)
            Fill_Records("SOTCART2", ORDR_GROUP_NO)
        End If
        EnforceConstraints(True)
        Sort_grdColumns(grdSOTCART1, "CART_NO")
        Sort_grdColumns(grdSOTCARTP, "CUST_DC_NO,PALLET_NO")

        grdSOTCART1.Text = "Cartons on All Shipments for Order Group " & ORDR_GROUP_NO

        If ORDR_GROUP_NO <> "" And ASCMAIN1.CLIENT = "VAN" Then
            Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value
            Dim CUST_DC_NO As String = grdSOTORDR0.ActiveRow.Cells("CUST_DC_NO").Value & ""
            If ORDR_CUST_PO <> "" Then
                grdSOTCART1.Text = grdSOTCART1.Text & ", Customer PO " & ORDR_CUST_PO
                PO_CARTON_ORIG = ORDR_CUST_PO
                If IsMultiPO <> "" Then
                    ASCMAIN1.sql = "select SOTORDR0.ORDR_CUST_PO from SOTSHIP1, SOTSHIP1 CONS, SOTORDR0
                                    where SOTSHIP1.ORDR_GROUP_NO = :PARM1
                                    and SOTSHIP1.SHIP_BOL_NO_CONS = CONS.SHIP_BOL_NO_CONS
                                    and SOTSHIP1.SHIP_BOL_NO <> CONS.SHIP_BOL_NO
                                    and SOTORDR0.ORDR_GROUP_NO = cons.ORDR_GROUP_NO"
                    Dim MULTI_PO As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {ORDR_GROUP_NO})
                    grdSOTCART1.Text = grdSOTCART1.Text & ", Combined PO " & MULTI_PO
                    PO_CARTON_COMBINED = MULTI_PO
                End If
            End If
            If CUST_DC_NO <> "" Then
                grdSOTCART1.Text = grdSOTCART1.Text & ", Customer DC " & CUST_DC_NO
            End If

            If CUST_CODE = "WALMART" Then ' walmart is the onlyone with RFID at this point

                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                    ASCMAIN1.sql = "Select WHTRFID1.CART_NO, ICVLUPC1.*, WHTRFID2.SCAN_QTY FROM WHTRFID2, ICVLUPC1," & vbCrLf _
                & "  (Select CART_NO, MAX(SCAN_NO) SCAN_NO From WHTRFID1 group by CART_NO) WHTRFID1 " & vbCrLf _
                & " Where WHTRFID1.CART_NO = :PARM1 and WHTRFID1.SCAN_NO = WHTRFID2.SCAN_NO and WHTRFID2.UPC_CODE = ICVLUPC1.UPC_CODE"
                    For Each rowRFID As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {rowSOTCART1("CART_NO")}).Select("")
                        For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select($"CART_NO = '{rowRFID("CART_NO")}' and STYLE_CODE = '{rowRFID("STYLE_CODE")}' and COLOR_CODE = '{rowRFID("COLOR_CODE")}'")
                            rowSOTCART2("QTY_RFID") = rowRFID("SCAN_QTY")
                        Next
                    Next
                Next
            End If
            ASCMAIN1.sql = "Select * FROM WHTPALT1, SOTSHIP1 where SOTSHIP1.SHIP_BOL_NO = WHTPALT1.SHIP_BOL_NO And SOTSHIP1.ORDR_GROUP_NO = :PARM1"
            For Each rowPALLET As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {ORDR_GROUP_NO}).Select("")
                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select($"PALLET_NO = '{rowPALLET.Item("PALLET_NO")}'")
                    rowSOTCART1.Item("SHIP_TRAILER_NO") = rowPALLET.Item("SHIP_TRAILER_NO")
                    rowSOTCART1.Item("PALLET_INIT_DATE") = rowPALLET.Item("INIT_DATE")
                    rowSOTCART1.Item("PALLET_INIT_OPER") = rowPALLET.Item("INIT_OPER")

                Next
                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCARTP").Select($"PALLET_NO = '{rowPALLET.Item("PALLET_NO")}'")
                    rowSOTCART1.Item("SHIP_TRAILER_NO") = rowPALLET.Item("SHIP_TRAILER_NO")
                    rowSOTCART1.Item("PALLET_INIT_DATE") = rowPALLET.Item("INIT_DATE")
                    rowSOTCART1.Item("PALLET_INIT_OPER") = rowPALLET.Item("INIT_OPER")

                Next
            Next
            UpdateCartStats()
            dst.Tables("SOTCART1").AcceptChanges()
            dst.Tables("SOTCART2").AcceptChanges()
            dst.Tables("SOTCARTP").AcceptChanges()
        End If
        grdSOTCARTP.Text = grdSOTCART1.Text
        grdSOTCARTX.Text = grdSOTCART1.Text
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub UpdateCartStats()
        Me.Cursor = Cursors.WaitCursor
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Cart Stats"), UltraWinToolbars.StateButtonTool)
        Dim Sql As New System.Text.StringBuilder With {.Length = 0}
        If tlb_sbt.Checked Then
            Dim PALLET_VALUE As Double = 0
            Dim PALLET_NO_LAST As String = ""
            Dim IsMultiPO = grdSOTORDR0.ActiveRow.Cells("EDI_CONS_NO").Value & ""

            For Each rowSOTCARTP As DataRow In dst.Tables("SOTCARTP").Select("", "PALLET_NO, CART_NO")
                Dim CART_NO As String = rowSOTCARTP.Item("CART_NO").ToString & String.Empty
                Dim PALLET_NO As String = rowSOTCARTP.Item("PALLET_NO").ToString & String.Empty

                Sql.Length = 0
                Sql.AppendLine("SELECT")
                Sql.AppendLine("C1.CART_TOTAL_UNITS,")
                Sql.AppendLine("C1.CART_TOTAL_WGT_ACTUAL,")
                Sql.AppendLine("C1.CART_TOTAL_WGT_CALC,")
                Sql.AppendLine("C1.CART_TRACKING_NO,")
                Sql.AppendLine("TO_CHAR (P1.INIT_DATE, 'HH12:MI:SS AM') AS SCAN_TIME,")
                Sql.AppendLine("S1.SHIP_LOAD_NO,")
                Sql.AppendLine("S1.BILL_OF_LADING_NO,")
                Sql.AppendLine("S1.MASTER_SHIP_BOL_NO")
                Sql.AppendLine("FROM SOTCART1 C1, WHTPALT1 P1, SOTSHIP1 S1")
                Sql.AppendLine("WHERE C1.PALLET_NO = P1.PALLET_NO (+)")
                Sql.AppendLine("AND P1.SHIP_BOL_NO = S1.SHIP_BOL_NO (+)")
                Sql.AppendLine($"AND C1.CART_NO = '{CART_NO}'")
                Dim tbl As DataTable = ASCDATA1.GetDataTable(Sql.ToString())
                If tbl.Rows.Count = 1 Then
                    rowSOTCARTP.Item("TRACKING_NO") = tbl.Rows(0).Item("CART_TRACKING_NO").ToString & String.Empty
                    Dim CARTON_WEIGHT As Double = 0
                    If Val(tbl.Rows(0).Item("CART_TOTAL_WGT_ACTUAL").ToString & String.Empty) > 0 Then
                        CARTON_WEIGHT = Format(Val(tbl.Rows(0).Item("CART_TOTAL_WGT_ACTUAL").ToString & String.Empty), "###,##0.00")
                    Else
                        CARTON_WEIGHT = Format(Val(tbl.Rows(0).Item("CART_TOTAL_WGT_CALC").ToString & String.Empty), "###,##0.00")
                    End If
                    rowSOTCARTP.Item("CARTON_WEIGHT") = CARTON_WEIGHT
                    rowSOTCARTP.Item("SCAN_TIME") = tbl.Rows(0).Item("SCAN_TIME").ToString & String.Empty
                    rowSOTCARTP.Item("SHIP_LOAD_NO") = tbl.Rows(0).Item("SHIP_LOAD_NO").ToString & String.Empty
                    rowSOTCARTP.Item("BILL_OF_LADING_NO") = tbl.Rows(0).Item("BILL_OF_LADING_NO").ToString & String.Empty
                    rowSOTCARTP.Item("MASTER_SHIP_BOL_NO") = tbl.Rows(0).Item("MASTER_SHIP_BOL_NO").ToString & String.Empty

                    Dim WORKCARTON As String = rowSOTCARTP.Item("CART_NO") & ""
                    Dim MULTIPOS As String = ""
                    If WORKCARTON <> "" Then
                        ASCMAIN1.sql = "SELECT DISTINCT SOTORDR1.ORDR_CUST_PO FROM SOTCARM2,SOTORDR1 WHERE CART_NO = '" & WORKCARTON & "'" _
                        & " AND SOTORDR1.ORDR_NO = SOTCARM2.ORDR_NO"
                        For Each rowMULTIPO As DataRow In ASCDATA1.GetDataTable.Rows
                            If MULTIPOS <> "" Then
                                MULTIPOS = MULTIPOS & ","
                            End If
                            MULTIPOS = MULTIPOS & rowMULTIPO.Item("ORDR_CUST_PO")
                        Next
                    End If
                    rowSOTCARTP.Item("MULTI_PO") = MULTIPOS
                End If
                Dim CARTON_POS As String = ""
                If IsMultiPO <> "" Then
                    Sql.Length = 0
                    Sql.AppendLine("SELECT")
                    Sql.AppendLine("SUM(C2.QTY_PACKED * O2.ORDR_UNIT_PRICE) AS PALLET_VALUE")
                    Sql.AppendLine("FROM SOTCARM1 C1, SOTCARM2 C2, SOTORDR2 O2")
                    Sql.AppendLine("WHERE C1.CART_NO = C2.CART_NO")
                    Sql.AppendLine("AND C2.ORDR_NO = O2.ORDR_NO")
                    Sql.AppendLine("AND C2.ORDR_LNO = O2.ORDR_LNO")
                    Sql.AppendLine($"AND C1.CART_NO = '{CART_NO}'")
                    CARTON_POS = PO_CARTON_ORIG & "," & PO_CARTON_COMBINED
                Else
                    Sql.Length = 0
                    Sql.AppendLine("SELECT")
                    Sql.AppendLine("SUM(C2.QTY_PACKED * O2.ORDR_UNIT_PRICE) AS PALLET_VALUE")
                    Sql.AppendLine("FROM SOTCART1 C1, SOTCART2 C2, SOTORDR2 O2")
                    Sql.AppendLine("WHERE C1.CART_NO = C2.CART_NO")
                    Sql.AppendLine("AND C2.ORDR_NO = O2.ORDR_NO")
                    Sql.AppendLine("AND C2.ORDR_LNO = O2.ORDR_LNO")
                    Sql.AppendLine($"AND C1.CART_NO = '{CART_NO}'")
                    CARTON_POS = PO_CARTON_ORIG
                End If
                ASCMAIN1.sql = Sql.ToString()
                Dim CARTON_VALUE As Double = Val(ASCDATA1.GetDataValue)
                rowSOTCARTP.Item("CARTON_VALUE") = CARTON_VALUE
                If rowSOTCARTP.Item("CART_TRACKING_NO") & "" = "" Then
                    rowSOTCARTP.Item("CART_TRACKING_NO") = "NA"
                End If
                If PALLET_NO_LAST <> PALLET_NO Then
                    If IsMultiPO <> "" Then
                        Sql.Length = 0
                        Sql.AppendLine("SELECT")
                        Sql.AppendLine("SUM(C2.QTY_PACKED * O2.ORDR_UNIT_PRICE) AS PALLET_VALUE")
                        Sql.AppendLine("FROM SOTCARM1 C1, SOTCARM2 C2, SOTORDR2 O2")
                        Sql.AppendLine("WHERE C1.CART_NO = C2.CART_NO")
                        Sql.AppendLine("AND C2.ORDR_NO = O2.ORDR_NO")
                        Sql.AppendLine("AND C2.ORDR_LNO = O2.ORDR_LNO")
                        Sql.AppendLine($"AND C1.PALLET_NO = '{PALLET_NO}'")
                    Else
                        Sql.Length = 0
                        Sql.AppendLine("SELECT")
                        Sql.AppendLine("SUM(C2.QTY_PACKED * O2.ORDR_UNIT_PRICE) AS PALLET_VALUE")
                        Sql.AppendLine("FROM SOTCART1 C1, SOTCART2 C2, SOTORDR2 O2")
                        Sql.AppendLine("WHERE C1.CART_NO = C2.CART_NO")
                        Sql.AppendLine("AND C2.ORDR_NO = O2.ORDR_NO")
                        Sql.AppendLine("AND C2.ORDR_LNO = O2.ORDR_LNO")
                        Sql.AppendLine($"AND C1.PALLET_NO = '{PALLET_NO}'")
                    End If
                    ASCMAIN1.sql = Sql.ToString()
                    PALLET_VALUE = Val(ASCDATA1.GetDataValue)
                    PALLET_NO_LAST = PALLET_NO
                End If
                rowSOTCARTP.Item("PALLET_VALUE") = PALLET_VALUE
            Next

            Sort_grdColumns(grdSOTCARTP, "CUST_DC_NO,PALLET_NO")
            PALLET_NO_LAST = ""
            Dim PALLET_COLOR As Drawing.Color = Drawing.Color.LightBlue
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTCARTP.Rows
                Dim PALLET_NO As String = grow.Cells.Item("PALLET_NO").Text & String.Empty
                If PALLET_NO_LAST <> PALLET_NO Then
                    If PALLET_COLOR = Drawing.Color.LightBlue Then
                        PALLET_COLOR = Drawing.Color.LightGray
                    Else
                        PALLET_COLOR = Drawing.Color.LightBlue
                    End If
                    PALLET_NO_LAST = PALLET_NO
                End If
                grow.Cells("PALLET_NO").Appearance.BackColor = PALLET_COLOR
                For Each CH As String In PALLET_NO
                    If CH = "0" Then
                        PALLET_NO = PALLET_NO.Substring(1)
                    Else
                        Exit For
                    End If
                Next
                grow.Cells.Item("PALLET_NO").Value = PALLET_NO
            Next

        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub optOrders_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optOrders.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTORDR0("", CUST_CODE)
    End Sub

    Private Sub chkReservations_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkReservations.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTORDR0("", CUST_CODE)
    End Sub

    Private Sub cmbSALES_DIVISION_CODE_ValueChanged(sender As Object, e As System.EventArgs) Handles cmbSALES_DIVISION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If ScreenMode Then
            Load_SOTORDR0(Absx1.txtFor("CUST_CODE").Text)
        Else
            Load_SOTORDR0()
        End If

    End Sub

    Private Sub grdSOTCORDD_DoubleClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdSOTCORDD.DoubleClickCell
        If e.Cell.Row.IsDataRow Then
            Dim SORT_SEQ As String = e.Cell.Row.Cells("SORT_SEQ").Value
            If SORT_SEQ = "1" Or SORT_SEQ = "8" Then
                Dim STYLE_CODE As String = e.Cell.Row.Cells("CODE_VALUE").Value
                Dim YP As String = e.Cell.Column.Tag
                Fill_Records("SOTCORDY", New String() {CUST_CODE, STYLE_CODE, YP})
                grdSOTCORDY.Text = grdSOTCORDX.ActiveRow.Cells("CODE_VALUE").Value & " Invoice Details for Customer " & CUST_CODE & ", Style " & STYLE_CODE & ", in " & e.Cell.Column.Header.Caption
                tabMonth.Tabs("Details").Visible = True
                tabMonth.SelectedTab = tabMonth.Tabs("Details")
            End If
        End If
    End Sub

    Sub CreateGraph_SATCSLS1_X()

        Dim chtIsVisible As Boolean = chtSATCSLS1_X.Visible
        chtSATCSLS1_X.Visible = False

        chtSATCSLS1_X.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String
        Dim Periods As Int32 = 12
        ReDim CL(Periods)

        'this will be necessary for line graph
        'For i As Integer = MOSMAX To 0 Step -1
        '    Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
        '    CL(MOSMAX - i) = Mid(L, 10, 6)
        '    grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        'Next
        For i As Integer = 1 To Periods
            'Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
            CL(i - 1) = grdSOTCORDX.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Header.Caption
            'grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        Next

        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtSATCSLS1_X.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATCSLS1_X.LabelHash = labelHash

        chtSATCSLS1_X.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATCSLS1_X.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To Periods
            DT.Columns.Add("V" & Format(P, "00"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SOTCORDX").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SOTCORDX").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1

            Dim rowDT As DataRow = DT.NewRow
            rowDT.Item("CODE_VALUE") = row("CODE_VALUE")
            'rowDT.Item("DESC_VALUE") = row("DESC_VALUE")
            For P As Integer = 1 To Periods
                rowDT.Item("V" & Format(P, "00")) = row("V" & Format(P, "00"))
            Next
            DT.Rows.Add(rowDT)
        Next
        chtSATCSLS1_X.Data.SetRowLabels(RL)
        chtSATCSLS1_X.Data.SetColumnLabels(CL)

        chtSATCSLS1_X.DataSource = DT
        'chtSATCSLS1_X.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("P00", False)

        chtSATCSLS1_X.DataBind()

        chtSATCSLS1_X.Visible = True ' chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Private Sub tabStyles_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabStyles.SelectedTabChanged
        'SplitContainer1.Panel2Collapsed = (tabStyles.SelectedTab.Key = "Summary")
        grdSOTORDRX.Visible = Not (tabStyles.SelectedTab.Key = "Summary")
        optQTY.Visible = (tabStyles.SelectedTab.Key = "Summary")
        Setup_Summary()
    End Sub

    Private Sub optQTY_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optQTY.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Summary()
    End Sub

    Sub Store_Configuration_Report(ByVal ORDR_GROUP_NOs As List(Of String))
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Store Configuration Report")

        Dim RPT As String = "SORCONF1"
        'Dim sqlw As String = " AND SOTORDR0.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
        Dim sqlw As String = " AND SOTORDR0.ORDR_GROUP_NO in ('" & Join(ORDR_GROUP_NOs.ToArray, "','") & "')"
        If Not REPORTS.ContainsKey(RPT) Then
            REPORTS.Add(RPT, Load_rptClass(RPT))
            REPORTS(RPT).Prepare_dst(True, sqlw)

        Else
            REPORTS(RPT).Fill_Records_RPT(sqlw)
        End If


        Dim FILENAME As String = ""
        With REPORTS(RPT).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            'Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", ORDR_GROUP_NO, False)
            'FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
            .Generate_Report(RPT, "Store Configuration Report", "", True)
            .Print_Report_End()
            ' .Print_Report_End(, True)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        ' Return FILENAME
    End Sub

    'Function Create_Invoice( _
    '                       SHIP_BOL_NO As String, _
    '                       PICK_NO As String, _
    '                       Optional makepdf As Boolean = False, _
    '                       Optional pro_forma As Boolean = False) As String

    '    Me.Cursor = Cursors.WaitCursor
    '    ASCMAIN1.Progress("Now Preparing Invoice")

    '    Dim REPORTFILE As String = "SORINVP1"
    '    If Not REPORTS.ContainsKey(REPORTFILE) Then
    '        REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
    '        REPORTS(REPORTFILE).Prepare_dst(False, "")
    '    End If

    '    'To fill the report's dataset with data from Oracle, 
    '    ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

    '    Dim sqlw As String = ""
    '    If SHIP_BOL_NO <> "" Then
    '        sqlw = " and SOTINVH1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
    '    Else
    '        sqlw = " and SOTINVH1.PICK_NO = '" & PICK_NO & "'"
    '    End If
    '    REPORTS(REPORTFILE).Fill_Records_RPT(New String() {sqlw, IIf(pro_forma, "1", "0")})

    '    With REPORTS(REPORTFILE).clsASCBASE1
    '        .Print_Report_Begin()
    '        .CR_params.Add("SUBT", "")
    '        .CR_params.Add("CONS_INV", "0")
    '        .CR_params.Add("EXPORT_INFO", "0")
    '        .Generate_Report("SORINVP1", "Sales Invoice", , True, , , , , False)
    '        .Print_Report_End()
    '    End With

    '    Me.Cursor = Cursors.WaitCursor
    '    ASCMAIN1.Progress("")

    '    Return ""

    'End Function

    'Sub Display_Raw(Vandale_SO As String)
    '    'we need to capture gen_doc_no in EDT850T1 table to do it correct way
    '    'Dim EDI_DOC_SEQ_NO As String
    '    'ASCMAIN1.sql = "Select GD.^DocumentBlobKEY^ RAW_DATA_FILE " _
    '    '    & " from GEN.^Document_tb^ GD, EDT850T1 " _
    '    '    & " where EDT850T1.GEN_DOC_NO = GD.^AppField1^" _
    '    '    & "   and GD.^TransactionSetID^ = '850'" _
    '    '    & "   and EDT850T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
    '    ASCMAIN1.sql = "Select ORDR_CUST_PO from SOTORDR1 " _
    '    & " where ORDR_NO = '" & Vandale_SO & "'" _
    '    & " Union " _
    '    & " Select ORDR_CUST_PO from SOTINVH1" _
    '    & " where ORDR_NO = '" & Vandale_SO & "'"

    '    ASCMAIN1.sql = "Select EDI_DOC_SEQ_NO from SOTORDR1 " _
    '    & " where ORDR_NO = '" & Vandale_SO & "'" 

    '    Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow
    '    If rowSOTORDR1 IsNot Nothing Then
    '        If rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & "" <> "" Then
    '            Dim RAW_DATA As String = TAC.SOCMAIN1.Get_Raw_EDI(rowSOTORDR1.Item("EDI_DOC_SEQ_NO"))

    '            Using F As New ASFMSGBF
    '                F.Show_Formatted_txt("Raw for Order: " & Vandale_SO, RAW_DATA, Me)
    '                ' F.Show_txt("Raw for Order: " & Vandale_SO, RAW_DATA, Me)
    '            End Using
    '            Exit Sub
    '        End If
    '        'If rowSOTORDR1.Item("ORDR_CUST_PO") & "" <> "" Then
    '        '    ASCMAIN1.sql = " Select GD.DocumentBlobKEY RAW_DATA_FILE " & vbCrLf _
    '        '    & "  from Document_tb GD " & vbCrLf _
    '        '    & "  where GD.TransactionSetID = '850'" & vbCrLf _
    '        '    & "    and GD.DOCUMENTNAME = '" & rowSOTORDR1.Item("ORDR_CUST_PO") & "'"
    '        '    '  If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.DBS_SERVER = "" Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GENAHA.")
    '        '    ASCMAIN1.sql = ASCMAIN1.sql
    '        '    Dim RAW_DATA_FILE As String = ASCDATA1.GetDataValue
    '        '    If RAW_DATA_FILE <> "" Then
    '        '        Dim RAW_DATA As String = ""
    '        '        Dim FILENAME As String = "\\EDServer1\GTranDoc\" & RAW_DATA_FILE & ".DOC"
    '        '        If My.Computer.FileSystem.FileExists(FILENAME) Then
    '        '            RAW_DATA = My.Computer.FileSystem.ReadAllText(FILENAME)

    '        '            Using F As New ASFMSGBF
    '        '                F.Show_Formatted_txt("Raw for Order: " & Vandale_SO, RAW_DATA, Me)
    '        '            End Using
    '        '            Exit Sub
    '        '        End If
    '        '    End If
    '        'End If
    '    End If
    '    MsgBox("Now Raw data found for SO No: " & Vandale_SO, MsgBoxStyle.OkOnly, "Warning")
    'End Sub

    Private Sub grdSOTORDRS_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORDRS.InitializeLayout

    End Sub

    Private Sub grdSOTORDRS_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRS.InitializeRow

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
            Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
            ASCMAIN1.sql = "Select STYLE_STATUS, STYLE_COLOR_STATUS" _
                & " from ICTSTYL1,ICTSTYC1" _
                & " where ICTSTYC1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" _
                & "   and ICTSTYC1.COLOR_CODE (+) = '" & COLOR_CODE & "'" _
                & "   and ICTSTYL1.STYLE_CODE = '" & STYLE_CODE & "'"
            Dim row As DataRow = ASCDATA1.GetDataRow ' (ASCMAIN1.sql, "VV", New String() {STYLE_CODE, COLOR_CODE})
            If row Is Nothing Then Exit Sub
            If row.Item("STYLE_STATUS") & "" = "D" Then
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("STYLE_CODE").ToolTipText = "Style is Discontinued"
            Else
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("STYLE_CODE").ToolTipText = ""
            End If
            If row.Item("STYLE_COLOR_STATUS") & "" = "D" Then
                e.Row.Cells("COLOR_CODE").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("COLOR_CODE").ToolTipText = "Color is Discontinued"
            Else
                e.Row.Cells("COLOR_CODE").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("COLOR_CODE").ToolTipText = ""
            End If
        Else
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("STYLE_CODE").ToolTipText = ""
            e.Row.Cells("COLOR_CODE").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("COLOR_CODE").ToolTipText = ""
        End If

        If e.Row.Cells("ORDR_RELEASE_AVAIL").Value & "" <> "" Then
            Dim ORDR_RELEASE_AVAIL As Date = e.Row.Cells("ORDR_RELEASE_AVAIL").Value
            Dim WIP_IND As String = e.Row.Cells("WIP_IND").Value & ""
            Dim TT As String = ""
            If WIP_IND = "P" Then
                TT = " (PO is Open)"
                e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.BackColor = Drawing.Color.Yellow
            ElseIf WIP_IND = "S" Then
                TT = " (PO is Shipped)"
            End If

            Dim ORDR_CANCEL_DATE As Date = grdSOTORDR0.ActiveRow.Cells("ORDR_CANCEL_DATE").Value
            If Format(ORDR_RELEASE_AVAIL, "yyyyMMdd") > Format(ORDR_CANCEL_DATE, "yyyyMMdd") Then
                e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = "Availability Date is Past Cancel Date" & TT
            Else
                e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = TT
            End If
        Else
            If Val(e.Row.Cells("ORDR_QTY_ALLO").Value & "") <> 0 Then
                e.Row.Cells("ORDR_RELEASE_AVAIL").Appearance.BackColor = Drawing.Color.LightGreen
                e.Row.Cells("ORDR_RELEASE_AVAIL").ToolTipText = "Green: Allocation is Drawing from Current Stock O/H"
            End If
        End If
    End Sub

    Private Sub optFindBy_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optFindBy.ValueChanged

    End Sub

    Private Sub chkReservations_FindCustomerBy_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkReservations_FindCustomerBy.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub

        Absx1.txtFor("SREP_CODE").Text = ""
        Load_SOTORDR0()
    End Sub

    Private Sub grdSOTORDR0_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORDR0.InitializeLayout

        For Each colname As String In New String() {"CUST_CITY", "CUST_STATE", "CUST_COUNTRY"}
            grdSOTORDR0.DisplayLayout.Bands(0).Columns(colname).Hidden = Not (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
        Next

    End Sub

    Private Sub grdSOTORDR1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR1.AfterRowActivate

        If grdSOTORDR1.ActiveRow.IsDataRow Then
            If tabDetails.SelectedTab.Key = "All Orders" Then

                Dim ORDR_GROUP_NO As String = grdSOTORDR1.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                If ORDR_GROUP_NO <> "" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                        If grow.IsDataRow AndAlso grow.Cells("ORDR_GROUP_NO").Value = ORDR_GROUP_NO Then
                            grow.Activate()
                            Exit For
                        End If
                    Next
                End If
            End If

        End If
    End Sub

    Private Sub grdSOTORDR1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORDR1.InitializeLayout

    End Sub

    Private Sub chkActionDate_CheckedChanged(sender As Object, e As EventArgs) Handles chkActionDate.CheckedChanged
        Load_SOTORDR0()
    End Sub

    Private Sub chkIfReceivedSince_CheckedChanged(sender As Object, e As EventArgs) Handles chkIfReceivedSince.CheckedChanged
        Load_SOTORDR0()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Load_SOTORDR0()
    End Sub

    Private Sub chkCOSTS_CheckedChanged(sender As Object, e As EventArgs) Handles chkCOSTS.CheckedChanged

    End Sub

    Private Sub chkDetails_CheckedChanged(sender As Object, e As EventArgs) Handles chkDetails.CheckedChanged
        toggle_Show_Details()
    End Sub

    Sub toggle_Show_Details()
        splSOTORDR0s.Panel2Collapsed = Not (chkDetails.Checked)
    End Sub

    Private Sub chkShortView_CheckedChanged(sender As Object, e As EventArgs) Handles chkShortView.CheckedChanged

        If COLUMN_NAMEs_All.Count = 0 Then
            For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORDR0.DisplayLayout.Bands(0).Columns
                If Not gcol.Hidden Then
                    COLUMN_NAMEs_All.Add(gcol.Key)
                End If
            Next

            COLUMN_NAMEs_Short.Add("CUST_CODE")
            COLUMN_NAMEs_Short.Add("CUST_NAME")
            COLUMN_NAMEs_Short.Add("ORDR_CUST_PO")
            COLUMN_NAMEs_Short.Add("ORDR_SHIP_DATE")
            COLUMN_NAMEs_Short.Add("ORDR_CANCEL_DATE")
            COLUMN_NAMEs_Short.Add("ORDR_QTY_OPEN")
            COLUMN_NAMEs_Short.Add("ORDR_AMT_OPEN")


        End If

        For Each COLUMN_NAME As String In COLUMN_NAMEs_All
            grdSOTORDR0.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = chkShortView.Checked And Not COLUMN_NAMEs_Short.Contains(COLUMN_NAME)
        Next
    End Sub

    Private Sub grdSOTCART1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART1.InitializeRow
        If e.Row.Band.Key = "SOTCART1" Then
            If e.Row.Cells("PICK_STATUS").Value & "" = "D" Then
                e.Row.Cells("PICK_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("PICK_STATUS").ToolTipText = "De-Released"
            End If
        End If
        If e.Row.Band.Key = "SOTCART1_SOTCART2" Then
            If e.Row.Cells("PICK_ORDR_NO").Value & "" <> e.Row.Cells("ORDR_NO").Value & "" Then
                e.Row.Appearance.ForeColor = Drawing.Color.Red
                e.Row.ToolTipText = "Different PO"
            End If
        End If
    End Sub

    Private Sub cmdVerifyBL_Click(sender As Object, e As EventArgs) Handles cmdVerifyBL.Click
        ' verify 7 digits
        ' prepend with 0194546000
        ' create sql with master or not based on checkbox
        ' make datatable
        ' show grd
        If Len(txtBL.Text) <> 7 Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed BOL Must be 7 Charcters")
            Exit Sub
        End If

        Dim BOLNO As String = "0194546000" & txtBL.Text
        Dim SQLWHERE As String = " WHERE SOTSHIP1.bill_of_lading_no in ('" & BOLNO & "')" '  ('01945460008174010')"
        If chkMasterBL.Checked = True Then
            SQLWHERE = " WHERE SOTSHIP1.bill_of_lading_no in (select bol_no from sotshipb WHERE MASTER_Bol_NO = '" & BOLNO & "')"
        End If

        ASCMAIN1.sql = "SELECT CASE WHEN PICK_QTY_CONF = QTY_PACKED THEN ' ' ELSE 'X' END ERROR" & vbCrLf _
            & " , Z.SHIP_BOL_NO, Z.QTY_PACKED, Z.PICK_QTY_CONF" & vbCrLf _
            & " FROM(" & vbCrLf _
            & " SELECT X.SHIP_BOL_NO" & vbCrLf _
            & " , SUM (QTY_PACKED) QTY_PACKED, SUM (QTY_REL) QTY_REL" & vbCrLf _
            & " , SUM (CART_TOTAL_UNITS) CART_TOTAL_UNITS" & vbCrLf _
            & " , SUM (CART_TOTAL_UNITS_REL) CART_TOTAL_UNITS_REL" & vbCrLf _
            & " , SUM (PICK_QTY) PICK_QTY" & vbCrLf _
            & " , SUM (PICK_QTY_CONF) PICK_QTY_CONF FROM (" & vbCrLf _
            & " Select SOTPICK1.SHIP_BOL_NO," & vbCrLf _
            & " sum(SOTCART2.QTY_PACKED) QTY_PACKED, sum (SOTCART2.QTY_REL) QTY_REL" & vbCrLf _
            & " From SOTCART1, SOTPICK1, SOTCART2" & vbCrLf _
            & " WHERE SOTPICK1.SHIP_BOL_NO IN (select sotship1.ship_bol_no " & vbCrLf _
            & " From sOTSHIP1, SOTORDR0, WHTWAVE3" & vbCrLf _
            & SQLWHERE & vbCrLf _
            & " And SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
            & " And WHTWAVE3.WAVE_NO = SOTSHIP1.WAVE_NO" & vbCrLf _
            & " And WHTWAVE3.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO)" & vbCrLf _
            & " And SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
            & " And SOTCART1.PICK_NO = SOTPICK1.PICK_NO GROUP BY SOTPICK1.SHIP_BOL_NO) W" & vbCrLf _
            & " , (" & vbCrLf _
            & " Select SOTPICK1.SHIP_BOL_NO," & vbCrLf _
            & " sum(SOTCART1.CART_TOTAL_UNITS) CART_TOTAL_UNITS, sum (SOTCART1.CART_TOTAL_UNITS_REL) CART_TOTAL_UNITS_REL, COUNT (*) CARTONS" & vbCrLf _
            & " From SOTCART1, SOTPICK1" & vbCrLf _
            & " Where SOTPICK1.SHIP_BOL_NO In (Select sotship1.ship_bol_no" & vbCrLf _
            & " From sOTSHIP1, SOTORDR0, WHTWAVE3" & vbCrLf _
            & SQLWHERE & vbCrLf _
            & " And SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
            & " And WHTWAVE3.WAVE_NO = SOTSHIP1.WAVE_NO" & vbCrLf _
            & " And WHTWAVE3.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO)" & vbCrLf _
            & " And SOTCART1.PICK_NO = SOTPICK1.PICK_NO GROUP BY SOTPICK1.SHIP_BOL_NO) X" & vbCrLf _
            & " , (" & vbCrLf _
            & " Select SOTPICK1.SHIP_BOL_NO," & vbCrLf _
            & " sum(SOTPICK2.PICK_QTY) PICK_QTY, sum (SOTPICK2.PICK_QTY_CONF) PICK_QTY_CONF, COUNT (DISTINCT SOTPICK1.PICK_NO) PICKS" & vbCrLf _
            & " From SOTPICK1, SOTPICK2" & vbCrLf _
            & " Where SOTPICK1.SHIP_BOL_NO In (Select sotship1.ship_bol_no" & vbCrLf _
            & " From sOTSHIP1, SOTORDR0, WHTWAVE3" & vbCrLf _
            & SQLWHERE & vbCrLf _
            & " And SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
            & " And WHTWAVE3.WAVE_NO = SOTSHIP1.WAVE_NO" & vbCrLf _
            & " And WHTWAVE3.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO)" & vbCrLf _
            & " And SOTPICK2.PICK_NO = SOTPICK1.PICK_NO GROUP BY SOTPICK1.SHIP_BOL_NO) Y" & vbCrLf _
            & " WHERE x.SHIP_BOL_NO = y.SHIP_BOL_NO And x.SHIP_BOL_NO = w.SHIP_BOL_NO" & vbCrLf _
            & " GROUP by X.SHIP_BOL_NO" & vbCrLf _
            & " ) Z"
        Dim DT As DataTable = ASCDATA1.GetDataTable




        If DT.Rows.Count <> 0 Then
            Using F As New ASFMSGBF
                F.Show_grd(DT, Me, "BOL lines To make sure quantities match up", "")
            End Using
        End If


    End Sub

    Public Overrides Function CustomSummary_End(
   ByVal summarySettings As UltraWinGrid.SummarySettings,
   ByVal rows As UltraWinGrid.RowsCollection,
   ByVal CustomValue As Double,
   ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdSOTCARTP"
                Dim KEY As String = summarySettings.Key
                If KEY = "PALLET_NO" Then
                    TOTALS.Add("PALLET_NO", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("PALLET_NO") <> 0 Then CustomValue = TOTALS("PALLET_NO")
                ElseIf KEY = "SHIP_TRAILER_NO" Then
                    TOTALS.Add("SHIP_TRAILER_NO", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("SHIP_TRAILER_NO") <> 0 Then CustomValue = TOTALS("SHIP_TRAILER_NO")
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Public Overrides Function CustomStringSummary_End(
        ByVal summarySettings As UltraWinGrid.SummarySettings,
        ByVal rows As UltraWinGrid.RowsCollection,
        ByVal CustomValue As String,
        ByVal grd As UltraWinGrid.UltraGrid) As String

        Select Case grd.Name
            Case "grdSOTCARTP"
                Dim KEY As String = summarySettings.Key
                CustomValue = "Palletized"
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub CustomSummary_Calculate_Totals(
       ByVal rows As UltraWinGrid.RowsCollection,
       ByRef TOTALS As Dictionary(Of String, Decimal),
       ByVal KEY As String)


        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                CustomSummary_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If Not grow2.IsFilteredOut Then
                    If KEY = "PALLET_NO" Then
                        TOTALS("PALLET_NO") += IIf(grow2.Cells("PALLET_NO").Value & "" <> "", 1, 0)
                    ElseIf KEY = "SHIP_TRAILER_NO" Then
                        TOTALS("SHIP_TRAILER_NO") += IIf(grow2.Cells("SHIP_TRAILER_NO").Value & "" <> "", 1, 0)
                    End If
                End If
            End If
        Next
    End Sub

    Sub Fill_ICTSTATA()

        'Dim ORDR_TYPE As String = grdSOTORDR0.ActiveRow.Cells("ORDR_TYPE").Value & ""
        'Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
        Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
        'Dim ORDR_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_NO").Value & ""

        Dim SQL As New System.Text.StringBuilder With {.Length = 0}
        SQL.AppendLine("Select")
        SQL.AppendLine("X.STYLE_CODE,")
        SQL.AppendLine("X.COLOR_CODE,")
        SQL.AppendLine("ICTSTYL1.STYLE_DESC,")
        SQL.AppendLine("ICTCOLR1.COLOR_DESC,")
        SQL.AppendLine("SUM(X.BEG) BEG, SUM(X.SHP) SHP, SUM(X.RTN) RTN, SUM(X.REC) REC,")
        SQL.AppendLine("SUM(X.ADJ) ADJ, SUM(X.XFR) XFR, SUM(X.PHY) PHY, SUM(X.ON_HAND) ON_HAND,")
        SQL.AppendLine("SUM(X.ON_ORDER) ON_ORDER, SUM(X.TRAN) TRAN, SUM(X.OPEN) OPEN,")
        SQL.AppendLine("SUM(X.PICK) PICK, SUM(X.ALLO) ALLO, SUM(X.COMM) COMM, SUM(X.PROD) PROD,")
        SQL.AppendLine("MAX(UPC_CODE) UPC_CODE, MAX(STYLE_COLOR_STATUS) STYLE_COLOR_STATUS from ICTCOLR1, ICTSTYL1,")
        SQL.AppendLine("(")
        SQL.AppendLine("    (")
        SQL.AppendLine("        Select")
        SQL.AppendLine("        ICTSTAT1.STYLE_CODE,")
        SQL.AppendLine("        ICTSTAT1.COLOR_CODE,")
        SQL.AppendLine("        SUM(ICTSTAT1.WHSE_QTY_BEG) BEG,")
        SQL.AppendLine("        SUM(ICTSTAT1.WHSE_QTY_SHP) SHP, SUM(ICTSTAT1.WHSE_QTY_RTN) RTN,")
        SQL.AppendLine("        SUM(ICTSTAT1.WHSE_QTY_REC) REC, SUM(ICTSTAT1.WHSE_QTY_ADJ) ADJ,")
        SQL.AppendLine("        SUM(ICTSTAT1.WHSE_QTY_XFR) XFR, SUM(ICTSTAT1.WHSE_QTY_PHY) PHY,")
        SQL.AppendLine("        SUM(0) ON_HAND, SUM (0) ON_ORDER, SUM (0) TRAN, SUM (0) OPEN, SUM (0) PICK, SUM (0) ALLO, SUM (0) COMM, SUM (0) PROD,")
        SQL.AppendLine("        NULL UPC_CODE, NULL STYLE_COLOR_STATUS from ICTSTAT1")
        SQL.AppendLine("        where (ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE) IN")
        SQL.AppendLine("        (")
        SQL.AppendLine("            SELECT")
        SQL.AppendLine("            DISTINCT O2.STYLE_CODE, O2.COLOR_CODE")
        SQL.AppendLine("            FROM SOTORDR1 O1, SOTORDR2 O2")
        SQL.AppendLine("            WHERE O1.ORDR_NO = O2.ORDR_NO")
        SQL.AppendLine($"            AND O1.ORDR_GROUP_NO = '{ORDR_GROUP_NO}'")
        SQL.AppendLine("        )")
        SQL.AppendLine($"        and ICTSTAT1.OPS_YYYYPP = '{ASCMAIN1.CYP}'")
        SQL.AppendLine("        group by ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE")
        SQL.AppendLine("    )")
        SQL.AppendLine("    union")
        SQL.AppendLine("    (")
        SQL.AppendLine("        Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE,")
        SQL.AppendLine("        SUM(0) BEG,")
        SQL.AppendLine("        SUM (0) SHP,")
        SQL.AppendLine("        SUM (0) RTN,")
        SQL.AppendLine("        SUM (0) REC,")
        SQL.AppendLine("        SUM (0) ADJ,")
        SQL.AppendLine("        SUM (0) XFR,")
        SQL.AppendLine("        SUM (0) PHY,")
        SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_ON_HAND) ON_HAND,")
        SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_ON_ORDER) ON_ORDER,")
        SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_TRAN) TRAN,")
        SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_OPEN) OPEN, SUM(ICTSTAT2.WHSE_QTY_PICK) PICK,")
        SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_ALLO) ALLO,")
        SQL.AppendLine("        SUM(ICTSTAT2.WHSE_QTY_COMM) COMM, SUM(ICTSTAT2.WHSE_QTY_PROD) PROD,")
        SQL.AppendLine("        NULL UPC_CODE, NULL STYLE_COLOR_STATUS from ICTSTAT2")
        SQL.AppendLine("        where (ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE) IN")
        SQL.AppendLine("        (")
        SQL.AppendLine("            SELECT")
        SQL.AppendLine("            DISTINCT O2.STYLE_CODE, O2.COLOR_CODE")
        SQL.AppendLine("            FROM SOTORDR1 O1, SOTORDR2 O2")
        SQL.AppendLine("            WHERE O1.ORDR_NO = O2.ORDR_NO")
        SQL.AppendLine($"            AND O1.ORDR_GROUP_NO = '{ORDR_GROUP_NO}'")
        SQL.AppendLine("        )")
        SQL.AppendLine("        group by ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE")
        SQL.AppendLine("    )")
        SQL.AppendLine("    union")
        SQL.AppendLine("    (")
        SQL.AppendLine("        Select")
        SQL.AppendLine("        ICTSTYC1.STYLE_CODE,")
        SQL.AppendLine("        ICTSTYC1.COLOR_CODE,")
        SQL.AppendLine("        0 BEG,")
        SQL.AppendLine("        0 SHP,")
        SQL.AppendLine("        0 RTN,")
        SQL.AppendLine("        0 REC,")
        SQL.AppendLine("        0 ADJ,")
        SQL.AppendLine("        0 XFR,")
        SQL.AppendLine("        0 PHY,")
        SQL.AppendLine("        0 ON_HAND,")
        SQL.AppendLine("        0 ON_ORDER,")
        SQL.AppendLine("        0 TRAN,")
        SQL.AppendLine("        0 OPEN,")
        SQL.AppendLine("        0 PICK,")
        SQL.AppendLine("        0 ALLO,")
        SQL.AppendLine("        0 COMM,")
        SQL.AppendLine("        0 PROD,")
        SQL.AppendLine("        ICTSTYC1.UPC_CODE,")
        SQL.AppendLine("        ICTSTYC1.STYLE_COLOR_STATUS")
        SQL.AppendLine("        from ICTSTYC1")
        SQL.AppendLine("        WHERE (ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE) IN")
        SQL.AppendLine("        (")
        SQL.AppendLine("            SELECT")
        SQL.AppendLine("            DISTINCT O2.STYLE_CODE, O2.COLOR_CODE")
        SQL.AppendLine("            FROM SOTORDR1 O1, SOTORDR2 O2")
        SQL.AppendLine("            WHERE O1.ORDR_NO = O2.ORDR_NO")
        SQL.AppendLine($"            AND O1.ORDR_GROUP_NO = '{ORDR_GROUP_NO}'")
        SQL.AppendLine("        )")
        SQL.AppendLine("    )")
        SQL.AppendLine(") X")
        SQL.AppendLine("where ICTCOLR1.COLOR_CODE (+) = X.COLOR_CODE")
        SQL.AppendLine("and ICTSTYL1.STYLE_CODE (+) = X.STYLE_CODE")
        SQL.AppendLine("group by X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC")
        ASCMAIN1.sql = SQL.ToString
        Fill_Records("ICTSTATA",,, SQL.ToString)

        For Each row As DataRow In dst.Tables.Item("ICTSTATA").Select
            row.Item("THIS_PO") = 0
            Dim STYLE_CODE As String = row.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = row.Item("COLOR_CODE").ToString & String.Empty
            Dim FLT As String = $"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'"

            For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select(FLT, "")
                row.Item("THIS_PO") = Val(row.Item("THIS_PO").ToString & String.Empty) + Val(rowSOTORDRS.Item("ORDR_QTY_OPEN").ToString & String.Empty) + Val(rowSOTORDRS.Item("ORDR_QTY_PICK").ToString & String.Empty)
            Next

        Next

    End Sub

    Private Sub UltraOptionSet1_ValueChanged(sender As Object, e As EventArgs) Handles optMULTIPOP.ValueChanged
        If SELECTION_NO = 0 Or ASCMAIN1.CLIENT <> "VAN" Then Exit Sub
        If optMULTIPOP.Value = "M" Then
            grdSOTCARTP.Visible = False
            grdSOTCARTX.Visible = True

        Else
            grdSOTCARTX.Visible = False
            grdSOTCARTP.Visible = True

        End If
    End Sub

    Private Sub btnRefreshOrders_Click(sender As Object, e As EventArgs) Handles btnRefreshOrders.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Calculating all Supply & Demand")

        ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDRS_BASE}")
        ASCMAIN1.sql = Get_SQL_SOTORDRS_BASE()
        ASCDATA1.ExecuteSQL($"Insert into {SOTORDRS_BASE} {ASCMAIN1.sql }")

        Fill_Records("SOTORDRS_ORDERS")
        Sort_grdColumns(grdSOTORDRS_ORDERS, "CUST_CODE, ORDR_NO")

        Fill_Records("SOTORDRS_SC")
        Sort_grdColumns(grdSOTORDRS_sc, "STYLE_CODE, COLOR_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Get_SQL_SOTORDRS_BASE(Optional initialize As Boolean = False) As String

        Dim sql1 As String = "" _
            & "Select SOTORDR2.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC" & vbCrLf _
            & ", SOTORDR2.STYLE_UOM, ICTCOLR1.COLOR_DESC, SOTORDR2.RANGE_STYLE_CODE" & vbCrLf _
            & ", SOTORDR2.CUST_STYLE_CODE, ICTSTYL1.VEND_CODE" & vbCrLf _
            & ", SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" & vbCrLf _
            & ", MAX (SOTORDR2.ORDR_UNIT_PRICE) ORDR_UNIT_PRICE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY, SUM (SOTORDR2.ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (SOTORDR2.ORDR_QTY_ALLO) ORDR_QTY_ALLO" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK, SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC, SUM (SOTORDR2.ORDR_EXTD_COST) ORDR_CGS" & vbCrLf _
            & ", MAX (SOTORDR2.ORDR_RELEASE_AVAIL) ORDR_RELEASE_AVAIL" & vbCrLf _
            & "from SOTORDR2,ICTSTYL1,ICTCOLR1,SOTORDR1" & vbCrLf _
            & "where ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
            & "AND SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
            & "AND SOTORDR1.WHSE_CODE = 'MS'" & vbCrLf _
            & "AND SOTORDR1.ORDR_TYPE_CODE = 'REG'" & vbCrLf _
            & "AND SOTORDR2.ORDR_QTY_OPEN > 0" & vbCrLf _
            & "and ICTSTYL1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf _
            & "and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "group by SOTORDR2.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", SOTORDR2.STYLE_DESC, SOTORDR2.STYLE_UOM, ICTSTYL1.VEND_CODE, ICTCOLR1.COLOR_DESC, SOTORDR2.RANGE_STYLE_CODE" & vbCrLf _
            & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" & vbCrLf

        Dim sql2 As String = "" _
            & "SELECT X.*, SOTORDRS.WIP_IND" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_ON_HAND ONHD_3PL" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_TRAN TRAN_3PL" & vbCrLf _
            & " from SOTORDRS, ICTSTAT2, (" & vbCrLf _
            & sql1 _
            & ") X" & vbCrLf _
            & "where SOTORDRS.ORDR_GROUP_NO (+) = X.ORDR_NO and SOTORDRS.STYLE_CODE (+) = X.STYLE_CODE and SOTORDRS.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
            & "and ICTSTAT2.WHSE_CODE (+) = 'US' and ICTSTAT2.STYLE_CODE (+) = X.STYLE_CODE and ICTSTAT2.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf

        If initialize Then
            sql2 &= " AND ROWNUM < 1"
        End If
        Return sql2

    End Function

    Function Get_SQL_SOTORDRS_ALL() As String

        Dim SQL As String = "" _
            & ", SUM (CASE WHEN ORDR_QTY_ALLO > 0 AND ORDR_RELEASE_AVAIL IS NULL THEN ORDR_QTY_ALLO ELSE 0 END) CUR_QTY_ONHD" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_QTY_ALLO > 0 AND ORDR_RELEASE_AVAIL IS NULL THEN ORDR_QTY_ALLO * ORDR_UNIT_PRICE ELSE 0 END) CUR_AMT_ONHD" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_QTY_ALLO > 0 AND ORDR_RELEASE_AVAIL IS NOT NULL AND WIP_IND = 'P' THEN ORDR_QTY_ALLO ELSE 0 END) FUT_QTY_ONPO" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_QTY_ALLO > 0 AND ORDR_RELEASE_AVAIL IS NOT NULL AND WIP_IND = 'P' THEN ORDR_QTY_ALLO * ORDR_UNIT_PRICE ELSE 0 END) FUT_AMT_ONPO" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_QTY_ALLO > 0 AND ORDR_RELEASE_AVAIL IS NOT NULL AND WIP_IND = 'S' THEN ORDR_QTY_ALLO ELSE 0 END) FUT_QTY_TRAN" & vbCrLf _
            & ", SUM (CASE WHEN ORDR_QTY_ALLO > 0 AND ORDR_RELEASE_AVAIL IS NOT NULL AND WIP_IND = 'S' THEN ORDR_QTY_ALLO * ORDR_UNIT_PRICE ELSE 0 END) FUT_AMT_TRAN" & vbCrLf _
            & ", SUM (CASE WHEN NVL(ORDR_QTY_ALLO,0) = 0 THEN ORDR_QTY_OPEN ELSE 0 END) UNAV_QTY" & vbCrLf _
            & ", SUM (CASE WHEN NVL(ORDR_QTY_ALLO,0) = 0 THEN ORDR_QTY_OPEN * ORDR_UNIT_PRICE ELSE 0 END) UNAV_AMT" & vbCrLf _
            & ", ONHD_3PL" & vbCrLf _
            & ", TRAN_3PL" & vbCrLf _
            & $" from {SOTORDRS_BASE} SOTORDRS_BASE" & vbCrLf
        Return SQL

    End Function

    Function Get_SQL_SOTORDRS_ORDERS() As String
        Dim sql1 As String = "Select ORDR_NO, CUST_CODE, ORDR_CUST_PO, ORDR_DATE, ORDR_SHIP_DATE, ORDR_CANCEL_DATE" & vbCrLf _
        & Replace(Replace(Get_SQL_SOTORDRS_ALL(), ", ONHD_3PL", ", SUM (ONHD_3PL) ONHD_3PL"), ", TRAN_3PL", ", SUM (TRAN_3PL) TRAN_3PL") _
        & " group by ORDR_NO, CUST_CODE, ORDR_CUST_PO, ORDR_DATE, ORDR_SHIP_DATE, ORDR_CANCEL_DATE"

        Dim sql As String = $"Select X.*, ARTCUST1.CUST_NAME from ({sql1}) X, ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = X.CUST_CODE"

        Return sql

    End Function

    Function Get_SQL_SOTORDRS_SC() As String
        Dim sql1 As String = "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
        & Get_SQL_SOTORDRS_ALL() _
        & " group by STYLE_CODE, COLOR_CODE, ONHD_3PL, TRAN_3PL"

        Dim sql As String = $"Select X.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC from ({sql1}) X, ICTSTYL1, ICTCOLR1" & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE and ICTCOLR1.COLOR_CODE = X.COLOR_CODE"

        Return sql

    End Function

    Private Sub grdSOTORDRS_ORDERS_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDRS_ORDERS.AfterRowActivate
        Setup_Synch()
    End Sub

    Sub Setup_Synch()
        Dim dvw As DataView = dst.Tables("SOTORDRS_SC").DefaultView
        If Not chkSynchronize.Checked Then
            dvw.RowFilter = ""
        Else
            If grdSOTORDRS_ORDERS.ActiveRow Is Nothing OrElse Not grdSOTORDRS_ORDERS.ActiveRow.IsDataRow Then
                dvw.RowFilter = ""
            Else
                dvw.RowFilter = ""
                For Each row As DataRow In dst.Tables("SOTORDRS_SC").Select("SEL = '1'")
                    row.Item("SEL") = DBNull.Value
                Next

                Dim ORDR_NO As String = grdSOTORDRS_ORDERS.ActiveRow.Cells("ORDR_NO").Value & ""
                ASCMAIN1.sql = $"Select STYLE_CODE, COLOR_CODE from SOTORDR2 where ORDR_NO = '{ORDR_NO}'"
                For Each row As DataRow In ASCDATA1.GetDataTable().Select("")
                    Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                    Dim rowSC As DataRow = dst.Tables("SOTORDRS_SC").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                    If rowSC IsNot Nothing Then
                        rowSC.Item("SEL") = "1"
                    End If
                Next

                Dim sql As String = "SEL = '1'"
                dvw.RowFilter = sql
            End If
        End If
    End Sub

    Private Sub chkSynchronize_CheckedChanged(sender As Object, e As EventArgs) Handles chkSynchronize.CheckedChanged
        Setup_Synch
    End Sub

    Private Sub btnSaveShort_Click(sender As Object, e As EventArgs) Handles btnSaveShort.Click
        grdSOTORDR0.DisplayLayout.SaveAsXml(SOFCORD1_LAYOUT_SHORT, UltraWinGrid.PropertyCategories.All)
    End Sub
End Class

Public Class MyCustomTooltip
    Implements IRenderLabel

    Public Sub New()

    End Sub 'New

    Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
        'Return Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        'Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        Return Context("SERIES_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))

    End Function 'ToString 
End Class 'MyCustomTooltip
