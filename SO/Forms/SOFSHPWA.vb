Imports System.Text
Imports Infragistics.Win.UltraWinGrid
'Imports Microsoft.Office.Interop.Excel

Public Class SOFSHPWA
    Private CUST_CODE As String = ""
    Private sqlSOTORDRS As String
    Private PODATA_TEMP As String = ""
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFSHPWI" Then
            InquiryMode = True
        End If

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, -12)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -24, 0, 0)
        'Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        Check_Form_Options()
        Dim SQLB As New System.Text.StringBuilder

        With dst
            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("O1.ORDR_YYYYPP_BOOKED,")
            SQLB.AppendLine("O1.CUST_CODE,")
            SQLB.AppendLine("MAX(O1.ORDR_STATUS) AS ORDR_STATUS,")
            SQLB.AppendLine("MAX(O1.ORDR_GROUP_NO) AS ORDR_GROUP_NO,")
            SQLB.AppendLine("COUNT(DISTINCT(O1.ORDR_NO)) AS ORDR_CNT,")
            SQLB.AppendLine("O1.ORDR_CUST_PO,")
            SQLB.AppendLine("O1.ORDR_DATE_RECD,")
            SQLB.AppendLine("O1.ORDR_SHIP_DATE,")
            SQLB.AppendLine("O1.ORDR_CANCEL_DATE,")
            SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY,0)) AS ORDR_QTY,")
            SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY_OPEN,0)) AS ORDR_QTY_OPEN,")
            SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY_PICK,0)) AS ORDR_QTY_PICK,")
            SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY_SHIP,0)) AS ORDR_QTY_SHIP,")
            SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY_CANC,0)) AS ORDR_QTY_CANC,")
            SQLB.AppendLine("I1.INV_DATE,")
            SQLB.AppendLine("I1.INV_NO_CONS,")
            SQLB.AppendLine("I1.INV_TOTAL_AMOUNT,")
            SQLB.AppendLine("C1.CART_CNT,")
            SQLB.AppendLine("C1.CART_TOTAL_UNITS AS QTY_PACKED_TOTAL,")
            SQLB.AppendLine("I1.INV_BALANCE,")
            SQLB.AppendLine("C1.SHIP_VIA_CODE,")
            SQLB.AppendLine("C1.SHIP_REF")
            SQLB.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2,")
            SQLB.AppendLine("(")
            SQLB.AppendLine("   SELECT")
            SQLB.AppendLine("   O1.ORDR_GROUP_NO,")
            SQLB.AppendLine("   COUNT(C1.CART_NO) AS CART_CNT,")
            SQLB.AppendLine("   SUM(NVL(C1.CART_TOTAL_UNITS, 0)) AS CART_TOTAL_UNITS,")
            SQLB.AppendLine("   MAX(S1.SHIP_VIA_CODE) AS SHIP_VIA_CODE,")
            SQLB.AppendLine("   MAX(S1.SHIP_REF) AS SHIP_REF")
            SQLB.AppendLine("   FROM SOTORDR1 O1, SOTPICK1 P1, SOTCART1 C1, SOTSHIP1 S1")
            SQLB.AppendLine("   WHERE O1.ORDR_NO = P1.ORDR_NO")
            SQLB.AppendLine("   AND P1.PICK_NO = C1.PICK_NO (+)")
            SQLB.AppendLine("   AND P1.SHIP_BOL_NO = S1.SHIP_BOL_NO (+)")
            SQLB.AppendLine("   AND P1.PICK_STATUS <> 'D'")
            SQLB.AppendLine("   GROUP BY O1.ORDR_GROUP_NO")
            SQLB.AppendLine(") C1,")
            SQLB.AppendLine("(")
            SQLB.AppendLine("  SELECT")
            SQLB.AppendLine("  O1.ORDR_GROUP_NO,")
            SQLB.AppendLine("  MAX(I1.INV_DATE) AS INV_DATE,")
            SQLB.AppendLine("  MAX(NVL(I1.INV_NO_CONS, I1.INV_NO || ':I')) AS INV_NO_CONS,")
            SQLB.AppendLine("  NVL(SUM(NVL(I1.INV_TOTAL_AMOUNT, 0)), 0) AS INV_TOTAL_AMOUNT,")
            SQLB.AppendLine("  NVL(SUM(NVL(A1.INV_BALANCE, 0)), 0) AS INV_BALANCE")
            SQLB.AppendLine("  FROM SOTINVH1 I1, SOTORDR1 O1, ARTOPEN1 A1")
            SQLB.AppendLine("  WHERE I1.ORDR_NO = O1.ORDR_NO")
            SQLB.AppendLine("  And I1.INV_NO = A1.INV_NUM (+)")
            SQLB.AppendLine("  GROUP BY O1.ORDR_GROUP_NO")
            SQLB.AppendLine(") I1")
            SQLB.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
            SQLB.AppendLine("AND O1.ORDR_GROUP_NO = C1.ORDR_GROUP_NO (+)")
            SQLB.AppendLine("AND O1.ORDR_GROUP_NO = I1.ORDR_GROUP_NO (+)")
            SQLB.AppendLine("AND O1.CUST_CODE IN ('WALMART','WALMARTCOM')")
            SQLB.AppendLine("AND O1.ORDR_YYYYPP_BOOKED >= '201702'")
            SQLB.AppendLine("AND ROWNUM <= 0")
            SQLB.AppendLine("GROUP BY")
            SQLB.AppendLine("O1.ORDR_YYYYPP_BOOKED,")
            SQLB.AppendLine("O1.CUST_CODE,")
            SQLB.AppendLine("O1.ORDR_CUST_PO,")
            SQLB.AppendLine("O1.ORDR_SHIP_DATE,")
            SQLB.AppendLine("O1.ORDR_CANCEL_DATE,")
            SQLB.AppendLine("O1.ORDR_DATE_RECD,")
            SQLB.AppendLine("I1.INV_DATE,")
            SQLB.AppendLine("I1.INV_NO_CONS,")
            SQLB.AppendLine("I1.INV_TOTAL_AMOUNT,")
            SQLB.AppendLine("I1.INV_BALANCE,")
            SQLB.AppendLine("C1.CART_CNT,")
            SQLB.AppendLine("C1.CART_TOTAL_UNITS,")
            SQLB.AppendLine("C1.SHIP_VIA_CODE,")
            SQLB.AppendLine("C1.SHIP_REF")
            ASCMAIN1.sql = SQLB.ToString
            Dim TABLES As String() = {"SOTSHPWA", "SOTSHPWH"}
            For Each TABLE As String In TABLES
                Create_TDA(.Tables.Add, TABLE, "**", 0, False)
                With .Tables(TABLE)
                    .Columns.Add("TANNER")
                    .Columns.Add("VEND_NO")
                    .Columns.Add("VEND_NO_DEPT")
                    .Columns.Add("VEND_SEQ_NO", GetType(System.Decimal))
                    .Columns.Add("WHSE_NO")
                    .Columns.Add("PO_TYPE")
                    .Columns.Add("PO_EVENT")
                    .Columns.Add("PO_DATE_CREATED", GetType(System.DateTime))
                    .Columns.Add("PO_SHIP_DATE", GetType(System.DateTime))
                    .Columns.Add("PO_CANCEL_DATE", GetType(System.DateTime))
                    .Columns.Add("ORIG_MABD_DATE", GetType(System.DateTime))
                    .Columns.Add("MABD_DATE", GetType(System.DateTime))
                    .Columns.Add("MABD_COMPLIANCE_DATE", GetType(System.DateTime))
                    .Columns.Add("WHSE_QTY_ORDR", GetType(System.Decimal))
                    .Columns.Add("WHSE_QTY_REC", GetType(System.Decimal))
                    .Columns.Add("PCT_REC", GetType(System.Decimal))
                    .Columns.Add("TOT_ORDR_COST", GetType(System.Decimal))
                    .Columns.Add("TOT_REC_COST", GetType(System.Decimal))
                    .Columns("WHSE_QTY_ORDR").DefaultValue = 0
                    .Columns("WHSE_QTY_REC").DefaultValue = 0
                    .Columns("PCT_REC").DefaultValue = 0
                    .Columns("TOT_ORDR_COST").DefaultValue = 0
                    .Columns("TOT_REC_COST").DefaultValue = 0
                    .Columns.Add("VARIANCE", GetType(System.Decimal), "ORDR_QTY_SHIP - WHSE_QTY_REC")
                    .Columns.Add("VARIANCE_COST", GetType(System.Decimal), "INV_TOTAL_AMOUNT - TOT_REC_COST")
                End With
            Next

            SQLB.Length = 0
            SQLB.AppendLine("Select")
            SQLB.AppendLine("SOTORDR1.ORDR_GROUP_NO")
            SQLB.AppendLine("FROM SOTORDR1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTGROUP", "**", 0, False)

            SQLB.Length = 0
            SQLB.AppendLine("Select")
            SQLB.AppendLine("*")
            SQLB.AppendLine("FROM SOTWMTD1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTWMTD1", "**", 0, True)

            SQLB.Length = 0
            SQLB.AppendLine("Select")
            SQLB.AppendLine("*")
            SQLB.AppendLine("FROM SOTWMPO1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTWMPO1", "**", 0, True)

            SQLB.Length = 0
            SQLB.AppendLine("Select")
            SQLB.AppendLine("*")
            SQLB.AppendLine("FROM SOTWMPOH")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTWMPOH", "**", 0, False)

            SQLB.Length = 0
            SQLB.AppendLine("Select ORDR_NO, ORDR_CUST_PO, ORDR_DATE")
            SQLB.AppendLine(", ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_ORIG_SHIP_DATE, ORDR_ORIG_CANCEL_DATE")
            SQLB.AppendLine(", CUST_STORE_NO, SALES_DIVISION_CODE, ORDR_SOURCE, ORDR_DEPT, ORDR_ADDR_TYPE_ST")
            SQLB.AppendLine(", ORDR_STATUS, CUST_STORE_NAME, SREP_CODE, ORDR_PRIORITY, ORDR_HOLD")
            SQLB.AppendLine(", ORDR_REL_HOLD_CODES, CUST_DC_NO, ORDR_PRE_ALLOC, WHSE_CODE, EDI_DOC_SEQ_NO, REASON_CODE")
            SQLB.AppendLine(", 'O' ORDR_TYPE, ORDR_GROUP_NO, X.ORDR_QTY, X.ORDR_QTY_OPEN, X.ORDR_QTY_PICK, X.ORDR_QTY_SHIP, X.ORDR_QTY_CANC, X.ORDR_AMT, X.ORDR_AMT_OPEN, X.ORDR_AMT_PICK, X.ORDR_AMT_SHIP, X.ORDR_AMT_CANC")
            SQLB.AppendLine(" from SOTORDR1, (Select ORDR_NO ORDR_NO_DTL, Sum (ORDR_QTY) ORDR_QTY, Sum (ORDR_QTY_OPEN) ORDR_QTY_OPEN, Sum (ORDR_QTY_PICK) ORDR_QTY_PICK, Sum (ORDR_QTY_SHIP) ORDR_QTY_SHIP, Sum (ORDR_QTY_CANC) ORDR_QTY_CANC, Sum (ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT, Sum (ORDR_QTY_OPEN * ORDR_UNIT_PRICE) ORDR_AMT_OPEN, Sum (ORDR_QTY_PICK * ORDR_UNIT_PRICE) ORDR_AMT_PICK, Sum (ORDR_QTY_SHIP * ORDR_UNIT_PRICE) ORDR_AMT_SHIP, Sum (ORDR_QTY_CANC * ORDR_UNIT_PRICE) ORDR_AMT_CANC from SOTORDR2 group by ORDR_NO) X where X.ORDR_NO_DTL = SOTORDR1.ORDR_NO")
            SQLB.AppendLine(" and ORDR_GROUP_NO = :PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V", 1)
            With .Tables("SOTORDR1")
                .Columns.Add("TOTAL_SHIPPED", GetType(System.Int64))
                .Columns.Add("WMT_REC", GetType(System.Int64))
                .Columns.Add("VARIANCE", GetType(System.Decimal), "TOTAL_SHIPPED - WMT_REC")
            End With

            SQLB.Length = 0
            SQLB.AppendLine("Select SOTPICK1.PICK_NO, SOTORDR1.CUST_STORE_NO, SOTPICK1.ORDR_NO")
            SQLB.AppendLine(", SOTPICK1.PICK_STATUS, SOTPICK1.PICK_RELEASED, SOTPICK1.PICK_FREIGHT")
            SQLB.AppendLine(", SOTPICK1.PICK_PICKER, SOTPICK1.PICK_NO_REV")
            SQLB.AppendLine(", SOTPICK1.PICK_PRINTED, SOTPICK1.PICK_PACKED, SOTPICK1.PICK_SHIPPED")
            SQLB.AppendLine(", SOTPICK1.PICK_BATCH_NO, SOTPICK1.SHIP_BOL_NO, SOTPICK1.INV_NO")
            SQLB.AppendLine(", SOTPICK1.PICK_CNT_CARTONS, SOTPICK1.PICK_TOTAL_WGT")
            SQLB.AppendLine(", SOTPICK1.INIT_OPER, SOTPICK1.LAST_OPER, SOTPICK1.INIT_DATE, SOTPICK1.LAST_DATE")
            SQLB.AppendLine(", SOTPICK0.PICK_FORCED")
            SQLB.AppendLine(" from SOTPICK1, SOTORDR1, SOTPICK0")
            SQLB.AppendLine(" where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO")
            SQLB.AppendLine("   And SOTPICK0.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO")
            SQLB.AppendLine("   And SOTORDR1.ORDR_GROUP_NO = : PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)
            With .Tables("SOTPICK1")
                .Columns.Add("WMT_REC", GetType(System.Int64))
                .Columns.Add("VARIANCE", GetType(System.Decimal), "PICK_SHIPPED - WMT_REC")
            End With

            SQLB.Length = 0
            SQLB.AppendLine("Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE")
            SQLB.AppendLine(", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC")
            SQLB.AppendLine(", SOTPICK1.SHIP_BOL_NO, EDT850T2.EDI_COLOR_CODE")
            SQLB.AppendLine(", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU")
            SQLB.AppendLine(" from SOTPICK1,SOTPICK2,SOTORDR2,ICTCOLR1,EDT850T2")
            SQLB.AppendLine(" where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO")
            SQLB.AppendLine("   And SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO")
            SQLB.AppendLine("   And EDT850T2.EDI_DOC_SEQ_NO (+) = SOTORDR2.EDI_DOC_SEQ_NO")
            SQLB.AppendLine("   And EDT850T2.EDI_DTL_SEQ (+) = SOTORDR2.EDI_DTL_SEQ")
            SQLB.AppendLine("   And ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE")
            SQLB.AppendLine("   And SOTPICK1.PICK_NO = SOTPICK2.PICK_NO")
            SQLB.AppendLine("   And SOTPICK2.PICK_NO = :PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)
            With .Tables("SOTPICK2")
                .Columns.Add("WMT_REC", GetType(System.Int64))
                .Columns.Add("VARIANCE", GetType(System.Decimal), "PICK_QTY_CONF - WMT_REC")
                .Columns.Add("FIRST_SHIP", GetType(System.Int64))
                .Columns.Add("LAST_SHIP", GetType(System.Int64))
            End With

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")

            SQLB.Length = 0
            SQLB.AppendLine("Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE")
            SQLB.AppendLine(", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC, SOTORDR2.RANGE_STYLE_CODE")
            SQLB.AppendLine(", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU")
            SQLB.AppendLine(", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY")
            SQLB.AppendLine(", SUM (SOTORDR2.ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT")
            SQLB.AppendLine(", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN")
            SQLB.AppendLine(", SUM (SOTORDR2.ORDR_QTY_ALLO) ORDR_QTY_ALLO")
            SQLB.AppendLine(", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK")
            SQLB.AppendLine(", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP")
            SQLB.AppendLine(", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC")
            SQLB.AppendLine(", SUM (SOTORDR2.ORDR_EXTD_COST) ORDR_CGS")
            SQLB.AppendLine(", MAX (SOTORDR2.ORDR_RELEASE_AVAIL) ORDR_RELEASE_AVAIL")
            SQLB.AppendLine(" from SOTORDR2,ICTCOLR1,SOTORDR1")
            SQLB.AppendLine(" where ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE")
            SQLB.AppendLine(" And SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO")
            SQLB.AppendLine(" group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE")
            SQLB.AppendLine(", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC, SOTORDR2.RANGE_STYLE_CODE")
            SQLB.AppendLine(", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU")
            sqlSOTORDRS = SQLB.ToString
            ASCMAIN1.sql = Replace(sqlSOTORDRS, " group by ", " And ROWNUM < 1 group by ")
            Create_TDA(.Tables.Add, "SOTORDRS", "**", 0, False, "V", 0)
            .Tables("SOTORDRS").Columns.Add("ORDR_UNIT_PRICE", GetType(System.Decimal), "IIF(ISNULL(ORDR_QTY,0)=0,0,ISNULL(ORDR_AMT,0) / ISNULL(ORDR_QTY,0))")
            .Tables("SOTORDRS").Columns.Add("ORDR_GP", GetType(System.Decimal), "ISNULL(ORDR_AMT,0)-ISNULL(ORDR_CGS,0)")
            .Tables("SOTORDRS").Columns.Add("ORDR_GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ORDR_AMT,0)=0,0,100*ORDR_GP/ISNULL(ORDR_AMT,0))")
            .Tables("SOTORDRS").Columns.Add("WMT_REC", GetType(System.Int64))
            .Tables("SOTORDRS").Columns.Add("VARIANCE", GetType(System.Int64), "ORDR_QTY_SHIP - WMT_REC")

            SQLB.Length = 0
            SQLB.AppendLine("Select SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_DATE_SHIPPED, SOTSHIP1.SHIP_VIA_CODE, SOTSHIP1.SHIP_REF")
            SQLB.AppendLine(", SOTSHIP1.SHIP_TOTAL_WGT, SOTSHIP1.SHIP_CNT_CARTONS, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE")
            SQLB.AppendLine(", SOTSHIP1.SHIP_PICK_PRINTED, SOTSHIP1.PICK_BATCH_NO, SOTSHIP1.SHIP_STATUS, SOTSHIP1.LP_STATUS")
            SQLB.AppendLine(", SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.FRT_TERMS, SOTSHIP1.SHIP_PULL_BY_STYLE")
            SQLB.AppendLine(", SOTSHIP1.SHIP_856_BATCH_NO, SOTSHIP1.SHIP_810_BATCH_NO, SOTSHIP1.WHSE_CODE, SOTSHIP1.INV_DATE, SOTSHIP1.SHIP_MANIFEST_NO")
            SQLB.AppendLine(", SOTSHIP1.SHIP_BOL_NO_REV, SOTSHIP1.SHIP_NOTES, SOTSHIP1.SHIPPED_ACTUAL, SOTSHIP1.SHIP_SEAL_NO")
            SQLB.AppendLine(", SOTSHIP1.SHIP_BOL_NO_ORIG, SOTSHIP1.SHIP_BOL_NO_SPLIT, SOTSHIP1.BOL_PRINTED, SOTSHIP1.SHIP_SPEC_INST")
            SQLB.AppendLine(", SOTSHIP1.MASTER_SHIP_BOL_NO, SOTSHIP1.SHIP_940_BATCH_NO, SOTSHIP1.SHIP_753_IND, SOTSHIP1.SHIP_DATE_PACKED")
            SQLB.AppendLine(", SOTSHIP1.INIT_DATE, SOTSHIP1.INIT_OPER, SOTSHIP1.SHIP_LOAD_NO, SOTSHIP1.SHIP_APPT_NO, SOTSHIP1.SHIP_WAVE_STATUS")
            SQLB.AppendLine(" from SOTSHIP1")
            SQLB.AppendLine(" where ORDR_GROUP_NO = : PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, False, "V", 1)

            SQLB.Length = 0
            SQLB.AppendLine("Select SOTPICK1.SHIP_BOL_NO")
            SQLB.AppendLine(", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE")
            SQLB.AppendLine(", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC")
            SQLB.AppendLine(", EDT850T2.EDI_COLOR_CODE")
            SQLB.AppendLine(", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU")
            SQLB.AppendLine(", SUM (SOTPICK2.PICK_QTY) PICK_QTY")
            SQLB.AppendLine(", SUM (SOTPICK2.PICK_QTY * SOTPICK2.PICK_UNIT_PRICE) PICK_AMT")
            SQLB.AppendLine(", SUM (SOTPICK2.PICK_QTY_CONF) PICK_QTY_CONF")
            SQLB.AppendLine(", SUM (SOTPICK2.PICK_QTY_CANC) PICK_QTY_CANC")
            SQLB.AppendLine(", SUM (SOTPICK2.PICK_QTY_BACK) PICK_QTY_BACK")
            SQLB.AppendLine(", SUM (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL")
            SQLB.AppendLine(", SUM (SOTPICK2.PICK_QTY_BACK_REL) PICK_QTY_BACK_REL")
            SQLB.AppendLine(" from SOTPICK1,SOTPICK2,SOTORDR2,ICTCOLR1,EDT850T2")
            SQLB.AppendLine(" where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO")
            SQLB.AppendLine("   And SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO")
            SQLB.AppendLine("   And EDT850T2.EDI_DOC_SEQ_NO (+) = SOTORDR2.EDI_DOC_SEQ_NO")
            SQLB.AppendLine("   And EDT850T2.EDI_DTL_SEQ (+) = SOTORDR2.EDI_DTL_SEQ")
            SQLB.AppendLine("   And ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE")
            SQLB.AppendLine("   And SOTPICK1.PICK_NO = SOTPICK2.PICK_NO")
            SQLB.AppendLine("   And SOTPICK1.SHIP_BOL_NO = :PARM1")
            SQLB.AppendLine(" group by SOTPICK1.SHIP_BOL_NO")
            SQLB.AppendLine(", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE")
            SQLB.AppendLine(", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC")
            SQLB.AppendLine(", EDT850T2.EDI_COLOR_CODE")
            SQLB.AppendLine(", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTSHIP2", "**", 0, False, "V", 0)
            With .Tables("SOTSHIP2")
                .Columns.Add("PICK_UNIT_PRICE", GetType(System.Decimal), "IIF(PICK_QTY=0,0,PICK_AMT/PICK_QTY)")
                .Columns.Add("WMT_REC", GetType(System.Int64))
                .Columns.Add("VARIANCE", GetType(System.Int64), "PICK_QTY_CONF - WMT_REC")
                .Columns.Add("FIRST_SHIP", GetType(System.DateTime))
                .Columns.Add("LAST_SHIP", GetType(System.DateTime))
                .Columns("PICK_QTY").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CONF").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CANC").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_BACK").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_CANC_REL").DataType = GetType(System.Int64)
                .Columns("PICK_QTY_BACK_REL").DataType = GetType(System.Int64)
            End With

            Create_Relation("SOTSHIP1", "SOTSHIP2", "SHIP_BOL_NO")

            SQLB.Length = 0
            SQLB.AppendLine("Select SOTCART1.*, SOTPICK1.SHIP_BOL_NO, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE, SOTORDR1.CUST_STORE_NO")
            SQLB.AppendLine(" from SOTCART1, SOTPICK1, SOTSHIP1, SOTORDR1")
            SQLB.AppendLine(" where SOTPICK1.PICK_NO = SOTCART1.PICK_NO")
            SQLB.AppendLine("   And SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO")
            SQLB.AppendLine("   And SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO")
            SQLB.AppendLine("   And SOTSHIP1.ORDR_GROUP_NO = : PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

            SQLB.Length = 0
            SQLB.AppendLine("Select SOTCART2.*")
            SQLB.AppendLine(" from SOTCART2, SOTCART1, SOTPICK1, SOTSHIP1")
            SQLB.AppendLine(" where SOTCART1.CART_NO = SOTCART2.CART_NO")
            SQLB.AppendLine("   And SOTPICK1.PICK_NO = SOTCART1.PICK_NO")
            SQLB.AppendLine("   And SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO")
            SQLB.AppendLine("   And SOTSHIP1.ORDR_GROUP_NO = : PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "V", 2)

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")

            SQLB.Length = 0
            SQLB.AppendLine("SELECT *")
            SQLB.AppendLine("FROM SOTWMTD1")
            SQLB.AppendLine("WHERE PO_NUMBER")
            SQLB.AppendLine("NOT IN")
            SQLB.AppendLine("(")
            SQLB.AppendLine("  SELECT DISTINCT ORDR_CUST_PO")
            SQLB.AppendLine("  FROM SOTORDR1")
            SQLB.AppendLine("  WHERE CUST_CODE IN ('WALMART','WALMARTCOM')")
            SQLB.AppendLine(")")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add("SOTNOMCH"), "SOTWMTD1", "**", 0, False)

            SQLB.Length = 0
            SQLB.AppendLine("SELECT BOL_NO")
            SQLB.AppendLine("FROM SOTSHIPB")
            SQLB.AppendLine("WHERE SHIP_TO_ZIP_CODE = '18372'")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add(), "SOTTANNR", "**", 0, False)

            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("I1.ORDR_CUST_PO,")
            SQLB.AppendLine("I1.CUST_CODE,")
            SQLB.AppendLine("I1.CUST_STORE_NO,")
            SQLB.AppendLine("I2.STYLE_CODE,")
            SQLB.AppendLine("I2.COLOR_CODE,")
            SQLB.AppendLine("SUM(ORDR_QTY_SHIP) AS ORDR_QTY_SHIP")
            SQLB.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
            SQLB.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
            SQLB.AppendLine("AND I1.INV_NO = I2.INV_NO")
            SQLB.AppendLine("AND I1.INV_TYPE = 'I'")
            SQLB.AppendLine("AND ROWNUM < 0")
            SQLB.AppendLine("GROUP BY")
            SQLB.AppendLine("I1.ORDR_CUST_PO,")
            SQLB.AppendLine("I1.CUST_CODE,")
            SQLB.AppendLine("I1.CUST_STORE_NO,")
            SQLB.AppendLine("I2.STYLE_CODE,")
            SQLB.AppendLine("I2.COLOR_CODE")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add(), "PODATA", "**", 0, False)
            With .Tables("PODATA")
                .Columns.Add("UPC")
                .Columns.Add("CUST_SKU")
                .Columns.Add("FST_DT_SHIPPED", GetType(System.DateTime))
                .Columns.Add("FST_PO_SHIPPED")
                .Columns.Add("FST_INV_SHIPPED")
                .Columns.Add("LAST_DT_SHIPPED", GetType(System.DateTime))
                .Columns.Add("LAST_PO_SHIPPED")
                .Columns.Add("LAST_INV_SHIPPED")
                .Columns.Add("TOTAL_SHIPPED", GetType(System.Int64))
            End With

            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("O1.PO_NUMBER,")
            SQLB.AppendLine("O1.CUST_STORE_NO,")
            SQLB.AppendLine("O1.PO_STATUS,")
            SQLB.AppendLine("O1.ST_EA_ORDR,")
            SQLB.AppendLine("O1.ST_EACH_RCD,")
            SQLB.AppendLine("O2.ORDR_QTY_SHIP,")
            SQLB.AppendLine("O1.EXCEL_FILE,")
            SQLB.AppendLine("O1.EXCEL_LINE")
            SQLB.AppendLine("FROM SOTWMPO1 O1, SOTWMPO2 O2")
            SQLB.AppendLine("WHERE O1.PO_NUMBER = O2.ORDR_CUST_PO")
            SQLB.AppendLine("AND O1.CUST_STORE_NO = O2.CUST_STORE_NO")
            SQLB.AppendLine("AND O1.PO_NUMBER = : PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "STOREPO1", "**", 0, False, "V", 2)
            With .Tables("STOREPO1")
                .Columns.Add("VARIANCE", GetType(System.Decimal), "ORDR_QTY_SHIP - ST_EACH_RCD")
            End With

            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("O1.PO_NUMBER,")
            SQLB.AppendLine("O1.CUST_STORE_NO,")
            SQLB.AppendLine("O1.PO_STATUS,")
            SQLB.AppendLine("O1.ST_EA_ORDR,")
            SQLB.AppendLine("O1.ST_EACH_RCD,")
            SQLB.AppendLine("O2.ORDR_QTY_SHIP,")
            SQLB.AppendLine("(O2.ORDR_QTY_SHIP - O1.ST_EACH_RCD) AS VARIANCE,")
            SQLB.AppendLine("O1.EXCEL_FILE,")
            SQLB.AppendLine("O1.EXCEL_LINE")
            SQLB.AppendLine("FROM SOTWMPO1 O1, SOTWMPO2 O2")
            SQLB.AppendLine("WHERE O1.PO_NUMBER = O2.ORDR_CUST_PO")
            SQLB.AppendLine("AND O1.CUST_STORE_NO = O2.CUST_STORE_NO")
            SQLB.AppendLine("AND (O2.ORDR_QTY_SHIP - O1.ST_EACH_RCD) <> 0")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "STOREPOV", "**", 0, False)

            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("NVL(O1.PO_NUMBER,'EMPTY') AS PO_NUMBER,")
            SQLB.AppendLine("NVL(O1.PO_STATUS,'N') AS PO_STATUS,")
            SQLB.AppendLine("SUM(O1.ST_EA_ORDR) AS ST_EA_ORDR,")
            SQLB.AppendLine("SUM(O1.ST_EACH_RCD) AS ST_EACH_RCD,")
            SQLB.AppendLine("SUM(O2.ORDR_QTY_SHIP) AS ORDR_QTY_SHIP")
            SQLB.AppendLine("FROM SOTWMPO1 O1, SOTWMPO2 O2")
            SQLB.AppendLine("WHERE O1.PO_NUMBER = O2.ORDR_CUST_PO")
            SQLB.AppendLine("AND O1.CUST_STORE_NO = O2.CUST_STORE_NO")
            SQLB.AppendLine("GROUP BY")
            SQLB.AppendLine("NVL(O1.PO_NUMBER,'EMPTY'),")
            SQLB.AppendLine("NVL(O1.PO_STATUS,'N')")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "STOREPOS", "**", 0, False, "V", 2)
            With .Tables("STOREPOS")
                .Columns.Add("VARIANCE", GetType(System.Decimal), "ORDR_QTY_SHIP - ST_EACH_RCD")
            End With
        End With

        grdSOTSHPWA.DataSource = dst.Tables("SOTSHPWA")
        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTORDRS.DataSource = dst.Tables("SOTORDRS")
        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTNOMCH.DataSource = dst.Tables("SOTNOMCH")
        grdPOSTYLES.DataSource = dst.Tables("PODATA")
        grdSTOREPO1.DataSource = dst.Tables("STOREPO1")
        grdSTOREPOS.DataSource = dst.Tables("STOREPOS")
        grdSTOREPOV.DataSource = dst.Tables("STOREPOV")

        Sort_grdColumns(grdSOTSHPWA, "ORDR_YYYYPP_BOOKED, ORDR_GROUP_NO", False)

        Sort_grdColumns(grdSTOREPO1, "PO_NUMBER, CUST_STORE_NO", False)
        Sort_grdColumns(grdSTOREPOS, "PO_NUMBER", False)
        Sort_grdColumns(grdSTOREPOV, "PO_NUMBER, CUST_STORE_NO", False)

        Create_Summary(grdSTOREPO1, "CUST_STORE_NO", "Count")
        Create_Summary(grdSTOREPOS, "PO_NUMBER", "Count")
        Create_Summary(grdSTOREPOV, "CUST_STORE_NO", "Count")

        'grdGROUPS.DataSource = dst.Tables("SOTGROUP")

        TABLE_NAME = "SOTSHPWA"

        EntryMode = "E"

        With grdSTOREPO1.DisplayLayout.Bands(0)
            .Columns("ST_EA_ORDR").Format = "#,###,##0"
            .Columns("ST_EACH_RCD").Format = "#,###,##0"
            .Columns("ORDR_QTY_SHIP").Format = "#,###,##0"
            .Columns("VARIANCE").Format = "#,###,##0"
            .Columns("EXCEL_LINE").Format = "######0"
        End With

        With grdSTOREPOV.DisplayLayout.Bands(0)
            .Columns("ST_EA_ORDR").Format = "#,###,##0"
            .Columns("ST_EACH_RCD").Format = "#,###,##0"
            .Columns("ORDR_QTY_SHIP").Format = "#,###,##0"
            .Columns("VARIANCE").Format = "#,###,##0"
            .Columns("EXCEL_LINE").Format = "######0"
        End With

        With grdSTOREPOS.DisplayLayout.Bands(0)
            .Columns("ST_EA_ORDR").Format = "#,###,##0"
            .Columns("ST_EACH_RCD").Format = "#,###,##0"
            .Columns("ORDR_QTY_SHIP").Format = "#,###,##0"
            .Columns("VARIANCE").Format = "#,###,##0"
        End With

        'Call Load_Record()
        grdPOSTYLES.DisplayLayout.Bands(0).Columns("ORDR_QTY_SHIP").Format = "#,###,##0"
        grdPOSTYLES.DisplayLayout.Bands(0).Columns("TOTAL_SHIPPED").Format = "#,###,##0"
        With grdSOTSHPWA.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_CNT", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "CART_CNT", "QTY_PACKED_TOTAL"}
                .Columns(COLUMN_NAME).Format = "#,###,##0"
            Next
            For Each COLUMN_NAME As String In New String() {"TANNER", "ORDR_YYYYPP_BOOKED", "CUST_CODE", "SHIP_VIA_CODE", "SHIP_REF", "ORDR_GROUP_NO", "ORDR_CNT", "ORDR_CUST_PO", "ORDR_DATE_RECD", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "INV_DATE", "INV_NO_CONS", "INV_TOTAL_AMOUNT", "INV_BALANCE", "CART_CNT", "QTY_PACKED_TOTAL", "ORDR_STATUS"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"VEND_NO", "VEND_NO_DEPT", "VEND_SEQ_NO", "WHSE_NO", "PO_TYPE", "PO_EVENT", "PO_DATE_CREATED", "PO_SHIP_DATE", "PO_CANCEL_DATE", "ORIG_MABD_DATE", "MABD_DATE", "MABD_COMPLIANCE_DATE", "WHSE_QTY_ORDR", "WHSE_QTY_REC", "PCT_REC", "TOT_ORDR_COST", "TOT_REC_COST"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.Blue
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"PO_SHIP_DATE", "PO_CANCEL_DATE", "ORIG_MABD_DATE", "MABD_DATE", "MABD_COMPLIANCE_DATE"}
                .Columns(COLUMN_NAME).Format = "MM/dd/yyyy"
            Next
            For Each COLUMN_NAME As String In New String() {"VEND_SEQ_NO", "WHSE_QTY_ORDR", "WHSE_QTY_REC", "VARIANCE"}
                .Columns(COLUMN_NAME).Format = "#,###,##0"
            Next
            For Each COLUMN_NAME As String In New String() {"TOT_ORDR_COST", "TOT_REC_COST", "VARIANCE_COST"}
                .Columns(COLUMN_NAME).Format = "#,###,##0.00"
            Next
            For Each COLUMN_NAME As String In New String() {"PCT_REC"}
                .Columns(COLUMN_NAME).Format = "#,##0.00%"
            Next
            For Each COLUMN_NAME As String In New String() {"VARIANCE", "VARIANCE_COST"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = System.Drawing.Color.Red
                End With
            Next
        End With

        Bind_Controls(grpHeader, "SOTSHPWH")
        Bind_Controls(grpVariance, "SOTSHPWH")

        'ASCMAIN1.Add_Value_List(grdSOFSHPWA, "ORDR_STATUS", , New String() {":", "C:Cancelled", "D:Deleted", "F:Final", "O:Open", "P:In Pick"})
        ASCMAIN1.Add_Value_List(grdSOTSHPWA, "ORDR_STATUS", , New String() {":", "C:Cancelled", "D:Deleted"})

        Call Mode_Settings(True)

        'SplitContainer2.SplitterDistance = 120

        CUST_CODE = "WALMART"
        Absx1.txtFor("CUST_CODE").Text = CUST_CODE

        Fill_Records("STOREPOS")
        Fill_Records("STOREPOV")

    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Sub Check_Form_Options()
        'With UltraExplorerBar1.Groups("Screen Control")
        '    .Items("New").Visible = (Me.Name = "PMFVIST1")
        'End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Update"
            Case "Import PO Data"
                Dim CC As String = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CUST_CODE.ToLower())
                Dim iResult As MsgBoxResult
                Dim iTitle As String = String.Format("Refresh {0} Data", CC)
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                Dim FileType As String = "An Excel"
                If chkImportCSV.Checked Then
                    FileType = "A CSV"
                End If
                iMSG.AppendLine(String.Format("This Will Ask For {0} File", FileType))
                iMSG.AppendLine(String.Format("To Refresh {0} Shipment Data.", CC))
                iMSG.AppendLine("")
                iMSG.AppendLine("Are You Ready?")
                iMSG.AppendLine("")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg = vbCrLf & "Import Aborted By User"
                End If
            Case "Import Store Data"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Refresh Data"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Ask For A CSV File")
                iMSG.AppendLine("To Refresh PO Store Data.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Are You Ready?")
                iMSG.AppendLine("")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg = vbCrLf & "Import Aborted By User"
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
            Case "Done"
                Call Mode_Settings(False)
                'Me.Close()
            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)
            Case "Import PO Data"
                ImportWalmartFile()
            Case "Import Store Data"
                ImportPOStoreFileCSV()
        End Select

    End Sub

    Private Sub ImportPOStoreFileCSV()
        Dim fDialog As New OpenFileDialog
        Dim FullFileName As String = ""
        Dim FileName As String = ""

        fDialog.Filter = "CSV Files|*.csv"
        fDialog.Title = "Select a CSV File To Import"
        If fDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            FullFileName = fDialog.FileNames(0)
            FileName = fDialog.SafeFileName

            Me.Cursor = Cursors.WaitCursor

            Dim RecsUpdated As Int64 = 0
            Dim RecsAdded As Int64 = 0
            Dim RecWarnings As Int64 = 0
            Dim FoundStartRow As Boolean = False
            Dim FoundStartPO As Boolean = False
            Dim FoundLastRow As Boolean = False
            Fill_Records("SOTWMPO1")
            Fill_Records("SOTWMPOH")
            Dim curRow As Int64 = 0
            Using MyReader As New FileIO.TextFieldParser(FullFileName)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                While Not MyReader.EndOfData
                    curRow += 1
                    If FoundStartRow = False And curRow > 1000 Then
                        MsgBox("Can Not Find Starting Row!", vbOKOnly, "Problem With File")
                        Exit While
                    End If
                    If (curRow Mod 10) = 0 Then
                        ASCMAIN1.Progress("Processing Row", curRow)
                    End If
                    Try
                        currentRow = MyReader.ReadFields()
                        If FoundStartRow Then
                            Dim PO_NUMBER As String = formatPO_NUMBER(currentRow(0).ToString & String.Empty)
                            Dim CUST_STORE_NO As String = (currentRow(13).ToString & String.Empty).ToString.PadLeft(6, "0")
                            If PO_NUMBER <> "" And CUST_STORE_NO <> "000000" Then
                                Dim filter As String = String.Format("PO_NUMBER = '{0}' AND CUST_STORE_NO = '{1}'", PO_NUMBER, CUST_STORE_NO)
                                Dim isAdding As Boolean = False
                                Dim rowSOTWMPO1 As DataRow = dst.Tables.Item("SOTWMPO1").Select(filter).FirstOrDefault
                                If IsNothing(rowSOTWMPO1) Then
                                    rowSOTWMPO1 = dst.Tables.Item("SOTWMPO1").NewRow
                                    isAdding = True
                                End If
                                'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                                '    If curRow = 27123 Then Stop
                                'End If

                                SetSOTWMPO1Data_CSV(rowSOTWMPO1, currentRow)
                                rowSOTWMPO1.Item("EXCEL_FILE") = FileName
                                rowSOTWMPO1.Item("EXCEL_LINE") = curRow
                                If isAdding Then
                                    dst.Tables.Item("SOTWMPO1").Rows.Add(rowSOTWMPO1)
                                    RecsAdded += 1
                                Else
                                    RecsUpdated += 1
                                End If
                            Else
                                'FoundLastRow = True
                            End If
                        Else
                            Dim chkVal As String = currentRow(0) & String.Empty
                            If chkVal = "PO Number" Then
                                FoundStartRow = True
                            End If
                        End If
                        If FoundLastRow Then
                            Exit While
                        End If
                    Catch ex As FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
            End Using
            Me.Cursor = Cursors.Default
            Update_Record_TDA("SOTWMPO1")
            Dim iMsg As New System.Text.StringBuilder With {.Length = 0}
            iMsg.AppendLine(String.Format("Records Updated: {0}", RecsUpdated))
            iMsg.AppendLine(String.Format("Records Added: {0}", RecsAdded))
            iMsg.AppendLine(String.Format("Warnings: {0}", RecWarnings))
            MsgBox(iMsg.ToString, vbOKOnly, "Import Complete!")
            ASCMAIN1.Progress("", "")

        End If
    End Sub

    Private Sub ImportPOStoreFileXLS()
        Dim fDialog As New OpenFileDialog
        Dim FullFileName As String = ""
        Dim FileName As String = ""

        fDialog.Filter = "Excel Files|*.xlsx"
        fDialog.Title = "Select an Excel File To Import"
        If fDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            FullFileName = fDialog.FileNames(0)
            FileName = fDialog.SafeFileName
            'Begin - Copied Code
            Me.Cursor = Cursors.WaitCursor
            Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Add
            Dim XWS As Microsoft.Office.Interop.Excel.Worksheet = XWB.Sheets(1)
            Try
                XWB = excel.Workbooks.Open(FullFileName)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Opening File")
                XWB.Close()
                XWB = Nothing
                excel = Nothing
                Exit Sub
            End Try
            Me.Cursor = Cursors.WaitCursor
            XWS = XWB.Worksheets(1)

            Dim RecsUpdated As Int64 = 0
            Dim RecsAdded As Int64 = 0
            Dim RecWarnings As Int64 = 0
            Dim FoundStartRow As Boolean = False
            Dim FoundStartPO As Boolean = False
            Dim FoundLastRow As Boolean = False
            Fill_Records("SOTWMPO1")
            Fill_Records("SOTWMPOH")

            For curRow As Integer = 1 To 20000
                If FoundStartRow = False And curRow > 1000 Then
                    MsgBox("Can Not Find Starting Row!", vbOKOnly, "Problem With File")
                    Exit For
                End If
                If (curRow Mod 10) = 0 Then
                    ASCMAIN1.Progress("Processing Row", curRow)
                End If

                'If curRow = 4012 Then Stop

                If FoundStartRow Then
                    Dim PO_NUMBER As String = formatPO_NUMBER(XWS.Cells(curRow, 1).text.ToString & String.Empty)
                    Dim CUST_STORE_NO As String = (XWS.Cells(curRow, 14).text.ToString & String.Empty).ToString.PadLeft(6, "0")
                    If PO_NUMBER <> "" Then
                        FoundStartPO = True
                        Dim fltr As String = $"PO_NUMBER = '{PO_NUMBER}' AND CUST_STORE_NO = '{CUST_STORE_NO}'"
                        Dim rowSOTWMPO1 As DataRow = dst.Tables.Item("SOTWMPO1").Select(fltr).FirstOrDefault
                        Dim newRow As Boolean = False
                        If IsNothing(rowSOTWMPO1) Then
                            rowSOTWMPO1 = dst.Tables("SOTWMPO1").NewRow
                            newRow = True
                        End If

                        For Each rowSOTWMPOH As DataRow In dst.Tables("SOTWMPOH").Select()
                            Dim EXL_COL As String = (rowSOTWMPOH.Item("EXL_COL").ToString & String.Empty).ToUpper
                            Dim EXL_COL_NUM As Int64 = Asc(EXL_COL) - 64
                            Dim EXL_COL_HEADER As String = rowSOTWMPOH.Item("EXL_COL_HEADER").ToString & String.Empty
                            Dim EXCEL_ORA_COL As String = rowSOTWMPOH.Item("EXCEL_ORA_COL").ToString & String.Empty
                            Dim DATA_TYPE As String = rowSOTWMPOH.Item("DATA_TYPE").ToString & String.Empty
                            Dim VAR_CUR As String = XWS.Cells(curRow, EXL_COL_NUM).text.ToString & String.Empty
                            'Current - Copied Code Changed
                            Select Case DATA_TYPE
                                Case "S"
                                    If EXCEL_ORA_COL = "PO_NUMBER" Then
                                        rowSOTWMPO1.Item(EXCEL_ORA_COL) = formatPO_NUMBER((XWS.Cells(curRow, EXL_COL_NUM).text.ToString & String.Empty).Trim)
                                    Else
                                        rowSOTWMPO1.Item(EXCEL_ORA_COL) = (XWS.Cells(curRow, EXL_COL_NUM).text.ToString & String.Empty).Trim
                                    End If
                                Case "V"
                                    rowSOTWMPO1.Item(EXCEL_ORA_COL) = Val((XWS.Cells(curRow, EXL_COL_NUM).text.ToString & String.Empty).Replace("$", "").Replace(",", ""))
                                Case Else
                                    Stop
                            End Select
                        Next
                        rowSOTWMPO1.Item("EXCEL_FILE") = FileName
                        rowSOTWMPO1.Item("EXCEL_LINE") = curRow
                        If newRow Then
                            dst.Tables.Item("SOTWMPO1").Rows.Add(rowSOTWMPO1)
                            RecsAdded += 1
                        Else
                            RecsUpdated += 1
                        End If
                    Else
                        If FoundStartPO Then
                            FoundLastRow = True
                        End If
                    End If
                Else
                    Dim POCHK As String = XWS.Cells(curRow, 1).text.ToString & String.Empty
                    If POCHK.ToUpper() = "PO NUMBER" Then
                        FoundStartRow = True
                        For Each rowSOTWMPOH As DataRow In dst.Tables("SOTWMPOH").Select()
                            Dim EXL_COL As String = (rowSOTWMPOH.Item("EXL_COL").ToString & String.Empty).ToUpper
                            Dim EXL_COL_NUM As Int64 = Asc(EXL_COL) - 64
                            Dim EXL_COL_HEADER As String = rowSOTWMPOH.Item("EXL_COL_HEADER").ToString & String.Empty
                            If XWS.Cells(curRow, EXL_COL_NUM).text.ToString & String.Empty <> EXL_COL_HEADER Then
                                MsgBox("Invalid Header Found In File", vbOKOnly, "Please Let Wayne Know")
                                XWB.Close()
                                XWB = Nothing
                                excel = Nothing
                                Me.Cursor = Cursors.Default
                                ASCMAIN1.Progress("", "")
                                Exit Sub
                            End If
                        Next
                    End If
                End If

                If FoundLastRow Then
                    Exit For
                End If
            Next
            XWB.Close()
            XWB = Nothing
            excel = Nothing
            Me.Cursor = Cursors.Default
            Update_Record_TDA("SOTWMPO1")
            Dim iMsg As New System.Text.StringBuilder With {.Length = 0}
            iMsg.AppendLine(String.Format("Records Updated: {0}", RecsUpdated))
            iMsg.AppendLine(String.Format("Records Added: {0}", RecsAdded))
            iMsg.AppendLine(String.Format("Warnings: {0}", RecWarnings))
            MsgBox(iMsg.ToString, vbOKOnly, "Import Complete!")
            ASCMAIN1.Progress("", "")
            'End - Copied Code
        End If
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        TabControl1.Visible = Not tf

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            '.Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Import PO Data").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Import Store Data").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("ORDR_GROUP_NO").Text = ""
        Absx1.txtFor("ORDR_CUST_PO").Text = ""

        dst.EnforceConstraints = False
        For Each TABLE_NAME As String In New String() _
            {"SOTSHPWH", "SOTORDR1", "SOTPICK1", "SOTPICK2", "SOTORDRS", "SOTSHIP1", "SOTSHIP2", "SOTCART1", "SOTCART2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        'Fill_Records("PMTVIST1")

        dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        'tab.Visible = ScreenMode

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        'dst.Tables("SOTROYLI").Rows.Clear()

        dst.EnforceConstraints = False

        fillSOTSHPWA()
        Dim ORDR_GROUP_NO As String = Absx1.txtFor("ORDR_GROUP_NO").Text.ToString & String.Empty
        If ORDR_GROUP_NO.Length > 0 Then
            For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTPICK1", "SOTSHIP1", "SOTCART1", "SOTCART2"}
                Fill_Records(TABLE_NAME, ORDR_GROUP_NO)
            Next
            'Setup_SOTORDRS()
        End If

        dst.EnforceConstraints = True

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()
        'BeginTrans()
        'INIT_LAST("PMTVIST1", True, "", True)
        'Update_Record_TDA("PMTVIST1")
        'CommitTrans("Update Complete")
    End Sub

    Sub Setup_Summary()
        grdSOTSHPWA.Update()
        grdSOTSHPWA.Refresh()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOTSHPWA, "SSB", "Show Filter", "Show GroupBox", "Customer Order Inquiry")
        Call Load_Popup_Menu(grdSTOREPO1, "SSB", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdSTOREPOS, "SSB", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdSTOREPOV, "SSBB", "Show Filter", "Show GroupBox", "Refresh")
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

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'Dim tlb_btn As UltraWinToolbars.ButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""

                    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
            Case "Customer Order Inquiry"
                Dim FIND_BY As String = CUST_CODE
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Text
                FIND_BY &= ":" & ORDR_GROUP_NO
                Context_Launch("Select", FIND_BY, e.Tool.Key, "SOFCORD1")
            Case "Refresh"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Refreshing", "")
                Application.DoEvents()
                Fill_Records("STOREPOV")
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
                Application.DoEvents()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"
    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

#End Region

#Region "Custom Methods"

    Private Function formatPO_NUMBER(ByVal PO_NUMBER As String) As String
        Dim retval As String = PO_NUMBER
        Dim PAD_NO As String = ""
        Select Case retval.Length
            Case 9
                PAD_NO = "0"
            Case 8
                PAD_NO = "00"
            Case 7
                PAD_NO = "000"
            Case 6
                PAD_NO = "0000"
        End Select
        retval = PAD_NO & retval
        Return retval
    End Function

    Private Sub fillSOTSHPWA()
        Dim ORDR_GROUP_NO As String = Absx1.txtFor("ORDR_GROUP_NO").Text & String.Empty
        Dim filter As String = String.Format("ORDR_GROUP_NO = {0}", ORDR_GROUP_NO)
        Dim rowSOTSHPWA As DataRow = dst.Tables("SOTSHPWA").Select(filter).FirstOrDefault
        If Not IsNothing(rowSOTSHPWA) Then
            Dim newSOTSHPWH As DataRow = dst.Tables("SOTSHPWH").NewRow
            For Each dc As DataColumn In dst.Tables.Item("SOTSHPWA").Columns
                Select Case dc.ColumnName
                    Case "VARIANCE", "VARIANCE_COST"
                    Case Else
                        newSOTSHPWH.Item(dc.ColumnName) = rowSOTSHPWA.Item(dc.ColumnName)
                End Select
            Next
            dst.Tables("SOTSHPWH").Rows.Add(newSOTSHPWH)
        End If
    End Sub

    Private Function GetTANNERSVILLE(ByRef ORDR_GROUP_NO As String) As Integer
        Dim RetVal As String = "0"
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT COUNT(*) FROM SOTSHIPB")
        SQLS.AppendLine("WHERE BOL_NO IN (")
        SQLS.AppendLine("  SELECT DISTINCT(MASTER_BILL_OF_LADING_NO) MASTER_BILL_OF_LADING_NO")
        SQLS.AppendLine("  FROM SOTSHIP1")
        SQLS.AppendLine(String.Format("  WHERE ORDR_GROUP_NO = '{0}'", ORDR_GROUP_NO))
        SQLS.AppendLine(")")
        ASCMAIN1.sql = SQLS.ToString()
        Dim TAN As Int16 = Val(ASCDATA1.GetDataValue)
        If TAN >= 1 Then
            RetVal = 1
        End If
        Return RetVal
    End Function

    Private Sub ImportWalmartFile()
        Dim fDialog As New OpenFileDialog
        Dim FullFileName As String = ""
        If chkImportCSV.Checked Then
            fDialog.Filter = "CSV Files|*.csv"
            fDialog.Title = "Select a CSV File To Import"
            If fDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FullFileName = fDialog.FileNames(0)
                ImportWalmartCSVFile(FullFileName)
            End If
        Else
            fDialog.Filter = "Excel Files|*.xlsx"
            fDialog.Title = "Select an Excel File To Import"
            If fDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FullFileName = fDialog.FileNames(0)
                ImportWalmartExcelFile(FullFileName)
            End If
        End If

    End Sub

    Private Sub ImportWalmartCSVFile(fullFileName As String)
        Me.Cursor = Cursors.WaitCursor
        Dim RecsUpdated As Int64 = 0
        Dim RecsAdded As Int64 = 0
        Dim RecWarnings As Int64 = 0
        Dim FoundStartRow As Boolean = False
        Dim FoundLastRow As Boolean = False
        Fill_Records("SOTWMTD1")
        Dim curRow As Int64 = 0
        Using MyReader As New FileIO.TextFieldParser(fullFileName)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()
            While Not MyReader.EndOfData
                curRow += 1
                If FoundStartRow = False And curRow > 1000 Then
                    MsgBox("Can Not Find Starting Row!", vbOKOnly, "Problem With File")
                    Exit While
                End If
                If (curRow Mod 100) = 0 Then
                    ASCMAIN1.Progress("Processing Row", curRow)
                End If
                Try
                    currentRow = MyReader.ReadFields()
                    If FoundStartRow Then
                        Dim PO_NUMBER As String = formatPO_NUMBER(currentRow(4).ToString & String.Empty)
                        If PO_NUMBER <> "" Then
                            If PO_EXISTS(PO_NUMBER) Then
                                Dim filter As String = String.Format("PO_NUMBER = '{0}'", PO_NUMBER)
                                Dim rowSOTWMTD1 As DataRow = dst.Tables.Item("SOTWMTD1").Select(filter).FirstOrDefault
                                SetSOTWMTD1Data_CSV(rowSOTWMTD1, currentRow)
                                RecsUpdated += 1
                            Else
                                Dim newSOTWMTD1 As DataRow = dst.Tables.Item("SOTWMTD1").NewRow
                                SetSOTWMTD1Data_CSV(newSOTWMTD1, currentRow)
                                dst.Tables.Item("SOTWMTD1").Rows.Add(newSOTWMTD1)
                                RecsAdded += 1
                            End If
                        Else
                            FoundLastRow = True
                        End If
                    Else
                        Dim chkVal As String = currentRow(0) & String.Empty
                        If chkVal = "Vendor Nbr" Then
                            FoundStartRow = True
                        End If
                    End If
                    If FoundLastRow Then
                        Exit While
                    End If
                Catch ex As FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
            End While
        End Using
        Me.Cursor = Cursors.Default
        Update_Record_TDA("SOTWMTD1")
        Dim iMsg As New System.Text.StringBuilder With {.Length = 0}
        iMsg.AppendLine(String.Format("Records Updated: {0}", RecsUpdated))
        iMsg.AppendLine(String.Format("Records Added: {0}", RecsAdded))
        iMsg.AppendLine(String.Format("Warnings: {0}", RecWarnings))
        MsgBox(iMsg.ToString, vbOKOnly, "Import Complete!")
        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub ImportWalmartExcelFile(ByVal FullFileName As String)
        Me.Cursor = Cursors.WaitCursor
        Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
        Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Add
        Dim XWS As Microsoft.Office.Interop.Excel.Worksheet = XWB.Sheets(1)
        Try
            XWB = excel.Workbooks.Open(FullFileName)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Opening File")
            XWB.Close()
            XWB = Nothing
            excel = Nothing
            Exit Sub
        End Try
        Me.Cursor = Cursors.WaitCursor
        XWS = XWB.Worksheets(1)

        Dim RecsUpdated As Int64 = 0
        Dim RecsAdded As Int64 = 0
        Dim RecWarnings As Int64 = 0
        Dim FoundStartRow As Boolean = False
        Dim FoundLastRow As Boolean = False
        Fill_Records("SOTWMTD1")
        For curRow As Integer = 1 To 20000
            If FoundStartRow = False And curRow > 1000 Then
                MsgBox("Can Not Find Starting Row!", vbOKOnly, "Problem With File")
                Exit For
            End If
            If (curRow Mod 100) = 0 Then
                ASCMAIN1.Progress("Processing Row", curRow)
            End If

            If FoundStartRow Then
                Dim PO_NUMBER As String = formatPO_NUMBER(XWS.Cells(curRow, 5).text.ToString & String.Empty)
                If PO_NUMBER <> "" Then
                    If PO_EXISTS(PO_NUMBER) Then
                        Dim filter As String = String.Format("PO_NUMBER = '{0}'", PO_NUMBER)
                        Dim rowSOTWMTD1 As DataRow = dst.Tables.Item("SOTWMTD1").Select(filter).FirstOrDefault
                        SetSOTWMTD1Data_EXCEL(rowSOTWMTD1, XWS, curRow)
                        RecsUpdated += 1
                    Else
                        Dim newSOTWMTD1 As DataRow = dst.Tables.Item("SOTWMTD1").NewRow
                        SetSOTWMTD1Data_EXCEL(newSOTWMTD1, XWS, curRow)
                        dst.Tables.Item("SOTWMTD1").Rows.Add(newSOTWMTD1)
                        RecsAdded += 1
                    End If
                Else
                    FoundLastRow = True
                End If
            Else
                Dim chkVal As String = XWS.Cells(curRow, 1).text.ToString & String.Empty
                If chkVal = "Vendor Nbr" Then
                    FoundStartRow = True
                End If
            End If

            If FoundLastRow Then
                Exit For
            End If
        Next
        XWB.Close()
        XWB = Nothing
        excel = Nothing
        Me.Cursor = Cursors.Default
        Update_Record_TDA("SOTWMTD1")
        Dim iMsg As New System.Text.StringBuilder With {.Length = 0}
        iMsg.AppendLine(String.Format("Records Updated: {0}", RecsUpdated))
        iMsg.AppendLine(String.Format("Records Added: {0}", RecsAdded))
        iMsg.AppendLine(String.Format("Warnings: {0}", RecWarnings))
        MsgBox(iMsg.ToString, vbOKOnly, "Import Complete!")
        ASCMAIN1.Progress("", "")

    End Sub

    Private Function PO_EXISTS(ByVal PO_NUMBER As String) As Boolean
        Dim retVal As Boolean = False
        Dim filter As String = String.Format("PO_NUMBER = '{0}'", PO_NUMBER)
        retVal = dst.Tables.Item("SOTWMTD1").Select(filter).Count > 0
        Return retVal
    End Function

    Private Function MakeNum(ByVal CellVal As String) As Double
        Dim Retval As Double = 0
        Dim StripString As String = CellVal.Replace(",", "").Replace("$", "").Replace("%", "")
        If IsNumeric(StripString) Then
            Retval = Val(StripString)
        End If
        Return Retval
    End Function

    Private Sub MatchPOData()
        Fill_Records("SOTWMTD1")
        For Each rowSOTSHPWA As DataRow In dst.Tables("SOTSHPWA").Select()
            Dim ORDR_CUST_PO As String = rowSOTSHPWA.Item("ORDR_CUST_PO").ToString & String.Empty
            Dim filter As String = String.Format("PO_NUMBER = '{0}'", ORDR_CUST_PO)
            Dim ORDR_GROUP_NO As String = rowSOTSHPWA.Item("ORDR_GROUP_NO").ToString & String.Empty
            Dim TAN As Integer = GetTANNERSVILLE(ORDR_GROUP_NO)
            rowSOTSHPWA.Item("TANNER") = TAN
            For Each rowSOTWMTD1 As DataRow In dst.Tables.Item("SOTWMTD1").Select(filter)
                If Not IsNothing(rowSOTWMTD1) Then
                    rowSOTSHPWA.Item("VEND_NO") = rowSOTWMTD1.Item("VEND_NO")
                    rowSOTSHPWA.Item("VEND_NO_DEPT") = rowSOTWMTD1.Item("VEND_NO_DEPT")
                    rowSOTSHPWA.Item("VEND_SEQ_NO") = rowSOTWMTD1.Item("VEND_SEQ_NO")
                    rowSOTSHPWA.Item("WHSE_NO") = rowSOTWMTD1.Item("WHSE_NO")
                    rowSOTSHPWA.Item("PO_TYPE") = rowSOTWMTD1.Item("PO_TYPE")
                    If rowSOTSHPWA.Item("PO_EVENT").ToString & String.Empty = "" Then
                        rowSOTSHPWA.Item("PO_EVENT") = rowSOTWMTD1.Item("PO_EVENT")
                    End If
                    rowSOTSHPWA.Item("PO_DATE_CREATED") = rowSOTWMTD1.Item("PO_DATE_CREATED")
                    rowSOTSHPWA.Item("PO_SHIP_DATE") = rowSOTWMTD1.Item("PO_SHIP_DATE")
                    rowSOTSHPWA.Item("PO_CANCEL_DATE") = rowSOTWMTD1.Item("PO_CANCEL_DATE")
                    rowSOTSHPWA.Item("ORIG_MABD_DATE") = rowSOTWMTD1.Item("ORIG_MABD_DATE")
                    rowSOTSHPWA.Item("MABD_DATE") = rowSOTWMTD1.Item("MABD_DATE")
                    rowSOTSHPWA.Item("MABD_COMPLIANCE_DATE") = rowSOTWMTD1.Item("MABD_COMPLIANCE_DATE")
                    rowSOTSHPWA.Item("WHSE_QTY_ORDR") =
                        Val(rowSOTSHPWA.Item("WHSE_QTY_ORDR").ToString & String.Empty) +
                        Val(rowSOTWMTD1.Item("WHSE_QTY_ORDR").ToString & String.Empty)
                    rowSOTSHPWA.Item("WHSE_QTY_REC") =
                        Val(rowSOTSHPWA.Item("WHSE_QTY_REC").ToString & String.Empty) +
                        Val(rowSOTWMTD1.Item("WHSE_QTY_REC").ToString & String.Empty)
                    rowSOTSHPWA.Item("PCT_REC") =
                        Val(rowSOTSHPWA.Item("PCT_REC").ToString & String.Empty) +
                        Val(rowSOTWMTD1.Item("PCT_REC").ToString & String.Empty)
                    rowSOTSHPWA.Item("TOT_ORDR_COST") =
                        Val(rowSOTSHPWA.Item("TOT_ORDR_COST").ToString & String.Empty) +
                        Val(rowSOTWMTD1.Item("TOT_ORDR_COST").ToString & String.Empty)
                    rowSOTSHPWA.Item("TOT_REC_COST") =
                        Val(rowSOTSHPWA.Item("TOT_REC_COST").ToString & String.Empty) +
                        Val(rowSOTWMTD1.Item("TOT_REC_COST").ToString & String.Empty)
                End If
            Next
            'Dim rowSOTWMTD1 As DataRow = dst.Tables.Item("SOTWMTD1").Select(filter).FirstOrDefault
            'If Not IsNothing(rowSOTWMTD1) Then
            '    rowSOTSHPWA.Item("VEND_NO") = rowSOTWMTD1.Item("VEND_NO")
            '    rowSOTSHPWA.Item("VEND_NO_DEPT") = rowSOTWMTD1.Item("VEND_NO_DEPT")
            '    rowSOTSHPWA.Item("VEND_SEQ_NO") = rowSOTWMTD1.Item("VEND_SEQ_NO")
            '    rowSOTSHPWA.Item("WHSE_NO") = rowSOTWMTD1.Item("WHSE_NO")
            '    rowSOTSHPWA.Item("PO_TYPE") = rowSOTWMTD1.Item("PO_TYPE")
            '    rowSOTSHPWA.Item("PO_EVENT") = rowSOTWMTD1.Item("PO_EVENT")
            '    rowSOTSHPWA.Item("PO_DATE_CREATED") = rowSOTWMTD1.Item("PO_DATE_CREATED")
            '    rowSOTSHPWA.Item("PO_SHIP_DATE") = rowSOTWMTD1.Item("PO_SHIP_DATE")
            '    rowSOTSHPWA.Item("PO_CANCEL_DATE") = rowSOTWMTD1.Item("PO_CANCEL_DATE")
            '    rowSOTSHPWA.Item("ORIG_MABD_DATE") = rowSOTWMTD1.Item("ORIG_MABD_DATE")
            '    rowSOTSHPWA.Item("MABD_DATE") = rowSOTWMTD1.Item("MABD_DATE")
            '    rowSOTSHPWA.Item("MABD_COMPLIANCE_DATE") = rowSOTWMTD1.Item("MABD_COMPLIANCE_DATE")
            '    rowSOTSHPWA.Item("WHSE_QTY_ORDR") = rowSOTWMTD1.Item("WHSE_QTY_ORDR")
            '    rowSOTSHPWA.Item("WHSE_QTY_REC") = rowSOTWMTD1.Item("WHSE_QTY_REC")
            '    rowSOTSHPWA.Item("PCT_REC") = rowSOTWMTD1.Item("PCT_REC")
            '    rowSOTSHPWA.Item("TOT_ORDR_COST") = rowSOTWMTD1.Item("TOT_ORDR_COST")
            '    rowSOTSHPWA.Item("TOT_REC_COST") = rowSOTWMTD1.Item("TOT_REC_COST")
            'End If
        Next
    End Sub

    Private Sub SetSOTWMPO1Data_CSV(ByRef rowSOTWMPO1 As DataRow, ByVal currentRow As String())
        If currentRow(0) & String.Empty <> "" Then
            '0','PO Number','PO_NUMBER'
            rowSOTWMPO1.Item("PO_NUMBER") = formatPO_NUMBER(currentRow(0).ToString & String.Empty)
            '13','Store Nbr','CUST_STORE_NO'
            rowSOTWMPO1.Item("CUST_STORE_NO") = (currentRow(13).ToString & String.Empty).ToString.PadLeft(6, "0")
            '1','PO Status','PO_STATUS'
            rowSOTWMPO1.Item("PO_STATUS") = (currentRow(1).ToString & String.Empty).ToString
            '4','Total Eaches Str Ordered','ST_EA_ORDR'
            rowSOTWMPO1.Item("ST_EA_ORDR") = Val((currentRow(4).ToString & String.Empty).ToString.Replace("$", "").Replace(",", ""))
            '5','Curr Str In Transit Qty','CURR_ST_TRAN'
            rowSOTWMPO1.Item("CURR_ST_TRAN") = Val((currentRow(5).ToString & String.Empty).ToString.Replace("$", "").Replace(",", ""))
            '6','Curr Str On Hand Qty','CURR_ST_ON_HAND'
            rowSOTWMPO1.Item("CURR_ST_ON_HAND") = Val((currentRow(6).ToString & String.Empty).ToString.Replace("$", "").Replace(",", ""))
            '7','Curr Str In Whse Qty','CURR_ST_WHS_QTY'
            rowSOTWMPO1.Item("CURR_ST_WHS_QTY") = Val((currentRow(7).ToString & String.Empty).ToString.Replace("$", "").Replace(",", ""))
            '8','Curr Str On Order Qty','CURR_ST_ORD_QTY'
            rowSOTWMPO1.Item("CURR_ST_ORD_QTY") = Val((currentRow(8).ToString & String.Empty).ToString.Replace("$", "").Replace(",", ""))
            '9','Gross Ship Qty','GROSS_SHIP_QTY'
            rowSOTWMPO1.Item("GROSS_SHIP_QTY") = Val((currentRow(9).ToString & String.Empty).ToString.Replace("$", "").Replace(",", ""))
            '10','Net Ship Qty','NET_SHIP_QTY'
            rowSOTWMPO1.Item("NET_SHIP_QTY") = Val((currentRow(10).ToString & String.Empty).ToString.Replace("$", "").Replace(",", ""))
            '11','Total Eaches Str Received','ST_EACH_RCD'
            rowSOTWMPO1.Item("ST_EACH_RCD") = Val((currentRow(12).ToString & String.Empty).ToString.Replace("$", "").Replace(",", ""))
        End If
    End Sub

    Private Sub SetSOTWMTD1Data_CSV(ByRef rowSOTWMTD1 As DataRow, ByVal currentRow As String())
        If currentRow(4) & String.Empty <> "" Then
            'PO_NUMBER
            rowSOTWMTD1.Item("PO_NUMBER") = formatPO_NUMBER(currentRow(4).ToString & String.Empty)
            'VEND_NO
            rowSOTWMTD1.Item("VEND_NO") = currentRow(0).ToString & String.Empty
            'VEND_NO_DEPT
            rowSOTWMTD1.Item("VEND_NO_DEPT") = currentRow(1).ToString & String.Empty
            'VEND_SEQ_NO
            rowSOTWMTD1.Item("VEND_SEQ_NO") = MakeNum(currentRow(2).ToString)
            'WHSE_NO
            rowSOTWMTD1.Item("WHSE_NO") = currentRow(3).ToString & String.Empty
            'PO_TYPE
            rowSOTWMTD1.Item("PO_TYPE") = currentRow(5).ToString & String.Empty
            'PO_EVENT
            If (currentRow(6).ToString & String.Empty).Length > 25 Then
                rowSOTWMTD1.Item("PO_EVENT") = (currentRow(6).ToString & String.Empty).Substring(0, 25)
            Else
                rowSOTWMTD1.Item("PO_EVENT") = currentRow(6).ToString & String.Empty
            End If
            'PO_DATE_CREATED
            If IsDate(currentRow(7).ToString & String.Empty) Then
                rowSOTWMTD1.Item("PO_DATE_CREATED") = CDate(currentRow(7).ToString & String.Empty)
            End If
            'PO_SHIP_DATE
            If IsDate(currentRow(8).ToString & String.Empty) Then
                rowSOTWMTD1.Item("PO_SHIP_DATE") = CDate(currentRow(8).ToString & String.Empty)
            End If
            'PO_CANCEL_DATE
            If IsDate(currentRow(9).ToString & String.Empty) Then
                rowSOTWMTD1.Item("PO_CANCEL_DATE") = CDate(currentRow(9).ToString & String.Empty)
            End If
            'ORIG_MABD_DATE
            If IsDate(currentRow(10).ToString & String.Empty) Then
                rowSOTWMTD1.Item("ORIG_MABD_DATE") = CDate(currentRow(10).ToString & String.Empty)
            End If
            'MABD_DATE
            If IsDate(currentRow(11).ToString & String.Empty) Then
                rowSOTWMTD1.Item("MABD_DATE") = CDate(currentRow(11).ToString & String.Empty)
            End If
            'MABD_COMPLIANCE_DATE
            If IsDate(currentRow(12).ToString & String.Empty) Then
                rowSOTWMTD1.Item("MABD_COMPLIANCE_DATE") = CDate(currentRow(12).ToString & String.Empty)
            End If
            'WHSE_QTY_ORDR
            rowSOTWMTD1.Item("WHSE_QTY_ORDR") = MakeNum(currentRow(13).ToString)
            'WHSE_QTY_REC
            rowSOTWMTD1.Item("WHSE_QTY_REC") = MakeNum(currentRow(14).ToString)
            'PCT_REC
            rowSOTWMTD1.Item("PCT_REC") = MakeNum(currentRow(15).ToString)
            'TOT_ORDR_COST
            rowSOTWMTD1.Item("TOT_ORDR_COST") = MakeNum(currentRow(16).ToString)
            'TOT_REC_COST
            rowSOTWMTD1.Item("TOT_REC_COST") = MakeNum(currentRow(17).ToString)
        End If
    End Sub

    Private Sub SetSOTWMTD1Data_EXCEL(ByRef rowSOTWMTD1 As DataRow, ByRef XWS As Microsoft.Office.Interop.Excel.Worksheet, ByVal curRow As Integer)
        If XWS.Cells(curRow, 5).text.ToString & String.Empty <> "" Then
            'PO_NUMBER
            rowSOTWMTD1.Item("PO_NUMBER") = formatPO_NUMBER(XWS.Cells(curRow, 5).text.ToString & String.Empty)
            'VEND_NO
            rowSOTWMTD1.Item("VEND_NO") = XWS.Cells(curRow, 1).text.ToString & String.Empty
            'VEND_NO_DEPT
            rowSOTWMTD1.Item("VEND_NO_DEPT") = XWS.Cells(curRow, 2).text.ToString & String.Empty
            'VEND_SEQ_NO
            rowSOTWMTD1.Item("VEND_SEQ_NO") = MakeNum(XWS.Cells(curRow, 3).text.ToString)
            'WHSE_NO
            rowSOTWMTD1.Item("WHSE_NO") = XWS.Cells(curRow, 4).text.ToString & String.Empty
            'PO_TYPE
            rowSOTWMTD1.Item("PO_TYPE") = XWS.Cells(curRow, 6).text.ToString & String.Empty
            'PO_EVENT
            If (XWS.Cells(curRow, 7).text.ToString & String.Empty).Length > 25 Then
                rowSOTWMTD1.Item("PO_EVENT") = (XWS.Cells(curRow, 7).text.ToString & String.Empty).Substring(0, 25)
            Else
                rowSOTWMTD1.Item("PO_EVENT") = XWS.Cells(curRow, 7).text.ToString & String.Empty
            End If
            'PO_DATE_CREATED
            If IsDate(XWS.Cells(curRow, 8).text.ToString & String.Empty) Then
                rowSOTWMTD1.Item("PO_DATE_CREATED") = CDate(XWS.Cells(curRow, 8).text.ToString & String.Empty)
            End If
            'PO_SHIP_DATE
            If IsDate(XWS.Cells(curRow, 9).text.ToString & String.Empty) Then
                rowSOTWMTD1.Item("PO_SHIP_DATE") = CDate(XWS.Cells(curRow, 9).text.ToString & String.Empty)
            End If
            'PO_CANCEL_DATE
            If IsDate(XWS.Cells(curRow, 10).text.ToString & String.Empty) Then
                rowSOTWMTD1.Item("PO_CANCEL_DATE") = CDate(XWS.Cells(curRow, 10).text.ToString & String.Empty)
            End If
            'ORIG_MABD_DATE
            If IsDate(XWS.Cells(curRow, 11).text.ToString & String.Empty) Then
                rowSOTWMTD1.Item("ORIG_MABD_DATE") = CDate(XWS.Cells(curRow, 11).text.ToString & String.Empty)
            End If
            'MABD_DATE
            If IsDate(XWS.Cells(curRow, 12).text.ToString & String.Empty) Then
                rowSOTWMTD1.Item("MABD_DATE") = CDate(XWS.Cells(curRow, 12).text.ToString & String.Empty)
            End If
            'MABD_COMPLIANCE_DATE
            If IsDate(XWS.Cells(curRow, 13).text.ToString & String.Empty) Then
                rowSOTWMTD1.Item("MABD_COMPLIANCE_DATE") = CDate(XWS.Cells(curRow, 13).text.ToString & String.Empty)
            End If
            'WHSE_QTY_ORDR
            rowSOTWMTD1.Item("WHSE_QTY_ORDR") = MakeNum(XWS.Cells(curRow, 14).text.ToString)
            'WHSE_QTY_REC
            rowSOTWMTD1.Item("WHSE_QTY_REC") = MakeNum(XWS.Cells(curRow, 15).text.ToString)
            'PCT_REC
            rowSOTWMTD1.Item("PCT_REC") = MakeNum(XWS.Cells(curRow, 16).text.ToString)
            'TOT_ORDR_COST
            rowSOTWMTD1.Item("TOT_ORDR_COST") = MakeNum(XWS.Cells(curRow, 17).text.ToString)
            'TOT_REC_COST
            rowSOTWMTD1.Item("TOT_REC_COST") = MakeNum(XWS.Cells(curRow, 18).text.ToString)
        End If
    End Sub

    Private Sub Setup_SOTORDRS()
        Dim ORDR_GROUP_NO As String = Absx1.txtFor("ORDR_GROUP_NO").Text & String.Empty
        Dim sql As String = ""

        sql = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' group by ")
        grdSOTORDRS.Text = "Style Summary for Order Group " & ORDR_GROUP_NO

        For Each COLUMN_NAME As String In New String() _
                {"STYLE_CODE", "COLOR_CODE", "CUST_UPC", "RANGE_STYLE_CODE",
                 "CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_SKU"}

            'grdSOTORDRS.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not Absx1.chkFor("SHOW_" & COLUMN_NAME).Checked

            'If Not Absx1.chkFor("SHOW_" & COLUMN_NAME).Checked Then
            '    sql = Replace(sql, "SOTORDR2." & COLUMN_NAME, "NULL " & COLUMN_NAME, , 1)
            '    sql = Replace(sql, "SOTORDR2." & COLUMN_NAME, "NULL")
            '    If COLUMN_NAME = "STYLE_CODE" Then
            '        sql = Replace(sql, "SOTORDR2.STYLE_DESC", "NULL " & "STYLE_DESC", , 1)
            '        sql = Replace(sql, "SOTORDR2.STYLE_DESC", "NULL")
            '    End If
            '    If COLUMN_NAME = "COLOR_CODE" Then
            '        sql = Replace(sql, "ICTCOLR1.COLOR_DESC", "NULL " & "COLOR_DESC", , 1)
            '        sql = Replace(sql, "ICTCOLR1.COLOR_DESC", "NULL")
            '    End If
            'End If
        Next

        sql = Replace(sql, "ICTCOLR1.COLOR_CODE (+) = NULL", "ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE")
        Fill_Records("SOTORDRS", "", True, sql)
        Sort_grdColumns(grdSOTORDRS, "STYLE_CODE, COLOR_CODE, RANGE_STYLE_CODE, CUST_STYLE_CODE, CUST_COLOR_CODE, CUST_SKU")
    End Sub
#End Region

#Region "Form Controls"
    Private Sub btnSelGroups_Click(sender As Object, e As EventArgs)
        If CUST_CODE = "" Then
            MsgBox("You Must Select A Customer First")
        Else
            dst.Tables("SOTGROUP").Clear()
            Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
            SQLS.AppendLine(" ORDR_GROUP_NO IN (SELECT ORDR_GROUP_NO FROM SOTORDR1 WHERE ORDR_STATUS <> 'D' AND CUST_CODE = '" & CUST_CODE & "')")
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ORDR_GROUP_NO", , SQLS.ToString)
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.SelectedCodes.Count > 0 Then
                    For Each SelCode As String In ASCMAIN1.CodeSelector.SelectedCodes
                        Dim newRec As DataRow = dst.Tables("SOTGROUP").NewRow
                        newRec.Item("ORDR_GROUP_NO") = SelCode
                        dst.Tables("SOTGROUP").Rows.Add(newRec)
                    Next
                End If
                F.Dispose()
            End If
        End If
    End Sub

    Private Sub btnFETCH_Click(sender As Object, e As EventArgs) Handles btnFETCH.Click
        Me.Cursor = Cursors.WaitCursor

        Dim RYPLEGEND0 As String = Absx1.cmbFor("RYP0", True).Value
        Dim RYP0 As String = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)
        Dim RYPLEGEND1 As String = Absx1.cmbFor("RYP1", True).Value
        Dim RYP1 As String = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)

        Dim SQLB As New System.Text.StringBuilder

        SQLB.Length = 0
        SQLB.AppendLine("SELECT")
        SQLB.AppendLine("O1.ORDR_YYYYPP_BOOKED,")
        SQLB.AppendLine("O1.CUST_CODE,")
        SQLB.AppendLine("MAX(O1.ORDR_STATUS) AS ORDR_STATUS,")
        SQLB.AppendLine("MAX(O1.ORDR_GROUP_NO) AS ORDR_GROUP_NO,")
        SQLB.AppendLine("COUNT(DISTINCT(O1.ORDR_NO)) AS ORDR_CNT,")
        SQLB.AppendLine("O1.ORDR_CUST_PO,")
        SQLB.AppendLine("O1.ORDR_DATE_RECD,")
        SQLB.AppendLine("O1.ORDR_SHIP_DATE,")
        SQLB.AppendLine("O1.ORDR_CANCEL_DATE,")
        SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY,0)) AS ORDR_QTY,")
        SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY_OPEN,0)) AS ORDR_QTY_OPEN,")
        SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY_PICK,0)) AS ORDR_QTY_PICK,")
        SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY_SHIP,0)) AS ORDR_QTY_SHIP,")
        SQLB.AppendLine("SUM(NVL(O2.ORDR_QTY_CANC,0)) AS ORDR_QTY_CANC,")
        SQLB.AppendLine("I1.INV_DATE,")
        SQLB.AppendLine("I1.INV_NO_CONS,")
        SQLB.AppendLine("I1.INV_TOTAL_AMOUNT,")
        SQLB.AppendLine("C1.CART_CNT,")
        SQLB.AppendLine("C1.CART_TOTAL_UNITS AS QTY_PACKED_TOTAL,")
        SQLB.AppendLine("I1.INV_BALANCE,")
        SQLB.AppendLine("C1.SHIP_VIA_CODE,")
        SQLB.AppendLine("C1.SHIP_REF")
        SQLB.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2,")
        SQLB.AppendLine("(")
        SQLB.AppendLine("   SELECT")
        SQLB.AppendLine("   O1.ORDR_GROUP_NO,")
        SQLB.AppendLine("   COUNT(C1.CART_NO) AS CART_CNT,")
        SQLB.AppendLine("   SUM(NVL(C1.CART_TOTAL_UNITS, 0)) AS CART_TOTAL_UNITS,")
        SQLB.AppendLine("   MAX(S1.SHIP_VIA_CODE) AS SHIP_VIA_CODE,")
        SQLB.AppendLine("   MAX(S1.SHIP_REF) AS SHIP_REF")
        SQLB.AppendLine("   FROM SOTORDR1 O1, SOTPICK1 P1, SOTCART1 C1, SOTSHIP1 S1")
        SQLB.AppendLine("   WHERE O1.ORDR_NO = P1.ORDR_NO")
        SQLB.AppendLine("   AND P1.PICK_NO = C1.PICK_NO (+)")
        SQLB.AppendLine("   AND P1.SHIP_BOL_NO = S1.SHIP_BOL_NO (+)")
        SQLB.AppendLine("   AND P1.PICK_STATUS <> 'D'")
        SQLB.AppendLine("   GROUP BY O1.ORDR_GROUP_NO")
        SQLB.AppendLine(") C1,")
        SQLB.AppendLine("(")
        SQLB.AppendLine("  SELECT")
        SQLB.AppendLine("  O1.ORDR_GROUP_NO,")
        SQLB.AppendLine("  MAX(I1.INV_DATE) AS INV_DATE,")
        SQLB.AppendLine("  MAX(NVL(I1.INV_NO_CONS, I1.INV_NO || ':I')) AS INV_NO_CONS,")
        SQLB.AppendLine("  NVL(SUM(NVL(I1.INV_TOTAL_AMOUNT, 0)), 0) AS INV_TOTAL_AMOUNT,")
        SQLB.AppendLine("  NVL(SUM(NVL(A1.INV_BALANCE, 0)), 0) AS INV_BALANCE")
        SQLB.AppendLine("  FROM SOTINVH1 I1, SOTORDR1 O1, ARTOPEN1 A1")
        SQLB.AppendLine("  WHERE I1.ORDR_NO = O1.ORDR_NO")
        SQLB.AppendLine("  And I1.INV_NO = A1.INV_NUM (+)")
        SQLB.AppendLine("  GROUP BY O1.ORDR_GROUP_NO")
        SQLB.AppendLine(") I1")
        SQLB.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
        SQLB.AppendLine("AND O1.ORDR_GROUP_NO = C1.ORDR_GROUP_NO (+)")
        SQLB.AppendLine("AND O1.ORDR_GROUP_NO = I1.ORDR_GROUP_NO (+)")
        SQLB.AppendLine("AND O1.CUST_CODE IN ('WALMART','WALMARTCOM')")
        SQLB.AppendLine(String.Format(" And o1.ORDR_YYYYPP_BOOKED >= '{0}'", RYP0))
        SQLB.AppendLine(String.Format("AND O1.ORDR_YYYYPP_BOOKED <= '{0}'", RYP1))
        SQLB.AppendLine("GROUP BY")
        SQLB.AppendLine("O1.ORDR_YYYYPP_BOOKED,")
        SQLB.AppendLine("O1.CUST_CODE,")
        SQLB.AppendLine("O1.ORDR_CUST_PO,")
        SQLB.AppendLine("O1.ORDR_SHIP_DATE,")
        SQLB.AppendLine("O1.ORDR_CANCEL_DATE,")
        SQLB.AppendLine("O1.ORDR_DATE_RECD,")
        SQLB.AppendLine("I1.INV_DATE,")
        SQLB.AppendLine("I1.INV_NO_CONS,")
        SQLB.AppendLine("I1.INV_TOTAL_AMOUNT,")
        SQLB.AppendLine("I1.INV_BALANCE,")
        SQLB.AppendLine("C1.CART_CNT,")
        SQLB.AppendLine("C1.CART_TOTAL_UNITS,")
        SQLB.AppendLine("C1.SHIP_VIA_CODE,")
        SQLB.AppendLine("C1.SHIP_REF")
        Fill_Records("SOTSHPWA",,, SQLB.ToString)

        MatchPOData()

        Fill_Records("SOTNOMCH")

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnBegEnd_Click(sender As Object, e As EventArgs)
        Dim BegGroup As String = ""
        Dim EndGroup As String = ""
        Dim BegGroupVal As Int64 = 0
        Dim EndGroupVal As Int64 = 0
        If CUST_CODE = "" Then
            MsgBox("You Must Select A Customer First")
        Else
            dst.Tables("SOTGROUP").Clear()
            Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
            SQLS.AppendLine(" ORDR_GROUP_NO IN (SELECT ORDR_GROUP_NO FROM SOTORDR1 WHERE ORDR_STATUS <> 'D' AND CUST_CODE = '" & CUST_CODE & "')")
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ORDR_GROUP_NO", , SQLS.ToString)
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                lblGather1.Visible = True
                lblGather2.Visible = True

                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                ASCMAIN1.CodeSelector.Caption = "Select Beginning Group #"
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.SelectedCodes.Count = 1 Then
                    BegGroup = ASCMAIN1.CodeSelector.SelectedCode
                Else
                    MsgBox("You Can Only Select One Group When Using the Begin/End Feature", vbOKOnly, "Begin/End")
                    F.Dispose()
                    Exit Sub
                End If
                F.Dispose()
                lblGather1.Visible = False
                lblGather2.Visible = False
            End If
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                lblGather1.Visible = True
                lblGather2.Visible = True
                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                ASCMAIN1.CodeSelector.Caption = "Select Ending Group #"
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.SelectedCodes.Count = 1 Then
                    EndGroup = ASCMAIN1.CodeSelector.SelectedCode
                Else
                    MsgBox("You Can Only Select One Group When Using the Begin/End Feature", vbOKOnly, "Begin/End")
                    F.Dispose()
                    Exit Sub
                End If
                F.Dispose()
                lblGather1.Visible = False
                lblGather2.Visible = False
            End If
            BegGroupVal = Val(BegGroup)
            EndGroupVal = Val(EndGroup)
            Dim GroupCount As Int64 = 0
            For i As Int64 = BegGroupVal To EndGroupVal
                GroupCount += 1
                Dim newRec As DataRow = dst.Tables("SOTGROUP").NewRow
                newRec.Item("ORDR_GROUP_NO") = Format(i, "0000000000")
                dst.Tables("SOTGROUP").Rows.Add(newRec)
            Next
            MsgBox(GroupCount & " Groups Selected", vbOKOnly, "Selected")
        End If
    End Sub

    Private Sub chkOnlyVariances_CheckedChanged(sender As Object, e As EventArgs) Handles chkOnlyVariances.CheckedChanged
        Dim filterby As String = ""
        If chkOnlyVariances.Checked Then
            filterby = "VARIANCE <> 0 OR VARIANCE_COST <> 0"
        End If
        Dim dvw As DataView = DirectCast(grdSOTSHPWA.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Format(filterby)
    End Sub

    Private Sub grdSOTSHPWA_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTSHPWA.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ORDR_GROUP_NO").Text = e.Row.Cells("ORDR_GROUP_NO").Value
            Absx1.txtFor("ORDR_CUST_PO").Text = e.Row.Cells("ORDR_CUST_PO").Value
            'Absx1.txtFor("ORDR_GROUP_NO").Text = ORDR_GROUP_NO
            Click_Command("View")
        End If
    End Sub

    Private Sub grdSOTPICK1_BeforeRowExpanded(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPICK1.BeforeRowExpanded
        Dim PICK_NO As String = e.Row.Cells("PICK_NO").Value
        Fill_Records("SOTPICK2", PICK_NO)
        Sort_grdColumns(grdSOTPICK1, "PICK_LNO", False, 1)
    End Sub

    Private Sub grdSOTSHIP1_BeforeRowExpanded(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSHIP1.BeforeRowExpanded
        Dim SHIP_BOL_NO As String = e.Row.Cells("SHIP_BOL_NO").Value
        Fill_Records("SOTSHIP2", SHIP_BOL_NO)
        Sort_grdColumns(grdSOTSHIP1, "STYLE_CODE,COLOR_CODE", False, 1)
    End Sub

    Private Sub btnPOFetch_Click(sender As Object, e As EventArgs) Handles btnPOFetch.Click
        dst.Tables("PODATA").Clear()

        Dim SE As New StringBuilder With {.Length = 0}
        Dim SQ As New StringBuilder With {.Length = 0}
        Dim err As New StringBuilder With {.Length = 0}

        Dim PO_STORES As New Dictionary(Of String, String)
        If txtPOPO_1.Text.ToString.Length > 0 Then
            If Not PO_STORES.Keys.Contains(txtPOPO_1.Text) Then
                PO_STORES.Add(ComaToInStr(txtPOPO_1.Text.ToString), ComaToInStr(txtPOStores_1.Text.ToString, True))
            End If
        End If
        If txtPOPO_2.Text.ToString.Length > 0 Then
            If Not PO_STORES.Keys.Contains(txtPOPO_2.Text) Then
                PO_STORES.Add(ComaToInStr(txtPOPO_2.Text.ToString), ComaToInStr(txtPOStores_2.Text.ToString, True))
            End If
        End If
        If txtPOPO_3.Text.ToString.Length > 0 Then
            If Not PO_STORES.Keys.Contains(txtPOPO_3.Text) Then
                PO_STORES.Add(ComaToInStr(txtPOPO_3.Text.ToString), ComaToInStr(txtPOStores_3.Text.ToString, True))
            End If
        End If
        If txtPOPO_4.Text.ToString.Length > 0 Then
            If Not PO_STORES.Keys.Contains(txtPOPO_4.Text) Then
                PO_STORES.Add(ComaToInStr(txtPOPO_4.Text.ToString), ComaToInStr(txtPOStores_4.Text.ToString, True))
            End If
        End If
        If txtPOPO_5.Text.ToString.Length > 0 Then
            If Not PO_STORES.Keys.Contains(txtPOPO_5.Text) Then
                PO_STORES.Add(ComaToInStr(txtPOPO_5.Text.ToString), ComaToInStr(txtPOStores_5.Text.ToString, True))
            End If
        End If
        If txtPOPO_6.Text.ToString.Length > 0 Then
            If Not PO_STORES.Keys.Contains(txtPOPO_6.Text) Then
                PO_STORES.Add(ComaToInStr(txtPOPO_6.Text.ToString), ComaToInStr(txtPOStores_6.Text.ToString, True))
            End If
        End If
        'Dim POPO As String = ComaToInStr(txtPOPO_1.Text.ToString)
        'Dim POStores As String = ComaToInStr(txtPOStores_1.Text.ToString, True)
        'Dim POStyles As String = ComaToInStr(txtPOStyles.Text.ToString)

        If PO_STORES.Count = 0 Then
            err.AppendLine("You Must Supply A PO.")
        Else
            Me.Cursor = Cursors.WaitCursor

            dst.Tables.Item("PODATA").Clear()

            For Each POS As KeyValuePair(Of String, String) In PO_STORES
                SQ.Length = 0
                SQ.AppendLine("SELECT")
                SQ.AppendLine("I1.ORDR_CUST_PO,")
                SQ.AppendLine("I1.CUST_CODE,")
                SQ.AppendLine("I1.CUST_STORE_NO,")
                SQ.AppendLine("I2.STYLE_CODE,")
                SQ.AppendLine("I2.COLOR_CODE,")
                SQ.AppendLine("SUM(ORDR_QTY_SHIP) AS ORDR_QTY_SHIP")
                SQ.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
                SQ.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
                SQ.AppendLine("AND I1.INV_NO = I2.INV_NO")
                SQ.AppendLine("AND I1.CUST_CODE IN ('WALMART','WALMARTCOM')")
                SQ.AppendLine($"AND I1.ORDR_CUST_PO = {POS.Key}")
                If POS.Value.Length > 0 Then
                    SQ.AppendLine($"AND I1.CUST_STORE_NO IN ({POS.Value})")
                End If
                SQ.AppendLine("GROUP BY")
                SQ.AppendLine("I1.ORDR_CUST_PO,")
                SQ.AppendLine("I1.CUST_CODE,")
                SQ.AppendLine("I1.CUST_STORE_NO,")
                SQ.AppendLine("I2.STYLE_CODE,")
                SQ.AppendLine("I2.COLOR_CODE")
                ASCMAIN1.sql = SQ.ToString
                Dim TEMP = ASCMAIN1.Temp_Table
                Fill_Records("PODATA",, False, $"SELECT * FROM {TEMP}")

            Next
            grdPOSTYLES.Text = "Styles On PO's" '& POPO
        End If

        If err.Length = 0 Then
            ASCMAIN1.Progress("Fetching Calculating Data", "")
            Me.Cursor = Cursors.WaitCursor

            Dim STYLE_CODE As String = ""
            Dim COLOR_CODE As String = ""
            Dim CUR_REC As Int64 = 0
            For Each rowPODATA As DataRow In dst.Tables("PODATA").Select("", "STYLE_CODE, COLOR_CODE")
                Dim POPO As String = rowPODATA.Item("ORDR_CUST_PO").ToString & String.Empty
                'Stop
                CUR_REC += 1
                ASCMAIN1.Progress("", CUR_REC.ToString)
                If (STYLE_CODE <> rowPODATA.Item("STYLE_CODE").ToString & String.Empty) Or (COLOR_CODE <> rowPODATA.Item("COLOR_CODE").ToString & String.Empty) Then
                    STYLE_CODE = rowPODATA.Item("STYLE_CODE").ToString & String.Empty
                    COLOR_CODE = rowPODATA.Item("COLOR_CODE").ToString & String.Empty
                End If
                Dim CUST_STORE_NO As String = rowPODATA.Item("CUST_STORE_NO").ToString & String.Empty
                'UPC & SKU
                Dim sd As New Text.StringBuilder With {.Length = 0}
                sd.AppendLine("SELECT")
                sd.AppendLine("MAX(O2.CUST_UPC) AS UPC,")
                sd.AppendLine("MAX(O2.CUST_SKU) AS CUST_SKU")
                sd.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
                sd.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
                sd.AppendLine($"AND O1.ORDR_CUST_PO = '{POPO}'")
                sd.AppendLine($"AND O2.STYLE_CODE = '{STYLE_CODE}'")
                sd.AppendLine($"AND O2.COLOR_CODE = '{COLOR_CODE}'")
                sd.AppendLine("AND O1.CUST_CODE IN ('WALMART','WALMARTCOM')")
                Dim tblSKUUPC As DataTable = ASCDATA1.GetDataTable(sd.ToString(), String.Empty)
                If tblSKUUPC.Rows.Count = 1 Then
                    rowPODATA.Item("UPC") = tblSKUUPC.Rows(0).Item("UPC").ToString & String.Empty
                    rowPODATA.Item("CUST_SKU") = tblSKUUPC.Rows(0).Item("CUST_SKU").ToString & String.Empty
                End If

                'First / Last ship Data
                sd.Length = 0
                sd.AppendLine("SELECT MIN(INV_NO) AS MIN_INV_NO, MAX(INV_NO) AS MAX_INV_NO")
                sd.AppendLine("FROM SOTINVH2 I2")
                sd.AppendLine($"WHERE I2.STYLE_CODE = '{STYLE_CODE}'")
                sd.AppendLine($"AND I2.COLOR_CODE = '{COLOR_CODE}'")
                sd.AppendLine("AND I2.CUST_CODE IN ('WALMART','WALMARTCOM')")
                Dim tblFSTLST As DataTable = ASCDATA1.GetDataTable(sd.ToString(), String.Empty)
                If tblFSTLST.Rows.Count = 1 Then
                    Dim FST_INV_SHIPPED As String = tblFSTLST.Rows(0).Item("MIN_INV_NO").ToString & String.Empty
                    Dim LAST_INV_SHIPPED As String = tblFSTLST.Rows(0).Item("MAX_INV_NO").ToString & String.Empty
                    If FST_INV_SHIPPED.Length > 0 Then
                        rowPODATA.Item("FST_INV_SHIPPED") = FST_INV_SHIPPED
                        sd.Length = 0
                        sd.AppendLine("SELECT ORDR_CUST_PO, INV_DATE")
                        sd.AppendLine("FROM SOTINVH1")
                        sd.AppendLine($"WHERE INV_NO = '{FST_INV_SHIPPED}'")
                        Dim tblFST As DataTable = ASCDATA1.GetDataTable(sd.ToString(), String.Empty)
                        If tblFST.Rows.Count = 1 Then
                            rowPODATA.Item("FST_DT_SHIPPED") = tblFST.Rows(0).Item("INV_DATE").ToString & String.Empty
                            rowPODATA.Item("FST_PO_SHIPPED") = tblFST.Rows(0).Item("ORDR_CUST_PO").ToString & String.Empty
                        End If
                    End If
                    If LAST_INV_SHIPPED.Length > 0 Then
                        rowPODATA.Item("LAST_INV_SHIPPED") = LAST_INV_SHIPPED
                        sd.Length = 0
                        sd.AppendLine("SELECT ORDR_CUST_PO, INV_DATE")
                        sd.AppendLine("FROM SOTINVH1")
                        sd.AppendLine($"WHERE INV_NO = '{LAST_INV_SHIPPED}'")
                        Dim tblLST As DataTable = ASCDATA1.GetDataTable(sd.ToString(), String.Empty)
                        If tblLST.Rows.Count = 1 Then
                            rowPODATA.Item("LAST_DT_SHIPPED") = tblLST.Rows(0).Item("INV_DATE").ToString & String.Empty
                            rowPODATA.Item("LAST_PO_SHIPPED") = tblLST.Rows(0).Item("ORDR_CUST_PO").ToString & String.Empty
                        End If
                    End If
                    'TOTAL Shipped
                    sd.Length = 0
                    sd.AppendLine("SELECT")
                    sd.AppendLine("SUM(ORDR_QTY_SHIP) AS ORDR_QTY_SHIP")
                    sd.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
                    sd.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
                    sd.AppendLine("AND I1.INV_NO = I2.INV_NO")
                    sd.AppendLine("AND I1.CUST_CODE IN ('WALMART','WALMARTCOM')")
                    sd.AppendLine("AND I1.ORDR_YYYYPP_UPDATED >= '201901'")
                    sd.AppendLine($"AND I1.CUST_STORE_NO = '{CUST_STORE_NO}'")
                    sd.AppendLine($"AND I2.STYLE_CODE  = '{STYLE_CODE}'")
                    sd.AppendLine($"AND I2.COLOR_CODE = '{COLOR_CODE}'")
                    ASCMAIN1.sql = sd.ToString()
                    Dim SHIP_TOTAL As Int64 = Val(ASCDATA1.GetDataValue)
                    rowPODATA.Item("TOTAL_SHIPPED") = SHIP_TOTAL
                End If
            Next
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
            MsgBox("Data Fetch Complete", vbOKOnly, "Complete")
        Else
            MsgBox(err.ToString(), vbOKOnly, "Please Fix The Following.")
        End If
    End Sub

    Private Function ComaToInStr(ByVal InStr As String, Optional PadZeros As Boolean = False) As String
        Dim Retval As String = ""
        If InStr.EndsWith(",") Then
            InStr = InStr.Substring(0, InStr.Length - 1)
        End If
        If InStr.StartsWith(",") Then
            InStr = InStr.Substring(1, InStr.Length - 1)
        End If
        InStr = InStr.Replace(",", "','")
        If InStr.Length > 0 Then
            If PadZeros Then
                Dim padSTR As String() = Split(InStr, "','")
                InStr = ""
                For Each pd As String In padSTR
                    If InStr.Length = 0 Then
                        InStr = pd.PadLeft(6, "0")
                    Else
                        InStr = InStr & "','" & pd.PadLeft(6, "0")
                    End If
                Next
            End If
            InStr = "'" & InStr & "'"
        End If
        Retval = InStr
        Return Retval
    End Function

    Private Sub btnPasteStyles_Click(sender As Object, e As EventArgs)
        PasteData(1)
    End Sub

    Private Sub btnPasteSKUs_Click(sender As Object, e As EventArgs)
        If txtPOPO_1.Text.Length = 0 Then
            MsgBox("You Must Provide A PO When Pasting SKUs", vbOKOnly, "PO Required")
        Else
            PasteData(2)
        End If
    End Sub

    Private Sub btnPasteStores_Click(sender As Object, e As EventArgs) Handles btnPasteStores_1.Click, btnPasteStores_2.Click, btnPasteStores_3.Click,
         btnPasteStores_4.Click, btnPasteStores_5.Click, btnPasteStores_6.Click
        Dim bNum As String = sender.name.ToString.Substring(sender.name.length - 1, 1)
        Dim storetxt As Int64 = 1
        If IsNumeric(bNum) Then
            storetxt = Val(bNum)
        End If
        PasteData(3, storetxt)
    End Sub

    Private Sub PasteData(ByVal TYPE As Integer, Optional Storetxt As Int64 = 1)
        '1 = Styles
        '2 = SKUS
        '3 = Stores
        '4 = Stores2
        Dim frmASFMSGBF As New ASFMSGBF
        Dim Results As String = frmASFMSGBF.Get_txtblock_from_User("Please Paste Your Data", "Paste Data", "", False)
        If Results.Length > 0 Then
            Dim ResultList As String() = Split(Results, vbCrLf)
            If ResultList.Length > 0 Then
                Dim NEWDATA As String = ""
                Dim DUPES As New List(Of String)
                For Each Str As String In ResultList
                    If TYPE = 3 Or TYPE = 4 Then
                        If Str.Length > 0 Then
                            Str = Str.PadLeft(6, "0")
                        End If
                    End If
                    If Not DUPES.Contains(Str) Then
                        If Str.Length > 0 Then
                            If TYPE = 2 Then
                                Str = SKU2STYLE(Str)
                            End If
                            DUPES.Add(Str)
                            NEWDATA = NEWDATA & "," & Str
                        End If
                    End If
                Next
                If NEWDATA.Length >= 3 Then
                    NEWDATA = NEWDATA.Substring(1, NEWDATA.Length - 1)
                End If
                Select Case TYPE
                    Case 1
                        'txtPOStyles.Text = NEWDATA
                    Case 2
                        'txtPOStyles.Text = NEWDATA
                    Case 3
                        Select Case Storetxt
                            Case 1
                                txtPOStores_1.Text = NEWDATA
                            Case 2
                                txtPOStores_2.Text = NEWDATA
                            Case 3
                                txtPOStores_3.Text = NEWDATA
                            Case 4
                                txtPOStores_4.Text = NEWDATA
                            Case 5
                                txtPOStores_5.Text = NEWDATA
                            Case 6
                                txtPOStores_6.Text = NEWDATA
                        End Select

                    Case 4
                        txtPOStores2.Text = NEWDATA
                End Select

            End If
        End If

    End Sub

    Private Function SKU2STYLE(ByVal SKU As String) As String
        Dim Retval As String = SKU
        Dim PO As String = txtPOPO_1.Text.ToString & String.Empty
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("MAX(S2.STYLE_CODE) AS STYLE_CODE")
        SQLS.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2")
        SQLS.AppendLine("WHERE S1.ORDR_NO = S2. ORDR_NO")
        SQLS.AppendLine($"AND S2.CUST_SKU = '{SKU}'")
        SQLS.AppendLine($"AND S1.ORDR_CUST_PO = '{PO}'")
        ASCMAIN1.sql = SQLS.ToString()
        Dim STYLE_CODE As String = ASCDATA1.GetDataValue
        If STYLE_CODE.Length > 0 Then
            Retval = STYLE_CODE
        End If
        Return Retval
    End Function

    Private Sub btnSTOREPO1_Click_1(sender As Object, e As EventArgs) Handles btnSTOREPO1.Click
        If (txtPOSTORE.Text.ToString & String.Empty).Length > 0 Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Getting Data", "")
            Application.DoEvents()

            Dim PO As String = txtPOSTORE.Text.ToString & String.Empty
            Dim STORES As String = ComaToInStr(txtPOStores2.Text.ToString, True)

            Dim SQ As New StringBuilder With {.Length = 0}
            SQ.AppendLine("SELECT")
            SQ.AppendLine("O1.PO_NUMBER,")
            SQ.AppendLine("O1.CUST_STORE_NO,")
            SQ.AppendLine("O1.PO_STATUS,")
            SQ.AppendLine("O1.ST_EA_ORDR,")
            SQ.AppendLine("O1.ST_EACH_RCD,")
            SQ.AppendLine("O2.ORDR_QTY_SHIP,")
            SQ.AppendLine("O1.EXCEL_FILE,")
            SQ.AppendLine("O1.EXCEL_LINE")
            SQ.AppendLine("FROM SOTWMPO1 O1, SOTWMPO2 O2")
            SQ.AppendLine("WHERE O1.PO_NUMBER = O2.ORDR_CUST_PO")
            SQ.AppendLine("AND O1.CUST_STORE_NO = O2.CUST_STORE_NO")
            SQ.AppendLine($"AND O1.PO_NUMBER = '{PO}'")
            If STORES.Length > 0 Then
                SQ.AppendLine($"AND O1.CUST_STORE_NO IN ({STORES})")
            End If

            Fill_Records("STOREPO1",,, SQ.ToString)
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
            Application.DoEvents()
        End If
    End Sub

    Private Sub btnSOTWMPO2R_Click(sender As Object, e As EventArgs) Handles btnSOTWMPO2R.Click
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Refreshing Data", "")
        Application.DoEvents()

        Dim LMP As String = ASCMAIN1.Get_YYYYMM(ASCMAIN1.CYP, -1)
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine($"DELETE FROM SOTWMPO2 WHERE ORDR_YYYYPP_UPDATED >= '{LMP}'")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        SQLS.Length = 0
        SQLS.AppendLine("INSERT INTO SOTWMPO2")
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("I1.ORDR_CUST_PO, I1.CUST_STORE_NO, I1.ORDR_YYYYPP_UPDATED,")
        SQLS.AppendLine("SUM(NVL(I2.ORDR_QTY_SHIP,0)) AS ORDR_QTY_SHIP")
        SQLS.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
        SQLS.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
        SQLS.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
        SQLS.AppendLine("AND I1.CUST_CODE = 'WALMART'")
        SQLS.AppendLine("AND NVL(I2.ORDR_QTY_SHIP,0) > 0")
        SQLS.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED >= '{LMP}'")
        SQLS.AppendLine("GROUP BY I1.ORDR_CUST_PO, I1.CUST_STORE_NO, I1.ORDR_YYYYPP_UPDATED")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        Application.DoEvents()
        MsgBox("Data Refreshed", vbOKOnly, "Done")

    End Sub

    Private Sub btnSTOREPOX_Click(sender As Object, e As EventArgs) Handles btnSTOREPOX.Click
        dst.Tables.Item("STOREPO1").Clear()
    End Sub

    Private Sub btnPasteStores2_Click(sender As Object, e As EventArgs) Handles btnPasteStores2.Click
        PasteData(4)
    End Sub


#End Region

End Class