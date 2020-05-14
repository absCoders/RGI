Public Class ICFLOTH1
    Dim ICTSTDL1 As String
    Dim WHSE_CODE As String
    Dim LOT_NO As String
    Dim LOT_SEQ_NO As Int64
    Dim PO_ORDER_NO As String
    Dim PACK_FACTOR As Decimal
    Dim rowICTLOTD1 As DataRow
    Dim rowICTCOSTZ As DataRow
    Dim rowICTIREC2 As DataRow
    Dim rowICTIRECX As DataRow
    Dim sqlICTLOTD1 As String

    Dim alreadyAuth As Boolean = False
    Dim adjustment_password_failures As Integer = 0

    ' RESERVE PRICE MAINTENANCE NOT SUPPORTED
    ' WRITE DOWN PRICE MAINT NOT SUPPORTED
    ' NO LONGER PERMITTING CHANGES TO SPG, CON/REG, OTHER COSTS

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")

        With dst
            'Create_TDA(.Tables.Add, "ASTAUDT1", "*")
            Create_TDA(.Tables.Add, "ICTIADJ1", "*")
            Create_TDA(.Tables.Add, "ICTLOTH1", "*")
            Create_TDA(.Tables.Add, "ICTLOTDW", "*")

            sqlICTLOTD1 = "Select ICTLOTD1.*,ICTPACK1.PACK_FACTOR,POTORDR1.PO_ORDER_NO,POTORDR1.MSC_NO " _
            & ", ICTPROD1.PROD_DESC" _
            & ", ICTSIZE1.SIZE_DESC" _
            & ", ICTPACK1.PACK_DESC" _
            & " from ICTLOTD1 ICTLOTD1,ICTPACK1,POTORDR1,ICTPROD1,ICTSIZE1" _
            & " where ICTPACK1.PACK_CODE = ICTLOTD1.PACK_CODE" _
            & "   and ICTPROD1.PROD_CODE = ICTLOTD1.PROD_CODE" _
            & "   and ICTSIZE1.SIZE_CODE = ICTLOTD1.SIZE_CODE" _
            & "   and POTORDR1.IMPORT_NO (+) = ICTLOTD1.IMPORT_NO"

            ASCMAIN1.sql = sqlICTLOTD1 _
            & " and ICTLOTD1.WHSE_CODE = :PARM1" _
            & " and ICTLOTD1.LOT_NO = :PARM2" _
            & " and ICTLOTD1.LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTLOTD1", "**", 0, False, "VVN", 3)
            Create_TDA(.Tables.Add, "ICTLOTDF", "**", 0, False, "VVN", 3)
            Create_TDA(.Tables.Add, "ICTLOTDV", "**", 0, False, "VVN", 3)
            For Each TABLE_NAME As String In New String() {"ICTLOTD1", "ICTLOTDF", "ICTLOTDV"}
                .Tables(TABLE_NAME).Columns.Add("QTY_AVAIL", GetType(System.Int64), "ISNULL(QTY_ON_HAND,0) - ISNULL(QTY_COMMITTED,0) - ISNULL(PROD_MGR_HOLD,0)")
            Next

            .Tables("ICTLOTDF").Columns.Add("RELATIONSHIP")

            Create_TDA(.Tables.Add("ICTLOTD1_ADJ"), "ICTLOTD1", "*", , , , , "QTY_ON_HAND,COOL_COMPLIANT,LOT_EXP_DATE,DATE_WHSE_ANNIV,DATE_LAST_TRAN,STANDARD_COST,ADJUSTED_COST")

            Dim rowICTLOTD1 As DataRow = ASCDATA1.GetDataRow _
                (ASCMAIN1.sql, , "VVN", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})

            ASCMAIN1.sql = "Select ICTIREC1.RECEIPT_NO, ICTIREC1.PO_ORDER_NO, ICTIREC1.VEND_CODE, " & vbCrLf _
            & " ICTIREC1.CON_REG_IND, ICTIREC1.CA_COMM_PCT, ICTIREC1.CA_VA_COMM_PCT, " & vbCrLf _
            & " ICTIREC1.CA_EXP_RECOVERY, ICTIREC1.CA_NO_OCEAN_FRT_RECOVER, ICTIREC1.CA_EXP_CAP_PCT, ICTIREC1.CA_ADV_PCT_SP, " & vbCrLf _
            & " ICTIREC1.CA_AMT_LB, ICTIREC1.JOINT_VENTURE, ICTIREC1.JOINT_VENTURE_PCT, " & vbCrLf _
            & " ICTIREC1.JOINT_VENTURE_PARTNER, ICTIREC1.INT_START_DATE, ICTIREC1.INT_PCT," & vbCrLf _
            & " ICTIREC2.REC_CASES, ICTIREC2.REC_UNITS,  " & vbCrLf _
            & " ICTIREC2.PURCHASE_COST, ICTIREC2.VALUATION_COST, ICTIREC2.ADJUSTED_COST, " & vbCrLf _
            & " ICTIREC2.DATE_WHSE_ANNIV, ICTIREC2.CA_VALUE_ADD_IND, ICTCOSTZ.OCEAN_FRT   " & vbCrLf _
            & " From ICTIREC2, ICTIREC1, ICTCOSTZ" & vbCrLf _
            & " WHERE ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "   and ICTCOSTZ.WHSE_CODE (+) = ICTIREC2.WHSE_CODE " & vbCrLf _
            & "   and ICTCOSTZ.LOT_NO (+) = ICTIREC2.LOT_NO " & vbCrLf _
            & "   and ICTCOSTZ.LOT_SEQ_NO (+) = ICTIREC2.LOT_SEQ_NO " & vbCrLf _
            & "   and ICTIREC2.WHSE_CODE = :PARM1 and ICTIREC2.LOT_NO = :PARM2 and ICTIREC2.LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTIRECX", "**", 0, False, "VVN", 0)
            .Tables("ICTIRECX").Columns.Add("SO_ORDER_NO")
            .Tables("ICTIRECX").Columns.Add("ORDR_INV_NO")

            ASCMAIN1.sql = "Select S3.WHSE_CODE, S3.LOT_NO, S3.LOT_SEQ_NO, S1.ORDR_INV_DATE TRAN_DATE" & vbCrLf _
            & ", S1.ORDR_INV_NO TRAN_NO, DECODE(S1.ORDR_TYPE_CODE,'T','X',S1.ORDR_TYPE_CODE) TRAN_TYPE" & vbCrLf _
            & ", S3.SO_LOT_CASES CASES, S3.SO_LOT_UNITS UNITS" & vbCrLf _
            & ", -1 * S3.SO_LOT_CASES CASES_INV, -1 * S3.SO_LOT_UNITS UNITS_INV" & vbCrLf _
            & ", S2.ORDR_PRICE_GRS, S2.ORDR_PRICE_NET, S2.REBATE, S2.FUND_RATE" & vbCrLf _
            & ", S2.BRKR_RATE, S1.FRT_RATE, S2.ALLOW_RATE, S2.SVC_CHG_RATE" & vbCrLf _
            & ", DECODE(S1.ORDR_TYPE_CODE,'X',S3.WHSE_CODE,NULL) X_WHSE_CODE" & vbCrLf _
            & ", DECODE(S1.ORDR_TYPE_CODE,'X',S3.LOT_NO,NULL) X_LOT_NO" & vbCrLf _
            & ", DECODE(S1.ORDR_TYPE_CODE,'X',S3.LOT_SEQ_NO,NULL) X_LOT_SEQ_NO" & vbCrLf _
            & ", DECODE(S1.ORDR_TYPE_CODE,'X','X',NULL) LOT_XY" & vbCrLf _
            & ", S1.CUST_CODE, S1.CUST_NAME, S1.CUST_ORDER_NO, S1.SO_ORDER_NO" & vbCrLf _
            & ", S3.ADJ_COST_EXT, S3.CGS_COST_EXT, S3.STD_COST_EXT, S3.CON_REG_IND" & vbCrLf _
            & " from SOTINVH3 S3,SOTINVH2 S2,SOTINVH1 S1" & vbCrLf _
            & " where S1.SO_ORDER_NO = S3.SO_ORDER_NO" & vbCrLf _
            & "   and S2.SO_ORDER_NO = S3.SO_ORDER_NO" & vbCrLf _
            & "   and S2.SO_ORDER_LNO = S3.SO_ORDER_LNO" & vbCrLf _
            & "   and S3.WHSE_CODE = :PARM1 and S3.LOT_NO = :PARM2 and S3.LOT_SEQ_NO = :PARM3" & vbCrLf _
            & " UNION " & vbCrLf _
            & " Select S6.ORIG_WHSE_CODE, S6.ORIG_LOT_NO, S6.ORIG_LOT_SEQ_NO, S1.ORDR_INV_DATE TRAN_DATE" & vbCrLf _
            & ", DECODE(S7.CLAIM_IND,'1',S1.CLAIM_NO,S1.ORDR_INV_NO) TRAN_NO" & vbCrLf _
            & ", DECODE(S7.CLAIM_IND,'1','D',S1.ORDR_TYPE_CODE) TRAN_TYPE" & vbCrLf _
            & ", S7.CHG_CASES CASES, S7.CHG_UNITS UNITS" & vbCrLf _
            & ", S7.INV_CASES CASES_INV, S7.INV_UNITS UNITS_INV" & vbCrLf _
            & ", S7.ORDR_PRICE_GRS, S7.ORDR_PRICE_NET, S7.REBATE, S7.FUND_RATE" & vbCrLf _
            & ", S7.BRKR_RATE, S1.FRT_RATE, S7.ALLOW_RATE, S7.SVC_CHG_RATE" & vbCrLf _
            & ", S6.ORIG_WHSE_CODE X_WHSE_CODE" & vbCrLf _
            & ", S6.ORIG_LOT_NO X_LOT_NO" & vbCrLf _
            & ", S6.ORIG_LOT_SEQ_NO X_LOT_SEQ_NO" & vbCrLf _
            & ", 'X' LOT_XY" & vbCrLf _
            & ", S1.CUST_CODE, S1.CUST_NAME, S1.CUST_ORDER_NO, S1.SO_ORDER_NO" & vbCrLf _
            & ", S7.ADJ_COST_EXT, S7.CGS_COST_EXT, S7.STD_COST_EXT, S7.CON_REG_IND" & vbCrLf _
            & " from SOTINVH7 S7,SOTINVH6 S6,SOTINVH1 S1" & vbCrLf _
            & " where S1.SO_ORDER_NO = S7.SO_ORDER_NO" & vbCrLf _
            & "   and S6.SO_ORDER_NO = S7.SO_ORDER_NO" & vbCrLf _
            & "   and S6.SO_ORDER_LNO = S7.SO_ORDER_LNO" & vbCrLf _
            & "   and S7.WHSE_CODE = :PARM1 and S7.LOT_NO = :PARM2 and S7.LOT_SEQ_NO = :PARM3" & vbCrLf _
            & "   and (S1.ORDR_TYPE_CODE <> 'D')"
            Create_TDA(.Tables.Add, "ICTLOTHX", "**", 0, False, "VVN", 0)
            .Tables("ICTLOTHX").Columns.Add("STANDARD_COST", GetType(System.Decimal), "IIF(ISNULL(UNITS,0) = 0, 0, STD_COST_EXT / UNITS)")
            .Tables("ICTLOTHX").Columns.Add("ADJUSTED_COST", GetType(System.Decimal), "IIF(ISNULL(UNITS,0) = 0, 0, ADJ_COST_EXT / UNITS)")
            .Tables("ICTLOTHX").Columns.Add("CGS_COST", GetType(System.Decimal), "IIF(ISNULL(UNITS,0) = 0, 0, CGS_COST_EXT / UNITS)")
            .Tables("ICTLOTHX").Columns.Add("RECORD_INDEX", GetType(System.Int64))

            '& ", DECODE(S1.ORDR_TYPE_CODE,'X',S7.WHSE_CODE,NULL) X_WHSE_CODE" & vbCrLf _
            '& ", DECODE(S1.ORDR_TYPE_CODE,'X',S7.LOT_NO,NULL) X_LOT_NO" & vbCrLf _
            '& ", DECODE(S1.ORDR_TYPE_CODE,'X',S7.LOT_SEQ_NO,NULL) X_LOT_SEQ_NO" & vbCrLf _


            ASCMAIN1.sql = "SELECT * FROM ICTLOTH1" _
            & " where X_WHSE_CODE = :PARM1 and X_LOT_NO = :PARM2 and X_LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTLOTHX_XREC", "**", 0, False, "VVN", 0)
            .Tables("ICTLOTHX_XREC").Columns.Add("RECORD_INDEX", GetType(System.Int64))

            Create_Relation("ICTLOTHX", "ICTLOTHX_XREC", "RECORD_INDEX")
            '.Relations.Add("ICTLOTHX_ICTLOTHX_XREC" _
            '               , New DataColumn() {.Tables("ICTLOTHX").Columns("WHSE_CODE"), .Tables("ICTLOTHX").Columns("LOT_NO"), .Tables("ICTLOTHX").Columns("LOT_SEQ_NO"), .Tables("ICTLOTHX").Columns("TRAN_TYPE"), .Tables("ICTLOTHX").Columns("TRAN_NO")} _
            '               , New DataColumn() {.Tables("ICTLOTHX_XREC").Columns("X_WHSE_CODE"), .Tables("ICTLOTHX_XREC").Columns("X_LOT_NO"), .Tables("ICTLOTHX_XREC").Columns("X_LOT_SEQ_NO"), .Tables("ICTLOTHX_XREC").Columns("TRAN_TYPE"), .Tables("ICTLOTHX_XREC").Columns("TRAN_NO")})

            ASCMAIN1.sql = "Select SOTORDR3.SO_ORDER_NO, SOTORDR3.SO_ORDER_LNO, SOTORDR3.SO_LOT_LNO, " & vbCrLf _
            & " SOTORDR1.ORDR_DIV_CODE, " & vbCrLf _
            & " ICTLOTD2.PROD_CODE, ICTLOTD2.SIZE_CODE, ICTLOTD2.BRAND_CODE, " & vbCrLf _
            & " ICTLOTD2.SP_GROUP, ICTLOTD2.ORIG_CODE, ICTLOTD2.GRADE_CODE, " & vbCrLf _
            & " ICTLOTD2.PACK_CODE, ICTLOTD2.DIVISION_CODE, ICTLOTD2.WHSE_CODE, " & vbCrLf _
            & " ICTLOTD2.TERR_CODE, ICTLOTD2.PROD_DIV_CODE, " & vbCrLf _
            & " SOTORDR3.LOT_ORDER_QTY QTY_CASES, SOTORDR3.LOT_ORDER_QTY * SOTORDR3.PACK_FACTOR QTY_UNITS, SOTORDR2.ORDR_PRICE_GRS, " & vbCrLf _
            & " SOTORDR1.CUST_NAME, SOTORDR1.CUST_ORDER_NO, SOTORDR1.CUST_PU_DATE, " & vbCrLf _
            & " SOTORDR1.ORDR_DATE, SOTORDR1.SREP_CODE," & vbCrLf _
            & " ICTPROD1.PROD_DESC, ICTSIZE1.SIZE_DESC, " & vbCrLf _
            & " ICTPACK1.PACK_DESC, ICTWHSE1.WHSE_NAME, " & vbCrLf _
            & " SOTORDR1.ORDR_TYPE_CODE, ICTLOTD2.LOT_NO, SOTORDR3.COOL_COMPLIANT " & vbCrLf _
            & " from SOTORDR1, SOTORDR2, SOTORDR3, ICTLOTD2, ICTPROD1, ICTSIZE1, ICTPACK1, ICTWHSE1" & vbCrLf _
            & " where SOTORDR1.SO_ORDER_NO = SOTORDR2.SO_ORDER_NO " & vbCrLf _
            & "   and SOTORDR2.SO_ORDER_NO = SOTORDR3.SO_ORDER_NO " & vbCrLf _
            & "   and SOTORDR2.SO_ORDER_LNO = SOTORDR3.SO_ORDER_LNO " & vbCrLf _
            & "   and SOTORDR3.WHSE_CODE = ICTLOTD2.WHSE_CODE " & vbCrLf _
            & "   and SOTORDR3.LOT_NO = ICTLOTD2.LOT_NO " & vbCrLf _
            & "   and SOTORDR3.LOT_SEQ_NO = ICTLOTD2.LOT_SEQ_NO " & vbCrLf _
            & "   and ICTPROD1.PROD_CODE = ICTLOTD2.PROD_CODE" & vbCrLf _
            & "   and ICTSIZE1.SIZE_CODE = ICTLOTD2.SIZE_CODE" & vbCrLf _
            & "   and ICTPACK1.PACK_CODE = ICTLOTD2.PACK_CODE" & vbCrLf _
            & "   and ICTWHSE1.WHSE_CODE = ICTLOTD2.WHSE_CODE" & vbCrLf _
            & "   and SOTORDR3.WHSE_CODE = :PARM1 and SOTORDR3.LOT_NO = :PARM2 and SOTORDR3.LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "VVN", 3)

            ASCMAIN1.sql = "Select VEND_CODE, COST_CATGY_CODE, '2' STAGE, INIT_DATE TRAN_DATE, TOTAL_SHP, TOTAL_INV" & vbCrLf _
            & ", TOTAL_INV - TOTAL_SHP TOTAL_VAR, UNITS, (TOTAL_INV - TOTAL_SHP) / UNITS VARIANCE" & vbCrLf _
            & ", CTL_NO, INIT_DATE, INIT_OPER, LAST_DATE, LAST_OPER, CTL_NO VOUCHER_NO" & vbCrLf _
            & " from POTTRAN0, " & vbCrLf _
            & " (SELECT SUM (REC_UNITS) UNITS FROM ICTIREC2 WHERE PO_ORDER_NO = '') X WHERE ROWNUM < 1"
            ICTSTDL1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("ALTER TABLE " & ICTSTDL1 & " MODIFY TOTAL_VAR NUMBER (13,6)")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & ICTSTDL1 & " MODIFY UNITS NUMBER (13,6)")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & ICTSTDL1 & " MODIFY VARIANCE NUMBER (13,6)")
            ASCMAIN1.sql = "Select * from " & ICTSTDL1
            Create_TDA(.Tables.Add, "ICTSTDL1", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select ICTLOTD5.* from ICTLOTD5" & vbCrLf _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTLOTD5", "**", 0, False, "VVN")

            ASCMAIN1.sql = "Select ICTLOTD3.* from ICTLOTD3" & vbCrLf _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTLOTD3", "**", 0, False, "VVN")
            .Tables("ICTLOTD3").Columns.Add("YP_LEGEND")
            .Tables("ICTLOTD3").Columns.Add("AVG_CASES", GetType(System.Int32))
            .Tables("ICTLOTD3").Columns.Add("AVG_UNITS", GetType(System.Decimal))

            ASCMAIN1.sql = "Select ICTLOTD6.TRAN_DATE, ICTLOTD3.* from ICTLOTD3,ICTLOTD6" _
            & " where ICTLOTD3.WHSE_CODE = :PARM1 and ICTLOTD3.LOT_NO = :PARM2 and ICTLOTD3.LOT_SEQ_NO = :PARM3" _
            & " and ICTLOTD6.OPS_YYYYPP = ICTLOTD3.OPS_YYYYPP" _
            & " and ICTLOTD6.WHSE_CODE = ICTLOTD3.WHSE_CODE" _
            & " and ICTLOTD6.LOT_NO = ICTLOTD3.LOT_NO" _
            & " and ICTLOTD6.LOT_SEQ_NO = ICTLOTD3.LOT_SEQ_NO"
            Create_TDA(.Tables.Add, "ICTLOTDD", "**", 0, False, "VVN")
            ' .Tables("ICTLOTDD").Columns.Add("YP_LEGEND")

            ASCMAIN1.sql = "Select ICTLOTDX.* from ICTLOTDX" & vbCrLf _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTLOTDX", "**", 0, False, "VVN", 0)
            .Tables("ICTLOTDX").Columns.Add("INV_CASES", GetType(System.Int64))
            .Tables("ICTLOTDX").Columns.Add("INV_UNITS", GetType(System.Decimal))
            .Tables("ICTLOTDX").Columns.Add("INV_AMT", GetType(System.Decimal))
            .Tables("ICTLOTDX").Columns.Add("SORTING_UNIT", GetType(System.Decimal), "IIF(ISNULL(UNITS,0)=0,0,SORTING/UNITS)")
            .Tables("ICTLOTDX").Columns.Add("HANDLING_UNIT", GetType(System.Decimal), "IIF(ISNULL(UNITS,0)=0,0,HANDLING/UNITS)")

            ASCMAIN1.sql = "Select ICTLOTDG.* from ICTLOTDG" & vbCrLf _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTLOTDG", "**", 0, False, "VVN")

            ASCMAIN1.sql = "Select ASTATTA2.* from ASTATTA2" & vbCrLf _
            & " where TABLE_NAME = :PARM1 and COLUMN_NAME = :PARM2 and CODE_VALUE = :PARM3" _
            & " and NVL(ATTACHMENT_STATUS,'O') <> 'D'"
            Create_TDA(.Tables.Add, "ASTATTA2", "**", 0, False, "VVV")

            ASCMAIN1.sql = "Select ICTLOTD2.WHSE_CODE, ICTLOTD2.LOT_NO, ICTLOTD2.LOT_SEQ_NO, ICTLOTD2.PROD_CODE, ICTLOTD2.SIZE_CODE, ICTLOTD2.ORIG_CODE, ICTLOTD2.BRAND_CODE, ICTLOTD2.PACK_CODE, ICTLOTD2.SP_GROUP, ICTLOTD2.GRADE_CODE, ICTLOTD2.QTY_ON_HAND, ICTLOTD2.QTY_COMMITTED, ICTLOTD2.PROD_MGR_HOLD" _
            & ", NVL(ICTLOTD2.QTY_ON_HAND,0) - NVL(ICTLOTD2.QTY_COMMITTED,0) - NVL(ICTLOTD2.PROD_MGR_HOLD,0) QTY_AVAIL" _
            & " from ICTLOTD2 where (WHSE_CODE, LOT_NO) in " _
            & " (Select Distinct WHSE_CODE, LOT_NO from ICTLOTD2 where QTY_ON_HAND < 0)"
            Create_TDA(.Tables.Add, "ICTLOTN1", "**", 0, False)

            ASCMAIN1.sql = "SELECT * FROM (SELECT ICTLOTD2.WHSE_CODE, ICTLOTD2.LOT_NO, ICTLOTD2.LOT_SEQ_NO" & vbCrLf _
            & ", ICTLOTD2.PROD_CODE, ICTLOTD2.SIZE_CODE, ICTLOTD2.ORIG_CODE, ICTLOTD2.BRAND_CODE" & vbCrLf _
            & ", ICTLOTD2.PACK_CODE, ICTLOTD2.SP_GROUP, ICTLOTD2.GRADE_CODE" & vbCrLf _
            & ", ICTLOTD2.QTY_COMMITTED LOT_QTY_COMM, SUM (SOTORDR3.SO_LOT_CASES) ORDER_QTY_COMM" & vbCrLf _
            & " from ICTLOTD1 ICTLOTD2, SOTORDR3 " & vbCrLf _
            & " where SOTORDR3.WHSE_CODE (+) = ICTLOTD2.WHSE_CODE" & vbCrLf _
            & "   and SOTORDR3.LOT_NO (+) = ICTLOTD2.LOT_NO" & vbCrLf _
            & "   and SOTORDR3.LOT_SEQ_NO (+) = ICTLOTD2.LOT_SEQ_NO" & vbCrLf _
            & " group by ICTLOTD2.WHSE_CODE, ICTLOTD2.LOT_NO, ICTLOTD2.LOT_SEQ_NO" & vbCrLf _
            & ", ICTLOTD2.PROD_CODE, ICTLOTD2.SIZE_CODE, ICTLOTD2.ORIG_CODE, ICTLOTD2.BRAND_CODE" & vbCrLf _
            & ", ICTLOTD2.PACK_CODE, ICTLOTD2.SP_GROUP, ICTLOTD2.GRADE_CODE, ICTLOTD2.QTY_COMMITTED)" & vbCrLf _
            & "WHERE NVL(LOT_QTY_COMM,0) <> NVL(ORDER_QTY_COMM,0) "
            Create_TDA(.Tables.Add, "ICTLOTO1", "**", 0, False)

            ASCMAIN1.sql = "Select EDT810IA.* from EDT810IA" & vbCrLf _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "EDT810IA", "**", 0, False, "VVN")

            ASCMAIN1.sql = "Select APTINVH7.* from APTINVH7" & vbCrLf _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "APTINVH7", "**", 0, False, "VVN")

            ASCMAIN1.sql = "Select POTTRAN0.* from POTTRAN0" & vbCrLf _
            & " where PO_ORDER_NO = :PARM1 and CTL_STATUS IN ('0','1')"
            Create_TDA(.Tables.Add, "POTTRAN0", "**", 0, False, "V")


            With .Tables.Add("ICTCOSTZ_LIST")
                .Columns.Add("COST_DESC")
                .Columns.Add("COST_VALUE", GetType(System.Decimal))
            End With


            'ASCMAIN1.sql = "SELECT X.ITEM_NUMBER, X.LOCATION, X.LOT_NUMBER_OR_SIZE, X.ON_HAND_IN_UM, X.RECEIPT_DATE" _
            '& ", X.RECEIVED_COST, X.LAST_COST, I.ACTUAL_COST, X.WEIGHT_RECEIVED" _
            '& ", X.LOT_VENDOR, X.PO_NUMBER_LINE_KEY, X.ACRONYM" _
            '& " FROM COSFF.InventoryLocationFILE X,CODEXREF CODEXREF_WHSE_CODE" _
            '& ",EMP.ICTCOSFO" _
            '& ",ITEMXREF,LOTNXREF,CODEXREF CODEXREF_X_WHSE_CODE, LOTNXREF LOTNXREFX" _
            '& ",COSFF.INVENTORYMASTERFILE I" _
            '& ", (SELECT DISTINCT DOCUMENT_NUMBER, ITEM_NUMBER, RECEIPT_PRICE, PRICING_UNIT" _
            '& " from COSFF.POapinterfacelineshistory) AP" _
            '& " WHERE CODEXREF_WHSE_CODE.TABLE_NAME (+) = 'ICTWHSE1'" _
            '& " AND CODEXREF_WHSE_CODE.COSFF_KEY (+) = TRIM(X.LOCATION)" _
            '& " AND ITEMXREF.ITEM_NUMBER (+) = X.ITEM_NUMBER" _
            '& " AND LOTNXREF.ITEM_NUMBER (+) = TRIM(X.ITEM_NUMBER)" _
            '& " AND LOTNXREF.LOCATION (+) = NVL(TRIM(X.LOCATION),'X')" _
            '& " AND LOTNXREF.LOT_NUMBER_OR_SIZE (+) = TRIM(X.LOT_NUMBER_OR_SIZE)" _
            '& " AND LOTNXREFX.ITEM_NUMBER (+) = CASE WHEN (X.TRANSFER_FROM_LOT IS NOT NULL OR X.TRANSFER_FROM_ITEM IS NOT NULL OR X.TRANSFER_FROM_LOCATION IS NOT NULL) THEN NVL(TRIM(X.TRANSFER_FROM_ITEM),TRIM(X.ITEM_NUMBER)) ELSE NULL END" _
            '& " AND LOTNXREFX.LOCATION (+) = NVL(CASE WHEN (X.TRANSFER_FROM_LOT IS NOT NULL OR X.TRANSFER_FROM_ITEM IS NOT NULL OR X.TRANSFER_FROM_LOCATION IS NOT NULL) THEN NVL(TRIM(X.TRANSFER_FROM_LOCATION),TRIM(X.LOCATION)) ELSE NULL END,'X')" _
            '& " AND LOTNXREFX.LOT_NUMBER_OR_SIZE (+) = CASE WHEN (X.TRANSFER_FROM_LOT IS NOT NULL OR X.TRANSFER_FROM_ITEM IS NOT NULL OR X.TRANSFER_FROM_LOCATION IS NOT NULL) THEN NVL(TRIM(X.TRANSFER_FROM_LOT),TRIM(X.LOT_NUMBER_OR_SIZE)) ELSE NULL END" _
            '& " AND CODEXREF_X_WHSE_CODE.TABLE_NAME (+) = 'ICTWHSE1'" _
            '& " AND CODEXREF_X_WHSE_CODE.COSFF_KEY (+) = TRIM(X.TRANSFER_FROM_LOCATION)" _
            '& " AND I.ITEM_NUMBER (+) = X.ITEM_NUMBER" _
            '& " AND ICTCOSFO.ORIGX (+) = TRIM(NVL(X.ORIGIN_CODE1,NVL(X.ORIGIN_CODE2,NVL(X.ORIGIN_CODE3,NVL(X.ORIGIN_CODE4,NVL(X.ORIGIN_CODE5,'?'))))))" _
            '& " AND TRIM(X.LOT_NUMBER_OR_SIZE) IS NOT NULL AND trim(X.location) is not null" _
            '& " AND TRIM(AP.DOCUMENT_NUMBER(+)) = SUBSTR(X.PO_NUMBER_LINE_KEY,1,6)" _
            '& " AND TRIM(AP.ITEM_NUMBER(+)) = TRIM(x.ITEM_NUMBER)" _
            '& " AND AP.RECEIPT_PRICE (+) = x.RECEIVED_COST" _
            '& " AND NVL(LOTNXREF.LOT_NO,'NOLOTX') = :PARM1"
            'Create_TDA(.Tables.Add, "SS", "**", 0, False, "V")

            ASCMAIN1.sql = "SELECT * FROM ICTLOTDI" _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"
            Create_TDA(.Tables.Add, "ICTLOTDI", "**", 0, False, "VVN", 0)
            .Tables("ICTLOTDI").Columns.Add("CUM_INTEREST", GetType(System.Decimal))
            .Tables("ICTLOTDI").Columns.Add("NO", GetType(System.Int32))
            .Tables("ICTLOTDI").Columns.Add("DAYS", GetType(System.Int32))
        End With

        grdSS.DataSource = dst.Tables("SS")
        If ASCMAIN1.DBS_COMPANY <> "COS" And ASCMAIN1.DBS_COMPANY <> "COT" Then
            tabLotDetails.Tabs("Conversion").Visible = False
        End If

        grdICTLOTDF.DataSource = dst.Tables("ICTLOTDF")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdICTLOTHX.DataSource = dst.Tables("ICTLOTHX")
        grdICTLOTDI.DataSource = dst.Tables("ICTLOTDI")

        grdICTLOTDX.DataSource = dst.Tables("ICTLOTDX")
        grdICTLOTD3.DataSource = dst.Tables("ICTLOTD3")
        grdICTLOTD5.DataSource = dst.Tables("ICTLOTD5")
        grdICTLOTDD.DataSource = dst.Tables("ICTLOTDD")
        grdICTSTDL1.DataSource = dst.Tables("ICTSTDL1")
        grdICTCOSTZ_LIST.DataSource = dst.Tables("ICTCOSTZ_LIST")

        With grdICTCOSTZ_LIST.DisplayLayout.Bands("ICTCOSTZ_LIST")
            .Columns("COST_VALUE").Format = "###.0000"
        End With

        With grdICTLOTD5.DisplayLayout.Bands(0)
            .Columns("PURCHASE_COST").Format = "###.0000"
            .Columns("VALUATION_COST").Format = "###.0000"
            .Columns("ADJUSTED_COST").Format = "###.0000"
            .Columns("STANDARD_COST").Format = "###.0000"
        End With

        grdICTLOTHX.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay

        '        grdICTLOTDG.DataSource = dst.Tables("ICTLOTDG")
        grdASTATTA2.DataSource = dst.Tables("ASTATTA2")
        For Each GC As UltraWinGrid.UltraGridColumn In grdASTATTA2.DisplayLayout.Bands(0).Columns
            If GC.Key <> "ATTACHMENT_DESC" Or GC.Key = "ATTACHMENT_TYPE" Then
                GC.CellAppearance.BackColor = Color.LightYellow ' .Drawing.Color.FromArgb(255, 255, 255, 192)
            End If
            If GC.Key = "ATTACHMENT_ORIGINATOR" Or GC.Key = "ATTACHMENT_DATETIME" Then
                GC.Header.Appearance.BackColor2 = Color.Green
            End If
            If GC.Key = "ATTACHMENT_FILENAME" Or GC.Key = "ATTACHMENT_EXT" Then
                GC.Header.Appearance.BackColor2 = Color.Purple
            End If
            If GC.Key = "INIT_OPER" Or GC.Key = "INIT_DATE" Then
                GC.Header.Appearance.BackColor2 = Color.HotPink
            End If
            GC.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
        Next
        With grdICTLOTDF.DisplayLayout.Bands("ICTLOTDF")
            .Columns("WHSE_CODE").Header.Fixed = True
            .Columns("LOT_NO").Header.Fixed = True
            .Columns("LOT_SEQ_NO").Header.Fixed = True

            .Columns("PURCHASE_COST").Format = "###.000000"
            .Columns("VALUATION_COST").Format = "###.000000"
            .Columns("ADJUSTED_COST").Format = "###.000000"
            .Columns("STANDARD_COST").Format = "###.000000"
        End With

        With grdICTLOTHX.DisplayLayout.Bands("ICTLOTHX")
            .Columns("TRAN_DATE").Header.Fixed = True
            .Columns("TRAN_TYPE").Header.Fixed = True
            .Columns("TRAN_NO").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
            .Columns("CASES").Header.Fixed = True
            .Columns("UNITS").Header.Fixed = True

            .Columns("CASES_INV").CellAppearance.BackColor = Color.Yellow

            .Columns("FRT_RATE").Format = "###.0000"
            .Columns("ORDR_PRICE_GRS").Format = "###.0000"
            .Columns("ORDR_PRICE_NET").Format = "###.0000"
            .Columns("REBATE").Format = "###.0000"
            .Columns("BRKR_RATE").Format = "###.0000"
            .Columns("FUND_RATE").Format = "###.0000"
            .Columns("ALLOW_RATE").Format = "###.0000"
            .Columns("SVC_CHG_RATE").Format = "###.0000"
            .Columns("STANDARD_COST").Format = "###.000000"
            .Columns("ADJUSTED_COST").Format = "###.000000"
            .Columns("CGS_COST").Format = "###.000000"

            .Columns("CASES").Format = "###,##0"
            .Columns("UNITS").Format = "###,##0"
            .Columns("CASES_INV").Format = "###,##0"
            .Columns("UNITS_INV").Format = "###,##0"

        End With

        With grdICTLOTDF.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"ON_HOLD_FLAG", "LOT_NOTES"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        'cbeDIVISION_CODE.DataSource = ASCDATA1.GetDataTable("Select DIVISION_CODE, DIVISION_NAME from SOTSDIV1 order by DIVISION_CODE")
        'cbeDIVISION_CODE.ValueMember = "DIVISION_CODE"
        'cbeDIVISION_CODE.DisplayMember = "DIVISION_NAME"

        Create_Summary(grdICTLOTDF, "LOT_NO", "Count")
        Create_Summary(grdICTLOTDF, "QTY_ON_HAND")
        Create_Summary(grdICTLOTDF, "QTY_COMMITTED")
        Create_Summary(grdICTLOTDF, "PROD_MGR_HOLD")

        Create_Summary(grdICTLOTHX, "TRAN_NO", "Count")
        Create_Summary(grdICTLOTHX, "CASES_INV")
        Create_Summary(grdICTLOTHX, "UNITS_INV")

        Create_Summary(grdSOTORDRX, "SO_ORDER_NO", "Count")
        Create_Summary(grdSOTORDRX, "QTY_CASES")
        Create_Summary(grdSOTORDRX, "QTY_UNITS")


        Create_Summary(grdICTSTDL1, "VEND_CODE", "Count")
        Create_Summary(grdICTSTDL1, "TOTAL_SHP")
        Create_Summary(grdICTSTDL1, "TOTAL_INV")
        Create_Summary(grdICTSTDL1, "VARIANCE")

        Create_Summary(grdICTLOTDX, "DATE_WHSE_ANNIV", "Count")
        Create_Summary(grdICTLOTDX, "ADDED_STORAGE")
        Create_Summary(grdICTLOTDX, "ADDED_STORAGE_TOTAL")
        Create_Summary(grdICTLOTDX, "SORTING")
        Create_Summary(grdICTLOTDX, "HANDLING")
        Create_Summary(grdICTLOTDX, "SORTING_UNIT")
        Create_Summary(grdICTLOTDX, "HANDLING_UNIT")
        Create_Summary(grdICTLOTDX, "ACT_STG")
        Create_Summary(grdICTLOTDX, "ACT_SORT")
        Create_Summary(grdICTLOTDX, "ACT_HAND")

        Create_Summary(grdICTLOTDI, "NO")

        'TABLE_NAME = "ICTLOTD1"
        'Bind_Controls(UltraGroupBox1, "ICTLOTD1")
        'Bind_Controls(grpOriginatingLot, "ICTLOTD1")
        ASCMAIN1.Add_Value_List(grdICTLOTHX, "TRAN_TYPE", , New String() {":", "C:Return", "S:Sale", "R:Receipt", "A:Adjust", "T:XfrRec", "X:XfrOut", "W:<Slack>", "P:Slack", "D:Claim"})
        ASCMAIN1.Add_Value_List(grdICTLOTDX, "WHSE_CHARGE_TYPE", , New String() {":", "R:Rec", "S:Stg"})
        ASCMAIN1.Add_Value_List(grdICTLOTDF, "ON_HOLD_FLAG")

        Bind_Controls(frmCGS, "ICTIRECX")

        Set_Read_Only(frmCGS, True)
        Set_Read_Only_for_ctl(optCGS, False)

        UltraExplorerBar1.Groups("Screen Control").Items("Maintain Lots").Visible = ASCMAIN1.USER_SECURITY_CODEs.Contains("TI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Find Lots"
                If Absx1.txtFor("PROD_CODE").Text = "" _
                And Absx1.txtFor("LOT_NO").Text = "" _
                And Absx1.txtFor("PO_ORDER_NO").Text = "" _
                And Absx1.txtFor("IMPORT_NO").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify Either a Product, Lot No, PO No or Import No"
                End If

            Case "Maintain Lots"

                If optFindLots.Value <> "S" Then
                    EMsg &= vbCr & "You must first Find Lots"
                Else
                    If grdICTLOTDF.Rows.Count = 0 Then
                        EMsg &= vbCr & "You must first Find Lots"
                    End If
                End If

            Case "Update"
                If dst.Tables("ICTLOTDF").Select("ON_HOLD_FLAG = 'N'").Length <> 0 Then
                    EMsg &= vbCr & "Cannot Set FDA Status to None"
                End If


            Case "Load"
                WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                LOT_NO = Absx1.txtFor("LOT_NO").Text
                LOT_SEQ_NO = Val(Absx1.numFor("LOT_SEQ_NO").Value & "")
                rowICTLOTD1 = Fill_Record("ICTLOTD1", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
                If rowICTLOTD1 Is Nothing Then
                    EMsg &= vbCr & "No Record on file for " & WHSE_CODE & ":" & LOT_NO & ":" & CStr(LOT_SEQ_NO)
                Else
                    PO_ORDER_NO = rowICTLOTD1.Item("PO_ORDER_NO") & ""
                    If rowICTLOTD1.Item("PACK_CODE") = TAC.TACMAIN1.CATCH_PACK Then
                        PACK_FACTOR = Val(rowICTLOTD1.Item("CATCH_WEIGHT") & "")
                    Else
                        Dim rowICTPACK1 As DataRow = LookUp("ICTPACK1", rowICTLOTD1.Item("PACK_CODE"))
                        PACK_FACTOR = Val(rowICTPACK1.Item("PACK_FACTOR") & "")
                    End If
                End If

            Case "Adjust Qty", "Adjust Cost", "Change Lot Info"

                If Not alreadyAuth Then
                    Dim pwd As String = ASCMAIN1.Get_txt_from_User("Enter Password", "Adjustments Password", True)
                    If pwd = "" Then Exit Sub
                    If pwd = ROWs("ICTPARM1").Item("IC_PARM_ADJ_PWD") & "" Then
                        If adjustment_password_failures > 3 Then
                            alreadyAuth = False
                            Exit Sub
                        Else
                            alreadyAuth = True
                        End If
                    Else
                        adjustment_password_failures += 1
                        Exit Sub
                    End If

                End If

                If Not ASCMAIN1.Logical_Lock("ICTLOTD1", WHSE_CODE & ":" & LOT_NO & ":" & CStr(LOT_SEQ_NO)) Then
                    Exit Sub
                End If


            Case "Fix Lot Qty Comm"
                Dim QTY_COMMITTED As Int64 = Val(dst.Tables("SOTORDRX").Compute("SUM(QTY_CASES)", "") & "")
                If Val(Absx1.numFor("QTY_COMMITTED").Value & "") = QTY_COMMITTED Then
                    EMsg &= vbCr & "Lot Qty Committed is Not Out of Balance"
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

            Case "Find Lots"
                Find_Lots()

            Case "Maintain Lots"
                For Each rowICTLOTDF As DataRow In dst.Tables("ICTLOTDF").Select
                    If rowICTLOTDF.Item("ON_HOLD_FLAG") & "" = "" Then
                        rowICTLOTDF.Item("ON_HOLD_FLAG") = "M"
                    End If
                Next

                Maintain_Lots(True)

            Case "Update"

                For Each rowICTLOTDF As DataRow In dst.Tables("ICTLOTDF").Select
                    If rowICTLOTDF.Item("ON_HOLD_FLAG") & "" = "M" Then
                        rowICTLOTDF.Item("ON_HOLD_FLAG") = DBNull.Value
                    End If
                    For Each TABLE_NAME As String In New String() {"ICTLOTD1", "ICTLOTD2"}
                        ASCMAIN1.sql = "Update " & TABLE_NAME & " set ON_HOLD_FLAG = :PARM1, LOT_NOTES = :PARM2" _
                        & " where WHSE_CODE = :PARM3 and LOT_NO = :PARM4 and LOT_SEQ_NO = :PARM5"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVN", New Object() _
                                            {rowICTLOTDF.Item("ON_HOLD_FLAG"), _
                                             rowICTLOTDF.Item("LOT_NOTES"), _
                                             rowICTLOTDF.Item("WHSE_CODE"), _
                                             rowICTLOTDF.Item("LOT_NO"), _
                                             rowICTLOTDF.Item("LOT_SEQ_NO")})
                    Next
                Next
                dst.Tables("ICTLOTDF").AcceptChanges()

                Maintain_Lots(False)

            Case "Cancel"
                dst.Tables("ICTLOTDF").RejectChanges()
                Maintain_Lots(False)

            Case "Clear"
                For Each COLUMN_NAME As String In New String() _
                    {"PROD_CODE", "SIZE_CODE", "PACK_CODE", "BRAND_CODE", "ORIG_CODE", _
                     "SP_GROUP", "GRADE_CODE", "WHSE_CODE", "PO_ORDER_NO", "ITEM_CODE", _
                     "LOT_NO", "PROD_DIV_CODE", "MOP", "IMPORT_NO", "SHIPPERS_CODE"}
                    Absx1.txtFor(COLUMN_NAME).Text = ""
                Next
                Absx1.CtlFor("LOT_SEQ_NO").Text = ""

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Refresh"

                EnforceConstraints(False)
                Load_Lot_Data()
                EnforceConstraints(True)

            Case "Lots with Neg Avail"
                Fill_Records("ICTLOTN1")
                Using frm As New ASFMSGBF
                    frm.Show_grd(dst.Tables("ICTLOTN1"), Me, "Lots with Negative On Hand")
                End Using

            Case "Lots with Comm OOBal"
                Fill_Records("ICTLOTO1")
                If dst.Tables("ICTLOTO1").Rows.Count = 0 Then
                    MsgBox("No Lots are Currently Out of Balance" & vbCrLf & "  when Comparing the Qty Committed with Open Sales Orders", MsgBoxStyle.OkOnly, "Verification")
                Else
                    Using frm As New ASFMSGBF
                        frm.Show_grd(dst.Tables("ICTLOTO1"), Me, "Lots with OOB Commitments")
                    End Using
                    If ASCMAIN1.USER_SECURITY_CODEs.Contains("OM") Then
                        If MsgBox("OK to Reset the Qty Committed for All Lots OOBAL?", _
                            MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                            BeginTrans()
                            For Each rowICTLOTO1 As DataRow In dst.Tables("ICTLOTO1").Select
                                Dim QTY_COMMITTED_old As Int64 = Val(rowICTLOTO1.Item("LOT_QTY_COMM") & "")
                                Dim QTY_COMMITTED As Int64 = Val(rowICTLOTO1.Item("ORDER_QTY_COMM") & "")
                                Dim WHSE_CODE As String = rowICTLOTO1.Item("WHSE_CODE")
                                Dim LOT_NO As String = rowICTLOTO1.Item("LOT_NO")
                                Dim LOT_SEQ_NO As Int32 = rowICTLOTO1.Item("LOT_SEQ_NO")
                                Reset_Lot_Qty_Committed(WHSE_CODE, LOT_NO, LOT_SEQ_NO, QTY_COMMITTED, QTY_COMMITTED_old)
                            Next
                            CommitTrans("Lot Qty Committed has been Reset for all Lots OOBAL")
                        End If
                    End If
                End If


            Case "Print"
                Print_Record()

            Case "Attachments w/Empty Lots"
                MsgBox("The feature you have requested has not yet been enabled", MsgBoxStyle.OkOnly, "Cannot Proceed")
                'ICFLOTDG.Show(1)

            Case "Adjust Qty", "Adjust Cost", "Change Lot Info"
                tabADJ.Tabs("Qty").Visible = (eItemKey = "Adjust Qty")
                tabADJ.Tabs("Cost").Visible = (eItemKey = "Adjust Cost")
                tabADJ.Tabs("Lot").Visible = (eItemKey = "Change Lot Info")
                Initialize_for_Adjustment(False)

            Case "Fix Lot Qty Comm"
                Dim QTY_COMMITTED As Int64 = Val(dst.Tables("SOTORDRX").Compute("SUM(QTY_CASES)", "") & "")
                If MsgBox("OK to Reset the Qty Committed of this lot to " & CStr(QTY_COMMITTED) & "?", _
                          MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                    BeginTrans()
                    Dim QTY_COMMITTED_old As Int64 = Val(rowICTLOTD1.Item("QTY_COMMITTED") & "")
                    Reset_Lot_Qty_Committed(WHSE_CODE, LOT_NO, LOT_SEQ_NO, QTY_COMMITTED, QTY_COMMITTED_old)
                    rowICTLOTD1.Item("QTY_COMMITTED") = QTY_COMMITTED
                    CommitTrans("Lot Qty Committed has been Set to " & CStr(QTY_COMMITTED))
                End If

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Find Lots").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Clear").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Report").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Maintain Lots").Settings.Enabled = not_iScreenMode

                .Groups("Integrity Checks").Visible = Not ScreenMode
                .Groups("Adjustments").Visible = ScreenMode And ASCMAIN1.USER_SECURITY_CODEs.Contains("OP")

                Maintain_Lots(False)
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("LOT_NOTES"), False)


        Absx1.txtFor("LOT_NOTES").Visible = ScreenMode
        cmdSaveNote.Visible = ScreenMode

        grdICTLOTDF.DisplayLayout.Bands(0).Columns("RELATIONSHIP").Hidden = Not ScreenMode
        If ScreenMode Then
            grdICTLOTDF.Parent = tabLotDetails.Tabs("Associated").TabPage
        Else
            grdICTLOTDF.Parent = spl.Panel2
        End If

        Setup_Screen()

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("ICTLOTD1").Rows.Clear()
        dst.Tables("ICTLOTDF").Rows.Clear()

        dst.Tables("SOTORDRX").Rows.Clear()
        dst.Tables("ICTLOTHX").Rows.Clear()
        dst.Tables("ICTLOTHX_XREC").Rows.Clear()

        dst.Tables("ICTIRECX").Rows.Clear()
        dst.Tables("ICTLOTDX").Rows.Clear()

        If optFindLots.Value = "S" Then
            optFindLots.Value = "V"
        Else
            Setup_Recently_Viewed()
        End If

        dst.EnforceConstraints = True

        Absx1.txtFor("PO_ORDER_NO").Text = ""
        Absx1.txtFor("LOT_NO").Focus()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Dim rowICTLOTDV As DataRow = dst.Tables("ICTLOTDV").Rows.Find(New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        If rowICTLOTDV Is Nothing Then
            dst.Tables("ICTLOTDV").Merge(dst.Tables("ICTLOTD1"))
        End If


        'Fill_Records("SS", LOT_NO)

        EnforceConstraints(False)
        Load_Lot_Data()
        EnforceConstraints(True)

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "SIZE_CODE", "ORIG_CODE", "PACK_CODE", "BRAND_CODE", "GRADE_CODE", "SP_GROUP"
                If Absx1.txtFor("PROD_CODE").Text <> "" Then
                    sql_where = COLUMN_NAME & " IN (Select Distinct " & COLUMN_NAME & " from ICTLOTD1 where PROD_CODE = '" & Absx1.txtFor("PROD_CODE").Text & "')"
                End If
        End Select

    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)
            Case "Find Lots"
                Dim PROD_CODE As String = Split(key, vbTab)(1)
                Dim SIZE_CODE As String = Split(key & vbTab, vbTab)(2)
                Dim PACK_CODE As String = Split(key & vbTab, vbTab)(3)
                Dim WHSE_CODE As String = Split(key & vbTab, vbTab)(4)
                Dim BRAND_CODE As String = Split(key & vbTab, vbTab)(5)

                Absx1.txtFor("PROD_CODE").Text = PROD_CODE
                Absx1.txtFor("SIZE_CODE").Text = SIZE_CODE
                Absx1.txtFor("PACK_CODE").Text = PACK_CODE
                Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
                Absx1.txtFor("BRAND_CODE").Text = BRAND_CODE
                chkLiveLotsOnly.Checked = True

                Click_Command(command)
            Case "Load"
                Dim WHSE_CODE As String = Split(key, vbTab)(0)
                Dim LOT_NO As String = Split(key & vbTab, vbTab)(1)
                Dim LOT_SEQ_NO As Int64 = Split(key & vbTab, vbTab)(2)
                Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
                Absx1.txtFor("LOT_NO").Text = LOT_NO
                Absx1.numFor("LOT_SEQ_NO").Value = LOT_SEQ_NO


                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            If rowICTLOTD1.Item("RECEIPT_NO") & "" <> "" Then
                E.TABLE_NAME = "ICTIREC2"
                E.COLUMN_NAME = "KEY"
                E.CODE_VALUE = rowICTLOTD1.Item("RECEIPT_NO") & ":" & rowICTLOTD1.Item("RECEIPT_LNO")

                Dim customSql As String = "SELECT * FROM (" & vbCrLf _
                & " SELECT * FROM ASTATTA2 WHERE TABLE_NAME = 'POTORDR1' AND COLUMN_NAME = 'PO_ORDER_NO' AND CODE_VALUE = '" & rowICTLOTD1.Item("PO_ORDER_NO") & "'" & vbCrLf _
                & " UNION " & vbCrLf _
                & " SELECT * FROM ASTATTA2 WHERE TABLE_NAME = 'ICTIREC2' AND COLUMN_NAME = '" & E.COLUMN_NAME & "' AND CODE_VALUE = '" & E.CODE_VALUE & "'" & vbCrLf _
                & " UNION " & vbCrLf _
                & " SELECT * FROM ASTATTA2 WHERE TABLE_NAME = 'ICTLOTD1' AND COLUMN_NAME = 'LOT_NO' AND CODE_VALUE = '" & E.CODE_VALUE & "'" & vbCrLf _
                & " ) where NVL(ATTACHMENT_STATUS,'O') <> 'D'"
                E.CUSTOM_SQL = customSql

            Else
                E.TABLE_NAME = "ICTLOTD1"
                E.COLUMN_NAME = "LOT_NO"
                E.CODE_VALUE = HFs("WHSE_CODE") & ":" & HFs("LOT_NO") & ":" & HFs("LOT_SEQ_NO")
            End If
            E.DESC_VALUE = ""
            E.ATTACHMENT_NOTES = ""
        End If

        Return E
    End Function

    Public Overrides Function Audit_Context() As Audit_Entity

        'Dim E As New Audit_Entity
        'If ScreenMode Then
        '    If tab1.SelectedTab.Key = "General Info" Then
        '        E.TABLE_NAME = "BATGRPM2"
        '        E.KEY_VALUE = HFs("GROUP_NO") & ":" & HFs("SUBLOC_NO")
        '        E.KEY_DESC = "Sub-Location"
        '    End If
        'End If
        'Return E
        Return Nothing
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTLOTDF, "SSBBBB", "Show Filter", "Show GroupBox", "Clear History", "Lot Inquiry", "Lot Inquiry X-Lot", "Set FDA Status")
        Load_Popup_Menu(grdICTSTDL1, "B", "Voucher Inquiry")
        Load_Popup_Menu(grdICTLOTHX, "BBBBB", "Lot Inquiry X-Lot", "PO Inquiry", "Sales Order Inquiry", "Show All SEQs", "DR/CR Memo Inquiry")
        Load_Popup_Menu(grdSOTORDRX, "BB", "Sales Order Inquiry", "Sales Order Entry")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If tlb_pop.Tools.Exists("Clear History") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Clear History"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = (optFindLots.Value = "V")
        End If
        If tlb_pop.Tools.Exists("Set FDA Status") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Set FDA Status"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = (grdICTLOTDF.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True)
        End If

        If tlb_pop.Tools.Exists("PO Inquiry") Then
            tlb_btn = DirectCast(tlb_pop.Tools("PO Inquiry"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Cells("TRAN_TYPE").Value = "R")
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdICTLOTHX"

                    If grd.ActiveRow IsNot Nothing Then
                        If grd.ActiveRow.Band.Index <> 0 Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If

                    tlb_pop.Tools("Sales Order Inquiry").SharedProps.Visible = (grd.ActiveRow.Cells("TRAN_TYPE").Text <> "Return")
                    tlb_pop.Tools("DR/CR Memo Inquiry").SharedProps.Visible = (grd.ActiveRow.Cells("TRAN_TYPE").Text = "Return")

                Case Else
            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Clear History"
                dst.Tables("ICTLOTDV").Rows.Clear()

            Case "Set FDA Status"
                Using F As New ASFMSGBF
                    Dim FDA As Integer = F.Get_opt_from_User( _
                    "Select FDA Status to apply to all Lots Shown", _
                    New String() {"May Proceed", "Rejected", "Examination", "Unknown"}, 0, "FDA Status")
                    If FDA <> -1 Then
                        For Each rowICTLOTDF As DataRow In dst.Tables("ICTLOTDF").Select
                            rowICTLOTDF.Item("ON_HOLD_FLAG") = New String() {"M", "R", "X", "N"}(FDA)
                        Next
                    End If
                End Using


            Case "Show All SEQs"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
                Dim dd As DataTable = dst.Tables("ICTLOTH1")
            Case "Sales Order Inquiry"
                Dim SO_ORDER_NO As String = grd.ActiveRow.Cells("SO_ORDER_NO").Text
                Context_Launch("Load", SO_ORDER_NO, e.Tool.Key, "SOFORDRI", "F", "SO")

            Case "DR/CR Memo Inquiry"
                Dim SO_ORDER_NO As String = grd.ActiveRow.Cells("SO_ORDER_NO").Text
                Context_Launch("Load", SO_ORDER_NO, e.Tool.Key, "SOFMEMOI", "F", "SOE")
            Case "Sales Order Entry"
                Dim SO_ORDER_NO As String = grd.ActiveRow.Cells("SO_ORDER_NO").Text
                Context_Launch("Edit", SO_ORDER_NO, e.Tool.Key, "SOFORDR1", "F", "SOE")

            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Text
                'Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                MsgBox("Not Yet")

            Case "PO Inquiry"
                Dim TRAN_TYPE As String = grd.ActiveRow.Cells("TRAN_TYPE").Value
                If TRAN_TYPE = "R" Then
                    Dim TRAN_NO As String = grd.ActiveRow.Cells("TRAN_NO").Text
                    Dim rowICTIREC1 As DataRow = LookUp("ICTIREC1", TRAN_NO)
                    Dim PO_ORDER_NO As String = rowICTIREC1.Item("PO_ORDER_NO")
                    Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")
                End If

            Case "Lot Inquiry"
                Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Text
                Dim LOT_NO As String = grd.ActiveRow.Cells("LOT_NO").Text
                If WHSE_CODE = "" Or LOT_NO = "" Then
                    MsgBox("No Lot to View", MsgBoxStyle.OkOnly, "Cannot View Lot")
                    Exit Sub
                End If
                Dim LOT_SEQ_NO As Int64 = Val(grd.ActiveRow.Cells("LOT_SEQ_NO").Text)
                Context_Launch("Load", WHSE_CODE & vbTab & LOT_NO & vbTab & CStr(LOT_SEQ_NO), e.Tool.Key, "ICFLOTH1")

            Case "Lot Inquiry X-Lot"

                Dim WHSE_CODE As String = grd.ActiveRow.Cells("X_WHSE_CODE").Text
                Dim LOT_NO As String = grd.ActiveRow.Cells("X_LOT_NO").Text
                If WHSE_CODE = "" Or LOT_NO = "" Then
                    MsgBox("No X-Lot to View", MsgBoxStyle.OkOnly, "Cannot View X-Lot")
                    Exit Sub
                End If
                Dim LOT_SEQ_NO As Int64 = Val(grd.ActiveRow.Cells("X_LOT_SEQ_NO").Text)
                Context_Launch("Load", WHSE_CODE & vbTab & LOT_NO & vbTab & CStr(LOT_SEQ_NO), e.Tool.Key, "ICFLOTH1")

        End Select


    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LOT_NO", "IMPORT_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Find Lots", e)
                End If

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "SREP_CODE"
            '    Call Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "ITEM_CODE"
                If Absx1.txtFor("ITEM_CODE").Text <> "" Then
                    Dim rowICTITEM0 As DataRow = LookUp("ICTITEM0", Absx1.txtFor("ITEM_CODE").Text)
                    If rowICTITEM0 IsNot Nothing Then
                        Absx1.txtFor("PROD_CODE").Text = rowICTITEM0.Item("PROD_CODE") & ""
                        Absx1.txtFor("SIZE_CODE").Text = rowICTITEM0.Item("SIZE_CODE") & ""
                        Absx1.txtFor("PACK_CODE").Text = rowICTITEM0.Item("PACK_CODE") & ""
                        Absx1.txtFor("BRAND_CODE").Text = rowICTITEM0.Item("BRAND_CODE") & ""
                        Absx1.txtFor("ORIG_CODE").Text = rowICTITEM0.Item("ORIG_CODE") & ""
                        Absx1.txtFor("SP_GROUP").Text = rowICTITEM0.Item("SP_GROUP") & ""
                        Absx1.txtFor("ITEM_CODE").Tag = "Y"
                    End If
                End If
        End Select

    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "ITEM_CODE"
                If Absx1.txtFor("ITEM_CODE").Tag = "Y" Then
                    Absx1.txtFor("ITEM_CODE").Tag = ""
                    Click_Command("Find Lots")
                End If
        End Select
    End Sub

    Public Overrides Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_Leave(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            'Case "ITEM_CODE"
            '    If Absx1.txtFor(COLUMN_NAME).Text <> "" Then
            '        Absx1.txtFor("PROD_CODE").Text = ""
            '    End If
        End Select

    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "ADJ_COST"
                Calculate_New_Costs()
        End Select
    End Sub


#End Region

    Sub Load_Lot_Data()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Views")

        Prepare_ICTLOTHX()

        Fill_Records("SOTORDRX", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Load_ICTLOTDX()

        Fill_Records("ICTLOTD3", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        For Each rowICTLOTD3 As DataRow In dst.Tables("ICTLOTD3").Rows
            rowICTLOTD3.Item("YP_LEGEND") = ASCMAIN1.Get_Legend(rowICTLOTD3.Item("OPS_YYYYPP"))
            Dim DAYS As Int32 = Val(rowICTLOTD3.Item("OH_DAYS") & "")
            If DAYS <> 0 Then
                rowICTLOTD3.Item("AVG_CASES") = Val(rowICTLOTD3.Item("OH_CASES_AVG") & "") / DAYS
                rowICTLOTD3.Item("AVG_UNITS") = Val(rowICTLOTD3.Item("OH_UNITS_AVG") & "") / DAYS
            End If
        Next
        dst.Tables("ICTLOTD3").AcceptChanges()
        grdICTLOTD3.Text = "Monthly Balances & Activity for " & WHSE_CODE & " " & LOT_NO & " " & CStr(LOT_SEQ_NO)
        Sort_grdColumns(grdICTLOTD3, "OPS_YYYYPP".ToLower)
        Setup_LOTD6()

        Fill_Records("ICTLOTD5", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Sort_grdColumns(grdICTLOTD5, "OPS_YYYYPP".ToLower)

        Load_Attachments()

        rowICTIRECX = Fill_Record("ICTIRECX", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})

        Dim INT_START_DATE As Date
        If rowICTIRECX IsNot Nothing Then
            If rowICTIRECX.Item("INT_START_DATE") & "" <> "" Then
                INT_START_DATE = rowICTIRECX.Item("INT_START_DATE")
            Else
                INT_START_DATE = rowICTLOTD1.Item("DATE_RECEIVED")
            End If
        End If
        Fill_Records("ICTLOTDI", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Dim CUM_INTEREST As Decimal = 0
        Dim NO As Integer = 0
        For Each rowICTLOTDI As DataRow In dst.Tables("ICTLOTDI").Select("", "INIT_DATE")
            CUM_INTEREST += Val(rowICTLOTDI.Item("INTEREST_COST") & "")
            rowICTLOTDI.Item("CUM_INTEREST") = CUM_INTEREST
            NO += 1
            rowICTLOTDI.Item("NO") = NO
            Dim INIT_DATE As Date = rowICTLOTDI.Item("INIT_DATE")
            Dim DAYS As Integer = INIT_DATE.Date.Subtract(INT_START_DATE).Days
        Next
        Sort_grdColumns(grdICTLOTDI, "INIT_DATE")

        Load_Costs()

        Initialize_for_Adjustment()

        rowICTLOTD1 = Fill_Record("ICTLOTD1", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})

        If rowICTLOTD1.Item("ON_HOLD_FLAG") & "" = "" Then
            txtON_HOLD_FLAG_DESC.Text = "May Proceed"
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Sub Load_Attachments()

        'Get attachments for PO
        Dim PO_ORDER_NO As String = rowICTLOTD1.Item("PO_ORDER_NO")
        Fill_Records("ASTATTA2", New Object() {"POTORDR1", "PO_ORDER_NO", PO_ORDER_NO})
        'Get attachments for Receipts
        'Get attachments for Lots
        Dim KEY As String = rowICTLOTD1.Item("RECEIPT_NO") & ":" & rowICTLOTD1.Item("RECEIPT_LNO")
        Fill_Records("ASTATTA2", New Object() {"ICTIREC2", "KEY", KEY}, False)
        Fill_Records("ASTATTA2", New Object() {"ICTLOTD1", "LOT_NO", KEY}, False)

    End Sub
    Sub Load_Costs()

        ' Put non-zero elements from ICTCOSTZ into Work Table

        ASCMAIN1.sql = "Select * from ICTIREC2 " _
        & " where WHSE_CODE = :WHSE_CODE" _
        & " and LOT_NO = :LOT_NO" _
        & " and LOT_SEQ_NO = :LOT_SEQ_NO"
        rowICTIREC2 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVN", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Dim REC_UNITS As Decimal = 0
        If rowICTIREC2 IsNot Nothing Then
            REC_UNITS = Val(rowICTIREC2.Item("REC_UNITS") & "")
        End If

        Dim cc(15, 2) As String
        cc(1, 1) = "INSURANCE" : cc(1, 2) = "MARINE"
        cc(2, 1) = "PROD_LIAB" : cc(2, 2) = "LIAB"
        cc(3, 1) = "LC_COMM" : cc(3, 2) = "LC"
        cc(4, 1) = "REJECT_INS" : cc(4, 2) = "REJECT"
        cc(5, 1) = "OCEAN_FRT" : cc(5, 2) = "OCEAN"
        cc(6, 1) = "LOCAL_CARTAGE" : cc(6, 2) = "LOCAL"
        cc(7, 1) = "DUTY" : cc(7, 2) = "DUTY"
        cc(8, 1) = "CUSTOMS" : cc(8, 2) = "CUSTOM"
        cc(9, 1) = "COMMISSION" : cc(9, 2) = "COMM"
        cc(10, 1) = "INSPECTION" : cc(10, 2) = "INSP"
        cc(11, 1) = "OTHER" : cc(11, 2) = "OTHER"
        cc(12, 1) = "WHSE_STORAGE" : cc(12, 2) = "STG"
        cc(13, 1) = "WHSE_HANDLING" : cc(13, 2) = "HAND"
        cc(14, 1) = "WHSE_SORTING" : cc(14, 2) = "SORT"
        cc(15, 1) = "PO_BRKR_COMM" : cc(15, 2) = "POBRKR"

        ASCMAIN1.sql = "Select * from ICTCOSTZ " _
        & " where WHSE_CODE = :WHSE_CODE" _
        & " and LOT_NO = :LOT_NO" _
        & " and LOT_SEQ_NO = :LOT_SEQ_NO"
        rowICTCOSTZ = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVN", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})

        Fill_Records("POTTRAN0", PO_ORDER_NO)


        Load_ICTSTDL1()

        Dim VARIANCE As Decimal

        dst.Tables("ICTCOSTZ_LIST").Rows.Clear()
        If rowICTCOSTZ IsNot Nothing Then
            For i As Integer = 1 To UBound(cc, 1)
                If Val(rowICTCOSTZ.Item(cc(i, 1)) & "") <> 0 Then
                    VARIANCE = Val(rowICTCOSTZ.Item(cc(i, 1)) & "")
                    If cc(i, 1) = "WHSE_STORAGE" Then
                        VARIANCE = VARIANCE / 2
                    End If
                    Dim rowICTSTDL1 As DataRow = dst.Tables("ICTSTDL1").NewRow
                    With rowICTSTDL1
                        .Item("STAGE") = "1"
                        .Item("COST_CATGY_CODE") = cc(i, 2)
                        .Item("VARIANCE") = VARIANCE
                        .Item("TOTAL_SHP") = VARIANCE * REC_UNITS
                        .Item("UNITS") = REC_UNITS
                    End With
                    dst.Tables("ICTSTDL1").Rows.Add(rowICTSTDL1)
                End If

                Dim rowICTCOSTZ_LIST As DataRow = dst.Tables("ICTCOSTZ_LIST").NewRow
                With rowICTCOSTZ_LIST
                    .Item("COST_DESC") = cc(i, 1)
                    .Item("COST_VALUE") = VARIANCE
                End With
                dst.Tables("ICTCOSTZ_LIST").Rows.Add(rowICTCOSTZ_LIST)
            Next i
        End If
     

        frmCGS.Visible = False
        optCGS.Visible = False


        ' NEED TO TEST A NON-EDI WHSE IN HERE

        '    Dim dynAPTINVH1_EDI As OraDynaset
        '    SQL = "Select EDI_DOC_SEQ_NO from APTINVH1 where VOUCHER_NO = :CODE"
        '    Set dynAPTINVH1_EDI = OraD.CreateDynaset(SQL, 8&)
        '
        '    SQL = "Select * from ICWSTDL1 where (STAGE = '4' or STAGE = '5' or (STAGE = '1' and COST_CATGY_CODE in ('STG','HAND','SORT')))"
        '    & " and VOUCHER_NO is Not Null"
        '    Set dynICWSTDL1 = AccD.OpenRecordset(SQL, dbOpenDynaset)
        '    Do While Not dynICWSTDL1.EOF
        '        OraD.Parameters("CODE").Value = dynICWSTDL1.Fields("VOUCHER_NO").Value & ""
        '        dynAPTINVH1_EDI.Refresh
        '        If dynAPTINVH1_EDI.Fields("EDI_DOC_SEQ_NO").Value & "" <> "" Then
        '            OraD.Parameters("CODE").Value = dynAPTINVH1_EDI.Fields("EDI_DOC_SEQ_NO").Value & ""
        '            dynEDT810IA.Refresh
        '            dynICWSTDL1.Edit
        '            dynICWSTDL1.Fields("TOTAL_INV").Value = dynEDT810IA.Fields("INV_AMT").Value
        '            dynICWSTDL1.Fields("TOTAL_VAR").Value = Val(dynICWSTDL1.Fields("TOTAL_INV").Value & "") - Val(dynICWSTDL1.Fields("TOTAL_SHP").Value & "")
        '            dynICWSTDL1.Update
        '        End If
        '        dynICWSTDL1.MoveNext
        '    Loop
        '    dynICWSTDL1.Close

        For Each rowICTSTDL1 As DataRow In dst.Tables("ICTSTDL1").Select _
            ("STAGE = '1' and COST_CATGY_CODE = 'STG'")
            Dim rowICTLOTDX() As DataRow = dst.Tables("ICTLOTDX").Select _
                ("ORIG_TRAN_TYPE = 'P' AND WHSE_CHARGE_TYPE = 'R'")
            If rowICTLOTDX.Length > 0 Then
                If rowICTLOTDX(0).Item("VOUCHER_NO") & "" <> "" Then
                    rowICTSTDL1.Item("VOUCHER_NO") = rowICTLOTDX(0).Item("VOUCHER_NO")
                    rowICTSTDL1.Item("TOTAL_INV") = rowICTLOTDX(0).Item("INV_AMT")
                    rowICTSTDL1.Item("TOTAL_VAR") = Val(rowICTSTDL1.Item("TOTAL_INV") & "") - Val(rowICTSTDL1.Item("TOTAL_SHP") & "")
                End If
            End If
        Next

        Dim rowICTSTDL1s() As DataRow = dst.Tables("ICTSTDL1").Select("COST_CATGY_CODE = 'PURCH'")
        If rowICTSTDL1s.Length > 0 Then
            ASCMAIN1.sql = "SELECT MIN (VOUCHER_NO) VOUCHER_NO, SUM (PO_UNITS_INV * INV_COST) TOTAL_INV" & vbCrLf _
            & " FROM APTINVH5 WHERE CTL_NO IN (SELECT CTL_NO FROM POTTRAN1" & vbCrLf _
            & " WHERE PO_ORDER_NO = :PARM1 AND TRAN_DL_STATUS = '1')" & vbCrLf _
            & " and PO_ORDER_LNO = (SELECT PO_ORDER_LNO from ICTIREC2 " & vbCrLf _
            & " where WHSE_CODE = :PARM2" & vbCrLf _
            & "   and LOT_NO = :PARM3" & vbCrLf _
            & "   and LOT_SEQ_NO =:PARM4)"
            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVVN", New Object() {PO_ORDER_NO, WHSE_CODE, LOT_NO, LOT_SEQ_NO})
            If row.Item("VOUCHER_NO") & "" <> "" Then
                rowICTSTDL1s(0).Item("VOUCHER_NO") = row.Item("VOUCHER_NO")
                rowICTSTDL1s(0).Item("TOTAL_INV") = row.Item("TOTAL_INV")
                rowICTSTDL1s(0).Item("TOTAL_VAR") = Val(rowICTSTDL1s(0).Item("TOTAL_INV") & "") _
                                                 - Val(rowICTSTDL1s(0).Item("TOTAL_SHP") & "")
            End If
        End If


        Dim VOUCHER_NO As String
        Dim TOTAL_VAR As Decimal
        Dim rowICTSTDL1_STG() As DataRow = dst.Tables("ICTSTDL1").Select("STAGE='1' and COST_CATGY_CODE='STG'")

        If rowICTSTDL1_STG.Length > 0 Then
            If rowICTSTDL1_STG(0).Item("VOUCHER_NO") & "" <> "" Then
                VOUCHER_NO = rowICTSTDL1_STG(0).Item("VOUCHER_NO") & ""
                TOTAL_VAR = Val(rowICTSTDL1_STG(0).Item("TOTAL_VAR") & "")
                Dim rowICTSTDL1_SORT() As DataRow = dst.Tables("ICTSTDL1").Select("STAGE='1' and COST_CATGY_CODE='SORT'")
                If rowICTSTDL1_SORT.Length > 0 Then
                    TOTAL_VAR = TOTAL_VAR - Val(rowICTSTDL1_SORT(0).Item("TOTAL_SHP") & "")
                    rowICTSTDL1_SORT(0).Item("VOUCHER_NO") = VOUCHER_NO
                End If
                Dim rowICTSTDL1_HAND() As DataRow = dst.Tables("ICTSTDL1").Select("STAGE='1' and COST_CATGY_CODE='HAND'")
                If rowICTSTDL1_HAND.Length > 0 Then
                    TOTAL_VAR = TOTAL_VAR - Val(rowICTSTDL1_HAND(0).Item("TOTAL_SHP") & "")
                    rowICTSTDL1_HAND(0).Item("VOUCHER_NO") = VOUCHER_NO
                End If
                rowICTSTDL1_STG(0).Item("TOTAL_VAR") = TOTAL_VAR
            End If
        End If

        For Each rowICTSTDL1 As DataRow In dst.Tables("ICTSTDL1").Select("VEND_CODE is Null and STAGE = '1'")
            Dim sqlw As String = "COST_CATGY_CODE = '" & rowICTSTDL1.Item("COST_CATGY_CODE") & "' and ACCRUED = '1'"
            Dim rows() As DataRow = dst.Tables("POTTRAN0").Select(sqlw)
            If rows.Length > 0 Then
                Dim rowPOTTRAN0 As DataRow = rows(0)
                rowICTSTDL1.Item("VEND_CODE") = rowPOTTRAN0.Item("VEND_CODE")
                rowICTSTDL1.Item("INIT_DATE") = rowPOTTRAN0.Item("INIT_DATE")
                rowICTSTDL1.Item("INIT_OPER") = rowPOTTRAN0.Item("INIT_OPER")
            End If
        Next

        If rowICTIREC2 IsNot Nothing Then
            ASCMAIN1.sql = "Select * from ICTIREC6 where RECEIPT_NO = :PARM1 and RECEIPT_LNO = :PARM2"
            For Each rowICTIREC6 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VN", New Object() {rowICTIREC2.Item("RECEIPT_NO"), rowICTIREC2.Item("RECEIPT_LNO")}).Rows
                Dim rowICTSTDL1 As DataRow = dst.Tables("ICTSTDL1").NewRow
                With rowICTSTDL1
                    .Item("STAGE") = "1"
                    .Item("COST_CATGY_CODE") = rowICTIREC6.Item("COST_CATGY_CODE")
                    .Item("TOTAL_SHP") = Val(rowICTIREC6.Item("COST_PER_UNIT") & "") * REC_UNITS
                    .Item("VARIANCE") = rowICTIREC6.Item("COST_PER_UNIT")
                    .Item("UNITS") = REC_UNITS

                End With
                dst.Tables("ICTSTDL1").Rows.Add(rowICTSTDL1)
            Next
        End If

        Dim TOTAL_COST As Decimal = Val(dst.Tables("ICTSTDL1").Compute("Sum (VARIANCE)", "") & "")

        Dim OOBAL_msg As String = ""
        Dim STANDARD_COST As Decimal = Val(rowICTLOTD1.Item("STANDARD_COST") & "") ' Absx1.numFor("STANDARD_COST").Value & "")
        If System.Math.Abs(TOTAL_COST - STANDARD_COST) > 0.01 Then
            OOBAL_msg = " *** OUT OF BALANCE BY MORE THAN .01 ***"
        End If
        grdICTSTDL1.Text = "Std Cost Detail (Total Costs = " & Format$(TOTAL_COST, "$##.000000") & ")" & OOBAL_msg
        Sort_grdColumns(grdICTSTDL1, "STAGE,INIT_DATE")
    End Sub

    Sub Load_ICTLOTDX()

        Fill_Records("ICTLOTDX", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Fill_Records("EDT810IA", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Fill_Records("APTINVH7", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})

        For Each rowICTLOTDX As DataRow In dst.Tables("ICTLOTDX").Select("VOUCHER_NO is Not Null")
            With rowICTLOTDX
                Dim EDI_DOC_SEQ_NO As String = .Item("EDI_DOC_SEQ_NO") & ""
                Dim VOUCHER_NO As String = .Item("VOUCHER_NO") & ""

                If EDI_DOC_SEQ_NO <> "" Then
                    Dim sqlw As String = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                    .Item("INV_CASES") = Val(dst.Tables("EDT810IA").Compute("SUM(QTY_ON_HAND)", sqlw) & "")
                    .Item("INV_UNITS") = Val(dst.Tables("EDT810IA").Compute("SUM(UNITS_ON_HAND)", sqlw) & "")
                    .Item("INV_AMT") = Val(dst.Tables("EDT810IA").Compute("SUM(LOT_SAC_AMOUNT)", sqlw) & "")
                Else
                    Dim sqlw As String = "VOUCHER_NO = '" & VOUCHER_NO & "'"
                    .Item("INV_AMT") = Val(dst.Tables("APTINVH7").Compute("SUM(LOT_INV_AMOUNT)", sqlw) & "")
                End If
            End With
        Next
    End Sub

    Sub Setup_Screen()
        tabLotDetails.Visible = ScreenMode
        grpLotHeader.Visible = ScreenMode

        With UltraExplorerBar1
            .Groups("Find Lots Options").Visible = Not ScreenMode
            .Groups("Adjustment").Visible = False
        End With
    End Sub

    Private Sub grdICTLOTD1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTLOTDF.AfterRowActivate

    End Sub

    Sub Find_Lots()

        Dim sql_where As String = ""
        For Each COLUMN_NAME In New String() _
        {"PROD_CODE", "SIZE_CODE", "ORIG_CODE", "PACK_CODE", "WHSE_CODE", "BRAND_CODE", "SP_GROUP", "GRADE_CODE", _
         "PROD_DIV_CODE", "PO_ORDER_NO", "IMPORT_NO", "SHIPPERS_CODE", "LOT_NO"}
            If Absx1.txtFor(COLUMN_NAME).Text <> "" Then
                Dim COLUMN_TABLE As String = "ICTLOTD1"
                If COLUMN_NAME = "PO_ORDER_NO" Then
                    COLUMN_TABLE = "POTORDR1"
                End If
                sql_where &= " and " & COLUMN_TABLE & "." & COLUMN_NAME & " = '" & Absx1.txtFor(COLUMN_NAME).Text & "'"
            End If
        Next

        If sql_where = "" Then
            MsgBox("You must select at least 1 code value before Finding Lots", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim TABLE_NAME As String = "ICTLOTD1"
        If chkLiveLotsOnly.Checked Then
            TABLE_NAME = "ICTLOTD2"
            sql_where &= " and ICTLOTD1.QTY_ON_HAND <> 0"
        End If

        If optFindLots.Value <> "S" Then
            optFindLots.Value = "S"
        End If


        ASCMAIN1.sql = Replace(sqlICTLOTD1, "ICTLOTD1 ICTLOTD1", TABLE_NAME & " ICTLOTD1") & sql_where
        Fill_Records("ICTLOTDF", "", , ASCMAIN1.sql)
        Sort_grdColumns(grdICTLOTDF, "DATE_RECEIVED".ToLower)

        ' GRABBING LOT AND GOING MEANS WE CANNOT CLICK THE MAINTAIN LOTS OPTION
        'If dst.Tables("ICTLOTDF").Rows.Count = 1 Then
        '    Dim rowICTLOTDF As DataRow = dst.Tables("ICTLOTDF").Rows(0)
        '    Absx1.txtFor("WHSE_CODE").Text = rowICTLOTDF.Item("WHSE_CODE")
        '    Absx1.txtFor("LOT_NO").Text = rowICTLOTDF.Item("LOT_NO")
        '    Absx1.numFor("LOT_SEQ_NO").Value = rowICTLOTDF.Item("LOT_SEQ_NO")
        '    Click_Command("Load")
        'End If

    End Sub

    Private Sub optFindLots_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optFindLots.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub

        grdICTLOTDF.Text = optFindLots.Text

        If optFindLots.Value = "V" Then
            Setup_Recently_Viewed()
        End If

        If optFindLots.Value = "S" Then
            dst.Tables("ICTLOTDF").Rows.Clear()
        End If


    End Sub

    Sub Setup_Recently_Viewed()
        dst.Tables("ICTLOTDF").Rows.Clear()
        dst.Tables("ICTLOTDF").Merge(dst.Tables("ICTLOTDV"))
    End Sub

    Private Sub grdICTLOTDF_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTLOTDF.DoubleClickRow

        If grdICTLOTDF.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True Then
            Exit Sub
        End If

        If e.Row.IsDataRow Then

            Dim WHSE_CODE As String = e.Row.Cells("WHSE_CODE").Text
            Dim LOT_NO As String = e.Row.Cells("LOT_NO").Text
            Dim LOT_SEQ_NO As Int64 = e.Row.Cells("LOT_SEQ_NO").Value

            If ScreenMode Then
                Click_Command("Done")
            End If

            Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
            Absx1.txtFor("LOT_NO").Text = LOT_NO
            Absx1.numFor("LOT_SEQ_NO").Value = LOT_SEQ_NO

            Click_Command("Load")
        End If
    End Sub

    Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Absx1.txtFor("X_WHSE_CODE").Text = "" _
        Or Absx1.txtFor("X_LOT_NO").Text = "" _
        Or Val(Absx1.numFor("X_LOT_SEQ_NO").Value & "") = 0 Then
            Exit Sub
        End If

        Absx1.txtFor("WHSE_CODE").Text = Absx1.txtFor("X_WHSE_CODE").Text
        Absx1.txtFor("LOT_NO").Text = Absx1.txtFor("X_LOT_NO").Text
        Absx1.numFor("LOT_SEQ_NO").Value = Absx1.numFor("X_LOT_SEQ_NO").Value
        Click_Command("Load")
    End Sub

    Sub Prepare_ICTLOTHX()

        Fill_Records("ICTLOTHX", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})

        ASCMAIN1.sql = "Select ICTLOTH1.WHSE_CODE, ICTLOTH1.LOT_NO, ICTLOTH1.LOT_SEQ_NO" & vbCrLf _
        & ", ICTLOTH1.TRAN_DATE, ICTLOTH1.TRAN_NO, ICTLOTH1.TRAN_TYPE" & vbCrLf _
        & ", ICTLOTH1.TRAN_CASES CASES, ICTLOTH1.TRAN_UNITS UNITS" & vbCrLf _
        & ", ICTLOTH1.TRAN_CASES CASES_INV, ICTLOTH1.TRAN_UNITS UNITS_INV" & vbCrLf _
        & ", Null CASES_REL, Null UNITS_REL" & vbCrLf _
        & ", Null ORDR_PRICE_GRS, Null ORDR_PRICE_NET, Null REBATE, Null FUND_RATE" & vbCrLf _
        & ", Null BRKR_RATE, ICTLOTH1.FRT_RATE, Null ALLOW_RATE, Null SVC_CHG_RATE" & vbCrLf _
        & ", ICTLOTH1.X_WHSE_CODE, ICTLOTH1.X_LOT_NO, ICTLOTH1.X_LOT_SEQ_NO, ICTLOTH1.LOT_XY" & vbCrLf _
        & ", ICTLOTH1.CUST_CODE, NULL CUST_NAME, NULL CUST_ORDER_NO, TRAN_NO SO_ORDER_NO " & vbCrLf _
        & ", NULL ADJ_COST_EXT, NULL CGS_COST_EXT, ICTLOTH1.STANDARD_COST * ICTLOTH1.TRAN_UNITS STD_COST_EXT" & vbCrLf _
        & ", NULL CON_REG_IND" & vbCrLf _
        & " from ICTLOTH1" & vbCrLf _
        & " where ICTLOTH1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
        & "   and ICTLOTH1.LOT_NO = '" & LOT_NO & "'"
        If Not chkCombineSeqs.Checked Then
            ASCMAIN1.sql &= "   and ICTLOTH1.LOT_SEQ_NO = " & CStr(LOT_SEQ_NO)
        Else
            ' force match of 7 codes
        End If
        ASCMAIN1.sql &= "   and ICTLOTH1.TRAN_TYPE <> 'C'"
        ASCMAIN1.sql &= "   and ICTLOTH1.TRAN_TYPE <> 'D'"
        Fill_Records("ICTLOTHX", "", False, ASCMAIN1.sql)

        If chkCombineSeqs.Checked Then
            grdICTLOTHX.DisplayLayout.Bands(0).Columns("LOT_SEQ_NO").Hidden = False
            ASCMAIN1.sql = "Select * from ICTLOTD1 " _
            & " where WHSE_CODE = '" & WHSE_CODE & "'" _
            & "   and LOT_NO = '" & LOT_NO & "'" _
            & "   and LOT_SEQ_NO <> " & CStr(LOT_SEQ_NO)
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim LOT_SEQ_NO_2 As Integer = Val(row.Item("LOT_SEQ_NO") & "")
                Fill_Records("ICTLOTHX", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO_2}, False)
            Next
        Else
            grdICTLOTHX.DisplayLayout.Bands(0).Columns("LOT_SEQ_NO").Hidden = True
        End If

        Sort_grdColumns(grdICTLOTHX, "TRAN_DATE") ' NEED TO SORT BY INIT_DATE

        For Each rowICTLOTHX As DataRow In dst.Tables("ICTLOTHX") _
            .Select("TRAN_TYPE = 'R'")
            Dim rowICTIREC1 As DataRow = LookUp("ICTIREC1", rowICTLOTHX.Item("TRAN_NO"), True)
            rowICTLOTHX.Item("CUST_CODE") = rowICTIREC1.Item("VEND_CODE") & ""
            rowICTLOTHX.Item("CUST_NAME") = "PO " & rowICTIREC1.Item("PO_ORDER_NO") & "; " & rowICTIREC1.Item("VEND_CODE")
        Next

        For Each rowICTLOTHX As DataRow In dst.Tables("ICTLOTHX").Select("CASES_INV = 0")
            rowICTLOTHX.Item("CASES_INV") = DBNull.Value
        Next
        For Each rowICTLOTHX As DataRow In dst.Tables("ICTLOTHX").Select("UNITS_INV = 0")
            rowICTLOTHX.Item("UNITS_INV") = DBNull.Value
        Next

        Set_XY()

        Dim RECORD_INDEX As Int64 = 0
        For Each rowICTLOTHX As DataRow In dst.Tables("ICTLOTHX").Rows
            RECORD_INDEX += 1
            rowICTLOTHX.Item("RECORD_INDEX") = RECORD_INDEX
        Next

        Fill_Records("ICTLOTHX_XREC", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        For Each rowICTLOTHX_XREC As DataRow In dst.Tables("ICTLOTHX_XREC").Rows
            ASCMAIN1.sql = "SELECT ORDR_INV_NO FROM ICTTREC1,SOTINVH1 " _
            & " where ICTTREC1.TRANSFER_REC_NO = :PARM1" _
            & " and SOTINVH1.SO_ORDER_NO = ICTTREC1.TRANSFER_NO"
            Dim ORDR_INV_NO As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", rowICTLOTHX_XREC.Item("TRAN_NO"))
            Dim rowICTLOTHX() As DataRow = dst.Tables("ICTLOTHX").Select("TRAN_TYPE = 'T' and TRAN_NO = '" & ORDR_INV_NO & "'")
            If rowICTLOTHX.Length = 1 Then
                rowICTLOTHX_XREC.Item("RECORD_INDEX") = rowICTLOTHX(0).Item("RECORD_INDEX")
            End If
        Next

        dst.Tables("ICTLOTHX").AcceptChanges()
        dst.Tables("ICTLOTHX_XREC").AcceptChanges()
    End Sub

    Sub Set_XY()
        Dim rowICTLOTD1_X As DataRow = dst.Tables("ICTLOTD1").NewRow
        Dim rowICTLOTD1_Y As DataRow = dst.Tables("ICTLOTD1").NewRow

        For Each rowICTLOTHX_XY As DataRow In dst.Tables("ICTLOTHX").Select _
            ("LOT_XY is Not Null AND TRAN_TYPE <> 'C'", "WHSE_CODE, LOT_NO, LOT_SEQ_NO")
            With rowICTLOTHX_XY
                If .Item("WHSE_CODE") <> rowICTLOTD1_Y.Item("WHSE_CODE") & "" _
                Or .Item("LOT_NO") <> rowICTLOTD1_Y.Item("LOT_NO") & "" _
                Or Val(.Item("LOT_SEQ_NO")) <> Val(rowICTLOTD1_Y.Item("LOT_SEQ_NO") & "") Then
                    rowICTLOTD1_Y = LookUp("ICTLOTD1", New String() {.Item("WHSE_CODE"), .Item("LOT_NO"), .Item("LOT_SEQ_NO")})
                End If
                rowICTLOTD1_X = LookUp("ICTLOTD1", New String() {.Item("X_WHSE_CODE"), .Item("X_LOT_NO"), .Item("X_LOT_SEQ_NO")})
                Dim TRAN_TYPE As String = .Item("TRAN_TYPE") & ""
                Dim z As String = .Item("LOT_XY") & " Lot " & .Item("X_LOT_NO")
                For Each COLUMN_NAME As String In New String() _
                    {"PROD_CODE", "SIZE_CODE", "PACK_CODE", "ORIG_CODE", "BRAND_CODE", _
                     "SP_GROUP", "GRADE_CODE", "WHSE_CODE", "PROD_DIV_CODE"}
                    If rowICTLOTD1_Y.Item(COLUMN_NAME) & "" <> rowICTLOTD1_X.Item(COLUMN_NAME) & "" Then
                        Dim NEW_VALUE As String = rowICTLOTD1_X.Item(COLUMN_NAME)
                        If COLUMN_NAME = "PROD_CODE" Then
                        ElseIf COLUMN_NAME = "SIZE_CODE" Then
                            Dim rowICTSIZE1 As DataRow = LookUp("ICTSIZE1", rowICTLOTD1_X.Item(COLUMN_NAME))
                            NEW_VALUE = rowICTSIZE1.Item("SIZE_DESC")
                        ElseIf COLUMN_NAME = "PACK_CODE" Then
                            Dim rowICTSIZE1 As DataRow = LookUp("ICTPACK1", rowICTLOTD1_X.Item(COLUMN_NAME))
                            NEW_VALUE = rowICTSIZE1.Item("PACK_DESC")
                        ElseIf COLUMN_NAME = "PROD_DIV_CODE" Then
                            NEW_VALUE = "Div" & " " & NEW_VALUE
                        Else
                            Dim COLUMN_DESC As String = ASCMAIN1.Make_Caption(Split(COLUMN_NAME & "_", "_")(0))
                            NEW_VALUE = COLUMN_DESC & " " & NEW_VALUE
                        End If
                        z &= " " & NEW_VALUE
                    End If
                Next
                If Len(z) > 35 Then z = Mid(z, 1, 35)
                .Item("CUST_NAME") = z
            End With
        Next
    End Sub

    Sub Print_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Report")

        Print_Report_Begin()
        CR_params.Add("SUBT", "Whse " & Absx1.txtFor("WHSE_CODE").Text & " Lot No " & Absx1.txtFor("LOT_NO").Text)

        For Each COLUMN_NAME As String In New String() _
            {"PROD_DESC", "SIZE_DESC", "PACK_DESC", "ORIG_DESC", "WHSE_NAME", "COST_DESC", "PO_ORDER_NO"}
            CR_params.Add(COLUMN_NAME, Absx1.txtFor(COLUMN_NAME).Text)
        Next
        CR_params.Add("ON_HOLD_FLAG_DESC", txtON_HOLD_FLAG_DESC.Text)

        Generate_Report("ICRLOTH1", "Lot History", , , , , False)
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdSaveNote_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSaveNote.Click

        For Each TABLE_NAME As String In New String() {"ICTLOTD1", "ICTLOTD2"}
            ASCMAIN1.sql = "Update " & TABLE_NAME & " Set LOT_NOTES = :PARM1" _
            & " where WHSE_CODE = :PARM2 and LOT_NO = :PARM3 and LOT_SEQ_NO = :PARM4"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVN", New Object() _
                                {Absx1.txtFor("LOT_NOTES").Text, WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Next

        MsgBox("Lot Notes have been Saved", MsgBoxStyle.OkCancel, "Verification")
    End Sub

    Private Sub grdICTSTDL1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTDL1.InitializeRow
        If (e.Row.Cells("STAGE").Value = "2" Or e.Row.Cells("STAGE").Text = "3") _
        And System.Math.Abs(Val(e.Row.Cells("TOTAL_SHP").Value & "") - Val(e.Row.Cells("TOTAL_INV").Value & "")) > 0.01 Then
            e.Row.Cells("VARIANCE").Appearance.ForeColor = Color.Red
        End If
    End Sub

    Sub Load_ICTSTDL1()
        dst.Tables("ICTSTDL1").Rows.Clear()
        ASCDATA1.ExecuteSQL("Truncate Table " & ICTSTDL1)

        ASCMAIN1.Progress("Stage 0")
        ASCMAIN1.sql = "" _
        & "SELECT ICTIREC1.VEND_CODE, 'PURCH' COST_CATGY_CODE, '0' STAGE, ICTIREC1.DATE_RECEIVED TRAN_DATE," & vbCrLf _
        & "       NVL(ICTIREC2.REC_UNITS,0) * NVL(ICTIREC2.PURCHASE_COST,0) TOTAL_SHP, 0 TOTAL_INV, 0 TOTAL_VAR," & vbCrLf _
        & "       NVL(ICTIREC2.REC_UNITS,0) UNITS, NVL(ICTIREC2.PURCHASE_COST,0) VARIANCE, ICTIREC2.RECEIPT_NO CTL_NO," & vbCrLf _
        & "       ICTIREC1.INIT_DATE, ICTIREC1.INIT_OPER, ICTIREC1.LAST_DATE, ICTIREC1.LAST_OPER, NULL VOUCHER_NO" & vbCrLf _
        & " FROM ICTIREC1, ICTIREC2" & vbCrLf _
        & "  WHERE ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
        & "    AND ICTIREC2.WHSE_CODE = '" & WHSE_CODE & "' AND ICTIREC2.LOT_NO = '" & LOT_NO & "' AND ICTIREC2.LOT_SEQ_NO = " & LOT_SEQ_NO
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)

        ASCMAIN1.Progress("Stage 1 A")
        ASCMAIN1.sql = "" _
        & "SELECT APTINVH1.VEND_CODE, 'PURCH' COST_CATGY_CODE, '0' STAGE, APTINVH1.INV_DATE TRAN_DATE," & vbCrLf _
        & "       NVL(APTINVHA.UNITS_ONH,0) * NVL(APTINVHA.PO_COST,0) TOTAL_SHP," & vbCrLf _
        & "       NVL(APTINVHA.UNITS_ONH,0) * NVL(APTINVHA.INV_COST,0) TOTAL_INV," & vbCrLf _
        & "       NVL(APTINVHA.UNITS_ONH,0) * (NVL(APTINVHA.INV_COST,0) - NVL(APTINVHA.PO_COST,0)) TOTAL_VAR," & vbCrLf _
        & "       NVL(APTINVHA.UNITS_ONH,0) UNITS, (NVL(APTINVHA.INV_COST,0) - NVL(APTINVHA.PO_COST,0)) VARIANCE, APTINVHA.VOUCHER_NO CTL_NO," & vbCrLf _
        & "       APTINVH1.INIT_DATE, APTINVH1.INIT_OPER, APTINVH1.LAST_DATE, APTINVH1.LAST_OPER, APTINVH1.VOUCHER_NO VOUCHER_NO" & vbCrLf _
        & " FROM APTINVH1, APTINVHA" & vbCrLf _
        & "  WHERE APTINVHA.VOUCHER_NO = APTINVH1.VOUCHER_NO and APTINVH1.INV_STATUS <> 'D' and ABS(NVL(APTINVHA.INV_COST,0) - NVL(APTINVHA.PO_COST,0)) > .001" & vbCrLf _
        & "    AND APTINVHA.WHSE_CODE = '" & WHSE_CODE & "' AND APTINVHA.LOT_NO = '" & LOT_NO & "' AND APTINVHA.LOT_SEQ_NO = " & LOT_SEQ_NO
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)

        ASCMAIN1.Progress("Stage 1 B")
        ASCMAIN1.sql = "" _
        & "SELECT NULL VEND_CODE, 'XFR' COST_CATGY_CODE, '0' STAGE, ICTTREC1.DATE_RECEIVED TRAN_DATE," & vbCrLf _
        & "       0 TOTAL_SHP, 0 TOTAL_INV, 0 TOTAL_VAR," & vbCrLf _
        & "       NVL(ICTTREC2.TRANSFER_REC_UNITS,0) UNITS, NVL(ICTTREC2.STANDARD_COST_TO,0) VARIANCE, ICTTREC2.TRANSFER_REC_NO CTL_NO," & vbCrLf _
        & "       ICTTREC1.INIT_DATE, ICTTREC1.INIT_OPER, ICTTREC1.LAST_DATE, ICTTREC1.LAST_OPER, NULL VOUCHER_NO" & vbCrLf _
        & " FROM ICTTREC1, ICTTREC2" & vbCrLf _
        & "  Where ICTTREC1.TRANSFER_REC_NO = ICTTREC2.TRANSFER_REC_NO" & vbCrLf _
        & "    AND ICTTREC2.WHSE_CODE_TO = '" & WHSE_CODE & "' AND ICTTREC2.LOT_NO_TO = '" & LOT_NO & "' AND ICTTREC2.LOT_SEQ_NO_TO = " & LOT_SEQ_NO
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)
        '& "       NVL(ICTTREC2.TRANSFER_UNITS,0) UNITS, NVL(ICTTREC2.STANDARD_COST_TO,0) VARIANCE, ICTTREC2.TRANSFER_REC_NO CTL_NO," & vbCrLf

        'SELECT * FROM ICTLOTH1 WHERE LOT_NO = '73628V' AND LOT_SEQ_NO = 3 AND TRAN_TYPE = 'P' AND TRAN_NO = '282035'
        'WHS LOT_NO LOT_SEQ_NO           TRAN_DATE  T TRAN_NO TRAN_CASES           TRAN_UNITS           X_W X_LOT_ X_LOT_SEQ_NO         FRT_RATE             ORDR_PRICE_GRS       ORDR_PRICE_NET       CUST_CODE  L OPS_YY INIT_DATE  INIT_OPE STANDARD_COST
        '--- ------ -------------------- ---------- - ------- -------------------- -------------------- --- ------ -------------------- -------------------- -------------------- -------------------- ---------- - ------ ---------- -------- --------------------
        '235 73628V 3                    09/25/2006 P 282035  1                    2.5                  235 73628V 1                                                                                              X 200609 09/25/2006 kayh     6.383201
        'SELECT * FROM SOTWORK1 WHERE WO_ORDER_NO = '006344'
        'WO_ORD WO_ORDER_D W SO_ORDE SO_ORDER_LNO         WO_ORDER_C WO_COMMENTS          WHS PICK_N
        '------ ---------- - ------- -------------------- ---------- -------------------- --- ------
        '006344 09/25/2006 F 282035  15                   09/25/2006                      235

        ' USING SOTWORK1 TO TRANSLATE A SO_ORDER_NO INTO A WO_ORDER_NO UNTIL WE USE TRAN_TYPE = 'W' AND TRAN_NO = WORK ORDER NO
        ' INSTEAD OF TRAN_TYPE = 'P' AND TRAN_TYPE = SO_ORDER_NO IN ICTLOTH1 WHEN RECORDING SLACK WORK ORDER CONFIRMATIONS

        ASCMAIN1.Progress("Stage 1 C")
        ASCMAIN1.sql = "" _
        & "SELECT NULL VEND_CODE, 'RTN' COST_CATGY_CODE, '0' STAGE, ICTLOTH1.TRAN_DATE," & vbCrLf _
        & "       0 TOTAL_SHP, 0 TOTAL_INV, 0 TOTAL_VAR," & vbCrLf _
        & "       NVL(ICTLOTH1.TRAN_UNITS,0) UNITS, NVL(ICTLOTH1.STANDARD_COST,0) VARIANCE, ICTLOTH1.TRAN_NO CTL_NO," & vbCrLf _
        & "       ICTLOTH1.INIT_DATE, ICTLOTH1.INIT_OPER, ICTLOTH1.INIT_DATE, ICTLOTH1.INIT_OPER, NULL VOUCHER_NO" & vbCrLf _
        & " FROM ICTLOTH1" & vbCrLf _
        & "  Where ICTLOTH1.WHSE_CODE = '" & WHSE_CODE & "' AND ICTLOTH1.LOT_NO = '" & LOT_NO & "' AND ICTLOTH1.LOT_SEQ_NO = " & LOT_SEQ_NO _
        & "    AND ICTLOTH1.TRAN_TYPE = 'C' AND ICTLOTH1.TRAN_NO = '" & rowICTLOTD1.Item("ORIG_TRAN_NO") & "'"
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)

        If rowICTLOTD1.ITEM("ORIG_TRAN_TYPE") & "" = "W" Then
            ASCMAIN1.Progress("Stage 1 D")
            ASCMAIN1.sql = "" _
            & "SELECT NULL VEND_CODE, 'W/O' COST_CATGY_CODE, '0' STAGE, ICTLOTH1.TRAN_DATE," & vbCrLf _
            & "       0 TOTAL_SHP, 0 TOTAL_INV, 0 TOTAL_VAR," & vbCrLf _
            & "       NVL(ICTLOTH1.TRAN_UNITS,0) UNITS, NVL(ICTLOTH1.STANDARD_COST,0) VARIANCE, SOTWORK1.WO_ORDER_NO CTL_NO," & vbCrLf _
            & "       ICTLOTH1.INIT_DATE, ICTLOTH1.INIT_OPER, ICTLOTH1.INIT_DATE, ICTLOTH1.INIT_OPER, NULL VOUCHER_NO" & vbCrLf _
            & " FROM ICTLOTH1, SOTWORK1" & vbCrLf _
            & "  Where ICTLOTH1.TRAN_NO = SOTWORK1.SO_ORDER_NO" & vbCrLf _
            & "    AND ICTLOTH1.WHSE_CODE = '" & WHSE_CODE & "' AND ICTLOTH1.LOT_NO = '" & LOT_NO & "' AND ICTLOTH1.LOT_SEQ_NO = " & LOT_SEQ_NO _
            & "    AND ICTLOTH1.TRAN_TYPE = 'P' AND SOTWORK1.WO_ORDER_NO = '" & rowICTLOTD1.ITEM("ORIG_TRAN_NO") & "'"
            ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.Progress("Stage 2")

        ASCMAIN1.sql = "" _
        & "SELECT POTTRAN0.VEND_CODE, POTTRAN0.COST_CATGY_CODE, '2' STAGE, POTTRAN0.PO_DATE_SHIPPED TRAN_DATE," & vbCrLf _
        & "       NVL(POTTRAN0.TOTAL_SHP,0) TOTAL_SHP, NVL(POTTRAN0.TOTAL_INV,0) TOTAL_INV, NVL(POTTRAN0.TOTAL_INV,0) - NVL(POTTRAN0.TOTAL_SHP,0) TOTAL_VAR," & vbCrLf _
        & "       NVL(X.UNITS,0) UNITS, (NVL(POTTRAN0.TOTAL_INV,0) - NVL(POTTRAN0.TOTAL_SHP,0)) / NVL(X.UNITS,0) VARIANCE, POTTRAN0.CTL_NO," & vbCrLf _
        & "       POTTRAN0.INIT_DATE, POTTRAN0.INIT_OPER, POTTRAN0.LAST_DATE, POTTRAN0.LAST_OPER, APTINVH8.VOUCHER_NO" & vbCrLf _
        & " FROM POTTRAN0, APTINVH1, APTINVH8, (SELECT SUM(REC_UNITS) UNITS FROM ICTIREC2 WHERE PO_ORDER_NO = '" & PO_ORDER_NO & "') X" & vbCrLf _
        & "  WHERE POTTRAN0.CTL_STATUS = '1' AND POTTRAN0.TOTAL_SHP <> POTTRAN0.TOTAL_INV" & vbCrLf _
        & "   AND POTTRAN0.LAST_DATE IS NOT NULL AND POTTRAN0.LAST_DATE < (SELECT MIN(INIT_DATE) FROM ICTIREC1 WHERE PO_ORDER_NO = '" & PO_ORDER_NO & "')" & vbCrLf _
        & "   AND POTTRAN0.PO_ORDER_NO = '" & PO_ORDER_NO & "'" _
        & "   AND APTINVH8.CTL_NO (+) = POTTRAN0.CTL_NO" _
        & "   AND APTINVH1.VOUCHER_NO (+) = APTINVH8.VOUCHER_NO" _
        & "   AND NVL(APTINVH1.INV_STATUS,'X') <> 'D'"
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)

        ASCMAIN1.Progress("Stage 3")
        ASCMAIN1.sql = "" _
        & "SELECT APTINVH1.VEND_CODE, POTTRAN0.COST_CATGY_CODE, '3' STAGE, APTINVH1.INV_DATE TRAN_DATE," & vbCrLf _
        & "       NVL(POTTRAN0.TOTAL_SHP,0) TOTAL_SHP, NVL(POTTRAN0.TOTAL_INV,0) TOTAL_INV, (NVL(POTTRAN0.TOTAL_INV,0) - NVL(POTTRAN0.TOTAL_SHP,0)) TOTAL_VAR," & vbCrLf _
        & "       NVL(APTINVHC.UNITS_ONH,0) UNITS, APTINVHC.VARIANCE VARIANCE, POTTRAN0.CTL_NO," & vbCrLf _
        & "       POTTRAN0.INIT_DATE, POTTRAN0.INIT_OPER, POTTRAN0.LAST_DATE, POTTRAN0.LAST_OPER, APTINVH8.VOUCHER_NO" & vbCrLf _
        & " FROM APTINVH1, APTINVHC, APTINVH8, POTTRAN0, (SELECT SUM(REC_UNITS) UNITS FROM ICTIREC2 WHERE PO_ORDER_NO = '" & PO_ORDER_NO & "') X" & vbCrLf _
        & "  WHERE APTINVHC.WHSE_CODE = '" & WHSE_CODE & "' AND APTINVHC.LOT_NO = '" & LOT_NO & "' AND APTINVHC.LOT_SEQ_NO = " & LOT_SEQ_NO & vbCrLf _
        & "    AND APTINVHC.VOUCHER_NO = APTINVH1.VOUCHER_NO AND POTTRAN0.CTL_NO = APTINVH8.CTL_NO" & vbCrLf _
        & "    AND APTINVH1.INV_STATUS <> 'D' AND APTINVH8.VOUCHER_NO = APTINVHC.VOUCHER_NO AND APTINVH8.VOUCHER_CLNO = APTINVHC.VOUCHER_CLNO" & vbCrLf
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)
        '& "       NVL(X.UNITS,0) UNITS, DECODE(NVL(X.UNITS,0),0,0,(NVL(POTTRAN0.TOTAL_INV,0) - NVL(POTTRAN0.TOTAL_SHP,0)) / NVL(X.UNITS,0)) VARIANCE, POTTRAN0.CTL_NO," & vbCrLf

        ASCMAIN1.Progress("Stage 4")
        ASCMAIN1.sql = "" _
        & "SELECT 'WHSE ' || ICTLOTDX.WHSE_CODE VEND_CODE, 'STG' COST_CATGY_CODE, '4' STAGE, ICTLOTDX.INIT_DATE TRAN_DATE," & vbCrLf _
        & "       NVL(ICTLOTDX.ADDED_STORAGE_TOTAL,0) TOTAL_SHP, 0 TOTAL_INV, 0 TOTAL_VAR," & vbCrLf _
        & "       NVL(ICTLOTDX.UNITS,0) UNITS, NVL(ICTLOTDX.ADDED_STORAGE,0) VARIANCE, ICTLOTDX.REGISTER_XNO CTL_NO," & vbCrLf _
        & "       ICTLOTDX.INIT_DATE, ICTLOTDX.INIT_OPER, NULL LAST_DATE, NULL LAST_OPER, ICTLOTDX.VOUCHER_NO" & vbCrLf _
        & " FROM ICTLOTDX" & vbCrLf _
        & "  WHERE ICTLOTDX.WHSE_CODE = '" & WHSE_CODE & "' AND ICTLOTDX.LOT_NO = '" & LOT_NO & "' AND ICTLOTDX.LOT_SEQ_NO = " & LOT_SEQ_NO & vbCrLf _
        & "    AND (NVL(ICTLOTDX.WHSE_CHARGE_TYPE,'?') <> 'R' AND (NVL(ICTLOTDX.ORIG_TRAN_TYPE,'?') <> 'T' AND NVL(ICTLOTDX.ORIG_TRAN_TYPE,'?') <> 'P'))"
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)

        ASCMAIN1.Progress("Stage 5")
        ASCMAIN1.sql = "" _
        & "SELECT 'WHSE ' || ICTLOTDX.WHSE_CODE VEND_CODE, 'REC' COST_CATGY_CODE, '5' STAGE, ICTLOTDX.INIT_DATE TRAN_DATE," & vbCrLf _
        & "       NVL(ICTLOTDX.ADDED_STORAGE_TOTAL,0) + NVL(ICTLOTDX.SORTING,0) + NVL(ICTLOTDX.HANDLING,0) TOTAL_SHP, 0 TOTAL_INV, 0 TOTAL_VAR," & vbCrLf _
        & "       NVL(ICTLOTDX.UNITS,0) UNITS, (NVL(ICTLOTDX.ADDED_STORAGE_TOTAL,0) + NVL(ICTLOTDX.SORTING,0) + NVL(ICTLOTDX.HANDLING,0)) / NVL(ICTLOTDX.UNITS,0) VARIANCE, ICTLOTDX.REGISTER_XNO CTL_NO," & vbCrLf _
        & "       ICTLOTDX.INIT_DATE, ICTLOTDX.INIT_OPER, NULL LAST_DATE, NULL LAST_OPER, ICTLOTDX.VOUCHER_NO" & vbCrLf _
        & " FROM ICTLOTDX" & vbCrLf _
        & "  WHERE ICTLOTDX.WHSE_CODE = '" & WHSE_CODE & "' AND ICTLOTDX.LOT_NO = '" & LOT_NO & "' AND ICTLOTDX.LOT_SEQ_NO = " & LOT_SEQ_NO & vbCrLf _
        & "    AND (NVL(ICTLOTDX.WHSE_CHARGE_TYPE,'?') = 'R' AND (NVL(ICTLOTDX.ORIG_TRAN_TYPE,'?') = 'T' OR NVL(ICTLOTDX.ORIG_TRAN_TYPE,'?') = 'C'))"
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)
        '& "    AND (NVL(ICTLOTDX.WHSE_CHARGE_TYPE,'?') = 'R' AND NVL(ICTLOTDX.ORIG_TRAN_TYPE,'?') = 'T')"

        ASCMAIN1.Progress("Stage 6")
        ASCMAIN1.sql = "" _
        & "SELECT 'ADJ ' || ICTIADJ1.ADJ_NO VEND_CODE, COST_CATGY_CODE, '5' STAGE, ICTIADJ1.INIT_DATE TRAN_DATE," & vbCrLf _
        & "       (NVL(ICTIADJ1.ADJ_COST_STD,0) - NVL(ICTIADJ1.ADJ_COST_STD_ORIG,0)) * NVL(ICTIADJ1.ADJ_QTY_UNITS_ORIG,0) TOTAL_SHP, 0 TOTAL_INV, (NVL(ICTIADJ1.ADJ_COST_STD,0) - NVL(ICTIADJ1.ADJ_COST_STD_ORIG,0)) * NVL(ICTIADJ1.ADJ_QTY_UNITS_ORIG,0) TOTAL_VAR," & vbCrLf _
        & "       NVL(ICTIADJ1.ADJ_QTY_UNITS_ORIG,0) UNITS, (NVL(ICTIADJ1.ADJ_COST_STD,0) - NVL(ICTIADJ1.ADJ_COST_STD_ORIG,0)) VARIANCE, ICTIADJ1.ADJ_NO CTL_NO," & vbCrLf _
        & "       ICTIADJ1.INIT_DATE, ICTIADJ1.INIT_OPER, NULL LAST_DATE, NULL LAST_OPER, ICTIADJ1.REGISTER_XNO" & vbCrLf _
        & " FROM ICTIADJ1" & vbCrLf _
        & "  WHERE ICTIADJ1.WHSE_CODE = '" & WHSE_CODE & "' AND ICTIADJ1.LOT_NO = '" & LOT_NO & "' AND ICTIADJ1.LOT_SEQ_NO = " & LOT_SEQ_NO & vbCrLf _
        & "    AND ICTIADJ1.ADJ_TYPE = 'S'"
        ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTDL1 & " " & ASCMAIN1.sql)

        Fill_Records("ICTSTDL1")
    End Sub

    Sub Setup_LOTD6()
        If grdICTLOTD3.ActiveRow Is Nothing Then
            grdICTLOTDD.Visible = False
            Exit Sub
        End If

        grdICTLOTDD.Visible = True
        grdICTLOTDD.Text = "Daily Movement for " & grdICTLOTD3.ActiveRow.Cells("YP_LEGEND").Text
        Dim YP As String = grdICTLOTD3.ActiveRow.Cells("OPS_YYYYPP").Text

        ASCMAIN1.sql = "SELECT TRAN_DATE, OPS_YYYYPP, WHSE_CODE, LOT_NO, LOT_SEQ_NO" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'B',TRAN_QTY_CASES,0)) MTD_BEG_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'R',TRAN_QTY_CASES,0)) MTD_REC_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'S',TRAN_QTY_CASES,0)) MTD_SLS_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'X',TRAN_QTY_CASES,0)) MTD_TRC_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'A',TRAN_QTY_CASES,0)) MTD_ADJ_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'W',TRAN_QTY_CASES,0)) MTD_SLA_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'T',TRAN_QTY_CASES,0)) MTD_TRO_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'C',TRAN_QTY_CASES,0)) MTD_RTN_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'P',TRAN_QTY_CASES,0)) MTD_PUR_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'E',TRAN_QTY_CASES,0)) MTD_EOH_CASES" & vbCrLf _
        & ", 0 MTD_OUT_CASES" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'B',TRAN_QTY_UNITS,0)) MTD_BEG_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'R',TRAN_QTY_UNITS,0)) MTD_REC_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'S',TRAN_QTY_UNITS,0)) MTD_SLS_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'X',TRAN_QTY_UNITS,0)) MTD_TRC_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'A',TRAN_QTY_UNITS,0)) MTD_ADJ_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'W',TRAN_QTY_UNITS,0)) MTD_SLA_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'T',TRAN_QTY_UNITS,0)) MTD_TRO_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'C',TRAN_QTY_UNITS,0)) MTD_RTN_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'P',TRAN_QTY_UNITS,0)) MTD_PUR_UNITS" & vbCrLf _
        & ", SUM (DECODE(TRAN_TYPE,'E',TRAN_QTY_UNITS,0)) MTD_EOH_UNITS" & vbCrLf _
        & ", 0 MTD_OUT_UNITS" & vbCrLf _
        & ", 0, 0, 0" & vbCrLf _
        & ", 0, 0, 0" & vbCrLf _
        & " From ICTLOTD6" & vbCrLf _
        & " Where OPS_YYYYPP = '" & YP & "'" & vbCrLf _
        & "  and WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
        & "  and LOT_NO = '" & LOT_NO & "'" & vbCrLf _
        & "  and LOT_SEQ_NO = " & CStr(LOT_SEQ_NO) & vbCrLf _
        & " GROUP BY TRAN_DATE, OPS_YYYYPP, WHSE_CODE, LOT_NO, LOT_SEQ_NO"
        Fill_Records("ICTLOTDD", "", , ASCMAIN1.sql)

        Dim QTY As Int64 = 0
        Dim i As Integer = 0
        For Each rowICTLOTDD As DataRow In dst.Tables("ICTLOTDD").Select("", "TRAN_DATE")
            With rowICTLOTDD
                i += 1
                If i > 1 Then
                    .Item("MTD_BEG_CASES") = QTY
                End If
                QTY = Val(.Item("MTD_BEG_CASES") & "") + Val(.Item("MTD_REC_CASES") & "") _
                    + Val(.Item("MTD_SLS_CASES") & "") + Val(.Item("MTD_TRC_CASES") & "") _
                    + Val(.Item("MTD_ADJ_CASES") & "") + Val(.Item("MTD_SLA_CASES") & "") _
                    + Val(.Item("MTD_TRO_CASES") & "") + Val(.Item("MTD_RTN_CASES") & "") _
                    + Val(.Item("MTD_PUR_CASES") & "")
                .Item("MTD_EOH_CASES") = QTY
            End With
        Next
        Sort_grdColumns(grdICTLOTDD, "TRAN_DATE")
    End Sub

    Private Sub grdICTLOTD3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTLOTD3.AfterRowActivate
        Setup_LOTD6()
    End Sub

#Region "Adjustments"

    Sub Initialize_for_Adjustment(Optional ByVal release_MT_Lock As Boolean = True)

        txtADJ_REASON.Text = ""
        UltraExplorerBar1.Groups("Adjustment").Visible = Not release_MT_Lock And ASCMAIN1.USER_SECURITY_CODEs.Contains("OP")
        UltraExplorerBar1.Groups("Adjustments").Visible = release_MT_Lock And ASCMAIN1.USER_SECURITY_CODEs.Contains("OP")

        Absx1.numFor("ADJ_COST").Value = 0
        Absx1.txtFor("COST_CATGY_CODE").Text = ""
        chkAdjADJ.Checked = False
        chkAdjSTD.Checked = False


        'Qty

        numADJ_QTY.Value = 0

        ' COOL

        If Not Absx1.chkFor("COOL_COMPLIANT").Checked Then
            chkCOOL.Text = "Make this Lot available to " _
            & "CPG Orders and Orders Invoiced by MOP. " _
            & "Be sure that this Lot Sequence is COOL Compliant"
        Else
            chkCOOL.Text = "Make this Lot UNavailable to " _
            & "CPG Orders and Orders Invoiced by MOP. " _
            & "Be sure that this Lot Sequence is NOT COOL Compliant"
        End If
        chkCOOL.Checked = False

        ' Lot

        dteADJ_LOT_EXP_DATE.Value = rowICTLOTD1.Item("LOT_EXP_DATE")
        dteADJ_DATE_WHSE_ANNIV.Value = rowICTLOTD1.Item("DATE_WHSE_ANNIV")

        If release_MT_Lock Then
            ASCMAIN1.MultiTask_Release()
        Else
            If tabADJ.Tabs("Cost").Visible Then
                Calculate_New_Costs()
            End If
        End If
    End Sub

    Private Sub numADJ_QTY_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numADJ_QTY.ValueChanged
        numADJ_ONH.Value = Val(Absx1.numFor("QTY_ON_HAND").Value & "") + Val(numADJ_QTY.Value & "")
        numADJ_VAL_QTY.Value = Val(numADJ_QTY.Value & "") * PACK_FACTOR * Val(Absx1.numFor("STANDARD_COST").Value & "")
    End Sub

    Private Sub cmdAdjUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdjUpdate.Click

        txtADJ_REASON.Text = Trim(txtADJ_REASON.Text)
        If txtADJ_REASON.Text = "" And txtADJ_REASON.Visible Then
            MsgBox("Reason for Adjustment Required", MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If

        dst.Tables("ICTLOTD1_ADJ").Rows.Clear()
        Dim rowICTLOTD1_ADJ As DataRow = Fill_Record("ICTLOTD1_ADJ", New String() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Dim QTY_ON_HAND As Int64 = Val(rowICTLOTD1_ADJ.Item("QTY_ON_HAND") & "")

        For Each COLUMN_NAME As String In New String() _
            {"QTY_ON_HAND", "LOT_EXP_DATE", "DATE_WHSE_ANNIV", "STANDARD_COST", "ADJUSTED_COST"}
            If rowICTLOTD1.Item(COLUMN_NAME) & "" <> rowICTLOTD1_ADJ.Item(COLUMN_NAME) & "" Then
                MsgBox("Lot Data has Changed" & vbCrLf & vbCrLf & "You Must Refresh this Lot in order to Adjust", _
                       MsgBoxStyle.OkOnly, "Lot Data has Changed since this Lot was loaded from the Database")
                Exit Sub
            End If
        Next
        If Val(rowICTLOTD1.Item("COOL_COMPLIANT") & "") <> Val(rowICTLOTD1_ADJ.Item("COOL_COMPLIANT") & "") Then
            MsgBox("Lot Data has Changed" & vbCrLf & vbCrLf & "You Must Refresh this Lot in order to Adjust", _
                   MsgBoxStyle.OkOnly, "Lot Data has Changed since this Lot was loaded from the Database")
            Exit Sub
        End If

        Dim ADJ_QTY_CASES As Integer = 0
        Dim ADJ_QTY_UNITS As Integer = 0

        Select Case tabADJ.SelectedTab.Key
            Case "Qty"
                ADJ_QTY_CASES = Val(numADJ_QTY.Value & "")
                ADJ_QTY_UNITS = ADJ_QTY_CASES * PACK_FACTOR

                If ADJ_QTY_CASES = 0 Then
                    MsgBox("A Non-Zero Value for QTY is Required", MsgBoxStyle.OkOnly, "Cannot Update")
                    Exit Sub
                End If

                If ADJ_QTY_CASES + QTY_ON_HAND < 0 Then
                    MsgBox("The adjustment causes (or leaves) a Negative Qty On Hand", MsgBoxStyle.OkOnly, "Cannot Update")
                    Exit Sub
                End If

            Case "Cost"
                If Absx1.txtFor("COST_CATGY_CODE").Text = "" Then
                    MsgBox("You Must Specify a Cost Category Code", MsgBoxStyle.OkOnly, "Cannot Update")
                    Exit Sub
                Else
                    Dim rowPOTCATG1 As DataRow = LookUp("POTCATG1", Absx1.txtFor("COST_CATGY_CODE").Text)
                    If rowPOTCATG1 Is Nothing Then
                        MsgBox("Invalid Value Specified for Cost Category", MsgBoxStyle.OkOnly, "Cannot Update")
                        Exit Sub
                    End If
                End If

                If Val(Absx1.numFor("ADJ_ADJ").Value) < 0 _
                Or Val(Absx1.numFor("ADJ_STD").Value) < 0 Then
                    MsgBox("Cannot Result in a Negative Adjusted or Standard Cost", MsgBoxStyle.OkOnly, "Cannot Update")
                    Exit Sub
                End If

                If chkAdjADJ.Checked Or chkAdjSTD.Checked Then
                Else
                    MsgBox("You Must Adjust at least One Cost (Standard, Adjusted, or Both)", MsgBoxStyle.OkOnly, "Cannot Update")
                    Exit Sub
                End If

                If System.Math.Abs(Val(Absx1.numFor("ADJ_COST").Value)) < 0.000001 Then
                    MsgBox("No Value Entered for Cost Adjustment", MsgBoxStyle.OkOnly, "Cannot Update")
                    Exit Sub
                End If


            Case "Lot"
                Dim rowICTPROD1 As DataRow = LookUp("ICTPROD1", Absx1.txtFor("PROD_CODE").Text)
                Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
                If rowICTPROD1.Item("EXP_DATE_REQD") & "" = "1" _
                Or rowICTBRAN1.Item("EXP_DATE_REQD") & "" = "1" Then
                    If dteADJ_LOT_EXP_DATE.Value & "" = "" Then
                        MsgBox("A Value for Lot Expiration Date is Required", vbOKOnly, "Cannot Update")
                        Exit Sub
                    End If
                End If

                Dim changes_were_made As Boolean = False
                If dteADJ_LOT_EXP_DATE.Value & "" = "" Then
                    rowICTLOTD1_ADJ.Item("LOT_EXP_DATE") = DBNull.Value
                Else
                    rowICTLOTD1_ADJ.Item("LOT_EXP_DATE") = dteADJ_LOT_EXP_DATE.Value
                End If
                If dteADJ_DATE_WHSE_ANNIV.Value & "" = "" Then
                    rowICTLOTD1_ADJ.Item("DATE_WHSE_ANNIV") = DBNull.Value
                Else
                    rowICTLOTD1_ADJ.Item("DATE_WHSE_ANNIV") = dteADJ_DATE_WHSE_ANNIV.Value
                End If
                If chkCOOL.Checked Then
                    If rowICTLOTD1_ADJ.Item("COOL_COMPLIANT") & "" = "1" Then
                        rowICTLOTD1_ADJ.Item("COOL_COMPLIANT") = "0"
                    Else
                        rowICTLOTD1_ADJ.Item("COOL_COMPLIANT") = "1"
                    End If
                End If
                For Each COLUMN_NAME As String In New String() {"LOT_EXP_DATE", "DATE_WHSE_ANNIV"}
                    If rowICTLOTD1_ADJ.Item(COLUMN_NAME) & "" <> rowICTLOTD1.Item(COLUMN_NAME) & "" Then
                        changes_were_made = True
                        Exit For
                    End If
                Next
                If Val(rowICTLOTD1_ADJ.Item("COOL_COMPLIANT") & "") <> Val(rowICTLOTD1.Item("COOL_COMPLIANT") & "") Then
                    changes_were_made = True
                End If
                If Not changes_were_made Then
                    MsgBox("No Changes were Made to Lot Info", vbOKOnly, "Cannot Update")
                    Exit Sub
                End If

        End Select

        ' Perform Update

        BeginTrans()

        Dim INIT_DATE As Date = Now + ASCMAIN1.NowTSD
        rowICTLOTD1_ADJ.Item("DATE_LAST_TRAN") = INIT_DATE

        Dim msg As String = ""

        Select Case tabADJ.SelectedTab.Key

            Case "Qty"
                Dim ADJ_NO As String = ""
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1")
                Else
                    ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
                End If
                Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow
                With rowICTIADJ1
                    .Item("ADJ_NO") = ADJ_NO
                    .Item("ADJ_DATE") = INIT_DATE.Date
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOT_NO") = LOT_NO
                    .Item("LOT_SEQ_NO") = LOT_SEQ_NO
                    .Item("ADJ_QTY_CASES") = ADJ_QTY_CASES
                    .Item("ADJ_QTY_UNITS") = ADJ_QTY_UNITS
                    .Item("ADJ_COST_PUR") = rowICTLOTD1.Item("PURCHASE_COST")
                    .Item("ADJ_COST_VAL") = rowICTLOTD1.Item("VALUATION_COST")
                    .Item("ADJ_COST_ADJ") = rowICTLOTD1.Item("ADJUSTED_COST")
                    .Item("ADJ_COST_STD") = rowICTLOTD1.Item("STANDARD_COST")
                    .Item("ADJ_CON_REG_IND") = rowICTLOTD1.Item("CON_REG_IND")
                    .Item("ADJ_COST_CODE") = rowICTLOTD1.Item("COST_CODE")
                    .Item("ADJ_SP_GROUP") = rowICTLOTD1.Item("SP_GROUP")
                    .Item("ADJ_GRADE_CODE") = rowICTLOTD1.Item("GRADE_CODE")
                    .Item("ADJ_REASON") = txtADJ_REASON.Text
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = INIT_DATE
                    .Item("CON_REG_IND") = rowICTLOTD1.Item("CON_REG_IND")
                    .Item("PROD_DIV_CODE") = rowICTLOTD1.Item("PROD_DIV_CODE")

                    .Item("ADJ_QTY_CASES_ORIG") = QTY_ON_HAND
                    .Item("ADJ_QTY_UNITS_ORIG") = QTY_ON_HAND * PACK_FACTOR
                    .Item("ADJ_COST_PUR_ORIG") = rowICTLOTD1.Item("PURCHASE_COST")
                    .Item("ADJ_COST_VAL_ORIG") = rowICTLOTD1.Item("VALUATION_COST")
                    .Item("ADJ_COST_ADJ_ORIG") = rowICTLOTD1.Item("ADJUSTED_COST")
                    .Item("ADJ_COST_STD_ORIG") = rowICTLOTD1.Item("STANDARD_COST")
                    .Item("ADJ_CON_REG_IND_ORIG") = rowICTLOTD1.Item("CON_REG_IND")
                    .Item("ADJ_COST_CODE_ORIG") = rowICTLOTD1.Item("COST_CODE")
                    .Item("ADJ_SP_GROUP_ORIG") = rowICTLOTD1.Item("SP_GROUP")
                    .Item("ADJ_GRADE_CODE_ORIG") = rowICTLOTD1.Item("GRADE_CODE")
                    .Item("ADJ_TARE_WEIGHT_ORIG") = rowICTLOTD1.Item("TARE_WEIGHT")

                    .Item("ADJ_TYPE") = "Q"
                End With

                rowICTLOTD1_ADJ.Item("QTY_ON_HAND") = QTY_ON_HAND + ADJ_QTY_CASES

                Update_Record_TDA("ICTIADJ1")

                Dim rowICTLOTH1 As DataRow = dst.Tables("ICTLOTH1").NewRow
                With rowICTLOTH1
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOT_NO") = LOT_NO
                    .Item("LOT_SEQ_NO") = LOT_SEQ_NO
                    .Item("TRAN_TYPE") = "A"
                    .Item("TRAN_NO") = ADJ_NO
                    .Item("TRAN_DATE") = INIT_DATE.Date
                    .Item("TRAN_CASES") = ADJ_QTY_CASES
                    .Item("TRAN_UNITS") = ADJ_QTY_UNITS
                    .Item("FRT_RATE") = 0
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("INIT_DATE") = INIT_DATE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("STANDARD_COST") = Val(rowICTLOTD1.Item("STANDARD_COST") & "")
                    .Item("CON_REG_IND") = rowICTLOTD1.Item("CON_REG_IND")
                End With
                dst.Tables("ICTLOTH1").Rows.Add(rowICTLOTH1)
                Update_Record_TDA("ICTLOTH1")

                msg = "Adjustment Record " & ADJ_NO & " Successfully Updated"

            Case "Cost"

                Dim ADJ_NO As String = ""
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1")
                Else
                    ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
                End If
                Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow
                With rowICTIADJ1
                    .Item("ADJ_NO") = ADJ_NO
                    .Item("ADJ_DATE") = INIT_DATE.Date
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOT_NO") = LOT_NO
                    .Item("LOT_SEQ_NO") = LOT_SEQ_NO
                    .Item("ADJ_QTY_CASES") = 0
                    .Item("ADJ_QTY_UNITS") = 0

                    .Item("COST_CATGY_CODE") = Absx1.txtFor("COST_CATGY_CODE").Text

                    .Item("ADJ_COST_PUR") = rowICTLOTD1.Item("PURCHASE_COST")
                    .Item("ADJ_COST_VAL") = rowICTLOTD1.Item("VALUATION_COST")
                    .Item("ADJ_COST_ADJ") = rowICTLOTD1.Item("ADJUSTED_COST")
                    .Item("ADJ_COST_STD") = rowICTLOTD1.Item("STANDARD_COST")
                    .Item("ADJ_CON_REG_IND") = rowICTLOTD1.Item("CON_REG_IND")
                    .Item("ADJ_COST_CODE") = rowICTLOTD1.Item("COST_CODE")
                    .Item("ADJ_SP_GROUP") = rowICTLOTD1.Item("SP_GROUP")
                    .Item("ADJ_GRADE_CODE") = rowICTLOTD1.Item("GRADE_CODE")
                    .Item("ADJ_REASON") = txtADJ_REASON.Text
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = INIT_DATE
                    .Item("CON_REG_IND") = rowICTLOTD1.Item("CON_REG_IND")
                    .Item("PROD_DIV_CODE") = rowICTLOTD1.Item("PROD_DIV_CODE")

                    .Item("ADJ_QTY_CASES_ORIG") = QTY_ON_HAND
                    .Item("ADJ_QTY_UNITS_ORIG") = QTY_ON_HAND * PACK_FACTOR
                    .Item("ADJ_COST_PUR_ORIG") = rowICTLOTD1.Item("PURCHASE_COST")
                    .Item("ADJ_COST_VAL_ORIG") = rowICTLOTD1.Item("VALUATION_COST")
                    .Item("ADJ_COST_ADJ_ORIG") = rowICTLOTD1.Item("ADJUSTED_COST")
                    .Item("ADJ_COST_STD_ORIG") = rowICTLOTD1.Item("STANDARD_COST")
                    .Item("ADJ_CON_REG_IND_ORIG") = rowICTLOTD1.Item("CON_REG_IND")
                    .Item("ADJ_COST_CODE_ORIG") = rowICTLOTD1.Item("COST_CODE")
                    .Item("ADJ_SP_GROUP_ORIG") = rowICTLOTD1.Item("SP_GROUP")
                    .Item("ADJ_GRADE_CODE_ORIG") = rowICTLOTD1.Item("GRADE_CODE")
                    .Item("ADJ_TARE_WEIGHT_ORIG") = rowICTLOTD1.Item("TARE_WEIGHT")

                    .Item("ADJ_TYPE") = "S"

                    If chkAdjSTD.Checked Then
                        rowICTLOTD1_ADJ.Item("STANDARD_COST") = Val(Absx1.numFor("ADJ_STD").Value)
                        .Item("ADJ_COST_STD") = rowICTLOTD1_ADJ.Item("STANDARD_COST")
                    End If
                    If chkAdjADJ.Checked Then
                        rowICTLOTD1_ADJ.Item("ADJUSTED_COST") = Val(Absx1.numFor("ADJ_ADJ").Value)
                        .Item("ADJ_COST_ADJ") = rowICTLOTD1_ADJ.Item("ADJUSTED_COST")
                    End If
                End With

                Update_Record_TDA("ICTIADJ1")


                'Stop ' CHECK CTL_NO
                'Dim CTL_NO As String = ASCMAIN1.Next_Control_No("SOTINVHW")
                'Dim rowSOTINVHW As DataRow = dst.Tables("SOTINVHW").NewRow
                'With rowSOTINVHW
                '    .Item("CTL_NO") = CTL_NO
                '    .Item("WHSE_CODE") = WHSE_CODE
                '    .Item("LOT_NO") = LOT_NO
                '    '.Item("LOT_SEQ_NO") = LOT_SEQ_NO
                '    .Item("WD_CHG_EXP") = Val(Absx1.numFor("ADJ_COST").Value) * Val(rowICTLOTD1.Item("QTY_ON_HAND")) * PACK_FACTOR
                '    .Item("ORDR_INV_DATE") = INIT_DATE.Date

                '    .Item("VEND_CODE") = Absx1.txtFor("ADJ_VEND_CODE").Text
                '    .Item("COST_CATGY_CODE") = Absx1.txtFor("COST_CATGY_CODE").Text

                '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                '    .Item("INIT_DATE") = INIT_DATE
                '    .Item("WD_CHG_ACTUAL") = 0
                '    .Item("ORDR_DIV_CODE") = rowICTLOTD1.Item("PROD_DIV_CODE")
                '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                '    .Item("ENTRY_TYPE") = "A"
                'End With

                'Update_Record_TDA("SOTINVHW")

                msg = "Adjustment Record " & ADJ_NO & " Successfully Updated"

            Case "Lot"
                Dim rowASTAUDT1 As DataRow = Nothing
                dst.Tables("ASTAUDT1").Rows.Clear()

                For Each COLUMN_NAME As String In New String() {"LOT_EXP_DATE", "DATE_WHSE_ANNIV", "COOL_COMPLIANT"}
                    If rowICTLOTD1_ADJ.Item(COLUMN_NAME) & "" <> rowICTLOTD1.Item(COLUMN_NAME) & "" Then
                        rowASTAUDT1 = dst.Tables("ASTAUDT1").NewRow
                        rowASTAUDT1.Item("TABLE_NAME") = "ICTLOTD1"
                        rowASTAUDT1.Item("KEY_VALUE") = WHSE_CODE & ":" & LOT_NO & ":" & CStr(LOT_SEQ_NO)
                        rowASTAUDT1.Item("COLUMN_NAME") = COLUMN_NAME
                        rowASTAUDT1.Item("USER_ID") = ASCMAIN1.USER_ID
                        rowASTAUDT1.Item("INIT_DATE") = INIT_DATE
                        rowASTAUDT1.Item("OLD_VALUE") = rowICTLOTD1.Item(COLUMN_NAME)
                        rowASTAUDT1.Item("NEW_VALUE") = rowICTLOTD1_ADJ.Item(COLUMN_NAME)
                        rowASTAUDT1.Item("FM_MODE") = "E"
                        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                    End If
                Next

                Update_Record_TDA("ASTAUDT1")

                msg = "Lot Info Successfully Updated"

        End Select

        Update_Record_TDA("ICTLOTD1_ADJ")
        'TAC.ICCMAIN1.Create_ICTLOTD2_by_LOT_NO(WHSE_CODE, LOT_NO, LOT_SEQ_NO)

        For i As Integer = 0 To rowICTLOTD1_ADJ.Table.Columns.Count - 1
            rowICTLOTD1.Item(i) = rowICTLOTD1_ADJ.Item(i)
        Next i

        CommitTrans(msg)

        Initialize_for_Adjustment()

    End Sub

    Private Sub cmdAdjCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdjCancel.Click
        Initialize_for_Adjustment()
    End Sub

#End Region

    Private Sub optCGS_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCGS.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub

        Dim FORMAT As String
        If optCGS.Value = "X" Then
            FORMAT = "###,##0.00"
        Else
            FORMAT = "#,##0.0000"
        End If

        FORMAT = Replace(FORMAT, "#", "n")
        FORMAT = Replace(FORMAT, "0", "n")

        Absx1.numFor("CGS_ORDR_PRICE_NET").MaskInput = FORMAT
        Absx1.numFor("CGS_COMM").MaskInput = FORMAT
        Absx1.numFor("CGS_VA_COMM").MaskInput = FORMAT
        Absx1.numFor("CGS_OCEAN_FRT").MaskInput = FORMAT
        Absx1.numFor("CGS_UNREC_EXP").MaskInput = FORMAT
        Absx1.numFor("CGS_CGS").MaskInput = FORMAT
        Absx1.numFor("GP_DOLLARS").MaskInput = FORMAT

        Calc_CGS()
    End Sub

    Private Sub grdICTLOTHX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTLOTHX.DoubleClickRow
        If e.Row.Band.Key = "ICTLOTHX" Then
            If e.Row.Cells("CON_REG_IND").Text = "C" And InStr("S", e.Row.Cells("TRAN_TYPE").Value) <> 0 Then
                rowICTIRECX = Fill_Record("ICTIRECX", New String() {WHSE_CODE, LOT_NO, CStr(LOT_SEQ_NO)})

                If rowICTIRECX IsNot Nothing Then
                    rowICTIRECX.Item("SO_ORDER_NO") = e.Row.Cells("SO_ORDER_NO").Value
                    rowICTIRECX.Item("ORDR_INV_NO") = e.Row.Cells("TRAN_NO").Value
                    Calc_CGS()

                    frmCGS.Visible = True
                    optCGS.Visible = True
                End If
            End If
            tabLotDetails.SelectedTab = tabLotDetails.Tabs("CGS")
        End If
    End Sub

    Sub Calc_CGS()

        ' this routine is not showing dtp used at time of sale
        ' this routine is not showing interest $

        If rowICTIRECX Is Nothing Or grdICTLOTHX.ActiveRow Is Nothing Then
            Exit Sub
        End If
        Dim CGS_VA_COMM As Decimal
        Dim CGS_UNREC_EXP As Decimal
        Dim CGS_OCEAN_FRT As Decimal
        Dim CGS_CGS As Decimal
        Dim CGS_DTP As Decimal
        Dim CGS_INT_AMT As Decimal
        ' Dim UNRECOVERED_EXPENSES As Decimal
        Dim int_dtp As Decimal

        With grdICTLOTHX.ActiveRow
            Dim ORDR_PRICE_NET As Decimal = Val(.Cells("ORDR_PRICE_NET").Value & "")
            Dim CGS_COMM As Decimal = ORDR_PRICE_NET * Val(rowICTIRECX.Item("CA_COMM_PCT") & "") / 100 _
                     + Val(rowICTIRECX.Item("CA_AMT_LB") & "")
            If rowICTIRECX.Item("CA_VALUE_ADD_IND") & "" = "1" Then
                CGS_VA_COMM = ORDR_PRICE_NET * Val(rowICTIRECX.Item("CA_VA_COMM_PCT") & "") / 100
            Else
                CGS_VA_COMM = 0
            End If

            Dim CHG_CASES As Int64 = Val(.Cells("CASES").Value & "")
            Dim CHG_UNITS As Decimal = Val(.Cells("UNITS").Value & "")
            Dim ADJ_COST_EXT As Decimal = Val(.Cells("ADJ_COST_EXT").Value & "")

            Dim ADJUSTED_COST As Decimal
            If CHG_UNITS = 0 Then
                ADJUSTED_COST = 0
            Else
                ADJUSTED_COST = ADJ_COST_EXT / CHG_UNITS
            End If

            Dim CGS_EXPENSES As Decimal = ADJUSTED_COST _
                     - Val(rowICTIRECX.Item("PURCHASE_COST") & "") _
                     - Val(rowICTIRECX.Item("OCEAN_FRT") & "") _
                     + int_dtp
            CGS_EXPENSES = System.Math.Round(CGS_EXPENSES, 4)
            If rowICTIRECX.Item("CA_EXP_RECOVERY") & "" <> "1" _
            Or (rowICTIRECX.Item("CA_EXP_RECOVERY") & "" = "1" And Val(rowICTIRECX.Item("CA_EXP_CAP_PCT") & "") <> 0) Then

                If CGS_EXPENSES < 0 Then
                    CGS_UNREC_EXP = 0
                Else
                    If rowICTIRECX.Item("CA_EXP_RECOVERY") & "" = "1" _
                    And Val(rowICTIRECX.Item("CA_EXP_CAP_PCT") & "") <> 0 Then
                        If CGS_EXPENSES > ORDR_PRICE_NET * Val(rowICTIRECX.Item("CA_EXP_CAP_PCT") & "") / 100 Then
                            CGS_UNREC_EXP = CGS_EXPENSES - ORDR_PRICE_NET * Val(rowICTIRECX.Item("CA_EXP_CAP_PCT") & "") / 100
                        Else
                            CGS_UNREC_EXP = 0
                        End If
                    Else
                        CGS_UNREC_EXP = CGS_EXPENSES
                    End If
                End If
            Else
                CGS_UNREC_EXP = 0
            End If

            If rowICTIRECX.Item("CA_NO_OCEAN_FRT_RECOVER") & "" = "1" Then
                CGS_CGS = ORDR_PRICE_NET - CGS_COMM - CGS_VA_COMM + CGS_UNREC_EXP + Val(rowICTIRECX.Item("OCEAN_FRT") & "")
                CGS_OCEAN_FRT = Val(rowICTIRECX.Item("OCEAN_FRT") & "")
            Else
                CGS_CGS = ORDR_PRICE_NET - CGS_COMM - CGS_VA_COMM + CGS_UNREC_EXP
                CGS_OCEAN_FRT = 0
            End If

            If rowICTIRECX.Item("JOINT_VENTURE") & "" = "1" Then
                If Val(rowICTIRECX.Item("CA_COMM_PCT") & "") = 0 Then
                    CGS_CGS = ADJUSTED_COST
                    CGS_UNREC_EXP = 0
                End If
                CGS_CGS = CGS_CGS + (ORDR_PRICE_NET - CGS_CGS) * (100 - Val(rowICTIRECX.Item("JOINT_VENTURE_PCT") & "")) / 100
            End If

            Dim GP_DOLLARS As Decimal = (ORDR_PRICE_NET - CGS_CGS)

            If optCGS.Value = "U" Then
                Absx1.numFor("CGS_ORDR_PRICE_NET").Value = ORDR_PRICE_NET
                Absx1.numFor("CGS_COMM").Value = CGS_COMM
                Absx1.numFor("CGS_VA_COMM").Value = CGS_VA_COMM
                Absx1.numFor("CGS_OCEAN_FRT").Value = CGS_OCEAN_FRT
                Absx1.numFor("CGS_UNREC_EXP").Value = CGS_UNREC_EXP
                Absx1.numFor("CGS_CGS").Value = CGS_CGS
                Absx1.numFor("GP_DOLLARS").Value = System.Math.Round(GP_DOLLARS, 4)
            Else
                Absx1.numFor("CGS_ORDR_PRICE_NET").Value = System.Math.Round(ORDR_PRICE_NET * CHG_UNITS, 2)
                Absx1.numFor("CGS_COMM").Value = System.Math.Round(CGS_COMM * CHG_UNITS, 2)
                Absx1.numFor("CGS_VA_COMM").Value = System.Math.Round(CGS_VA_COMM * CHG_UNITS, 2)
                Absx1.numFor("CGS_OCEAN_FRT").Value = System.Math.Round(CGS_OCEAN_FRT * CHG_UNITS, 2)
                Absx1.numFor("CGS_UNREC_EXP").Value = System.Math.Round(CGS_UNREC_EXP * CHG_UNITS, 2)
                Absx1.numFor("CGS_CGS").Value = System.Math.Round(CGS_CGS * CHG_UNITS, 2)
                Absx1.numFor("GP_DOLLARS").Value = System.Math.Round(GP_DOLLARS * CHG_UNITS, 2)
            End If

            Absx1.numFor("CGS_CASES").Value = CHG_CASES
            Absx1.numFor("CGS_UNITS").Value = CHG_UNITS

            Absx1.numFor("CGS_DTP").Value = CGS_DTP
            Absx1.numFor("CGS_INT_AMT").Value = CGS_INT_AMT
            Absx1.numFor("CGS_EXPENSES").Value = CGS_EXPENSES
            Absx1.numFor("CGS_ADJUSTED_COST").Value = ADJUSTED_COST


            Dim GP_PERCENT As Decimal
            If CHG_UNITS * ORDR_PRICE_NET <> 0 Then
                GP_PERCENT = GP_DOLLARS / ORDR_PRICE_NET * 100
            Else
                GP_PERCENT = 0
            End If

            Absx1.numFor("GP_PERCENT").Value = GP_PERCENT

        End With
    End Sub

#Region "grdASTATTA2"

    Private Sub grdASTATTA2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTATTA2.AfterRowActivate
        If grdASTATTA2.ActiveRow.IsDataRow Then
            With grdASTATTA2.DisplayLayout.Bands(0)
                .Columns("ATTACHMENT_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("ATTACHMENT_TYPE").CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        End If
    End Sub

    Private Sub grdASTATTA2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTATTA2.ClickCellButton
        Dim ATTACHMENT_NO As String = grdASTATTA2.ActiveRow.Cells("ATTACHMENT_NO").Text
        Dim ATTACHMENT_EXT As String = grdASTATTA2.ActiveRow.Cells("ATTACHMENT_EXT").Text.ToUpper
        Call ASCMAIN1.Launch_Attachment(ATTACHMENT_NO, ATTACHMENT_EXT)
    End Sub


    Private Sub grdICTLOTDG_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTATTA2.InitializeRow
        Dim FILENAME As String = ""

        Select Case e.Row.Cells("ATTACHMENT_EXT").Text.ToUpper
            Case "TXT"
            Case "XLS", "XLSX"
                FILENAME = "EXCEL"
            Case "PDF"
                FILENAME = "PDF"
            Case "MSG"
                FILENAME = "MAIL"
            Case "DOC", "DOCX"
                FILENAME = "WORD"
            Case Else
        End Select
        If FILENAME <> "" Then
            e.Row.Cells("ATTACHMENT_EXT").ButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", FILENAME)
        End If

        Dim i As Int32 = e.Row.ListIndex
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Select Case dst.Tables("ASTATTA2").Rows(i).RowState
            Case DataRowState.Added
                'e.Row.Appearance.BackColor = Color.LightGreen
                e.Row.RowSelectorAppearance.BackColor = Color.LightGreen
                e.Row.RowSelectorAppearance.BackColor2 = Color.Green
            Case DataRowState.Modified
                'e.Row.Appearance.BackColor = Color.LightSkyBlue
                e.Row.RowSelectorAppearance.BackColor = Color.LightSkyBlue
                e.Row.RowSelectorAppearance.BackColor2 = Color.Blue
        End Select

    End Sub
#End Region

    '#Region "VB6"
    '    Private Const SAVE_PATH As String = "C:\MY_SAVE_DIR\"
    '    Private Const TEMP_PATH As String = "C:\MY_TEMP_DIR\"
    '    Private Const STR_FileGroupDescriptor As String = "FileGroupDescriptor"
    '    Private Const OUTLOOK_CLASS As String = "rctrl_renwnd32"
    '    Private Const OUTLOOK_EXPL As String = "Explorer"
    '    Private Const OUTLOOK_INSP As String = "Inspector"


    '    Private Sub Image1_OLEDragDrop(ByVal data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    '        On Error GoTo finish_line

    '        Dim i As Integer
    '        Dim FAs(3, 0) As String

    '        'Check if data is dragged from Outlook
    '        Dim sWindowClass As String
    '        sWindowClass = GetWindowClass(GetForegroundWindow)
    '        If InStr(1, sWindowClass, OUTLOOK_CLASS) <> 0 Then
    '            FAs = Process_Outlook_Attachments(data)
    '        Else
    '            If data.GetFormat(vbCFFiles) = True Then
    '                If data.Files.Count > 0 Then
    '                    ReDim FAs(3, data.Files.Count)
    '                    For i = 1 To data.Files.Count
    '                        FAs(0, i) = data.Files.item(i)
    '                    Next i
    '                End If
    '            End If
    '        End If

    '        If UBound(FAs, 2) > 0 Then

    '            Dim FS As FileSystemObject
    '            FS = New FileSystemObject

    '            Dim FOLDER As String
    '            FOLDER = AS_PARM_ARCHIVE & "LOTATTACHMENTS\"

    '            Dim dynICWLOTDG As Recordset
    '            dynICWLOTDG = AccD.OpenRecordset("ICWLOTDG", dbOpenDynaset)

    '            Dim INIT_DATE As Date
    '            INIT_DATE = Now + NowTSD

    '            For i = 1 To UBound(FAs, 2)
    '                With dynICWLOTDG

    '                    Dim ORIG_FILE_NAME As String
    '                    ORIG_FILE_NAME = FAs(0, i)
    '                    Dim FILE_TYPE As String

    '                    Dim EMP_TYPE As String
    '                    EMP_TYPE = "No Type Definition"


    '                    Dim FSO As File
    '                    FSO = FS.GetFile(ORIG_FILE_NAME)
    '                    Dim zz() As String
    '                    zz = Split(FSO.ShortName, ".")

    '                    FILE_TYPE = UCase(zz(UBound(zz)))

    '                    Dim FILE_DESC As String
    '                    FILE_DESC = FAs(1, i)

    '                    If FILE_DESC = "" Then
    '                        FILE_DESC = FSO.ShortName
    '                        'FILE_DESC = "{Enter a Description Here}"
    '                    End If

    '                    Dim FILE_ORIGINATOR As String
    '                    FILE_ORIGINATOR = FAs(2, i)
    '                    If FILE_ORIGINATOR = "" Then
    '                        FILE_ORIGINATOR = UserID
    '                    End If

    '                    Dim FILE_DATETIME As Date
    '                    If FAs(3, i) <> "" Then
    '                        FILE_DATETIME = FAs(3, i)
    '                    Else
    '                        FILE_DATETIME = FSO.DateLastModified
    '                    End If

    '                    If FILE_TYPE = "XLS" Or FILE_TYPE = "XLSX" Or FILE_TYPE = "MSG" Or FILE_TYPE = "HTM" Or FILE_TYPE = "BMP" Or FILE_TYPE = "DOC" Or FILE_TYPE = "PDF" Or FILE_TYPE = "JPG" Or FILE_TYPE = "GIF" Or FILE_TYPE = "PNG" Then

    '                        Dim FILE_NAME As String
    '                        FILE_NAME = CTLNO("ICTLOTDG.FILE_NAME")

    '                        .AddNew()
    '                        .Fields("WHSE_CODE").value = WHSE_CODE
    '                        .Fields("LOT_NO").value = LOT_NO
    '                        .Fields("LOT_SEQ_NO").value = LOT_SEQ_NO
    '                        .Fields("FILE_NAME").value = FILE_NAME
    '                        .Fields("FILE_DESC").value = FILE_DESC
    '                        .Fields("ORIG_FILE_NAME").value = ORIG_FILE_NAME
    '                        .Fields("FILE_TYPE").value = FILE_TYPE
    '                        .Fields("FILE_TYPE_DESC").value = FSO.Type
    '                        .Fields("INIT_OPER").value = UserID
    '                        .Fields("INIT_DATE").value = INIT_DATE
    '                        .Fields("RECEIPT_NO").value = absx1.txtfor("RECEIPT_NO")).Text
    '                        .Fields("RECEIPT_LNO").value = absx1.txtfor("RECEIPT_LNO")).Text
    '                        .Fields("FILE_ORIGINATOR").value = FILE_ORIGINATOR
    '                        .Fields("FILE_DATETIME").value = FILE_DATETIME
    '                        .Fields("EMP_TYPE").value = EMP_TYPE

    '                        dynICTLOTDG.AddNew()
    '                        Dim f As Integer
    '                        For f = 0 To dynICWLOTDG.Fields.Count - 1
    '                            dynICTLOTDG.Fields(f).value = dynICWLOTDG.Fields(f).value
    '                        Next f
    '                        dynICTLOTDG.Update()

    '                        .Update()

    '                        FS.CopyFile(ORIG_FILE_NAME, FOLDER & FILE_NAME & "." & FILE_TYPE)
    '                    Else
    '                        MsgBox("Unsupported File Type (" & FILE_TYPE & ")", vbOKOnly, "Cannot Attach file " & ORIG_FILE_NAME)
    '                    End If
    '                End With

    '            Next i

    '            dynICWLOTDG.Close()
    '            datICWLOTDG.Refresh()
    '            FS = Nothing

    '        End If

    'finish_line:
    '        On Error GoTo 0


    '    End Sub

    '    Private Function Process_Outlook_Attachments(ByVal data As DataObject) As String()

    '        On Error GoTo ErrHandler
    '        Dim oOutlook As Object
    '        Dim oExplorer As Object
    '        Dim oInspector As Object
    '        Dim oSelection As Object
    '        Dim oItem As Outlook.MailItem
    '        Dim oAttachments() As Object
    '        Dim sFileNames() As String
    '        Dim iCount As Integer
    '        Dim iCounter As Integer
    '        Dim i As Integer, j As Integer
    '        Dim sName As String, sExt As String
    '        Dim bEqualNames As Boolean
    '        Dim lCF_FILE As Long
    '        Dim CF_FILE As Integer
    '        Dim bData() As Byte
    '        Dim sNameString As String
    '        Dim sNames() As String

    '        Dim FAs() As String
    '        ReDim FAs(3, 0)
    '        Dim FILE_NAME As String
    '        Dim q As Integer

    '        'Get ID of FileGroupDescriptor Data
    '        lCF_FILE = RegisterClipboardFormat(STR_FileGroupDescriptor)
    '        MoveMemory(CF_FILE, lCF_FILE, 2)

    '        'If data not available exit sub
    '        If Not data.GetFormat(CF_FILE) Then Exit Function

    '        'Assign values to variables
    '        iCounter = 1
    '        bEqualNames = False
    '        oOutlook = GetObject(, "Outlook.Application")
    '        Select Case TypeName(oOutlook.ActiveWindow)
    '            Case OUTLOOK_EXPL
    '                oExplorer = oOutlook.ActiveExplorer
    '                oSelection = oExplorer.Selection
    '                oItem = oSelection(1)

    '                q = UBound(FAs, 2) + 1

    '                'oItem.SaveAs App.Path & "\temp\mailitem.htm", olHTML
    '                'ReDim Preserve FAs(q)
    '                'FAs(q) = App.Path & "\temp\mailitem.htm"

    '                FILE_NAME = "mailitem.msg"
    '                oItem.SaveAs(App.Path & "\temp\" & FILE_NAME)
    '                ReDim Preserve FAs(3, q)
    '                FAs(0, q) = App.Path & "\temp\" & FILE_NAME
    '                FAs(1, q) = oItem.Subject
    '                FAs(2, q) = oItem.SenderName
    '                FAs(3, q) = oItem.SentOn

    '                Process_Outlook_Attachments = FAs
    '                Exit Function
    '            Case OUTLOOK_INSP
    '                oInspector = oOutlook.ActiveInspector
    '                oItem = oInspector.CurrentItem
    '        End Select

    '        iCount = oItem.attachments.Count

    '        'Fill array with attachment filenames
    '        'and check if there are attachments with equal filenames
    '    ReDim sFileNames(1 To iCount, 1)
    '    ReDim oAttachments(1 To iCount)

    '        For i = 1 To iCount
    '            oAttachments(i) = oItem.attachments(i)
    '            sFileNames(i, 0) = oAttachments(i).FileName
    '            sFileNames(i, 1) = oAttachments(i).FileName
    '            iCounter = 1
    '        Next i


    '        'Retrieve filedescriptor data and strip Null-chars -> convert to FileNames
    '        bData() = data.GetData(CF_FILE)
    '        sNameString = StrConv(bData, vbUnicode)
    '        Do While InStr(1, sNameString, vbNullChar & vbNullChar) > 0
    '            sNameString = Replace(sNameString, vbNullChar & vbNullChar, vbNullChar)
    '        Loop
    '        sNames = Split(sNameString, vbNullChar)

    '        'Query files and save in save directory
    '        For i = 1 To UBound(sNames) - 1
    '            For j = 1 To UBound(oAttachments)
    '                If sNames(i) Like oAttachments(j).FileName Then
    '                    q = UBound(FAs, 2) + 1
    '                    FILE_NAME = sNames(i)
    '                    oAttachments(j).SaveAsFile(App.Path & "\temp\" & FILE_NAME)
    '                    ReDim Preserve FAs(3, q)
    '                    FAs(0, q) = App.Path & "\temp\" & FILE_NAME
    '                    FAs(1, q) = FILE_NAME
    '                    FAs(2, q) = oItem.SenderName
    '                    FAs(3, q) = oItem.SentOn ' ReceivedTime
    '                    Exit For
    '                End If
    '            Next j
    '        Next i

    '        Process_Outlook_Attachments = FAs

    '        Exit Function
    'ErrHandler:
    '        MsgBox(Err.Description)
    '    End Function


    '    Private Sub List1_OLEDragDrop(ByVal data As DataObject, ByVal Effect As Long, ByVal Button As _
    '    Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '        On Error GoTo ErrHandler
    '        Dim oOutlook As Object
    '        Dim oExplorer As Object
    '        Dim oInspector As Object
    '        Dim oSelection As Object
    '        Dim oItem As Object
    '        Dim oAttachments() As Object
    '        Dim sFileNames() As String
    '        Dim iCount As Integer
    '        Dim iCounter As Integer
    '        Dim i As Integer, j As Integer
    '        Dim sName As String, sExt As String
    '        Dim bEqualNames As Boolean
    '        Dim lCF_FILE As Long
    '        Dim CF_FILE As Integer
    '        Dim bData() As Byte
    '        Dim sNameString As String
    '        Dim sNames() As String
    '        Dim sWindowClass As String


    '        'Check if data is dragged from Outlook
    '        sWindowClass = GetWindowClass(GetForegroundWindow)
    '        If InStr(1, sWindowClass, OUTLOOK_CLASS) = 0 Then Exit Sub

    '        'Get ID of FileGroupDescriptor Data
    '        lCF_FILE = RegisterClipboardFormat(STR_FileGroupDescriptor)
    '        MoveMemory(CF_FILE, lCF_FILE, 2)

    '        'If data not available exit sub
    '        If Not data.GetFormat(CF_FILE) Then Exit Sub

    '        'Make temp and save dir
    '        MakeDir(TEMP_PATH)
    '        MakeDir(SAVE_PATH)

    '        'Assign values to variables
    '        iCounter = 1
    '        bEqualNames = False
    '        oOutlook = GetObject(, "Outlook.Application")
    '        Select Case TypeName(oOutlook.ActiveWindow)
    '            Case OUTLOOK_EXPL
    '                oExplorer = oOutlook.ActiveExplorer
    '                oSelection = oExplorer.Selection
    '                oItem = oSelection(1)
    '            Case OUTLOOK_INSP
    '                oInspector = oOutlook.ActiveInspector
    '                oItem = oInspector.CurrentItem
    '        End Select

    '        iCount = oItem.attachments.Count

    '        'Fill array with attachment filenames
    '        'and check if there are attachments with equal filenames
    '    ReDim sFileNames(1 To iCount, 1)
    '    ReDim oAttachments(1 To iCount)

    '        For i = 1 To iCount
    '            oAttachments(i) = oItem.attachments(i)
    '            sFileNames(i, 0) = oAttachments(i).FileName
    '            sFileNames(i, 1) = oAttachments(i).FileName
    '            iCounter = 1
    'StartOver:
    '            For j = 1 To i - 1
    '                If sFileNames(j, 1) Like sFileNames(i, 1) Then
    '                    bEqualNames = True
    '                    iCounter = iCounter + 1
    '                    SplitFileName(sFileNames(i, 0), sName, sExt)
    '                    sFileNames(i, 1) = sName & " (" & iCounter & ")" & sExt
    '                    GoTo StartOver
    '                End If
    '            Next j
    '        Next i

    '        'If equal filenames exist ask to rename them or not and exit sub
    '        'Otherwise continue DragDrop
    '        If bEqualNames Then
    '            If MsgBox("This item contains attachments with equal names. " & _
    '                    "Do you want to rename these items", vbQuestion + vbYesNo, _
    '                    "Drag attachments to listbox") = vbNo Then Exit Sub

    '            For i = 1 To iCount
    '                If Not sFileNames(i, 0) Like sFileNames(i, 1) Then
    '                    oAttachments(i).SaveAsFile(TEMP_PATH & sFileNames(i, 1))
    '                    oAttachments(i).Delete()
    '                    oItem.attachments.Add(TEMP_PATH & sFileNames(i, 1))
    '                    Kill(TEMP_PATH & sFileNames(i, 1))
    '                End If
    '            Next
    '            Exit Sub
    '        End If

    '        'Retrieve filedescriptor data and strip Null-chars -> convert to FileNames
    '        bData() = data.GetData(CF_FILE)
    '        sNameString = StrConv(bData, vbUnicode)
    '        Do While InStr(1, sNameString, vbNullChar & vbNullChar) > 0
    '            sNameString = Replace(sNameString, vbNullChar & vbNullChar, vbNullChar)
    '        Loop
    '        sNames = Split(sNameString, vbNullChar)

    '        'Query files and save in save directory
    '        For i = 1 To UBound(sNames) - 1
    '            For j = 1 To UBound(oAttachments)
    '                If sNames(i) Like oAttachments(j).FileName Then
    '                    oAttachments(j).SaveAsFile(SAVE_PATH & sNames(i))
    '                    List1.AddItem(sNames(i))
    '                    Exit For
    '                End If
    '            Next j
    '        Next i
    '        Exit Sub
    'ErrHandler:
    '        MsgBox(Err.Description)
    '    End Sub

    '    Private Sub MakeDir(ByVal sDir As String)
    '        On Error Resume Next
    '        MkDir(sDir)
    '    End Sub

    '    Private Sub SplitFileName(ByVal sFullName As String, ByRef sShortName As String, _
    '                                ByRef sExtension As String)
    '        Dim sParts() As String
    '        sParts = Split(sFullName, ".")
    '        If UBound(sParts) = 0 Then
    '            sShortName = sParts(0)
    '            sExtension = ""
    '        Else
    '            sExtension = "." & sParts(UBound(sParts))
    '            sShortName = Left(sFullName, Len(sFullName) - Len(sExtension))
    '        End If
    '    End Sub

    '    ' Return the class name of the specified window
    '    Function GetWindowClass(ByVal hWnd As Long) As String
    '        Dim sClass As String
    '        sClass = Space$(256)
    '        GetClassName(hWnd, sClass, 255)
    '        GetWindowClass = Left$(sClass, InStr(sClass, vbNullChar) - 1)
    '    End Function

    '#End Region

    Sub Maintain_Lots(ByVal mode As Boolean)
        UltraExplorerBar1.Groups("Screen Control").Visible = Not mode
        UltraExplorerBar1.Groups("Find Lots Options").Visible = Not mode
        UltraExplorerBar1.Groups("Maintain Lots").Visible = mode

        spl.Panel1Collapsed = mode

        With grdICTLOTDF.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.CellActivation = UltraWinGrid.Activation.AllowEdit And mode Then
                    gcol.CellAppearance.BackColor = Color.Yellow
                Else
                    gcol.CellAppearance.BackColor = Color.Empty
                End If
            Next
        End With

        If mode Then
            grdICTLOTDF.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdICTLOTDF.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Else
            grdICTLOTDF.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdICTLOTDF.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect
        End If
    End Sub

    Private Sub chkAdjSTD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAdjSTD.CheckedChanged
        Calculate_New_Costs()
    End Sub

    Private Sub chkAdjADJ_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAdjADJ.CheckedChanged
        Calculate_New_Costs()
    End Sub

    Sub Calculate_New_Costs()
        If chkAdjADJ.Checked Then
            Absx1.numFor("ADJ_ADJ").Value = Val(Absx1.numFor("ADJUSTED_COST").Value) + Val(Absx1.numFor("ADJ_COST").Value)
        Else
            Absx1.numFor("ADJ_ADJ").Value = Absx1.numFor("ADJUSTED_COST").Value
        End If

        If chkAdjSTD.Checked Then
            Absx1.numFor("ADJ_STD").Value = Val(Absx1.numFor("STANDARD_COST").Value) + Val(Absx1.numFor("ADJ_COST").Value)
        Else
            Absx1.numFor("ADJ_STD").Value = Absx1.numFor("STANDARD_COST").Value
        End If
    End Sub

    Sub Reset_Lot_Qty_Committed( _
    ByVal WHSE_CODE As String, _
    ByVal LOT_NO As String, _
    ByVal LOT_SEQ_NO As Int32, _
    ByVal QTY_COMMITTED As Int64, _
    ByVal QTY_COMMITTED_old As Int64)

        For Each TABLE_NAME As String In New String() {"ICTLOTD1", "ICTLOTD2"}
            ASCMAIN1.sql = "Update " & TABLE_NAME & " Set QTY_COMMITTED = " & CStr(QTY_COMMITTED) _
            & " where WHSE_CODE = :PARM1 and LOT_NO = :PARM2 and LOT_SEQ_NO = :PARM3"

            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVN", New Object() {WHSE_CODE, LOT_NO, LOT_SEQ_NO})
        Next

        Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
        rowASTAUDT1.Item("TABLE_NAME") = "ICTLOTD1"
        rowASTAUDT1.Item("KEY_VALUE") = WHSE_CODE & ":" & LOT_NO & ":" & CStr(LOT_SEQ_NO)
        rowASTAUDT1.Item("COLUMN_NAME") = "QTY_COMMITTED"
        rowASTAUDT1.Item("USER_ID") = ASCMAIN1.USER_ID
        rowASTAUDT1.Item("INIT_DATE") = DATETIME_STAMP
        rowASTAUDT1.Item("OLD_VALUE") = QTY_COMMITTED_old
        rowASTAUDT1.Item("NEW_VALUE") = QTY_COMMITTED
        rowASTAUDT1.Item("FM_MODE") = "E"
        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)

        Update_Record_TDA("ASTAUDT1")

    End Sub

    Private Sub tabLotDetails_SelectedTabChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabLotDetails.SelectedTabChanged
        If tabLotDetails.SelectedTab.Key = "Attachments" Then
            Load_Attachments()
        End If
    End Sub


End Class