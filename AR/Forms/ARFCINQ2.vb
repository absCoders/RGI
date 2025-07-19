Public Class ARFCINQ2
     
    Dim AGED_TOTALS() As Decimal
    Dim DUE_TOTALS() As Decimal
    Dim AGE_DAYS(4) As Integer
    Dim DUE_DAYS(4) As Integer
    Dim DUE_DATE(4) As String
    Dim AGE_DATE(4) As String

    Dim rowARTCUST1 As DataRow
    Dim rowARTCUST1_orig As DataRow
    Dim TOTAL_DUE As Double
    Dim LogEditingEnabled As Boolean = False

    Dim sqlSOTORDR1 As String = ""
    Dim SATCUSTS As String
    Dim sqlSATCUSTS As String = ""
    Dim sqlSATCUSTI As String = ""
    Dim sqlARTOPEN1 As String = ""
 
    Dim showZBI As Boolean = True

    Dim ARTCUST0 As String = ""
    Dim ARTCUSTS As String = ""
    Dim ARTCUSTX As String = ""

    Dim sqlSOTAEBP1 As String = ""
    Dim sqlSOTAEBP2 As String = ""

    Dim INV_NUM_column As String = "INV_NUM"

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            INV_NUM_column = "INV_NO"
        End If

        With dst

            Get_PARM("SOTPARM1")
            Get_PARM("ARTPARM1")
            Get_PARM("ASTPARM1")

            For i As Integer = 2 To 4
                AGE_DAYS(i) = Val(ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_" & CStr(i)) & "")
                DUE_DAYS(i) = Val(ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_" & CStr(i)) & "")
            Next

            ASCMAIN1.sql = "SELECT '999999' OPS_YYYYPP, AR_PARM_AGE_CATG_DESC_1, AR_PARM_AGE_CATG_DESC_2, AR_PARM_AGE_CATG_DESC_3, AR_PARM_AGE_CATG_DESC_4" _
            & ", AR_PARM_DUE_CATG_DESC_1, AR_PARM_DUE_CATG_DESC_2, AR_PARM_DUE_CATG_DESC_3, AR_PARM_DUE_CATG_DESC_4" _
            & " FROM ARTPARM1 WHERE AR_PARM_KEY ='Z'"
            Create_TDA(.Tables.Add, "ARTSTMT1_DESC", "**", 0, False, "", 1)

            Dim DUE_DATE_ORA(4) As String
            For i As Integer = 1 To 4
                Dim PRD_END_DATE As Date = Now.Date.AddDays(-1 * DUE_DAYS(i))
                DUE_DATE(i) = "'" & Format(PRD_END_DATE, "MM/dd/yyyy") & "'"
                DUE_DATE_ORA(i) = "'" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
            Next
            Dim AGE_DATE_ORA(4) As String
            For i As Integer = 1 To 4
                Dim PRD_END_DATE As Date = Now.Date.AddDays(-1 * AGE_DAYS(i))
                AGE_DATE(i) = "'" & Format(PRD_END_DATE, "MM/dd/yyyy") & "'"
                AGE_DATE_ORA(i) = "'" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
            Next

            ASCMAIN1.sql = "Select CUST_CODE from ARTCUST1 where CUST_CODE = ''"
            ARTCUST0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST0 & " Add Primary Key (CUST_CODE)")

            ASCMAIN1.sql = "SELECT ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.CUST_ADDR1" _
            & ", ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_ZIP_CODE" _
            & ", ARTCUST1.CUST_PHONE, ARTCUST1.CUST_EXT, ARTCUST1.CUST_FAX" _
            & ", ARTCUST1.SREP_CODE, ARTCUST1.CUST_BILL_TO_CUST, ARTCUST1.CUST_CREDIT_GROUP_CUST" _
            & " from ARTCUST1," & ARTCUST0 & " ARTCUST0 " _
            & " where ARTCUST1.CUST_CODE = ARTCUST0.CUST_CODE"
            Create_TDA(.Tables.Add, "ARTCUST0", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select CUST_CODE from ARTCUST1 where CUST_CODE = ''"
            ARTCUSTS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTCUSTS & " Add Primary Key (CUST_CODE)")

            ASCMAIN1.sql = "SELECT ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.CUST_ADDR1" _
            & ", ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_ZIP_CODE" _
            & ", ARTCUST1.CUST_PHONE, ARTCUST1.CUST_EXT, ARTCUST1.CUST_FAX" _
            & ", ARTCUST1.SREP_CODE, ARTCUST1.CUST_BILL_TO_CUST, ARTCUST1.CUST_CREDIT_GROUP_CUST" _
            & " from ARTCUST1," & ARTCUSTS & " ARTCUSTS " _
            & " where ARTCUST1.CUST_CODE = ARTCUSTS.CUST_CODE"
            Create_TDA(.Tables.Add, "ARTCUSTS", "**", 0, False, "", 1)

            sqlARTOPEN1 = "SELECT ARTOPEN1.*" _
            & ", DECODE(ARTOPEN1.OPS_YYYYPP_PAID,NULL,ARTOPEN1.DATE_PAID,TRUNC(SYSDATE)) - ARTOPEN1.INV_DATE DAYS" _
            & ", (CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(2) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(1) & " THEN '1' ELSE" _
            & "  CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(3) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(2) & " THEN '2' ELSE" _
            & "  CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(4) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(3) & " THEN '3' ELSE" _
            & "  '4' END END END) AGE_BUCKET" _
            & ", TO_NUMBER(TRUNC(SYSDATE) - ARTOPEN1.INV_DATE) AGE" _
            & ", (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(2) & " THEN '1' ELSE" _
            & "  CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(3) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(2) & " THEN '2' ELSE" _
            & "  CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(4) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(3) & " THEN '3' ELSE" _
            & "  '4' END END END) DUE_BUCKET" _
            & ", TO_NUMBER(TRUNC(SYSDATE) - ARTOPEN1.INV_DUE_DATE) DUE"
            ASCMAIN1.sql = sqlARTOPEN1 _
            & " from ARTOPEN1," & ARTCUST0 & " ARTCUST0 where ARTOPEN1.CUST_CODE = ARTCUST0.CUST_CODE"
            Create_TDA(.Tables.Add, "ARTOPEN1", "**", 0, False, "", 3)
            '.Tables("ARTOPEN1").Columns("AGE").DataType = GetType(System.Int32)

            '& ", DECODE(DTP,NULL,TRUNC(SYSDATE) - ARTOPEN1.INV_DATE,NULL) AGE" _


            ASCMAIN1.sql = "Select ARTPYMT2.*, ARTPYMT1.PYMT_BATCH_DATE" _
            & ", ARTPYMT1.BANK_CODE, ARTPYMT1.PYMT_APPL_ONLY" _
            & ", ARTPYMT1.PYMT_SOURCE, ARTPYMT1.OPS_YYYYPP" _
            & " from ARTPYMT2,ARTPYMT1," & ARTCUST0 & " ARTCUST0 " _
            & " where ARTPYMT2.CUST_CODE = ARTCUST0.CUST_CODE" _
            & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & "   and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
            & "   and ARTPYMT1.OPS_YYYYPP >= :PARM1"
            Create_TDA(.Tables.Add, "ARTPYMT2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT ARTPYMT3.PYMT_BATCH_NO, ARTPYMT3.PYMT_BATCH_LNO, ARTPYMT3.PYMT_BATCH_ILNO " _
            & ", ARTPYMT3.INV_TYPE, ARTPYMT3.INV_NUM, ARTPYMT3.INV_DATE, ARTPYMT3.INV_DUE_DATE" _
            & ", ARTPYMT3.INV_BALANCE, ARTPYMT3.INV_PMT PMT" _
            & ", NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0) DED, ARTPYMT3.INV_BALANCE_NEW" _
            & ", INV_CUST_PO CUST_REF, ARTPYMT3.REASON_CODE, ARTREAS1.REASON_DESC" _
            & ", ARTPYMT3.SEG2_CODE, ARTPYMT3.SEG3_CODE, ARTPYMT3.SEG4_CODE" _
            & " FROM ARTPYMT3, ARTREAS1" _
            & " WHERE ARTREAS1.REASON_CODE (+) = ARTPYMT3.REASON_CODE " _
            & " AND ARTPYMT3.PYMT_BATCH_NO = :PARM1 AND ARTPYMT3.PYMT_BATCH_LNO = :PARM2" _
            & " UNION" _
            & " SELECT ARTPYMT4.PYMT_BATCH_NO, ARTPYMT4.PYMT_BATCH_LNO, ARTPYMT4.PYMT_BATCH_GLNO PYMT_BATCH_ILNO" _
            & ", 'G' INV_TYPE, 'GL W/Off' INV_NUM, NULL INV_DATE, NULL INV_DUE_DATE" _
            & ", NULL INV_BALANCE, NULL PMT" _
            & ", NVL(ARTPYMT4.GL_DIST_AMT,0) DED, NULL INV_BALANCE_NEW" _
            & ", ARTPYMT4.GL_DIST_REF CUST_REF, ARTPYMT4.ACCT_CODE REASON_CODE, GLTACCT1.ACCT_DESC REASON_DESC" _
            & ", ARTPYMT4.SEG2_CODE, ARTPYMT4.SEG3_CODE, ARTPYMT4.SEG4_CODE" _
            & " FROM ARTPYMT4, GLTACCT1" _
            & " WHERE GLTACCT1.ACCT_CODE (+) = ARTPYMT4.ACCT_CODE" _
            & " AND ARTPYMT4.PYMT_BATCH_NO = :PARM1 AND ARTPYMT4.PYMT_BATCH_LNO = :PARM2" _
            & " UNION" _
            & " SELECT ARTPYMT5.PYMT_BATCH_NO, ARTPYMT5.PYMT_BATCH_LNO, ARTPYMT5.PYMT_BATCH_DLNO PYMT_BATCH_ILNO " _
            & ", DECODE (ARTPYMT5.CHARGEBACK_IND,'1',ARTPYMT5.INV_TYPE_CB,'X') INV_TYPE" _
            & ", DECODE (ARTPYMT5.CHARGEBACK_IND,'1',ARTPYMT5.CHARGEBACK_NO,'Deduct OK') INV_NUM" _
            & ", NULL INV_DATE, NULL INV_DUE_DATE" _
            & ", NULL INV_BALANCE, NULL PMT" _
            & ", NVL(ARTPYMT5.GL_DIST_AMT,0) DED, NULL INV_BALANCE_NEW" _
            & ", ARTPYMT5.CUST_REFERENCE CUST_REF, ARTPYMT5.REASON_CODE, ARTREAS1.REASON_DESC" _
            & ", ARTPYMT5.SEG2_CODE, ARTPYMT5.SEG3_CODE, ARTPYMT5.SEG4_CODE" _
            & " FROM ARTPYMT5, ARTREAS1" _
            & " WHERE ARTREAS1.REASON_CODE (+) = ARTPYMT5.REASON_CODE" _
            & " AND ARTPYMT5.PYMT_BATCH_NO = :PARM1 AND ARTPYMT5.PYMT_BATCH_LNO = :PARM2"
            Create_TDA(.Tables.Add, "ARTPYMT3", "**", 0, False, "VI", 0)

            ASCMAIN1.sql = "SELECT ARTPYMT3.INV_TYPE, ARTPYMT3.INV_NUM" _
            & ", ARTPYMT1.PYMT_BATCH_NO, ARTPYMT1.PYMT_BATCH_DATE" _
            & ", ARTPYMT2.CUST_PYMT_REF_NO, ARTPYMT2.CUST_PYMT_REF_DATE, ARTPYMT2.CUST_PYMT_AMT " _
            & ", ARTPYMT3.INV_PMT" _
            & " from ARTPYMT3, ARTPYMT2, ARTPYMT1" _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT3.PYMT_BATCH_NO" _
            & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT3.PYMT_BATCH_NO" _
            & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT3.PYMT_BATCH_LNO" _
            & "   and ARTPYMT3.INV_TYPE = :PARM1" _
            & "   and ARTPYMT3.INV_NUM = :PARM2"
            Create_TDA(.Tables.Add, "ARTPYMTD", "**", 0, False, "VV", 0)

            Create_Relation("ARTOPEN1", "ARTPYMTD", "INV_TYPE," & INV_NUM_column, "INV_TYPE,INV_NUM")

            If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                Create_TDA(.Tables.Add, "EDTTRPM1", "*")
            End If

            Create_TDA(.Tables.Add, "ARTCCPA1", "*")
            .Tables("ARTCCPA1").Columns.Add("CUST_NAME")
            Create_TDA(.Tables.Add, "ARTCCPA2", "*")

            Create_TDA(.Tables.Add, "ARTCUST1", "*", , False)
            ASCMAIN1.sql = "SELECT * FROM ARTCUST1"
            Create_TDA(.Tables.Add, "ARTCUST1_CREDIT", "**", 1, False)

            'If ASCMAIN1.CLIENT_CODE = "VAN" Then
            '    ASCMAIN1.sql = "SELECT CUST_CODE," _
            '             & "CUST_ADDR_CODE CUST_STORE_NO," _
            '             & "CUST_NAME CUST_STORE_NAME," _
            '             & "CUST_ADDR1 CUST_STORE_ADDR1," _
            '             & "CUST_ADDR2 CUST_STORE_ADDR2," _
            '             & "CUST_ADDR_TYPE  CUST_STORE_ADDR3," _
            '             & "CUST_CITY CUST_STORE_CITY," _
            '             & "CUST_STATE CUST_STORE_STATE," _
            '             & "CUST_ZIP_CODE CUST_STORE_ZIP_CODE," _
            '             & "CUST_COUNTRY CUST_STORE_COUNTRY," _
            '             & "CUST_CONTACT CUST_STORE_CONTACT," _
            '             & "CUST_PHONE CUST_STORE_PHONE," _
            '             & "CUST_EXT CUST_STORE_EXT," _
            '             & "CUST_FAX CUST_STORE_FAX," _
            '             & "CUST_EMAIL CUST_STORE_EMAIL," _
            '             & "CUST_ADDR_STATUS CUST_STORE_STATUS," _
            '     & "INIT_OPER," _
            '     & "LAST_OPER," _
            '     & "INIT_DATE," _
            '     & "LAST_DATE," _
            '     & "GLOBAL_LOCATION_NUMBER," _
            '     & "NULL CUST_ROUTING_INST" _
            '     & " FROM ARTCUST2 WHERE CUST_CODE = :PARM1"
            '    Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "V", 0)

            'Else
            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, False)

            'End If
            ASCMAIN1.sql = "SELECT TATCONV1.*, ARTCUST1.CUST_NAME " _
            & " from TATCONV1,ARTCUST1 " _
            & " where ARTCUST1.CUST_CODE = TATCONV1.TABLE_KEY " _
            & "   and TATCONV1.TABLE_NAME = 'ARTCUST1'" _
            & "   and TATCONV1.CONV_FOLLOWUP_BY = '" & ASCMAIN1.USER_ID & "'" _
            & "   and (TATCONV1.CONV_STATUS = '1' or (TATCONV1.CONV_STATUS = '2' and TATCONV1.LAST_DATE > TRUNC(SYSDATE )))"

            Create_TDA(.Tables.Add, "ARTCUSTT_FUPS", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ARTCUST6.*" _
            & ", ARTCUST1.CUST_NAME, ARTCUST1.CUST_ADDR1, ARTCUST1.CUST_ADDR2, ARTCUST1.CUST_ADDR3" & vbCrLf _
            & ", ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_ZIP_CODE, ARTCUST1.CUST_COUNTRY" & vbCrLf _
            & ", ARTCUST1.CUST_PHONE, ARTCUST1.CUST_FAX, ARTCUST1.CUST_CONTACT, ARTCUST1.CUST_EMAIL" & vbCrLf _
            & ", ARTCUST1.SREP_CODE,  ARTCUST1.SREP2_CODE, TRUNC(ARTCUST1.INIT_DATE) INIT_DATE" & vbCrLf _
            & ", ARTCUST1.CUST_STATUS, ARTCUST1.CUST_STATUS_DATE, ARTCUST1.CUST_STATUS_COMMENT" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.CUST_CREDIT_SCORE, ARTCUST1.CUST_CREDIT_SCORE_DATE" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_LIMIT_NOTES, ARTCUST1.CUST_CRED_LIMIT_REV, ARTCUST1.CUST_CRED_LIMIT_EST " & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_HOLD, ARTCUST1.CUST_CREDIT_RELEASE" & vbCrLf _
            & ", ARTCUST1.CUST_GROUP_CODE, ARTCUST1.CUST_BILL_TO_CUST, ARTCUST1.CUST_CREDIT_GROUP_CUST" & vbCrLf _
            & ", ARTCUST1.TERM_CODE, ARTCUST1.CUST_CLASS_CODE, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ", ARTCUST1.CUST_DUNS, ARTCUST1.CUST_PD_GRACE_DAYS, ARTCUST1.CUST_BUS_ESTAB, ARTCUST1.CUST_FACTOR_IND" & vbCrLf _
            & ", ARTCUST1.CUST_INS_AMT, ARTCUST1.CUST_INS_DATE, ARTCUST1.CUST_TERMS_NOTE" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_LIMIT_APPR_BY, ARTCUST1.CUST_CREDIT_RATING, ARTCUST1.CUST_CREDIT_RATING_DATE" & vbCrLf _
            & ", X.INV_BALANCE" & vbCrLf _
            & " from ARTCUST6,ARTCUST1," & vbCrLf _
            & " (SELECT CUST_CODE, SUM (INV_BALANCE) INV_BALANCE " & vbCrLf _
            & "  from ARTOPEN1 group by CUST_CODE) X " & vbCrLf _
            & " where ARTCUST1.CUST_CODE = ARTCUST6.CUST_CODE " & vbCrLf _
            & " and X.CUST_CODE (+) = ARTCUST6.CUST_CODE"
            Create_TDA(.Tables.Add, "ARTCUST6", "**", 0, True, , 1)

            .Tables.Add("ARTCUST1_CODES")
            With .Tables("ARTCUST1_CODES")
                .Columns.Add("CODE_CATEGORY", GetType(System.String))
                .Columns.Add("CODE_TYPE", GetType(System.String))
                .Columns.Add("CODE_VALUE", GetType(System.String))
                .Columns.Add("DESC_VALUE", GetType(System.String))
            End With

            .Tables.Add("ARTCUST1_STATS")
            With .Tables("ARTCUST1_STATS")
                .Columns.Add("STAT_TYPE", GetType(System.String))
                .Columns.Add("STAT_MTD", GetType(System.Int32))
                .Columns.Add("STAT_YTD", GetType(System.Int32))
                .Columns.Add("STAT_LYR", GetType(System.Int32))
            End With

            '.Tables.Add("ARTCUST1_AGING")
            'With .Tables("ARTCUST1_AGING")
            '    .Columns.Add("AGE_CATGY", GetType(System.String))
            '    .Columns.Add("AGE_AMT", GetType(System.Double))
            'End With

            .Tables.Add("ARTCUST1_AGEDAR")
            With .Tables("ARTCUST1_AGEDAR")
                .Columns.Add("AGE_DATE", GetType(System.DateTime))
                .Columns.Add("AGE_AMT", GetType(System.Double))
                .PrimaryKey = New DataColumn() {.Columns("AGE_DATE")}
            End With

            .Tables.Add("ARTCUST1_AR_TYPE")
            With .Tables("ARTCUST1_AR_TYPE")
                .Columns.Add("AR_TYPE", GetType(System.String))
                .Columns.Add("AR_AMT", GetType(System.Double))
            End With

            .Tables.Add("ARTCUST1_FL")
            With .Tables("ARTCUST1_FL")
                .Columns.Add("EVENT_DESC", GetType(System.String))
                .Columns.Add("EVENT_ITEM", GetType(System.String))
                .Columns.Add("EVENT_DATE", GetType(System.DateTime))
                .Columns.Add("EVENT_AMT", GetType(System.Double))
            End With

            .Tables.Add("ARTCUST1_OPEN")
            With .Tables("ARTCUST1_OPEN")
                .Columns.Add("OPEN_CATGY", GetType(System.String))
                .Columns.Add("OPEN_AMT", GetType(System.Double))
            End With

            ASCMAIN1.sql = "Select ARTSTMT1.OPS_YYYYPP, GLTPARM2.LEGEND" & vbCrLf _
            & ",SUM(ARTSTMT1.BALFWD) BALFWD" & vbCrLf _
            & ",SUM(ARTSTMT1.TYP_I) TYP_I,SUM(ARTSTMT1.TYP_R) TYP_R,SUM(ARTSTMT1.TYP_C) TYP_C" & vbCrLf _
            & ",SUM(ARTSTMT1.TYP_D) TYP_D,SUM(ARTSTMT1.TYP_B) TYP_B,SUM(ARTSTMT1.TYP_O) TYP_O" & vbCrLf _
            & ",SUM(ARTSTMT1.PYMTS) PYMTS,SUM(ARTSTMT1.WOFFS) WOFFS" & vbCrLf _
            & ",SUM(ARTSTMT1.AGE_1) AGE_1,SUM(ARTSTMT1.AGE_2) AGE_2,SUM(ARTSTMT1.AGE_3) AGE_3,SUM(ARTSTMT1.AGE_4) AGE_4" & vbCrLf _
            & ",SUM(ARTSTMT1.TYP_I_OPEN) TYP_I_OPEN,SUM(ARTSTMT1.TYP_R_OPEN) TYP_R_OPEN" & vbCrLf _
            & ",SUM(ARTSTMT1.TYP_C_OPEN) TYP_C_OPEN,SUM(ARTSTMT1.TYP_D_OPEN) TYP_D_OPEN" & vbCrLf _
            & ",SUM(ARTSTMT1.TYP_B_OPEN) TYP_B_OPEN,SUM(ARTSTMT1.TYP_O_OPEN) TYP_O_OPEN" & vbCrLf _
            & ",SUM(ARTSTMT1.TOTAL_OPEN_AMT) TOTAL_OPEN_AMT,SUM(ARTSTMT1.TOTAL_OPEN_DDS) TOTAL_OPEN_DDS" & vbCrLf _
            & ",MAX(ARTSTMT1.CUST_HIGH_BAL_DATE) CUST_HIGH_BAL_DATE,MAX(ARTSTMT1.CUST_HIGH_BAL_AMT) CUST_HIGH_BAL_AMT" & vbCrLf _
            & ",SUM(ARTSTMT1.TOTAL_CLSD_AMT) TOTAL_CLSD_AMT,SUM(ARTSTMT1.TOTAL_CLSD_DDS) TOTAL_CLSD_DDS" & vbCrLf _
            & ",SUM(ARTSTMT1.DUE_1) DUE_1,SUM(ARTSTMT1.DUE_2) DUE_2,SUM(ARTSTMT1.DUE_3) DUE_3,SUM(ARTSTMT1.DUE_4) DUE_4" & vbCrLf _
            & ",SUM(ARTSTMT1.INV_CLSD_DYS) INV_CLSD_DYS,SUM(ARTSTMT1.INV_CLSD_CNT) INV_CLSD_CNT" & vbCrLf _
            & ",SUM(ARTSTMT1.INV_CLSD_AMT) INV_CLSD_AMT,SUM(ARTSTMT1.INV_CLSD_DDS) INV_CLSD_DDS" & vbCrLf _
            & ",SUM(ARTSTMT1.TYP_I_CNT) TYP_I_CNT,SUM(ARTSTMT1.TYP_R_CNT) TYP_R_CNT" & vbCrLf _
            & ",SUM(ARTSTMT1.TYP_C_CNT) TYP_C_CNT,SUM(ARTSTMT1.TYP_D_CNT) TYP_D_CNT" & vbCrLf _
            & ",SUM(ARTSTMT1.TYP_B_CNT) TYP_B_CNT,SUM(ARTSTMT1.TYP_O_CNT) TYP_O_CNT" & vbCrLf _
            & " from ARTSTMT1, GLTPARM2 " & vbCrLf _
            & " where GLTPARM2.OPS_YYYYPP (+) = ARTSTMT1.OPS_YYYYPP" & vbCrLf _
            & "   and ARTSTMT1.CUST_CODE in (Select CUST_CODE from " & ARTCUST0 & ")" & vbCrLf _
            & " group by ARTSTMT1.OPS_YYYYPP, GLTPARM2.LEGEND"

            'ASCMAIN1.sql = "SELECT ARTSTMT1.*, GLTPARM2.LEGEND" _
            '& " from ARTSTMT1, GLTPARM2 where GLTPARM2.OPS_YYYYPP (+) = ARTSTMT1.OPS_YYYYPP " _
            '& " and ARTSTMT1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTSTMT1", "**", 0, False, "", 1)
            .Tables("ARTSTMT1").Columns.Add("DAYS_OPEN", GetType(System.Decimal), "IIF(ISNULL(TOTAL_OPEN_AMT,0)=0,0,ISNULL(TOTAL_OPEN_DDS,0) / ISNULL(TOTAL_OPEN_AMT,0))")
            .Tables("ARTSTMT1").Columns.Add("DAYS_PAID", GetType(System.Decimal), "IIF(ISNULL(INV_CLSD_AMT,0)=0,0,ISNULL(INV_CLSD_DDS,0) / ISNULL(INV_CLSD_AMT,0))")
            .Tables("ARTSTMT1").Columns.Add("DAYS_PAID_SAVG", GetType(System.Decimal), "IIF(ISNULL(INV_CLSD_CNT,0)=0,0,ISNULL(INV_CLSD_DYS,0) / ISNULL(INV_CLSD_CNT,0))")

            ASCMAIN1.sql = "Select ARTCUST5.*" _
            & " from ARTCUST5 " _
            & " where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST5", "**", 0, True, "V")

            ASCMAIN1.sql = "Select SOTORDR0.* from SOTORDR0 where CUST_CODE = :PARM1" _
            & " and ORDR_DATE >= :PARM2"
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "VD", 1)

            ASCMAIN1.sql = "Select SOTINVH1.* from SOTINVH1"
            Create_TDA(.Tables.Add, "SOTINVH1", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTINVH2.*,ICTSTYL1.STYLE_DESC,ICTCOLR1.COLOR_DESC" _
                & " from SOTINVH2,ICTSTYL1,ICTCOLR1" _
                & " where ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE" _
                & " and ICTCOLR1.COLOR_CODE = SOTINVH2.COLOR_CODE" _
                & " and SOTINVH2.INV_TYPE = :PARM1 and SOTINVH2.INV_NO = :PARM2"
            Create_TDA(.Tables.Add, "SOTINVH2", "**", 0, False, "VV", 3)
            .Tables("SOTINVH2").Columns.Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")

            ASCMAIN1.sql = "Select ORDR_NO, ORDR_DATE, ORDR_CUST_PO, CUST_STORE_NO" & vbCrLf _
                & ", ORDR_HOLD, ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_TYPE_CODE" & vbCrLf _
                & ", FRT_TERMS, SHIP_VIA_CODE, SREP_CODE, SALES_DIVISION_CODE" & vbCrLf _
                & ", INIT_DATE, INIT_OPER, ORDR_DATE_RECD, ORDR_SOURCE, WHSE_CODE_TO, ORDR_GROUP_NO" & vbCrLf _
                & " from SOTORDR1"
            sqlSOTORDR1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)

            Create_Relation("SOTORDR0", "SOTORDR1", "ORDR_GROUP_NO")

            With .Tables("SOTORDR0").Columns
                .Add("FRT_TERMS", GetType(System.String), "MIN(CHILD.FRT_TERMS)")
                .Add("SHIP_VIA_CODE", GetType(System.String), "MIN(CHILD.SHIP_VIA_CODE)")
                '.Add("ORDR_TYPE_CODE", GetType(System.String), "MIN(CHILD.ORDR_TYPE_CODE)")
                .Add("INIT_DATE", GetType(System.DateTime), "MIN(CHILD.INIT_DATE)")
                .Add("INIT_OPER", GetType(System.String), "MIN(CHILD.INIT_OPER)")
                If Not .Contains("ORDR_DATE_RECD") Then
                    .Add("ORDR_DATE_RECD", GetType(System.DateTime), "MIN(CHILD.ORDR_DATE_RECD)")
                End If
                '.Add("ORDR_SOURCE", GetType(System.String), "MIN(CHILD.ORDR_SOURCE)")
                .Add("ORDR_NO", GetType(System.String), "MIN(CHILD.ORDR_NO)")
                .Add("ORDR_HOLD", GetType(System.String), "MIN(CHILD.ORDR_HOLD)")
            End With

            ASCMAIN1.sql = "Select SOTORDR2.* from SOTORDR2 where SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "V", 0)
            .Tables("SOTORDR2").Columns("ORDR_STATUS").AllowDBNull = True

            SetUp12Month(True)

            ASCMAIN1.sql = "Select SALES_DIVISION_CODE, STYLE_CLASS_CODE"
            For I As Integer = 1 To 12
                ASCMAIN1.sql &= ", SUM (QTY" & Format(I, "00") & ") QTY" & Format(I, "00")
                ASCMAIN1.sql &= ", SUM (SLS" & Format(I, "00") & ") SLS" & Format(I, "00")
                ASCMAIN1.sql &= ", SUM (CGS" & Format(I, "00") & ") CGS" & Format(I, "00")
            Next
            ASCMAIN1.sql &= " from " & SATCUSTS _
                & " group by SALES_DIVISION_CODE, STYLE_CLASS_CODE"
            Create_TDA(.Tables.Add, "SATCUSTS", "**", 0, False, "", 0)
            With .Tables("SATCUSTS").Columns
                Dim V As String = ""
                For I As Integer = 1 To 12
                    .Add("GP" & Format(I, "00"), GetType(System.Decimal), "ISNULL(SLS" & Format(I, "00") & ",0) - ISNULL(CGS" & Format(I, "00") & ",0)")
                    V &= "+ISNULL(V" & Format(I, "00") & ",0)"
                Next
                .Add("QTY13", GetType(System.Decimal), Replace(Mid(V, 2), "V", "QTY"))
                .Add("SLS13", GetType(System.Decimal), Replace(Mid(V, 2), "V", "SLS"))
                .Add("CGS13", GetType(System.Decimal), Replace(Mid(V, 2), "V", "CGS"))
                .Add("GP13", GetType(System.Decimal), Replace(Mid(V, 2), "V", "GP"))
                For I As Integer = 1 To 13
                    .Add("V" & Format(I, "00"), GetType(System.Decimal))
                Next
            End With

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE"
            For I As Integer = 1 To 12
                ASCMAIN1.sql &= ", SUM (QTY" & Format(I, "00") & ") QTY" & Format(I, "00")
                ASCMAIN1.sql &= ", SUM (SLS" & Format(I, "00") & ") SLS" & Format(I, "00")
                ASCMAIN1.sql &= ", SUM (CGS" & Format(I, "00") & ") CGS" & Format(I, "00")
            Next
            ASCMAIN1.sql &= " from " & SATCUSTS _
                & " where SALES_DIVISION_CODE = :PARM1" _
                & "   and STYLE_CLASS_CODE = :PARM2" _
                & " group by STYLE_CODE, COLOR_CODE"
            Create_TDA(.Tables.Add, "SATCUSTI", "**", 0, False, "VV", 0)
            With .Tables("SATCUSTI").Columns
                Dim V As String = ""
                For I As Integer = 1 To 12
                    .Add("GP" & Format(I, "00"), GetType(System.Decimal), "SLS" & Format(I, "00") & " - CGS" & Format(I, "00"))
                    V &= "+ISNULL(V" & Format(I, "00") & ",0)"
                Next
                .Add("QTY13", GetType(System.Decimal), Replace(Mid(V, 2), "V", "QTY"))
                .Add("SLS13", GetType(System.Decimal), Replace(Mid(V, 2), "V", "SLS"))
                .Add("CGS13", GetType(System.Decimal), Replace(Mid(V, 2), "V", "CGS"))
                .Add("GP13", GetType(System.Decimal), Replace(Mid(V, 2), "V", "GP"))
                For I As Integer = 1 To 13
                    .Add("V" & Format(I, "00"), GetType(System.Decimal), "SLS" & Format(I, "00") & " - CGS" & Format(I, "00"))
                Next
            End With

            Setup_optSales()

            ASCMAIN1.sql = "Select TATEVNT1.* " _
            & " from TATEVNT1 " _
            & " where TABLE_NAME = 'ARTCUST1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select TATCONV1.* from TATCONV1 " _
            & " where TABLE_NAME = 'ARTCUST1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATCONV1", "**", 0, , "V", 1)
            .Tables("TATCONV1").Columns.Add("CONV_ATTACHMENTS", GetType(System.Int64))

            Dim TYP01 As String = Mid(ASCMAIN1.CYP, 1, 4) & "01"
            Dim LYP01 As String = Format(Val(Mid(ASCMAIN1.CYP, 1, 4)) - 1, "0000") & "01"
            Dim TY As String = Mid(ASCMAIN1.CYP, 1, 4)

            ASCMAIN1.sql = "" _
            & "SELECT X.CUST_CODE, X.OPS_YYYYPP, SUBSTR(GLTPARM2.LEGEND,10,6) LEGEND" & vbCr _
            & ", SUM (X.TY_SLS) TY_SLS, SUM (X.TY_CGS) TY_CGS, SUM (X.TY_PYMT) TY_PYMT, SUM (X.TY_DISC) TY_DISC" & vbCr _
            & ", SUM (X.LY_SLS) LY_SLS, SUM (X.LY_CGS) LY_CGS, SUM (X.LY_PYMT) LY_PYMT, SUM (X.LY_DISC) LY_DISC" & vbCr _
            & " FROM (" & vbCr _
            & " SELECT 'YP' TYPE, :PARM1 CUST_CODE, OPS_YYYYPP " & vbCr _
            & ", 0 TY_SLS, 0 TY_CGS, 0 LY_SLS, 0 LY_CGS, 0 TY_PYMT, 0 LY_PYMT, 0 TY_DISC, 0 LY_DISC" & vbCr _
            & " FROM GLTPARM2 WHERE OPS_YYYYPP >= '" & TYP01 & "' AND OPS_YYYYPP <= '" & TY & "12'" & vbCr _
            & " UNION" & vbCr _
            & " SELECT 'SL' TYPE, CUST_CODE, '" & TY & "' || SUBSTR(ORDR_YYYYPP_UPDATED,5,2) OPS_YYYYPP" & vbCr _
            & ", SUM (CASE WHEN ORDR_YYYYPP_UPDATED >= '" & TY & "' THEN NVL(ORDR_QTY_SHIP,0) * NVL(ORDR_UNIT_PRICE,0) ELSE 0 END) TY_SLS" & vbCr _
            & ", SUM (CASE WHEN ORDR_YYYYPP_UPDATED >= '" & TY & "' THEN NVL(ORDR_QTY_SHIP,0) * NVL(ORDR_UNIT_COST,0) ELSE 0 END) TY_CGS" & vbCr _
            & ", SUM (CASE WHEN ORDR_YYYYPP_UPDATED < '" & TY & "' THEN NVL(ORDR_QTY_SHIP,0) * NVL(ORDR_UNIT_PRICE,0) ELSE 0 END) LY_SLS" & vbCr _
            & ", SUM (CASE WHEN ORDR_YYYYPP_UPDATED < '" & TY & "' THEN NVL(ORDR_QTY_SHIP,0) * NVL(ORDR_UNIT_COST,0) ELSE 0 END) LY_CGS" & vbCr _
            & ", 0 TY_PYMT, 0 LY_PYMT, 0 TY_DISC, 0 LY_DISC" & vbCr _
            & " FROM SOTINVH2 WHERE CUST_CODE = :PARM1 " & vbCr _
            & " AND ORDR_YYYYPP_UPDATED >= '" & LYP01 & "' AND ORDR_YYYYPP_UPDATED <= '" & TY & "12'" & vbCr _
            & " GROUP BY CUST_CODE, '" & TY & "' || SUBSTR(ORDR_YYYYPP_UPDATED,5,2)" & vbCr _
            & " UNION" & vbCr _
            & " SELECT 'PY' TYPE, ARTPYMT2.CUST_CODE, '" & TY & "' || SUBSTR(ARTPYMT1.OPS_YYYYPP,5,2) OPS_YYYYPP" & vbCr _
            & ", 0 TY_SLS, 0 TY_CGS, 0 LY_SLS, 0 LY_CGS" & vbCr _
            & ", SUM (CASE WHEN ARTPYMT1.OPS_YYYYPP >= '" & TY & "' THEN ARTPYMT2.CUST_PYMT_AMT ELSE 0 END) TY_PYMT" & vbCr _
            & ", SUM (CASE WHEN ARTPYMT1.OPS_YYYYPP < '" & TY & "' THEN ARTPYMT2.CUST_PYMT_AMT ELSE 0 END) LY_PYMT" & vbCr _
            & ", 0 TY_DISC, 0 LY_DISC" & vbCr _
            & " FROM ARTPYMT1,ARTPYMT2" & vbCr _
            & " WHERE ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCr _
            & " AND ARTPYMT1.OPS_YYYYPP >= '" & LYP01 & "'" & vbCr _
            & " AND ARTPYMT2.CUST_CODE = :PARM1" & vbCr _
            & " GROUP BY ARTPYMT2.CUST_CODE, '" & TY & "' || SUBSTR(ARTPYMT1.OPS_YYYYPP,5,2)" & vbCr _
            & " UNION" & vbCr _
            & " SELECT 'DW' TYPE, ARTPYMT2.CUST_CODE, '" & TY & "' || SUBSTR(ARTPYMT1.OPS_YYYYPP,5,2) OPS_YYYYPP" & vbCr _
            & ", 0 TY_SLS, 0 TY_CGS, 0 LY_SLS, 0 LY_CGS, 0 TY_PYMT, 0 LY_PYMT" & vbCr _
            & ", SUM (CASE WHEN ARTPYMT1.OPS_YYYYPP >= '" & TY & "' THEN NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0) ELSE 0 END) TY_DISC" & vbCr _
            & ", SUM (CASE WHEN ARTPYMT1.OPS_YYYYPP < '" & TY & "' THEN NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0) ELSE 0 END) LY_DISC" & vbCr _
            & " FROM ARTPYMT1,ARTPYMT2,ARTPYMT3" & vbCr _
            & " WHERE ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCr _
            & " AND ARTPYMT1.OPS_YYYYPP >= '" & LYP01 & "'" & vbCr _
            & " AND ARTPYMT2.CUST_CODE = :PARM1" & vbCr _
            & " AND ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCr _
            & " AND ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCr _
            & " AND (NVL(ARTPYMT3.INV_DISC_TAKEN,0) <> 0 OR NVL(ARTPYMT3.INV_WRITE_OFF,0) <> 0)" & vbCr _
            & " GROUP BY ARTPYMT2.CUST_CODE, '" & TY & "' || SUBSTR(ARTPYMT1.OPS_YYYYPP,5,2)" & vbCr _
            & " UNION" & vbCr _
            & " SELECT 'GL' TYPE, ARTPYMT2.CUST_CODE, '" & TY & "' || SUBSTR(ARTPYMT1.OPS_YYYYPP,5,2) OPS_YYYYPP" & vbCr _
            & ", 0 TY_SLS, 0 TY_CGS, 0 LY_SLS, 0 LY_CGS, 0 TY_PYMT, 0 LY_PYMT" & vbCr _
            & ", SUM (CASE WHEN ARTPYMT1.OPS_YYYYPP >= '" & TY & "' THEN NVL(ARTPYMT4.GL_DIST_AMT,0) ELSE 0 END) TY_DISC" & vbCr _
            & ", SUM (CASE WHEN ARTPYMT1.OPS_YYYYPP < '" & TY & "' THEN NVL(ARTPYMT4.GL_DIST_AMT,0) ELSE 0 END) LY_DISC" & vbCr _
            & " FROM ARTPYMT1,ARTPYMT2,ARTPYMT4" & vbCr _
            & " WHERE ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCr _
            & " AND ARTPYMT1.OPS_YYYYPP >= '" & LYP01 & "'" & vbCr _
            & " AND ARTPYMT2.CUST_CODE = :PARM1" & vbCr _
            & " AND ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCr _
            & " AND ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCr _
            & " GROUP BY ARTPYMT2.CUST_CODE, '" & TY & "' || SUBSTR(ARTPYMT1.OPS_YYYYPP,5,2)" & vbCr _
            & ") X, GLTPARM2 WHERE GLTPARM2.OPS_YYYYPP = X.OPS_YYYYPP" & vbCr _
            & " GROUP BY X.CUST_CODE, X.OPS_YYYYPP, SUBSTR(GLTPARM2.LEGEND,10,6)"
            Create_TDA(.Tables.Add, "ARTSLSMA", "**", 0, False, "V", 2)
            With .Tables("ARTSLSMA")
                .Columns.Add("TY_GP", GetType(System.Decimal), "TY_SLS - TY_CGS")
                .Columns.Add("LY_GP", GetType(System.Decimal), "LY_SLS - LY_CGS")
            End With

            'If ASCMAIN1.CLIENT_CODE = "VAN" Then
            ASCMAIN1.sql = "SELECT X.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.SALES_DIVISION_CODE, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTINVH1.REASON_CODE, SOTINVH1.TERM_CODE, SOTINVH1.SREP_CODE" & vbCrLf _
                & "FROM (" & vbCrLf _
                & "SELECT SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
                & ", COUNT (*) LINES" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP) QTY" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) SLS" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST) CGS" & vbCrLf _
                & ", MAX (DECODE(SOTINVH2.INV_TYPE,'I',SOTINVH2.INV_NO,NULL)) INV_NO" & vbCrLf _
                & " from SOTINVH2,SOTINVH1" & vbCrLf _
                & " where SOTINVH1.CUST_CODE = :PARM1" & vbCrLf _
                & "   and SOTINVH1.INV_DATE >= :PARM2" & vbCrLf _
                & "   and SOTINVH1.INV_DATE <= :PARM3" & vbCrLf _
                & "   and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
                & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
                & " group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
                & ") X, SOTINVH1, ICTSTYL1, ICTCOLR1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE and ICTCOLR1.COLOR_CODE = X.COLOR_CODE and SOTINVH1.INV_TYPE (+) = 'I' AND SOTINVH1.INV_NO (+) = X.INV_NO"
            'Else
            '    ASCMAIN1.sql = "SELECT X.*, SOTINVH1.INV_DATE, SOTINVH1.ORDR_CUST_PO, ICTITEM1.ITEM_DESC " & vbCrLf _
            '        & " from (" & vbCrLf _
            '        & "SELECT SOTINVH2.ITEM_CODE " & vbCrLf _
            '        & ", COUNT (*) LINES" & vbCrLf _
            '        & ", SUM (SOTINVH2.ORDR_QTY_SHIP) QTY" & vbCrLf _
            '        & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) SLS" & vbCrLf _
            '        & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ITEM_UNIT_COST) CGS" & vbCrLf _
            '        & ", MAX (DECODE(SOTINVH2.INV_TYPE,'I',SOTINVH2.INV_NO,NULL)) INV_NO" & vbCrLf _
            '        & " FROM SOTINVH2,SOTINVH1" & vbCrLf _
            '        & "WHERE SOTINVH1.CUST_CODE  = :PARM1" & vbCrLf _
            '        & "  AND SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            '        & "  AND SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            '        & "GROUP BY SOTINVH2.ITEM_CODE " & vbCrLf _
            '        & ") X, SOTINVH1, ICTITEM1" & vbCrLf _
            '        & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE and SOTINVH1.INV_TYPE = 'I' AND SOTINVH1.INV_NO = X.INV_NO"
            'End If

            sqlSOTAEBP1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTAEBP1", "**", 0, False, "VDD", 2)
            .Tables("SOTAEBP1").Columns.Add("PRICE", GetType(System.Decimal), "IIF(QTY=0,0,SLS / QTY)")
            .Tables("SOTAEBP1").Columns.Add("GP", GetType(System.Decimal), "SLS - CGS")
            .Tables("SOTAEBP1").Columns.Add("GPPCT", GetType(System.Decimal), "IIF(SLS=0,0,100 * GP / SLS)")

            ASCMAIN1.sql = "Select SOTINVH2.*, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_NO" & vbCrLf _
                & " from SOTINVH2,SOTINVH1,ICTCOLR1" & vbCrLf _
                & " where SOTINVH1.CUST_CODE = :PARM1" & vbCrLf _
                & "   and SOTINVH1.INV_DATE >= :PARM2" & vbCrLf _
                & "   and SOTINVH1.INV_DATE <= :PARM3" & vbCrLf _
                & "   and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
                & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = SOTINVH2.COLOR_CODE" & vbCrLf _
                & "   and SOTINVH2.STYLE_CODE = :PARM4" & vbCrLf _
                & "   and SOTINVH2.COLOR_CODE = :PARM5" & vbCrLf
            sqlSOTAEBP2 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTAEBP2", "**", 0, False, "VDDVV", 3)
            .Tables("SOTAEBP2").Columns.Add("SLS", GetType(System.Decimal), "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")
            .Tables("SOTAEBP2").Columns.Add("CGS", GetType(System.Decimal), "ORDR_QTY_SHIP * ORDR_UNIT_COST")
            .Tables("SOTAEBP2").Columns.Add("GP", GetType(System.Decimal), "SLS - CGS")
            .Tables("SOTAEBP2").Columns.Add("GPPCT", GetType(System.Decimal), "IIF(SLS=0,0,100 * GP / SLS)")

            Create_TDA(.Tables.Add, "TATTERM1", "*", 0, False)
            Create_TDA(.Tables.Add, "SOTSREP1", "*", 0, False)

            ASCMAIN1.sql = "Select ARTCUSTD.* " _
            & " from ARTCUSTD " _
            & " where ARTCUSTD.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "V", 2)

        End With

        cbeYP_PYMTs.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP_PYMTs.SelectedItem = cbeYP_PYMTs.Items(3)

        cbeSales.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeSales.SelectedItem = cbeSales.Items(0)

        dteOrderHistory.DateTime = Now.AddDays(-365)
        dteInvoiceHistory.DateTime = Now.AddDays(-365)

        grdARTCUSTD.DataSource = dst.Tables("ARTCUSTD")
        grdARTCUSTS.DataSource = dst.Tables("ARTCUSTS")
        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")
        grdARTSLSMA.DataSource = dst.Tables("ARTSLSMA")
        'grdARTCUST4.DataSource = dst.Tables("ARTCUST4")
        grdARTCUST5.DataSource = dst.Tables("ARTCUST5")
        grdARTCUSTT_FUPS.DataSource = dst.Tables("ARTCUSTT_FUPS")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")

        grdSOTAEBP1.DataSource = dst.Tables("SOTAEBP1")
        grdSOTAEBP2.DataSource = dst.Tables("SOTAEBP2")

        grdSATCUSTS.DataSource = dst.Tables("SATCUSTS")
        grdSATCUSTI.DataSource = dst.Tables("SATCUSTI")

        grdARTCUST1_FL.DataSource = dst.Tables("ARTCUST1_FL")
        grdARTCUST1_AR_TYPE.DataSource = dst.Tables("ARTCUST1_AR_TYPE")
        grdARTCUST1_OPEN.DataSource = dst.Tables("ARTCUST1_OPEN")
        grdARTCUST1_CODES.DataSource = dst.Tables("ARTCUST1_CODES")
        grdARTCUST1_STATS.DataSource = dst.Tables("ARTCUST1_STATS")

        grdARTOPEN1.DataSource = New DataView(dst.Tables("ARTOPEN1"), "", "", DataViewRowState.CurrentRows)
        grdARTPYMT2.DataSource = dst.Tables("ARTPYMT2")
        grdARTPYMT3.DataSource = dst.Tables("ARTPYMT3")

        grdARTSTMT1.DataSource = dst.Tables("ARTSTMT1")

        grdARTCUST6.DataSource = dst.Tables("ARTCUST6")

        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")

        grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")
        grdSOTINVH2.DataSource = dst.Tables("SOTINVH2")

        grdSOTORDR0.DisplayLayout.GroupByBox.Hidden = True ' False
        grdARTCUST6.DisplayLayout.GroupByBox.Hidden = False

        'Show_Filter(grdSOTORDR1, True)
        Show_Filter(grdARTCUST6, True)

        With grdARTPYMT2.DisplayLayout.Bands("ARTPYMT2")
            .Columns("PYMT_BATCH_NO").Header.Fixed = True
        End With

        With grdARTCUST6.DisplayLayout.Bands("ARTCUST6")
            .Groups("Customer").Header.Fixed = True
            .Columns("CUST_BILL_TO_CUST").Hidden = True
        End With


        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDR0, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})
        Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})

        Create_Summary(grdSOTINVH1, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1, New String() {"INV_SALES", "INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT"})

        Create_Summary(grdSOTINVH2, "INV_LNO", "Count")
        Create_Summary(grdSOTINVH2, New String() {"ORDR_QTY_SHIP", "ORDR_AMT_SHIP"})

        Create_Summary(grdARTPYMT2, "PYMT_BATCH_NO", "Count")
        Create_Summary(grdARTPYMT2, "CUST_PYMT_AMT")

        Create_Summary(grdARTPYMT3, "INV_NUM", "Count")
        Create_Summary(grdARTPYMT3, "PMT")
        Create_Summary(grdARTPYMT3, "DED")

        If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then
            Create_Summary(grdSOTAEBP1, "STYLE_CODE", "Count")
        Else
            Create_Summary(grdSOTAEBP1, "ITEM_CODE", "Count")
        End If
        Create_Summary(grdSOTAEBP1, New String() {"LINES", "QTY", "SLS", "CGS", "GP"})

        Create_Summary(grdSOTAEBP2, "INV_NO", "Count")
        Create_Summary(grdSOTAEBP2, New String() {"ORDR_QTY_SHIP", "SLS", "CGS", "GP"})

        'Create_Summary(grdARTCUST2, "CUST_SHIP_TO_CODE", "Count")
        grdARTCUST2.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement

        grdARTSLSMA.DisplayLayout.UseFixedHeaders = True
        With grdARTSLSMA.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"OPS_YYYYPP", "LEGEND"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Drawing.Color.White
                GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"OPS_YYYYPP", "LEGEND"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    GCOL.Width = 60
                ElseIf New String() {"TY_SLS", "TY_CGS", "TY_GP", "TY_PYMT", "TY_DISC"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    GCOL.Format = "#,##0"
                    GCOL.Width = 90
                    GCOL.Header.Caption = "TY " & ASCMAIN1.Make_Caption(Split(GCOL.Key, "_")(1))
                    Create_Summary(grdARTSLSMA, GCOL.Key)
                ElseIf New String() {"LY_SLS", "LY_CGS", "LY_GP", "LY_PYMT", "LY_DISC"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    GCOL.Format = "#,##0"
                    GCOL.Width = 90
                    GCOL.Header.Caption = "LY " & ASCMAIN1.Make_Caption(Split(GCOL.Key, "_")(1))
                    Create_Summary(grdARTSLSMA, GCOL.Key)
                Else
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Orange
                End If
            Next
        End With

        grdSOTORDR0.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_GROUP_NO", "ORDR_DATE", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Drawing.Color.White
                GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"ORDR_GROUP_NO", "ORDR_DATE", "ORDR_CUST_PO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    GCOL.Width = 120
                ElseIf New String() {"ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    GCOL.Format = "MM/dd"
                    GCOL.Width = 60
                ElseIf New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    GCOL.Format = "#,##0.00"
                    GCOL.Width = 85
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    GCOL.Format = "#,##0"
                    GCOL.Width = 60
                ElseIf New String() {"SHIP_VIA_CODE", "ORDR_TYPE_CODE", "FRT_TERMS", "WHSE_CODE", "ORDR_SHIP_INSTR"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                Else
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Orange
                End If
            Next
        End With

        With grdSOTORDR0.DisplayLayout.Bands(1)
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "ORDR_DATE", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Drawing.Color.White
                GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"ORDR_GROUP_NO", "ORDR_DATE", "ORDR_CUST_PO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    GCOL.Format = "MM/dd"
                ElseIf New String() {"SHIP_VIA_CODE", "ORDR_TYPE_CODE", "FRT_TERMS", "WHSE_CODE", "ORDR_SHIP_INSTR"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                Else
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Orange
                End If
            Next
        End With

        grdSOTINVH1.DisplayLayout.UseFixedHeaders = True
        With grdSOTINVH1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"INV_TYPE", "INV_NO", "INV_DATE", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With


        grdSATCUSTS.DisplayLayout.UseFixedHeaders = True
        With grdSATCUSTS.DisplayLayout.Bands(0)
            With .Columns("SALES_DIVISION_CODE")
                .Header.Fixed = True
                .Header.Caption = "Div"
                .Width = 40
            End With
            With .Columns("STYLE_CLASS_CODE")
                .Header.Fixed = True
                .Header.Caption = "Class"
                .Width = 50
            End With
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key.StartsWith("QTY") _
                Or GCOL.Key.StartsWith("SLS") _
                Or GCOL.Key.StartsWith("CGS") _
                Or GCOL.Key.StartsWith("GP") Then
                    GCOL.Hidden = True
                Else
                    GCOL.Header.Appearance.BackColor = Drawing.Color.White
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"SALES_DIVISION_CODE", "STYLE_CLASS_CODE"}.Contains(GCOL.Key) Then
                        GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        GCOL.Width = 70
                    ElseIf GCOL.Key.StartsWith("V") Then
                        GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        GCOL.Format = "#,##0"
                        GCOL.Width = 70
                        Create_Summary(grdSATCUSTS, GCOL.Key)
                        If GCOL.Key = "V13" Then
                            GCOL.Header.Caption = "Total"
                        End If
                    End If
                End If
            Next

        End With

        grdSATCUSTI.DisplayLayout.UseFixedHeaders = True
        With grdSATCUSTI.DisplayLayout.Bands(0)
            With .Columns("STYLE_CODE")
                .Header.Fixed = True
                .Header.Caption = "Style"
                .Width = 80
            End With
            With .Columns("COLOR_CODE")
                .Header.Fixed = True
                .Header.Caption = "Color"
                .Width = 50
            End With
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key.StartsWith("QTY") _
               Or GCOL.Key.StartsWith("SLS") _
               Or GCOL.Key.StartsWith("CGS") _
               Or GCOL.Key.StartsWith("GP") Then
                    GCOL.Hidden = True
                Else
                    GCOL.Header.Appearance.BackColor = Drawing.Color.White
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"STYLE_CODE", "COLOR_CODE"}.Contains(GCOL.Key) Then
                        GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        GCOL.Width = 70
                    ElseIf GCOL.Key.StartsWith("V") Then
                        GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        GCOL.Format = "#,##0"
                        GCOL.Width = 70
                        Create_Summary(grdSATCUSTI, GCOL.Key)
                        If GCOL.Key = "V13" Then
                            GCOL.Header.Caption = "Total"
                        End If
                    End If
                End If
            Next
        End With

        Create_Summary(grdARTCUST6, "CUST_CODE", "Count")
        Create_Summary(grdARTCUST6, "CUST_SALES_MTD")
        Create_Summary(grdARTCUST6, "CUST_SALES_YTD")
        Create_Summary(grdARTCUST6, "CUST_SALES_LYR")
        Create_Summary(grdARTCUST6, "CUST_CASH_MTD")
        Create_Summary(grdARTCUST6, "CUST_CASH_YTD")
        Create_Summary(grdARTCUST6, "CUST_CASH_LYR")
        Create_Summary(grdARTCUST6, "CUST_CRED_MTD")
        Create_Summary(grdARTCUST6, "CUST_CRED_YTD")
        Create_Summary(grdARTCUST6, "CUST_CRED_LYR")
        Create_Summary(grdARTCUST6, "CUST_NUM_INV_MTD")
        Create_Summary(grdARTCUST6, "CUST_NUM_INV_YTD")
        Create_Summary(grdARTCUST6, "CUST_NUM_INV_LYR")
        Create_Summary(grdARTCUST6, "INV_BALANCE")


        Create_Summary(grdARTOPEN1, INV_NUM_column, "Count")

        Create_Summary(grdARTOPEN1, "INV_BALANCE")
        Create_Summary(grdARTOPEN1, "INV_TOTAL_AMOUNT")
        Create_Summary(grdARTOPEN1, "INV_PMT")
        Create_Summary(grdARTOPEN1, "INV_DISC_TAKEN")
        Create_Summary(grdARTOPEN1, "INV_WRITE_OFF")

        grdARTCUSTT_FUPS.DisplayLayout.UseFixedHeaders = True
        With grdARTCUSTT_FUPS.DisplayLayout.Bands(0)
            .Columns("TABLE_KEY").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
        End With

        With grdARTOPEN1.DisplayLayout.Bands("ARTOPEN1")
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns(INV_NUM_column).Header.Fixed = True
            .Columns("AGE_BUCKET").GroupByMode = UltraWinGrid.GroupByMode.Value
        End With
        grdARTOPEN1.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay

        'With grdARTCUST2.DisplayLayout.Bands("ARTCUST2")
        '    .Columns("CUST_SHIP_TO_CODE").Header.Fixed = True
        '    .Columns("CUST_SHIP_TO_NAME").Header.Fixed = True
        'End With

        For i As Integer = 1 To 4
            With grdARTSTMT1.DisplayLayout.Bands(0)
                .Columns("AGE_" & CStr(i)).Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_" & CStr(i)) & ""
                .Columns("DUE_" & CStr(i)).Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_" & CStr(i)) & ""
            End With
        Next

        grdARTSTMT1.DisplayLayout.UseFixedHeaders = True
        With grdARTSTMT1.DisplayLayout.Bands(0)
            .Columns("LEGEND").Header.Fixed = True
        End With

        'grdARTCUST2.DisplayLayout.UseFixedHeaders = True
        'With grdARTCUST2.DisplayLayout.Bands(0)
        '    .Columns("CUST_SHIP_TO_CODE").Header.Fixed = True
        'End With

        grdSOTAEBP1.DisplayLayout.UseFixedHeaders = True
        With grdSOTAEBP1.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"LINES", "QTY", "SLS", "GP", "GPPCT"}.Contains(GCOL.Key) Then
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    .Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
                ElseIf New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC"}.Contains(GCOL.Key) Then
                    '.headerFixed = True
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    .Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
                    'GCOL.Hidden = True
                Else
                    .Header.Appearance.BackColor2 = Drawing.Color.DeepPink
                    .Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
                End If
            Next
            ' .Columns("LEGEND").Header.Fixed = True
        End With

        grdSOTORDR2.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_LNO", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Drawing.Color.White
                GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"ORDR_LNO", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC", "STYLE_UOM", "CARTON_PACK_QTY"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_ALLO", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"ORDR_UNIT_PRICE", "ORDR_UNIT_PRICE_CALC", "ORDR_UNIT_PRICE_MANUAL", "STYLE_PRICE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Orange
                End If
            Next
        End With



        Set_Read_Only(grpARTCUST1, True)
        Set_Read_Only(grpCustomerOP, True)
        Set_Read_Only(grpARCredit, True)

        If InStr(ASCMAIN1.USER_SECURITY_CODEs, "CR") Then
            chkEditCredit.Visible = True
            ToggleCreditEditable()
            grpBankPymtEdit.Visible = True
        Else
            grpBankPymtEdit.Visible = False
            chkEditCredit.Visible = False
            btnCreditUpdate.Visible = False
            btnCreditCancel.Visible = False
        End If

        ASCMAIN1.Add_Value_List(grdARTOPEN1, "AGE_BUCKET", , _
        New String() {":" _
                , "0:Zero Balance" _
                , "1:" & ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1") _
                , "2:" & ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2") _
                , "3:" & ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3") _
                , "4:" & ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4")})
        ASCMAIN1.Add_Value_List(grdARTOPEN1, "DUE_BUCKET", , _
        New String() {":" _
                , "0:Zero Balance" _
                , "1:" & ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_1") _
                , "2:" & ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_2") _
                , "3:" & ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_3") _
                , "4:" & ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_4")})

        With grdARTOPEN1.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"INV_LAST_PMT", "INV_LAST_PMT_REF", "INV_LAST_PMT_REF_DT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Yellow
            Next
            For Each COLUMN_NAME As String In New String() {"INV_TOTAL_AMOUNT", "INV_PMT", "INV_DISC_TAKEN", "INV_WRITE_OFF", "INV_BALANCE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"INV_TOTAL_AMOUNT", "INV_BALANCE"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_CREDIT_APPR_BY", "ORDR_CREDIT_APPR_DATE", "INV_NOTES"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Turquoise
            Next
            For Each COLUMN_NAME As String In New String() {"DTP", "AGE_BUCKET", "DUE_BUCKET", "AGE", "DUE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Pink
            Next
            For Each COLUMN_NAME As String In New String() {"OPS_YYYYPP_PAID", "AMT_PAID", "DATE_PAID"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next

        End With

        SetControlPanel()

        dteAEBFrom.Value = Now.Date.AddYears(-1)
        dteAEBTo.Value = Now.Date
        optAEBDate.Value = "D"

        tabMain.Visible = False

        '   ASCMAIN1.Add_Value_List(grdARTCUST5, "CUST_CREDIT_RELEASE", Nothing, New String() {":", "0:Manual Always", "1:Auto Normal Rules", "2:Auto ignore PD"})
        ASCMAIN1.Add_Value_List(grdARTCUSTT_FUPS, "CONV_STATUS", Nothing, New String() {":", "1:Action Reqd", "2:Complete"})
        ASCMAIN1.Add_Value_List(grdARTCUST6, "CUST_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive", "C:Closed"})
        ASCMAIN1.Add_Value_List(grdARTCUST6, "CUST_CREDIT_RELEASE", Nothing, New String() {":", "M:Manual", "R:Normal", "N:No-Check", "I:IgnorePD"})

        ASCMAIN1.Add_Value_List(grdSOTAEBP1, "STYLE_GROUP_CODE")
        Show_Filter(grdSOTAEBP1, True)
        grdSOTAEBP1.DisplayLayout.GroupByBox.Hidden = False
        Toggle_AEB_GP()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select Customer"
                Validate_Code("CUST_CODE")

            Case "Done"
                If chkEditCredit.Checked Then
                    EMsg &= "You must first either Update or Cancel the Edit to Credit Info"
                End If

            Case "Update"

                If Absx1.txtFor("CONV_LOG").Text = "" Then
                    EMsg &= vbCr & "You Must Enter Something as a Note"
                End If

                If Absx1.txtFor("CONV_FOLLOWUP_WITH").Text <> "" Then
                    If LookUp("ASTUSER1", Absx1.txtFor("CONV_FOLLOWUP_WITH").Text) Is Nothing Then
                        EMsg &= vbCr & "Follow Ups require a Valid User ID"
                    End If
                End If

                If Absx1.txtFor("CONV_FOLLOWUP_WITH").Text <> "" And Absx1.dteFor("CONV_FOLLOWUP_DATE").Text = "" _
                Or Absx1.txtFor("CONV_FOLLOWUP_WITH").Text = "" And Absx1.dteFor("CONV_FOLLOWUP_DATE").Text <> "" Then
                    EMsg &= vbCr & "Follow Ups require a Name and a Date"
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

            Case "Select Customer"
                EntryMode = "I"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Customer Activity"
                Customer_Activity()

            Case "Refresh Follow Ups"
                Refresh_FollowUps()

            Case "Print"
                Print_Hard_Copy()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Select Customer").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Customer Activity").Visible = Not ScreenMode
                    .Items("Refresh Follow Ups").Visible = Not ScreenMode

                    .Items("Print").Visible = False

                End With
                .Groups("Customer Log").Visible = False
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = ScreenMode
        grdARTCUST6.Visible = Not ScreenMode And (grdARTCUST6.Tag = "*")

        grdSATCUSTS.Visible = False
        splSATCUSTS.Panel2Collapsed = True

        chkEditCredit.Checked = False
        grdARTCUSTT_FUPS.Visible = Not ScreenMode And (grdARTCUST6.Tag <> "*")

        SetControlPanel()

        grpContact.Visible = ScreenMode
        lblSREP.Visible = ScreenMode

        If ScreenMode Then
            tabMain.ActiveTab = tabMain.Tabs("General")
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show $0 Balance Items"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = True
            Filter_ARTOPEN1()
        Else
            Clear_Record()
            splSATCUSTS.Visible = False
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("CSZ").Text = ""

        txtINV_NO_PYMT.Text = ""

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() _
        {"ARTPYMT2", "ARTPYMT3", "ARTCUST1", "ARTCUST1_CREDIT", "ARTCUST0", "ARTCUSTS", _
         "SOTAEBP1", "SOTAEBP2", "SOTORDR0", "SOTORDR1", "SOTORDR2", "SATCUSTS", "SATCUSTI", _
         "ARTCUST2", "ARTCUST5", "ARTOPEN1", "ARTSTMT1", "ARTCUST1_AGEDAR", _
         "SOTSREP1", "TATCONV1", "TATTERM1", "ARTPYMTD", "ARTSTMT1_DESC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            dst.Tables("EDTTRPM1").Clear()
        End If

        If grdARTCUST6.Tag = "" Then
            Refresh_FollowUps()
        End If

        EnforceConstraints(True)

        grdARTPYMT3.DisplayLayout.Bands("ARTPYMT3").SummaryFooterCaption = ""

        dst.AcceptChanges()
        tabMain.SelectedTab = tabMain.Tabs(0)
        SetControlPanel()

        splSOTAEBP1.Visible = False

        For Each grdname As String In GRDs.Keys
            Dim grd As UltraWinGrid.UltraGrid = GRDs(grdname)
            grd.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        Next

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading " & HFs("CUST_CODE"))

        EnforceConstraints(False)

        grdARTPYMT2.Tag = "*"
        grdSOTORDR0.Tag = "*"
        grdSOTINVH1.Tag = "*"
        grdSATCUSTS.Tag = "*"

        ASCDATA1.ExecuteSQL("Delete from " & ARTCUST0)
        ASCDATA1.ExecuteSQL("Insert into " & ARTCUST0 _
            & " Select Distinct CUST_CODE from ARTCUST1 " _
            & " where CUST_CODE = '" & HFs("CUST_CODE") & "'" _
            & " or CUST_CREDIT_GROUP_CUST = '" & HFs("CUST_CODE") & "'")

        ASCDATA1.ExecuteSQL("Delete from " & ARTCUSTS)
        ASCDATA1.ExecuteSQL("Insert into " & ARTCUSTS _
        & " Select Distinct CUST_CODE from ARTCUST1 " _
        & " where CUST_CODE = '" & HFs("CUST_CODE") & "'" _
        & " or CUST_BILL_TO_CUST = '" & HFs("CUST_CODE") & "'" _
        & " or CUST_CREDIT_GROUP_CUST = '" & HFs("CUST_CODE") & "'")

        ASCMAIN1.sql = "SELECT DISTINCT CUST_CODE, CUST_NAME FROM (" _
         & " SELECT CUST_CODE, CUST_NAME FROM ARTCUST1 WHERE CUST_CODE = '" & HFs("CUST_CODE") & "'" _
         & " UNION" _
         & " SELECT CUST_CODE, CUST_NAME FROM ARTCUST1 WHERE CUST_BILL_TO_CUST = '" & HFs("CUST_CODE") & "'" _
         & " UNION" _
         & " SELECT CUST_CODE, CUST_NAME FROM ARTCUST1 WHERE CUST_CREDIT_GROUP_CUST = '" & HFs("CUST_CODE") & "'" _
         & ") ORDER BY CUST_CODE"

        Load_Customer_and_Credit_History()

        Fill_Records("ARTCUSTD", HFs("CUST_CODE"))
        Sort_grdColumns(grdARTCUSTD, "CUST_CODE")
        Fill_Records("ARTCUSTS")
        Sort_grdColumns(grdARTCUSTS, "CUST_CODE")

        Fill_Records("ARTCUST2", HFs("CUST_CODE"))
        '  Sort_grdColumns(grdARTCUST2, "CUST_SHIP_TO_CODE")

        'Fill_Records("ARTCUST4", HFs("CUST_CODE"))
        'Sort_grdColumns(grdARTCUST4, "CUST_BILL_TO_CUST")

        Load_ARTOPEN1()

        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            ASCMAIN1.sql = "SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '810' AND CUST_CODE = '" & HFs("CUST_CODE") & "' AND EDI_STATUS = 'P'"
            Fill_Records("EDTTRPM1", String.Empty, True, ASCMAIN1.sql)
        End If

        grdARTSTMT1.DisplayLayout.Rows.FixedRows.Clear()

        Fill_Records("ARTSTMT1") ', HFs("CUST_CODE"))
        Fill_Records("ARTSTMT1_DESC")
        Dim rowARTSTMT1 As DataRow = dst.Tables("ARTSTMT1").NewRow
        'rowARTSTMT1.Item("CUST_CODE") = HFs("CUST_CODE")
        rowARTSTMT1.Item("OPS_YYYYPP") = "999999"
        rowARTSTMT1.Item("LEGEND") = "Open AR"

        rowARTSTMT1.Item("TYP_I_OPEN") = dst.Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'I'")
        rowARTSTMT1.Item("TYP_R_OPEN") = dst.Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'R'")
        rowARTSTMT1.Item("TYP_C_OPEN") = dst.Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'C'")
        rowARTSTMT1.Item("TYP_D_OPEN") = dst.Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'D'")
        rowARTSTMT1.Item("TYP_B_OPEN") = dst.Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'B'")
        rowARTSTMT1.Item("TYP_O_OPEN") = dst.Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'O'")

        Dim INV_CLSD_DYS As Integer = 0
        Dim INV_CLSD_CNT As Integer = 0
        Dim INV_CLSD_AMT As Decimal = 0
        Dim INV_CLSD_DDS As Decimal = 0
        Dim TOTAL_OPEN_DDS As Decimal = 0
        For Each row As DataRow In dst.Tables("ARTOPEN1").Select
            Dim DAYS As Integer = Val(row.Item("AGE") & "")
            Dim INV_BALANCE As Decimal = Val(row.Item("INV_BALANCE") & "")
            TOTAL_OPEN_DDS += DAYS * INV_BALANCE
            If row.Item("INV_TYPE") = "I" Then
                If row.Item("OPS_YYYYPP_PAID") & "" = "" Then
                    ' STILL OPEN
                Else
                    INV_CLSD_DYS += Val(row.Item("DTP") & "")
                    INV_CLSD_CNT += 1
                    INV_CLSD_AMT += Val(row.Item("AMT_PAID") & "")
                    INV_CLSD_DDS += Val(row.Item("DTP") & "") * Val(row.Item("AMT_PAID") & "")
                End If
            End If
        Next

        rowARTSTMT1.Item("AGE_1") = AGED_TOTALS(1)
        rowARTSTMT1.Item("AGE_2") = AGED_TOTALS(2)
        rowARTSTMT1.Item("AGE_3") = AGED_TOTALS(3)
        rowARTSTMT1.Item("AGE_4") = AGED_TOTALS(4)
        rowARTSTMT1.Item("TOTAL_OPEN_AMT") = AGED_TOTALS(0)
        rowARTSTMT1.Item("TOTAL_OPEN_DDS") = TOTAL_OPEN_DDS
        rowARTSTMT1.Item("DUE_1") = DUE_TOTALS(1)
        rowARTSTMT1.Item("DUE_2") = DUE_TOTALS(2)
        rowARTSTMT1.Item("DUE_3") = DUE_TOTALS(3)
        rowARTSTMT1.Item("DUE_4") = DUE_TOTALS(4)

        rowARTSTMT1.Item("INV_CLSD_DYS") = INV_CLSD_DYS
        rowARTSTMT1.Item("INV_CLSD_CNT") = INV_CLSD_CNT
        rowARTSTMT1.Item("INV_CLSD_AMT") = INV_CLSD_AMT
        rowARTSTMT1.Item("INV_CLSD_DDS") = INV_CLSD_DDS


        dst.Tables("ARTSTMT1").Rows.Add(rowARTSTMT1)

        Sort_grdColumns(grdARTSTMT1, "OPS_YYYYPP".ToLower)
        grdARTSTMT1.Text = "AR Aging History for " & HFs("CUST_CODE") & ":" & rowARTCUST1.Item("CUST_NAME")

        If grdARTSTMT1.Rows.Count > 0 Then
            grdARTSTMT1.DisplayLayout.Rows.FixedRows.Add(grdARTSTMT1.Rows(0))
        End If

        Fill_Records("ARTSLSMA", HFs("CUST_CODE"))
        Sort_grdColumns(grdARTSLSMA, "OPS_YYYYPP")
        grdARTSLSMA.Tag = ""

        Load_Order_Guide()

        If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then
            Sort_grdColumns(grdSOTAEBP1, "STYLE_CODE,COLOR_CODE")
        Else
            Sort_grdColumns(grdSOTAEBP1, "ITEM_CODE")
        End If
        splSOTAEBP1.Visible = True

        EnforceConstraints(True)

        Fill_Records("SOTSREP1")
        Fill_Records("TATTERM1")

        Dim SREP_CODE As String = rowARTCUST1.Item("SREP_CODE") & ""
        Dim SREP_NAME As String = LookUp("SOTSREP1", SREP_CODE, True).Item("SREP_NAME") & ""
        lblSREP.Text = "SRep: " & SREP_CODE & " - " & SREP_NAME
        'Mary requested removal
        'If rowARTCUST1.Item("SREP_CODE_OS") & "" <> "" Then
        '    SREP_CODE = rowARTCUST1.Item("SREP_CODE_OS") & ""
        '    SREP_NAME = LookUp("SOTSREP1", SREP_CODE, True).Item("SREP_NAME") & ""
        '    lblSREP.Text &= " / " & SREP_CODE & ":" & SREP_NAME
        'End If

        Absx1.txtFor("CSZ").Text = Absx1.txtFor("CUST_CITY").Text & ", " & Absx1.txtFor("CUST_STATE").Text

        Setup_ARTCUST1_tables()
        tabMain.SelectedTab = tabMain.Tabs("Info")

        'Absx1.txtFor("LAST_3612").Text = TAC.ARCMAIN1.Last_3612_SAvg(HFs("CUST_CODE"))

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Delete_Record()
        Stop
    End Sub

    Sub Update_Record()
        'BeginTrans()

        'CommitTrans("Update")
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "Select Customer"
                Absx1.txtFor("CUST_CODE").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ARTCUST1"
            E.COLUMN_NAME = "CUST_CODE"
            E.CODE_VALUE = HFs("CUST_CODE")
            E.DESC_VALUE = HFs("CUST_NAME")
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "ARTCUST1"
        E.TABLE_KEY_CAPTION = "Customer"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("CUST_CODE").Text '  HFs("CUST_CODE")
            E.TABLE_KEY_DESC = Absx1.txtFor("CUST_NAME").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E" Or EntryMode = "A")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

        Load_Popup_Menu(grdARTSTMT1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdARTOPEN1, "SSSBBBBBBBBBBBB", "Show Filter", "Show GroupBox", "Show $0 Balance Items", _
                        "email", "Fax", "Show", "Sales Order Inquiry", "Customer Returns Inquiry", "Show Pymt Applications", _
                        "Retrieve Paid Invoices", "Create Log", "Total Balance", "Change Terms", "Credit Card", "Sales Order Entry")
        Load_Popup_Menu(grdSOTORDR0, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Order Inquiry", "Sales Order Entry", "Print Selected")
        Load_Popup_Menu(grdSOTORDR2, "B", "Style Status Inquiry")
        Load_Popup_Menu(grdSATCUSTS, "SSB", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdARTCUST6, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTAEBP1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
        Load_Popup_Menu(grdSOTAEBP2, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice")
        Load_Popup_Menu(grdARTCUSTS, "B", "Customer Inquiry")

        If (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") Then
            Load_Popup_Menu(grdSOTINVH1, "SSSPBPBBBPB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice", "Resend EDI Invoice")
        ElseIf (ASCMAIN1.DBS_COMPANY = "RGI" AndAlso ASCMAIN1.DBS_SERVER = "RGI") Then
            Load_Popup_Menu(grdSOTINVH1, "SSSPBPBBBPB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice", "Send Invoice to Web")
        Else
            Load_Popup_Menu(grdSOTINVH1, "SSSPBPBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice")
        End If

        'If Val(ASCDATA1.GetDataValue("SELECT COUNT(*) FROM ARTCCPRC") & String.Empty) = 0 Then
        '    Load_Popup_Menu(grdSOTINVH1, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice")
        'Else
        '    Load_Popup_Menu(grdSOTINVH1, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice", "Charge Credit Card")
        'End If

        Load_Popup_Menu(grdSOTINVH2, "B", "Style Status Inquiry")
        ' Load_Popup_Menu(grdARTCUST4, "B", "Customer Inquiry")

        If ASCMAIN1.USER_SECURITY_CODEs.Contains("AR") Then
            Load_Popup_Menu(grdARTPYMT2, "BSS", "Issue CC Credit", "Show Filter", "Show GroupBox")
        Else
            Load_Popup_Menu(grdARTPYMT2, "SS", "Show Filter", "Show GroupBox")
        End If
        Load_Popup_Menu(grdTATEVNT1, "B", "Show email")
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
        Dim tlb_btn As UltraWinToolbars.ButtonTool

        Select Case grd.Name

            Case "grdTATEVNT1"
                tlb_btn = DirectCast(tlb_pop.Tools("Show email"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("EVENT_TYPE").Value = "QUOEML"))

        End Select

        If tlb_pop.Tools.Exists("Add to Log") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Add to Log"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Enabled = ScreenMode
            tlb_btn = DirectCast(tlb_pop.Tools("Show Log"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Enabled = False
            tlb_btn = DirectCast(tlb_pop.Tools("Follow-Up"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Enabled = False
        End If

        If tlb_pop.Tools.Exists("Show All Follow-Ups") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show All Follow-Ups"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = Not ScreenMode
        End If

        If tlb_pop.Tools.Exists("Show Outline") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Show Outline"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = ScreenMode
        End If


        If tlb_pop.Tools.Exists("Change Terms") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Change Terms"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = ASCMAIN1.USER_SECURITY_CODEs.Contains("CR")
        End If


        If tlb_pop.Tools.Exists("Credit Card") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Credit Card"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = ASCMAIN1.USER_SECURITY_CODEs.Contains("CR")

            Dim TOTAL_BALANCE As Decimal = 0
            For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                TOTAL_BALANCE += Val(grow.Cells("INV_BALANCE").Value & "")
            Next
            tlb_btn.Tag = TOTAL_BALANCE

            If TOTAL_BALANCE >= 0 Then
                tlb_btn.SharedProps.Caption = "Credit Card Payment"
            Else
                tlb_btn.SharedProps.Caption = "Credit Card Refund"
            End If
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdARTSTMT1"
                    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                        e.Cancel = True
                    End If

                Case "grdARTPYMT2"
                    If grd.ActiveRow.Cells("CCPA_NO").Text = "" _
                    Or grd.ActiveRow.Cells("CCPA_NO_CREDIT").Text <> "" _
                    Or Val(grd.ActiveRow.Cells("CUST_PYMT_AMT").Text) <= 0 _
                    Then
                        tlb_pop.Tools("Issue CC Credit").SharedProps.Visible = False
                        'e.Cancel = True
                    Else
                        tlb_pop.Tools("Issue CC Credit").SharedProps.Visible = True
                    End If

                Case "grdARTOPEN1"

                    If grd.ActiveRow IsNot Nothing Then
                        If grd.ActiveRow.Band.Index <> 0 Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    tlb_pop.Tools("Sales Order Inquiry").SharedProps.Visible = (grd.ActiveRow.Cells("INV_TYPE").Text = "I")
                    tlb_pop.Tools("Customer Returns Inquiry").SharedProps.Visible = (grd.ActiveRow.Cells("INV_TYPE").Text = "R")

                    If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
                        tlb_pop.Tools("Sales Order Entry").SharedProps.Visible = (grd.ActiveRow.Cells("INV_TYPE").Text = "I")
                    Else
                        tlb_pop.Tools("Sales Order Entry").SharedProps.Visible = False
                    End If

                Case "grdSOTORDR0"
                    If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
                        tlb_pop.Tools("Sales Order Entry").SharedProps.Visible = True
                        tlb_pop.Tools("Print Selected").SharedProps.Visible = True
                    Else
                        tlb_pop.Tools("Sales Order Entry").SharedProps.Visible = False
                        tlb_pop.Tools("Print Selected").SharedProps.Visible = False
                    End If

                Case "grdSOTINVH1"
                    If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then

                        Dim Sql As String = "SELECT EDTSLSP1.EDI_ID_810 "
                        Sql &= "  FROM EDTSLSP1, EDTTRPM1"
                        Sql &= "  WHERE EDTSLSP1.EDI_ID_810=EDTTRPM1.EDI_TP_ID"
                        Sql &= "  AND EDTSLSP1.CUST_CODE = :PARM1"
                        Sql &= "  AND EDI_DOC_NO = '810'"
                        Dim rowEDTSLSP1 As DataRow = ASCDATA1.GetDataRow(Sql, "V", HFs("CUST_CODE"))

                        'tlb_pop.Tools("Resend EDI Invoice").SharedProps.Visible = dst.Tables("EDTTRPM1").Rows.Count > 0 AndAlso ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                        tlb_pop.Tools("Resend EDI Invoice").SharedProps.Visible = rowEDTSLSP1 IsNot Nothing AndAlso ASCMAIN1.USER_SECURITY_CODEs.Contains("ED")
                    End If
            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Show $0 Balance Items"
                Filter_ARTOPEN1()
            Case "Retrieve Paid Invoices"
                Dim numDays As Double = 0
                Using FRM As New ASFMSGBF
                    numDays = FRM.Get_numint_from_User("Days to Retrieve", "Retrieve Paid Invoices")
                End Using
                Retrieve_Paid_Invoices(numDays)

            Case "Expand All"
                grd.Rows.ExpandAll(True)
            Case "Collapse All"
                grd.Rows.CollapseAll(True)

            Case "Credit Card"
                Dim TOTAL_BALANCE As Decimal = Val(e.Tool.Tag)
                Dim msg As String = ""
              
                If grdARTOPEN1.Selected.Rows.Count = 0 Then
                    msg = "No AR Items have been Selected." _
                            & vbCrLf & vbCrLf & "The next screen will ask you to enter the customer's credit card information," _
                            & vbCrLf & " charge the customers credit card for the amount specified," _
                            & vbCrLf & " and record the payment as an on/account credit."
                Else
                    If grdARTOPEN1.Selected.Rows.Count = 1 Then
                        msg = "You have selected 1 AR Items totaling " & Format(TOTAL_BALANCE, "$#,##0.00") & "." _
                             & vbCrLf & vbCrLf & "The next screen will ask you to enter the customer's credit card information," _
                             & vbCrLf & IIf(TOTAL_BALANCE < 0, "credit", "charge") & " the customers credit card" _
                             & vbCrLf & " for up to " & Format(TOTAL_BALANCE, "$#,##0.00") & " and apply it to the item selected."
                    Else
                        msg = "You have selected " & CStr(grdARTOPEN1.Selected.Rows.Count) & " AR Items totaling " & Format(TOTAL_BALANCE, "$#,##0.00") & "." _
                             & vbCrLf & vbCrLf & "The next screen will ask you to enter the customer's credit card information," _
                             & vbCrLf & IIf(TOTAL_BALANCE < 0, "credit", "charge") & " the customers credit card" _
                             & vbCrLf & " for exactly" & Format(TOTAL_BALANCE, "$#,##0.00") & " and apply it to the items selected."
                    End If
                End If

                If MsgBox(msg & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                If ASCMAIN1.Logical_Open("R", "ARRPYMT2", , , , 2) Then
                    Credit_Card(TOTAL_BALANCE, grdARTOPEN1.Selected.Rows.Count)
                    ASCMAIN1.MultiTask_Release(, , 2)
                End If


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Print Selected"
                Try
                    If grdSOTORDR0.Selected.Rows.Count = 0 Then
                        MessageBox.Show("You must Select at least One Sales Order.", "Print Selected", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Dim ordrGroupNoList As New List(Of String)
                    For Each grdRow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                        Select Case grdRow.Band.Key
                            Case grdSOTORDR0.DisplayLayout.Bands(0).Key
                                If Not ordrGroupNoList.Contains(grdRow.Cells("ORDR_GROUP_NO").Value & String.Empty) Then
                                    ordrGroupNoList.Add(grdRow.Cells("ORDR_GROUP_NO").Value & String.Empty)
                                End If
                        End Select
                    Next

                    If ordrGroupNoList.Count = 0 Then
                        MessageBox.Show("You must Select at least One Sales Order.", "Print Selected", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Preparing Order")

                    Dim REPORT_NAME As String = "SORORDR1"
                    If Not REPORTS.ContainsKey(REPORT_NAME) Then
                        REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
                        REPORTS(REPORT_NAME).Prepare_dst(False, "")
                    End If

                    REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {" and SOTORDR1.ORDR_NO in ('" & String.Join("', '", ordrGroupNoList) & "')"})
                    With REPORTS(REPORT_NAME).clsASCBASE1
                        .Print_Report_Begin()
                        Dim SUBT As String = ""
                        .CR_params.Add("SUBT", SUBT)
                        .Generate_Report(REPORT_NAME, "Sales Order", SUBT, True, , , , , False)
                        .Print_Report_End()
                    End With

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Print Selected", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                End Try

            Case "Resend EDI Invoice"
                Dim sql As String = String.Empty
                sql = "SELECT COMPANY_CODE, EDI_OUTBOUND_DOC_NO FROM EDTSYSIH "
                sql &= " WHERE (COMPANY_CODE, EDI_OUTBOUND_DOC_NO) IN "
                sql &= " (SELECT O1.COMPANY_CODE, MAX(O1.EDI_OUTBOUND_DOC_NO) "
                sql &= " FROM EDT810O1 O1, EDTSYSIH IH "
                sql &= " WHERE O1.COMPANY_CODE=IH.COMPANY_CODE "
                sql &= " AND O1.EDI_OUTBOUND_DOC_NO=IH.EDI_OUTBOUND_DOC_NO "
                sql &= " AND IH.EDI_PROCESS_IND = '2' "
                sql &= " AND O1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
                sql &= " AND EDI_INVOICE_NUMBER = '" & grd.ActiveRow.Cells("INV_NO").Value & "' "

                'sql &= " AND TRIM(EDI_TP_ID) NOT IN (Select EDI_TP_ID From EDTTRPM1 Where cust_code In (Select ED_PARM_FACTOR From EDTPARM1) And EDI_DOC_NO = '810')"
                sql &= " AND TRIM(EDI_TP_ID)"
                sql &= " IN (SELECT EDTSLSP1.EDI_ID_810"
                sql &= " FROM EDTSLSP1, EDTTRPM1"
                sql &= " WHERE EDTSLSP1.EDI_ID_810=EDTTRPM1.EDI_TP_ID AND EDTSLSP1.CUST_CODE =  '" & HFs("CUST_CODE") & "' AND EDI_DOC_NO = '810')"

                sql &= " GROUP BY O1.COMPANY_CODE) "
                Dim rowEDTSYSIH As DataRow = ASCDATA1.GetDataRow(sql)

                If rowEDTSYSIH Is Nothing Then
                    MessageBox.Show("The selected invoice is either not an EDI Invoice or the EDI Document has not been sent.", "Resend EDI Invoice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Try
                    sql = "Update EDTSYSIH set EDI_PROCESS_IND = '1' Where (COMPANY_CODE, EDI_OUTBOUND_DOC_NO) IN (" & sql & ")"
                    ASCDATA1.ExecuteSQL(sql)
                    MessageBox.Show("The Invoice's 810 has been reset to be transmitted.", "Resend EDI Invoice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("The following error occurred: " & ex.Message, "Resend EDI Invoice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try

            Case "email", "email Invoice", "email Statement", "Fax", "Fax Invoice", "Fax Statement"

                Dim FILENAME As String = ""
                Dim ATTACHMENT As String = ""
                Dim SUBJECT As String = ""
                Dim INV_NO As String = ""

                If e.Tool.OwningMenu.Key = "grdARTSTMT1" Then
                    Dim RYP As String = grd.ActiveRow.Cells("OPS_YYYYPP").Value
                    Dim STMT_NO As String = grd.ActiveRow.Cells("STMT_NO").Value & ""
                    If ASCMAIN1.useUNCPath Then
                        FILENAME = $"{ASCMAIN1.Folders("SharedRoot")}\OSG\" & RYP & "\PDF\" & STMT_NO & ".PDF"
                    Else
                        FILENAME = "S:\OSG\" & RYP & "\PDF\" & STMT_NO & ".PDF"
                    End If
                    ATTACHMENT = ASCMAIN1.Folders("Temp") & STMT_NO & "." & "PDF"
                    SUBJECT = "Statement for " & Mid(RYP, 5, 2) & "/" & Mid(RYP, 1, 4) & _
                            " (Acct# " & grd.ActiveRow.Cells("CUST_CODE").Value & " " & rowARTCUST1.Item("CUST_NAME") & ")"

                ElseIf e.Tool.OwningMenu.Key = "grdARTOPEN1" Then
                    'Dim INV_TYPE As String = grdARTOPEN1.ActiveRow.Cells("INV_TYPE").Value & ""

                    'Dim ORDR_NO As String = grdARTOPEN1.ActiveRow.Cells("ORDR_NO").Value & ""
                    If grd.Selected.Rows.Count <= 1 Then
                        INV_NO = grd.ActiveRow.Cells(INV_NUM_column).Value & ""
                    Else
                        For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Selected.Rows
                            INV_NO &= "," & grdRow.Cells(INV_NUM_column).Value & ""
                        Next
                        INV_NO = INV_NO.Substring(1).Trim
                    End If

                    FILENAME = TAC.SOCMAIN1.Create_Invoice(Me, INV_NO)
                    SUBJECT = "Invoice " & INV_NO

                ElseIf e.Tool.OwningMenu.Key = "grdSOTINVH1" Then
                    If grd.Selected.Rows.Count <= 1 Then
                        INV_NO = grd.ActiveRow.Cells("INV_NO").Value & ""
                    Else
                        For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Selected.Rows
                            INV_NO &= "," & grdRow.Cells("INV_NO").Value & ""
                        Next
                        INV_NO = INV_NO.Substring(1).Trim
                    End If
                    FILENAME = TAC.SOCMAIN1.Create_Invoice(Me, INV_NO)
                    SUBJECT = "Invoice " & INV_NO

                End If

                If e.Tool.Key Like "email*" Then
                    TAC.SOCMAIN1.email_Invoice(Me, _
                        Absx1.txtFor("CUST_CODE").Text, _
                        Absx1.txtFor("CUST_NAME").Text, _
                        Absx1.txtFor("CUST_EMAIL").Text, _
                        Absx1.txtFor("CUST_EMAIL").Text, _
                        FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT, INV_NO)

                    ' Send_email(FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT, INV_NO) ' WILL NEED SOMETHING FOR STATEMENTS
                ElseIf e.Tool.Key Like "Fax*" Then
                    Send_fax(FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT)
                End If

            Case "Show", "Show Invoice", "Show Statement", "Show Credit"
                Dim FILENAME As String = ""

                If e.Tool.OwningMenu.Key = "grdARTSTMT1" Then
                    Try
                        Dim RYP As String = grdARTSTMT1.ActiveRow.Cells("OPS_YYYYPP").Value
                        Dim STMT_NO As String = grdARTSTMT1.ActiveRow.Cells("STMT_NO").Value & ""
                        If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
                            FILENAME = $"{ASCMAIN1.Folders("SharedRoot")}\OSG\" & RYP & "\PDF\" & STMT_NO & ".PDF"
                        Else
                            FILENAME = "S:\OSG\" & RYP & "\PDF\" & STMT_NO & ".PDF"
                        End If
                    Catch ex As Exception

                    End Try

                ElseIf e.Tool.OwningMenu.Key = "grdARTOPEN1" Or e.Tool.OwningMenu.Key = "grdSOTORDR1" Or e.Tool.OwningMenu.Key = "grdSOTINVH1" Or e.Tool.OwningMenu.Key = "grdSOTAEBP2" Then

                    If grd.ActiveRow IsNot Nothing Then
                        If Not grd.ActiveRow.Selected Then
                            grd.Selected.Rows.Clear()
                            grd.ActiveRow.Selected = True
                        End If
                    End If

                    If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.Selected Then
                        Exit Sub
                    End If

                    Dim INV_TYPE As String = ""
                    If e.Tool.OwningMenu.Key = "grdARTOPEN1" Then
                        INV_TYPE = grd.ActiveRow.Cells("INV_TYPE").Value & ""
                    Else
                        INV_TYPE = grd.ActiveRow.Cells("INV_TYPE").Value & ""
                    End If
                    Dim INV_NO As String ' = IIf(e.Tool.OwningMenu.Key = "grdARTOPEN1", grd.ActiveRow.Cells("INV_NUM").Value & "", grd.ActiveRow.Cells("INV_NO").Value & "")

                    If INV_TYPE <> "I" And INV_TYPE <> "C" And INV_TYPE <> "R" Then
                        Exit Sub
                    End If

                    If e.Tool.OwningMenu.Key = "grdARTOPEN1" Then
                        INV_NO = grd.ActiveRow.Cells(INV_NUM_column).Value & ""
                    Else
                        INV_NO = grd.ActiveRow.Cells("INV_NO").Value & ""
                    End If

                    Dim INV_NOs As String = ""
                    Dim INV_NO_count As Int32 = grd.Selected.Rows.Count

                    If INV_NO_count = 0 Then
                        INV_NOs = INV_NO
                    Else
                        For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                            If e.Tool.OwningMenu.Key = "grdARTOPEN1" Then
                                INV_NO = grow.Cells(INV_NUM_column).Text
                            Else
                                INV_NO = grow.Cells("INV_NO").Text
                            End If

                            If INV_NO <> "" Then
                                INV_NOs &= "," & INV_NO
                            End If
                        Next
                        INV_NOs = Mid(INV_NOs, 2)
                    End If
                    If INV_NOs <> "" Then
                        FILENAME = TAC.SOCMAIN1.Create_Invoice(Me, INV_NOs)
                    End If

                End If

                Show_Document(FILENAME)

            Case "See Entire Order"
                Dim SO_ORDER_NO As String = grd.ActiveRow.Cells("SO_ORDER_NO").Text
                Dim row As DataRow = dst.Tables("SOTORDR1").Rows.Find(New Object() {SO_ORDER_NO})
                If row Is Nothing Then
                    Fill_Records("SOTORDR1", "", False, "Select * from SOTORDR1 where SO_ORDER_NO = '" & SO_ORDER_NO & "'")
                End If
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                    If grow.Cells("SO_ORDER_NO").Text = SO_ORDER_NO Then
                        grdSOTORDR0.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
                        grdSOTORDR0.Selected.Rows.Clear()
                        grdSOTORDR0.ActiveRow = grow
                        grow.Selected = True
                        Exit For
                    End If
                Next

            Case "Track Shipment"
                If grd.ActiveRow.Cells("SHIP_REF").Text <> "" Then
                    'SOCMAIN1.Track_Shipment(grd.ActiveRow.Cells("SHIP_VIA_CODE").Text, grd.ActiveRow.Cells("SHIP_REF").Text)
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = ""
                If e.Tool.OwningMenu.Key = "grdARTOPEN1" Then
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                Else
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                End If
                If ORDR_NO <> "" Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI", "F", "SO")
                End If
            Case "Sales Order Entry"
                If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
                    Dim ORDR_NO As String = ""
                    If e.Tool.OwningMenu.Key = "grdARTOPEN1" Then
                        ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                    Else
                        ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                    End If
                    If ORDR_NO <> "" Then
                        Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDR1", "F", "SO")
                    End If
                End If
            Case "Show Pymt Applications"
                Dim INV_NO As String = ""
                Dim INV_TYPE As String = ""
                Dim CUST_CODE_PYMT As String = ""
                If grd.Name = "grdARTOPEN1" Then
                    INV_NO = grd.ActiveRow.Cells(INV_NUM_column).Text
                    INV_TYPE = grd.ActiveRow.Cells("INV_TYPE").Text
                    If dst.Tables("ARTPYMTD").Select("INV_TYPE = '" & INV_TYPE & "' AND INV_NUM = '" & INV_NO & "'").Length = 0 Then
                        'If dst.Tables("ARTPYMTD").Select("INV_TYPE = '" & INV_TYPE & "' AND " & INV_NUM_column & " = '" & INV_NO & "'").Length = 0 Then
                        Fill_Records("ARTPYMTD", New String() {INV_TYPE, INV_NO}, False)
                    End If
                    grd.DisplayLayout.Bands(1).Hidden = False
                    grd.ActiveRow.Expanded = True

                Else
                    INV_NO = grd.ActiveRow.Cells(INV_NUM_column).Text
                    INV_TYPE = grd.ActiveRow.Cells("INV_TYPE").Text
                    txtINV_NO_PYMT.Text = INV_NO
                    txtINV_NO_PYMT.Tag = INV_TYPE
                    Process_INV_NO_PYMT()
                End If

            Case "Create Log"

                Dim invNum As String = ""
                Dim invDate As String = ""
                Dim invBalance As String = ""
                Dim logText As String = "Promise to Pay" & vbCrLf
                Dim invTotal As Double = 0

                If grd.Selected.Rows.Count <> 0 Then
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        invNum = grow.Cells("INV_NUM").Text
                        invDate = grow.Cells("INV_DATE").Text
                        invBalance = grow.Cells("INV_BALANCE").Text
                        invTotal = invTotal + grow.Cells("INV_BALANCE").Value
                        logText = logText & invNum & " " & invDate & " " & invBalance & vbCrLf
                    Next
                Else
                    invNum = grd.ActiveRow.Cells("INV_NUM").Text
                    invDate = grd.ActiveRow.Cells("INV_DATE").Text
                    invBalance = grd.ActiveRow.Cells("INV_BALANCE").Text
                    invTotal = invTotal + grd.ActiveRow.Cells("INV_BALANCE").Value
                    logText = logText & invNum & " " & invDate & " " & invBalance & vbCrLf
                End If

                logText = logText & "Total: " & String.Format("{0:c}", invTotal)

                If Len(logText) > 1000 Then

                    MsgBox("The length of the auto-generated log (" & Len(logText) & ") for the selected items" _
                           & " exceeds the maximum size (1000). Please reduce the number of selected items.", _
                            MsgBoxStyle.OkOnly, "Cannot Proceed")

                Else
                    Dim F As New ASFCONV2(Me, "ARTCUST1", HFs("CUST_CODE"), logText)
                    F.EntryMode = "N"
                    F.ShowDialog()
                    If F.result = "U" Then
                        dst.Tables("TATCONV1").Rows.Add(F.rowTATCONV1.ItemArray)
                        Update_Record_TDA("TATCONV1")
                    End If
                    F.Dispose()

                End If

            Case "Total Balance"

                If grd.Selected.Rows.Count <> 0 Then
                    Dim invTotal As Decimal = 0
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        invTotal = invTotal + grow.Cells("INV_BALANCE").Value
                    Next
                    MsgBox("Total Balance: " & String.Format("{0:c}", invTotal), _
                       MsgBoxStyle.OkOnly, _
                       String.Format("Total Balance for {0} Item(s) Selected", grd.Selected.Rows.Count))
                Else
                    MsgBox("You must select the rows that you want totaled", _
                            MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                End If

            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                Context_Launch("Select Customer", CUST_CODE, "Customer Inquiry", "ARFCINQ1")

            Case "Customer Order Inquiry"
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                Context_Launch("Select", CUST_CODE & ":" & ORDR_GROUP_NO, e.Tool.Key, "SOFCORD1")

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Show email"
                If grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "QUOEML" Then
                    Dim FILENAME As String = grd.ActiveRow.Cells("EVENT_KEY").Value & ".EML"
                    Show_Document(ASCMAIN1.Folders("Archive") & "\email\Sent\" & FILENAME)
                End If

            Case "Change Terms"

                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("TERM_CODE")
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False

                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    If ASCMAIN1.CodeSelector.SelectedCodes.Count <> 0 Then
                        Dim TERM_CODE As String = ASCMAIN1.CodeSelector.SelectedCode

                        Dim TERM_CODE_orig As String = grdARTOPEN1.ActiveRow.Cells("TERM_CODE").Value & ""
                        Dim INV_TYPE As String = grdARTOPEN1.ActiveRow.Cells("INV_TYPE").Value & ""
                        Dim INV_NUM As String = grdARTOPEN1.ActiveRow.Cells("INV_NUM").Value & ""
                        Dim CUST_CODE As String = grdARTOPEN1.ActiveRow.Cells("CUST_CODE").Value & ""

                        Dim INV_DATE As Date = grdARTOPEN1.ActiveRow.Cells("INV_DATE").Value & ""
                        Dim INV_DUE_DATE_orig As Date = grdARTOPEN1.ActiveRow.Cells("INV_DUE_DATE").Value & ""

                        Dim rowTATTERM1_orig As DataRow = LookUp("TATTERM1", TERM_CODE_orig)
                        Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE)

                        Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, TERM_CODE, rowTATTERM1, INV_DATE)

                        If MsgBox("Click Yes to change the Terms Code of" _
                                  & vbCrLf & " AR Item " & grdARTOPEN1.ActiveRow.Cells("INV_NUM").Value & " with a Document Date of " & Format(INV_DATE, "MM/dd/yy") _
                                  & vbCrLf & " from " & TERM_CODE_orig & " with a Due Date of " & Format(INV_DUE_DATE_orig, "MM/dd/yy") _
                                  & vbCrLf & " to " & TERM_CODE & " with a Due Date of " & Format(INV_DUE_DATE, "MM/dd/yy") _
                                    & vbCrLf & vbCrLf & "An Audit Trail Record will be Recorded.", _
                                    MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If

                        BeginTrans()

                        ASCMAIN1.sql = "Update ARTOPEN1 Set TERM_CODE = :PARM1, INV_DUE_DATE = :PARM2 where INV_TYPE = :PARM3 and INV_NUM = :PARM4 and CUST_CODE = :PARM5"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VDVVV", New Object() {TERM_CODE, INV_DUE_DATE, INV_TYPE, INV_NUM, CUST_CODE})

                        TAC.TACMAIN1.Record_Event("SOTINVH1", CUST_CODE & ":" & INV_TYPE & ":" & INV_NUM, _
                                                  DATETIME_STAMP, ASCMAIN1.USER_ID, "CHGTERMS", _
                                                  "AR Terms Code Changed from " & TERM_CODE_orig & " to " & TERM_CODE, "", Me.Name)

                        CommitTrans()

                        With grdARTOPEN1.ActiveRow
                            .Cells("TERM_CODE").Value = TERM_CODE
                            .Cells("INV_DUE_DATE").Value = INV_DUE_DATE
                            .Update()
                        End With
                    End If
                End If

            Case "Send Invoice to Web"
                Try
                    If grd.Selected.Rows.Count <= 0 Then
                        Exit Sub
                    End If

                    If MessageBox.Show("Do you want to send the " & grd.Selected.Rows.Count & " selected invoice(s) to the Web?", "Send Invoice to Web", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If

                    ASCMAIN1.Progress("Sending Invoice to Web", "")
                    Me.Cursor = Cursors.WaitCursor
                    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        ASCMAIN1.Progress("Sending Invoice to Web", row.Cells("INV_NO").Value)
                        TAC.SOCMAIN1.CreateWebInvoice(Me, row.Cells("INV_TYPE").Value, row.Cells("INV_NO").Value)
                    Next

                    MessageBox.Show("Invoices sent to Web.", "Send Invoice to Web", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception

                Finally
                    ASCMAIN1.Progress("", "")
                    Me.Cursor = Cursors.Default
                End Try
        End Select
    End Sub

#End Region

#Region "Popup Menu Functions"

    'Sub Send_email(ByVal FILENAME As String, ByVal ATTACHMENT As String, ByVal SUBJECT As String, ByVal INV_NO As String)

    '    Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
    '    If Absx1.txtFor("CUST_EMAIL").Text <> "" Then
    '        EMAIL_ADDRESSs.Add(Absx1.txtFor("CUST_EMAIL").Text, Absx1.txtFor("CUST_EMAIL").Text)
    '    End If

    '    Dim ATTACHMENTs As New Dictionary(Of String, String)
    '    ATTACHMENTs.Add(ATTACHMENT, ATTACHMENT)

    '    Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
    '           (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
    '            SUBJECT, "INV", False, True, Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("CUST_NAME").Text, "Customer")

    '    If SEND_NO <> "" Then
    '        Dim ORDR_NO As String = ASCDATA1.GetDataValue("Select ORDR_NO from SOTINVH1 where INV_NO = '" & INV_NO & "'")
    '        If ORDR_NO <> "" Then
    '            TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "EML", "email Invoice " & INV_NO, SEND_NO)
    '        End If
    '    End If
    'End Sub

    Sub Send_fax(ByVal FILENAME As String, ByVal ATTACHMENT As String, ByVal SUBJECT As String)

        If FILENAME <> ATTACHMENT Then
            Try
                If My.Computer.FileSystem.FileExists(ATTACHMENT) Then
                    My.Computer.FileSystem.DeleteFile(ATTACHMENT)
                End If
                My.Computer.FileSystem.CopyFile(FILENAME, ATTACHMENT)
            Catch ex As Exception
                MsgBox("Error Processing File " & ATTACHMENT & vbCr & vbCr & ex.Message, MsgBoxStyle.OkOnly, "Cannot Copy Original Document")
                Exit Sub
            End Try
        End If

        Dim CUST_CONTACT As String = Absx1.txtFor("CUST_CONTACT").Text

        Dim frmTAFSEND1 As New TAFSEND1(Me)
        frmTAFSEND1.SEND_FROM = ASCMAIN1.USER_NAME
        frmTAFSEND1.SEND_FROM_NAME = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & ""
        'frmTAFSEND1.SEND_FROM_SIGNATURE = ""
        frmTAFSEND1.SEND_TO = Absx1.medFor("CUST_FAX").Text
        frmTAFSEND1.SEND_TO_NAME = CUST_CONTACT
        'frmTAFSEND1.SEND_CC = ""
        'frmTAFSEND1.SEND_CC_NAME = ""
        'frmTAFSEND1.SEND_BCC = ""
        'frmTAFSEND1.SEND_BCC_NAME = ""
        frmTAFSEND1.SEND_SUBJECT = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & " " & SUBJECT
        frmTAFSEND1.SEND_BODY = "Attached is the file that you have requested." & vbCrLf & "Please contact Customer Service if you have any further Questions."
        frmTAFSEND1.SEND_ENTITY_KEY = HFs("CUST_CODE")
        frmTAFSEND1.SEND_ENTITY_NAME = Absx1.txtFor("CUST_NAME").Text
        frmTAFSEND1.SEND_METHOD = "F"
        frmTAFSEND1.SEND_ENTITY_CAPTION = "Customer"
        frmTAFSEND1.SEND_ATTACHMENT = ATTACHMENT

        frmTAFSEND1.ShowDialog()

        frmTAFSEND1.Dispose()
        frmTAFSEND1 = Nothing
    End Sub

    Private Sub Charge_Credit_Card(ByVal INV_NO As String)

        Try
            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_TYPE = 'I' AND INV_NO = :PARM1", "V", New Object() {INV_NO})

            If rowSOTINVH1 Is Nothing Then
                MessageBox.Show("Unable to locate Invoice " & INV_NO, "Charge Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow("Select * from artccpa1 where ordr_no = '" & rowSOTINVH1.Item("ORDR_NO") & "' AND " _
                                                              & " CCPA_STATUS = 'A'")
            If rowSOTINVH1.Item("CCPA_NO") & String.Empty <> String.Empty Then
                ' Double check to Sale has been processed. Some CC orders are pushed through then charged
                If rowARTCCPA1 IsNot Nothing Then
                    MessageBox.Show("Invoice " & INV_NO & " already has a credit card charge against it.", "Charge Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            End If

            Dim inv_total_amount As Decimal = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty)
            If inv_total_amount <= 0 Then
                MessageBox.Show("Invoice Total Amount must be greater than $0.01", "Charge Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If MessageBox.Show("Do you want to make a Credit Card Charge in the amount of " & Format(inv_total_amount, "$#,##0.00"), "Charge Credit Card", _
                                 MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            Using frmCCProcessor As New TAC.TAFCARDF(Me)
                frmCCProcessor.test_mode = ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
                frmCCProcessor.CUST_CODE = rowSOTINVH1.Item("CUST_CODE")
                frmCCProcessor.CCPA_REASON = "C" ' Sale Captured
                frmCCProcessor.ORDR_NO = rowSOTINVH1.Item("ORDR_NO") & String.Empty
                frmCCProcessor.INV_NO = rowSOTINVH1.Item("INV_NO") & String.Empty
                frmCCProcessor.TRAN_TYPE = "S" ' Sale

                With frmCCProcessor.rowARTCCPA1
                    .Item("CUST_CODE") = rowSOTINVH1.Item("CUST_CODE")
                    .Item("CCPA_AMT") = inv_total_amount
                    .Item("CCPA_NOTE") = "CC Payment"
                End With

                Try
                    frmCCProcessor.ShowDialog()
                    Dim row As DataRow = ASCDATA1.GetDataRow("select * from ARTCCPA1 where CCPA_NO = :PARM1", "V", New Object() {frmCCProcessor.CCPA_NO & String.Empty})
                    If row IsNot Nothing AndAlso row.Item("CCPA_STATUS") & String.Empty = "A" Then
                        ASCDATA1.ExecuteSQL("UPDATE SOTINVH1 SET CCPA_NO = '" & frmCCProcessor.CCPA_NO & "' WHERE INV_NO = '" & INV_NO & "'")
                        If rowSOTINVH1.Item("ORDR_NO") & String.Empty <> String.Empty Then
                            TAC.TACMAIN1.Record_Event("SOTORDR1", rowSOTINVH1.Item("ORDR_NO"), DATETIME_STAMP, ASCMAIN1.USER_ID, "CCCHG", "Credit card charged: " & Format(inv_total_amount, "#,##0.00"))
                        End If
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            End Using

        Catch ex As Exception

        End Try
    End Sub

    Sub Credit_Card(TOTAL_BALANCE As Decimal, open_items_count As Integer)

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

        Try

            Using frmCCProcessor As New TAC.TAFCARDF(Me)
                frmCCProcessor.test_mode = ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
                frmCCProcessor.CUST_CODE = CUST_CODE
                frmCCProcessor.CCPA_REASON = "C" ' Sale Captured ?
                frmCCProcessor.ORDR_NO = ""
                frmCCProcessor.INV_NO = ""
                frmCCProcessor.TRAN_TYPE = IIf(TOTAL_BALANCE >= 0, "S", "C") ' Sale or Credit

                With frmCCProcessor.rowARTCCPA1
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CCPA_AMT") = System.Math.Abs(TOTAL_BALANCE)
                    .Item("CCPA_NOTE") = "CC Payment"
                End With

                ' need a way to indicate that the amount is locked
                ' need control if the amount is not locked that positives don't go to negatives, and negatives don't go to positives, and 0 goes only positive
                Try
                    frmCCProcessor.ShowDialog()

                    Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow("select * from ARTCCPA1 where CCPA_NO = :PARM1", "V", New Object() {frmCCProcessor.CCPA_NO & String.Empty})
                    If rowARTCCPA1 IsNot Nothing AndAlso rowARTCCPA1.Item("CCPA_STATUS") & String.Empty = "A" Then
                        ' do stuff
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            End Using

        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "grdARTPYMT2"

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If ScreenMode Then
            Exit Sub
        End If

        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    e.Handled = True
                    Me.ProcessTabKey(Not e.Shift)
                    Call Click_Command("Select Customer", e)
                End If

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Call Click_Command("Select Customer")
        End Select
    End Sub

    Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.SELECTION_NO = 0 Then Exit Sub

        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CREDIT_LIMIT"
                If Not Me.ScreenMode Then Exit Sub
                Me.CalculateRemainingBalances()
                If chkEditCredit.Checked Then
                    Absx1.dteFor("CUST_CRED_LIMIT_EST").Value = Now.Date
                End If
        End Select

    End Sub

#End Region

    Private Sub grdARTCUST6_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTCUST6.DoubleClickRow

        If grdARTCUST6.ActiveRow.IsDataRow Then
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()
            Absx1.txtFor("CUST_CODE").Focus()
            Absx1.txtFor("CUST_CODE").Text = grdARTCUST6.ActiveRow.Cells("CUST_CODE").Text
            Me.ProcessTabKey(True)
            Call Click_Command("Select Customer")
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Sub Setup_ARTCUST1_tables()

        Dim rowARTCUST6 As DataRow = LookUp("ARTCUST6", HFs("CUST_CODE"), True)

        grdARTCUST1_CODES.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("ARTCUST1_CODES")
            .Rows.Clear()
            .Rows.Add(New Object() {"SALES", "Trade", rowARTCUST1.Item("TRADE_CLASS_CODE") & "", LookUp("SOTTCLS1", rowARTCUST1.Item("TRADE_CLASS_CODE") & "", True).Item("TRADE_CLASS_DESC")})
            .Rows.Add(New Object() {"SALES", "Class", rowARTCUST1.Item("CUST_CLASS_CODE") & "", LookUp("ARTCLAS1", rowARTCUST1.Item("CUST_CLASS_CODE") & "", True).Item("CUST_CLASS_DESC")})
            .Rows.Add(New Object() {"SALES", "Frt Terms", rowARTCUST1.Item("FRT_TERMS") & "", LookUp("ASTCODE1", New String() {"SOTORDR1", "FRT_TERMS", rowARTCUST1.Item("FRT_TERMS") & ""}, True).Item("T_DESC")})
            .Rows.Add(New Object() {"SALES", "Ship Via", rowARTCUST1.Item("SHIP_VIA_CODE") & "", LookUp("SOTSVIA1", rowARTCUST1.Item("SHIP_VIA_CODE") & "", True).Item("SHIP_VIA_CODE")})
            .Rows.Add(New Object() {"SALES", "Sales Rep", rowARTCUST1.Item("SREP_CODE") & "", LookUp("SOTSREP1", rowARTCUST1.Item("SREP_CODE") & "", True).Item("SREP_NAME")})

            'If rowARTCUST1.Item("BRKR_CODE") & "" <> "" Then
            '    .Rows.Add(New Object() {"SALES", "Broker", rowARTCUST1.Item("BRKR_CODE") & "", LookUp("SOTSREP1", rowARTCUST1.Item("BRKR_CODE") & "", True).Item("SREP_NAME")})
            'End If

            If rowARTCUST1.Item("CUST_BILL_TO_CUST") & "" <> "" Then
                .Rows.Add(New Object() {"ACCTG", "Bill To", rowARTCUST1.Item("CUST_BILL_TO_CUST") & "", LookUp("ARTCUST1", rowARTCUST1.Item("CUST_BILL_TO_CUST") & "", True).Item("CUST_NAME")})
            Else
                .Rows.Add(New Object() {"ACCTG", "Bill To", "Sold-To", LookUp("ARTCUST1", rowARTCUST1.Item("CUST_CODE") & "", True).Item("CUST_NAME")})
            End If
            If rowARTCUST1.Item("CUST_CREDIT_GROUP_CUST") & "" <> "" Then
                .Rows.Add(New Object() {"ACCTG", "Credit Group", rowARTCUST1.Item("CUST_CREDIT_GROUP_CUST"), LookUp("ARTCUST1", rowARTCUST1.Item("CUST_CREDIT_GROUP_CUST") & "", True).Item("CUST_NAME")})
            Else
                If rowARTCUST1.Item("CUST_BILL_TO_CUST") & "" <> "" Then
                    .Rows.Add(New Object() {"ACCTG", "Credit Group", "Bill-To", LookUp("ARTCUST1", rowARTCUST1.Item("CUST_BILL_TO_CUST") & "", True).Item("CUST_NAME")})
                Else
                    .Rows.Add(New Object() {"ACCTG", "Credit Group", "Sold-To", LookUp("ARTCUST1", rowARTCUST1.Item("CUST_CODE") & "", True).Item("CUST_NAME")})
                End If
            End If

            If rowARTCUST1.Item("CUST_CODE_ALLO") & "" <> "" Then
                .Rows.Add(New Object() {"SALES", "Allocation", rowARTCUST1.Item("CUST_CODE_ALLO") & "", LookUp("ARTCUST1", rowARTCUST1.Item("CUST_CODE_ALLO") & "", True).Item("CUST_NAME")})
            End If
            If rowARTCUST1.Item("VEND_CODE") & "" <> "" Then
                .Rows.Add(New Object() {"ACCTG", "Vendor", rowARTCUST1.Item("VEND_CODE") & "", LookUp("APTVEND1", rowARTCUST1.Item("VEND_CODE") & "", True).Item("VEND_NAME")})
            End If

            If rowARTCUST1.Item("COLLECTOR_CODE") & "" <> "" Then
                .Rows.Add(New Object() {"ACCTG", "Collector", rowARTCUST1.Item("COLLECTOR_CODE") & "", LookUp("ARTCOLL1", rowARTCUST1.Item("COLLECTOR_CODE") & "", True).Item("COLLECTOR_NAME")})
            End If
        End With

        grdARTCUST1_STATS.DisplayLayout.Bands(0).SortedColumns.Clear()
        dst.Tables("ARTCUST1_STATS").Rows.Clear()
        With rowARTCUST6
            dst.Tables("ARTCUST1_STATS").Rows.Add(New Object() {"Sales", .Item("CUST_SALES_MTD"), .Item("CUST_SALES_YTD"), .Item("CUST_SALES_LYR")})
            dst.Tables("ARTCUST1_STATS").Rows.Add(New Object() {"Credits", .Item("CUST_CRED_MTD"), .Item("CUST_CRED_YTD"), .Item("CUST_CRED_LYR")})
            dst.Tables("ARTCUST1_STATS").Rows.Add(New Object() {"Net Sales", Val(.Item("CUST_SALES_MTD") & "") + Val(.Item("CUST_CRED_MTD") & ""), Val(.Item("CUST_SALES_YTD") & "") + Val(.Item("CUST_CRED_YTD") & ""), Val(.Item("CUST_SALES_LYR") & "") + Val(.Item("CUST_CRED_LYR") & "")})
            dst.Tables("ARTCUST1_STATS").Rows.Add(New Object() {"Payments", .Item("CUST_CASH_MTD"), .Item("CUST_CASH_YTD"), .Item("CUST_CASH_LYR")})
            dst.Tables("ARTCUST1_STATS").Rows.Add(New Object() {"#Invoices", .Item("CUST_NUM_INV_MTD"), .Item("CUST_NUM_INV_YTD"), .Item("CUST_NUM_INV_LYR")})
        End With

        ASCMAIN1.sql = "Select " _
        & "  Sum (INV_BALANCE) OPEN" _
        & ", Sum (Decode(INV_TYPE,'I',INV_BALANCE,0)) OPEN_I" _
        & ", Sum (Decode(INV_TYPE,'C',INV_BALANCE,0)) OPEN_C" _
        & ", Sum (Decode(INV_TYPE,'D',INV_BALANCE,0)) OPEN_D" _
        & ", Sum (Decode(INV_TYPE,'O',INV_BALANCE,0)) OPEN_O" _
        & ", Sum (Decode(INV_TYPE,'B',INV_BALANCE,0)) OPEN_B" _
        & ", Sum (Decode(INV_TYPE,'R',INV_BALANCE,0)) OPEN_R" _
        & " from ARTOPEN1 where CUST_CODE = '" & HFs("CUST_CODE") & "'"
        Dim rowOPENAR As DataRow = ASCDATA1.GetDataRow
        TOTAL_DUE = Val(rowOPENAR.Item("OPEN") & "")

        grdARTCUST1_AR_TYPE.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("ARTCUST1_AR_TYPE")
            .Rows.Clear()
            .Rows.Add(New Object() {"Invoices", Val(rowOPENAR.Item("OPEN_I") & "")})
            .Rows.Add(New Object() {"CR Memos", Val(rowOPENAR.Item("OPEN_C") & "")})
            .Rows.Add(New Object() {"DR Memos", Val(rowOPENAR.Item("OPEN_D") & "")})
            .Rows.Add(New Object() {"On/Acct", Val(rowOPENAR.Item("OPEN_O") & "")})
            .Rows.Add(New Object() {"Chargebacks", Val(rowOPENAR.Item("OPEN_B") & "")})
            .Rows.Add(New Object() {"Returns", Val(rowOPENAR.Item("OPEN_R") & "")})
            .Rows.Add(New Object() {"Totals", Val(rowOPENAR.Item("OPEN") & "")})
        End With

        CalculateRemainingBalances()

        grdARTCUST1_FL.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("ARTCUST1_FL")
            .Rows.Clear()
            .Rows.Add(New Object() {"1st Purchase", Null, rowARTCUST6.Item("CUST_FIRST_PURCH"), Null})
            .Rows.Add(New Object() {"Last Invoice", rowARTCUST6.Item("CUST_LAST_INV_NUM"), rowARTCUST6.Item("CUST_LAST_INV_DATE"), rowARTCUST6.Item("CUST_LAST_INV_AMT")})
            .Rows.Add(New Object() {"Last Payment", rowARTCUST6.Item("CUST_LAST_PMT_REF"), rowARTCUST6.Item("CUST_LAST_PMT_DATE"), rowARTCUST6.Item("CUST_LAST_PMT_AMT")})
            .Rows.Add(New Object() {"Bus Established", Null, rowARTCUST1.Item("CUST_BUS_ESTAB"), Null})
        End With

    End Sub

    Function Calc_Date(ByVal F As Integer, ByVal AP_PARM_AGE_CATG As Integer) As String
        Return "'" & Format(Now.AddDays(F * AGE_DAYS(AP_PARM_AGE_CATG)), "MM/dd/yyyy") & "'"
    End Function

    Sub Setup_SOTORDR2()
        If SELECTION_NO = 0 Then Exit Sub
        If grdSOTORDR0.ActiveRow Is Nothing OrElse Not grdSOTORDR0.ActiveRow.IsDataRow Then
            grdSOTORDR2.Visible = False
        Else
            If grdSOTORDR0.ActiveRow.Band.Key = "SOTORDR0" AndAlso Val(grdSOTORDR0.ActiveRow.Cells("ORDR_CNT").Value & "") > 1 Then
                Dim ORDR_GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value
                ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                    & ", MIN (SOTORDR2.ORDR_UNIT_PRICE) ORDR_UNIT_PRICE" & vbCrLf _
                    & " from SOTORDR2,SOTORDR1" & vbCrLf _
                    & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                    & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                    & " group by SOTORDR1.ORDR_GROUP_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
                ASCMAIN1.sql = "Select X.*, ROWNUM ORDR_LNO, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                    & " from (" & ASCMAIN1.sql & ") X, ICTSTYL1, ICTCOLR1" & vbCrLf _
                    & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                    & "   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE"
                grdSOTORDR2.Text = "Details for Order Group No " & ORDR_GROUP_NO
                Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)
            Else
                Dim ORDR_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_NO").Value & ""
                ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
                grdSOTORDR2.Text = "Details for Order No " & ORDR_NO
                Fill_Records("SOTORDR2", ORDR_NO)
            End If
            grdSOTORDR2.Visible = True
        End If
    End Sub

    Sub Setup_SOTINVH2()
        If SELECTION_NO = 0 Then Exit Sub
        If grdSOTINVH1.ActiveRow Is Nothing OrElse Not grdSOTINVH1.ActiveRow.IsDataRow Then
            grdSOTINVH2.Visible = False
        Else
            Dim INV_TYPE As String = grdSOTINVH1.ActiveRow.Cells("INV_TYPE").Value
            Dim INV_NO As String = grdSOTINVH1.ActiveRow.Cells("INV_NO").Value
            grdSOTINVH2.Text = "Details for Invoice No " & INV_NO
            Fill_Records("SOTINVH2", New String() {INV_TYPE, INV_NO})
            grdSOTINVH2.Visible = True
        End If
    End Sub

    Sub Setup_ARTPYMT2()

    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
        Me.SetControlPanel()
    End Sub

    Sub Load_ARTPYMT2(Optional ByVal INV_NO As String = "", Optional ByVal CUST_CODE_P As String = "", _
                      Optional ByVal PYMT_BATCH_NO As String = "")

        Call ASCMAIN1.Progress("Now Loading Payment History")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Dim CUST_CODE_PYMT As String = ""

        grdARTPYMT3.Visible = False
        If CUST_CODE_P <> "" Then
            CUST_CODE_PYMT = CUST_CODE_P
        Else
            CUST_CODE_PYMT = HFs("CUST_CODE")
        End If
        If INV_NO <> "" Then
            ASCMAIN1.sql = "Select ARTPYMT2.*, ARTPYMT1.PYMT_BATCH_DATE" _
            & " from ARTPYMT2,ARTPYMT1 " _
            & " where (ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO) in " _
            & " (Select PYMT_BATCH_NO, PYMT_BATCH_LNO from ARTPYMT3 " _
            & "  where INV_NUM = '" & txtINV_NO_PYMT.Text & "'" & IIf(txtINV_NO_PYMT.Tag <> "", " and INV_TYPE = '" & txtINV_NO_PYMT.Tag & "'", "") & ")" _
            & " and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
            & " and ARTPYMT2.CUST_CODE = '" & CUST_CODE_PYMT & "'"

            Call Fill_Records("ARTPYMT2", "", True, ASCMAIN1.sql)
            grdARTPYMT2.Text = "Payment Applications involving Invoice No " & INV_NO
            chkPaymentsShowZero.Checked = True
        ElseIf PYMT_BATCH_NO <> "" Then
            ASCMAIN1.sql = "Select ARTPYMT2.*, ARTPYMT1.PYMT_BATCH_DATE" _
            & " from ARTPYMT2,ARTPYMT1 " _
            & " where (ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO) in " _
            & " (Select PYMT_BATCH_NO, PYMT_BATCH_LNO from ARTPYMT3 " _
            & "  where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "')" _
            & " and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
            & " and ARTPYMT2.CUST_CODE = '" & CUST_CODE_PYMT & "'"

            Call Fill_Records("ARTPYMT2", "", True, ASCMAIN1.sql)
            grdARTPYMT2.Text = "Payment Applications involving Batch " & PYMT_BATCH_NO
            chkPaymentsShowZero.Checked = True

        Else
            Call Fill_Records("ARTPYMT2", New String() {cbeYP_PYMTs.Value})
            grdARTPYMT2.Text = "Payments Received from Customer " & CUST_CODE_PYMT & " since " & cbeYP_PYMTs.Text
        End If

        Sort_grdColumns(grdARTPYMT2, "PYMT_BATCH_NO,PYMT_BATCH_LNO".ToLower)

        Payments_Show_Zero()

        txtINV_NO_PYMT.Tag = ""

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Load_ARTOPEN1()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Open AR Items")
        Fill_Records("ARTOPEN1")
        For Each rowARTOPEN1 As DataRow In dst.Tables("ARTOPEN1").Select("INV_BALANCE = 0")
            rowARTOPEN1.Item("AGE") = DBNull.Value
        Next
        Sort_grdColumns(grdARTOPEN1, "INV_DATE")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        Age_Open_Items_by_Date()
        Filter_ARTOPEN1()
    End Sub

    Sub Retrieve_Paid_Invoices(ByVal numDays As Integer)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Retrieving Paid Invoices")

        ASCMAIN1.sql = sqlARTOPEN1 _
        & " from ARTOPENX ARTOPEN1," & ARTCUST0 & " ARTCUST0 " _
        & " where ARTOPEN1.CUST_CODE = ARTCUST0.CUST_CODE" _
        & "   and ARTOPEN1.DATE_PAID > TRUNC(SYSDATE) - " & CStr(numDays)
        For Each rowARTOPENX As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").Rows.Find _
                (New Object() {rowARTOPENX.Item("CUST_CODE"), rowARTOPENX.Item("INV_TYPE"), rowARTOPENX.Item("INV_NUM")})
            If rowARTOPEN1 Is Nothing Then
                rowARTOPEN1 = dst.Tables("ARTOPEN1").NewRow
                With rowARTOPEN1
                    For i As Integer = 0 To rowARTOPENX.Table.Columns.Count - 1
                        .Item(i) = rowARTOPENX.Item(i)
                    Next i
                End With
                dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
            End If
        Next

        Sort_grdColumns(grdARTOPEN1, "INV_DATE")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        Age_Open_Items_by_Date()
        Filter_ARTOPEN1()
    End Sub

    Sub Filter_ARTOPEN1()
        Dim tlb_sbt_ZBI As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show $0 Balance Items"), UltraWinToolbars.StateButtonTool)
        Dim filterSql As String = ""
        Dim dvw As DataView = DirectCast(grdARTOPEN1.DataSource, DataView) ' DirectCast(grdARTOPEN1.DataSource, DataTable).DefaultView

        If tlb_sbt_ZBI.Checked Then
            showZBI = True
        Else
            showZBI = False
            filterSql = "INV_BALANCE <> 0"
        End If

        dvw.RowFilter = filterSql
    End Sub

    Sub Print_Hard_Copy()

        Dim RPT_TITLE As String = "Customer Inquiry - Hard Copy"
        Dim reportFile As String = "ARRCINQ1"
        Dim FILENAME As String = ""

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing " & RPT_TITLE)

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        If showZBI Then
            CR_params.Add("OPEN_ONLY", "0")
        Else
            CR_params.Add("OPEN_ONLY", "1")
        End If
        Generate_Report(reportFile, RPT_TITLE)
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Show_Document(FILENAME)

    End Sub

    Sub Load_SOTORDR1()

        Me.Cursor = Cursors.WaitCursor
        Dim CUST_CODE As String = HFs("CUST_CODE")
        Dim CUST_NAME As String = HFs("CUST_NAME")
        splSOTORDR0.Panel2Collapsed = False

        dst.Tables("SOTORDR2").Rows.Clear()
        dst.Tables("SOTORDR1").Rows.Clear()
        dst.Tables("SOTORDR0").Rows.Clear()

        If chkOpenOrdersOnly.Checked Then
            ASCMAIN1.Progress("Now Loading Open Order Data")
            grdSOTORDR0.Text = "Open Orders for " & CUST_CODE & ":" & CUST_NAME & " (by Order)"

            ASCMAIN1.sql = "Select SOTORDR0.* from SOTORDR0 where CUST_CODE = '" & CUST_CODE & "' and (ORDR_CNT_OPEN <> 0 or ORDR_CNT_PICK <> 0)"
            Fill_Records("SOTORDR0", "", , ASCMAIN1.sql)

        Else
            ASCMAIN1.Progress("Now Loading Sales Order History")
            grdSOTORDR0.Text = "Order History for " & CUST_CODE & ":" & CUST_NAME & " since " & Format(dteOrderHistory.DateTime, "MM/dd/yyyy")

            ASCMAIN1.sql = "Select SOTORDR0.* from SOTORDR0 where CUST_CODE = '" & CUST_CODE & "' and ORDR_DATE > '" & Format(dteOrderHistory.DateTime, "dd-MMM-yyyy") & "'"
            Fill_Records("SOTORDR0", "", , ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select SOTORDR1.* from SOTORDR1 where ORDR_GROUP_NO in (" & Replace(ASCMAIN1.sql, ".*", ".ORDR_GROUP_NO") & ")"
        Fill_Records("SOTORDR1", "", , ASCMAIN1.sql)
        Sort_grdColumns(grdSOTORDR0, "ORDR_GROUP_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Setup_SOTORDR2()
    End Sub

    Sub Load_SOTINVH1()

        Me.Cursor = Cursors.WaitCursor
        Dim CUST_CODE As String = HFs("CUST_CODE")
        Dim CUST_NAME As String = HFs("CUST_NAME")
        splSOTINVH1.Panel2Collapsed = False

        ASCMAIN1.Progress("Now Loading Sales Invoice History")
        grdSOTORDR0.Text = "Invoice History for " & CUST_CODE & ":" & CUST_NAME & " since " & Format(dteInvoiceHistory.DateTime, "MM/dd/yyyy")

        ASCMAIN1.sql = "Select SOTINVH1.* from SOTINVH1 where CUST_CODE = '" & CUST_CODE & "' and INV_DATE > '" & Format(dteInvoiceHistory.DateTime, "dd-MMM-yyyy") & "'"
        Fill_Records("SOTINVH1", "", , ASCMAIN1.sql)

        Sort_grdColumns(grdSOTINVH1, "INV_DATE".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Setup_SOTINVH2()
    End Sub

    Sub Calc_Total_Wgt()
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Rows

            Dim SO_ORDER_NO As String = rowSOTORDR1.Item("SO_ORDER_NO")

            ASCMAIN1.sql = "SELECT SUM(TOTAL_WGT) TOTAL_WGT FROM " _
            & "(SELECT SUM(QTY_UNITS) TOTAL_WGT FROM SOTORDR2 WHERE SO_ORDER_NO = '" & SO_ORDER_NO & "'" _
            & " UNION " _
            & " SELECT SUM(QTY_UNITS) TOTAL_WGT FROM SOTINVH2 WHERE SO_ORDER_NO = '" & SO_ORDER_NO & "')"
            Dim totWgt As Integer = Val(ASCDATA1.GetDataValue)
            rowSOTORDR1.Item("TOTAL_WGT") = totWgt
        Next
    End Sub

    Sub Count_BFS()
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Rows

            Dim SO_ORDER_NO As String = rowSOTORDR1.Item("SO_ORDER_NO")

            ASCMAIN1.sql = "SELECT SUM(BFS) BFS FROM " _
            & " (SELECT COUNT(*) BFS FROM SOTORDR2 WHERE CUST_BOUGHT_FOR IS NOT NULL AND SO_ORDER_NO = '" & SO_ORDER_NO & "' " _
            & " UNION " _
            & " SELECT COUNT(*) BFS FROM SOTINVH2 WHERE CUST_BOUGHT_FOR IS NOT NULL AND SO_ORDER_NO = '" & SO_ORDER_NO & "') "
            Dim bfCount As Integer = Val(ASCDATA1.GetDataValue)
            rowSOTORDR1.Item("BFS") = bfCount
        Next
    End Sub

    Private Sub grdARTCUST1_AR_TYPE_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUST1_AR_TYPE.InitializeRow
        If e.Row.Cells(0).Text Like "Total*" Then
            e.Row.Appearance.BackColor = Drawing.Color.Beige
            e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
        End If
    End Sub

    Private Sub grdARTCUST1_OPEN_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUST1_OPEN.InitializeRow
        If e.Row.Cells(0).Text Like "Total*" Then
            e.Row.Appearance.BackColor = Drawing.Color.Beige
            e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
        End If
    End Sub

    Sub Age_Open_Items_by_Date()
        If SELECTION_NO = 0 Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Calculating Aging")

        Dim DATE_COLUMN As String

        'With dst.Tables("ARTCUST1_AGING")
        '.Rows.Clear()

        ReDim DUE_TOTALS(4)
        Dim DUE_AMT As Decimal = 0
        Dim DUE_WHERE As String
        DATE_COLUMN = "INV_DUE_DATE"

        For i As Integer = 1 To 4
            If i = 1 Then
                DUE_WHERE = DATE_COLUMN & " >= " & DUE_DATE(i + 1)
            ElseIf i = 4 Then
                DUE_WHERE = DATE_COLUMN & " < " & DUE_DATE(i)
            Else
                DUE_WHERE = DATE_COLUMN & " >= " & DUE_DATE(i + 1) & " and " & DATE_COLUMN & " < " & DUE_DATE(i)
            End If
            DUE_AMT = Val(dst.Tables("ARTOPEN1").Compute("SUM (INV_BALANCE)", DUE_WHERE) & "")
            DUE_TOTALS(i) = DUE_AMT
        Next

        Dim T As Decimal = 0

        ReDim AGED_TOTALS(4)
        Dim AGE_AMT As Decimal = 0
        Dim AGE_WHERE As String
        DATE_COLUMN = "INV_DATE"

        For i As Integer = 1 To 4
            If i = 1 Then
                AGE_WHERE = DATE_COLUMN & " >= " & AGE_DATE(i + 1)
            ElseIf i = 4 Then
                AGE_WHERE = DATE_COLUMN & " < " & AGE_DATE(i)
            Else
                AGE_WHERE = DATE_COLUMN & " >= " & AGE_DATE(i + 1) & " and " & DATE_COLUMN & " < " & AGE_DATE(i)
            End If
            AGE_AMT = Val(dst.Tables("ARTOPEN1").Compute("SUM (INV_BALANCE)", AGE_WHERE) & "")
            T = T + AGE_AMT

            AGED_TOTALS(i) = AGE_AMT

            'Dim AGE_CATGY As String = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_" & CStr(i))
            '.Rows.Add(New Object() {AGE_CATGY, AGE_AMT})
        Next
        AGED_TOTALS(0) = T
        '.Rows.Add(New Object() {"Total", T})
        'End With

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub

    Sub SetControlPanel()
        UltraExplorerBar1.Groups("Payment History").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Accts Rec") And (UltraTabControl5.SelectedTab.Key = "Payment History")
        UltraExplorerBar1.Groups("Order History").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Orders") And (tabOrders.SelectedTab.Key = "Order History")
        UltraExplorerBar1.Groups("Invoice History").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Orders") And (tabOrders.SelectedTab.Key = "Invoice History")
        UltraExplorerBar1.Groups("Summary by Style").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Orders") And (tabOrders.SelectedTab.Key = "Summary by Style")
        UltraExplorerBar1.Groups("Sales").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Sales")
        UltraExplorerBar1.Groups("Open Balances").Visible = ScreenMode And (tabMain.SelectedTab.Key = "General" Or tabMain.SelectedTab.Key = "Info" Or (tabMain.SelectedTab.Key = "Accts Rec" And UltraTabControl5.SelectedTab.Key = "Open AR && Aging History"))
        UltraExplorerBar1.Groups("Open AR by Type").Visible = ScreenMode And (tabMain.SelectedTab.Key = "General" Or tabMain.SelectedTab.Key = "Info" Or (tabMain.SelectedTab.Key = "Accts Rec" And UltraTabControl5.SelectedTab.Key = "Open AR && Aging History"))
        'UltraExplorerBar1.Groups("Aged Open AR").Visible = ScreenMode And (UltraTabControl1.SelectedTab.Key = "General" Or UltraTabControl1.SelectedTab.Key = "Info" Or (UltraTabControl1.SelectedTab.Key = "Accts Rec" And UltraTabControl5.SelectedTab.Key = "Open AR && Aging History"))
        UltraExplorerBar1.Groups("Freight").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Freight")
    End Sub

    Sub Setup_tabMain()

        UltraExplorerBar1.Groups("Customer Log").Visible = _
        (tabMain.SelectedTab.Key = "Log")

        Select Case tabMain.SelectedTab.Key
            Case "Name && Address"
            Case "Info"
            Case "Accts Rec"
                Call Setup_Tab_AR()
            Case "Orders"
                Setup_tabOrders()
            Case "Sales"
                If grdARTSLSMA.Tag = "" Then
                    Fill_Records("ARTSLSMA", HFs("CUST_CODE"))
                    Sort_grdColumns(grdARTSLSMA, "OPS_YYYYPP")
                    grdARTSLSMA.Tag = HFs("CUST_CODE")
                End If
                If grdSATCUSTS.Tag = "*" Then
                    SetUp12Month()
                End If
        End Select

    End Sub

    Private Sub grdARTPYMT2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT2.AfterRowActivate
        Load_ARTPYMT3()
    End Sub

    Sub Load_ARTPYMT3()

        With grdARTPYMT2
            If .ActiveRow Is Nothing OrElse .ActiveRow.IsGroupByRow Then
                grdARTPYMT3.Visible = False
            Else
                grdARTPYMT3.Visible = True

                Dim PYMT_BATCH_NO As String = .ActiveRow.Cells("PYMT_BATCH_NO").Text
                Dim PYMT_BATCH_LNO As Integer = .ActiveRow.Cells("PYMT_BATCH_LNO").Text
                grdARTPYMT3.Text = "Payment Details for Batch-Lno " & PYMT_BATCH_NO & "-" & CStr(PYMT_BATCH_LNO) & "; " & .ActiveRow.Cells("CUST_CODE").Text & ":" & .ActiveRow.Cells("CUST_NAME").Text
                Call Fill_Records("ARTPYMT3", New Object() {PYMT_BATCH_NO, PYMT_BATCH_LNO})
                Call Sort_grdColumns(grdARTPYMT3, "PYMT_BATCH_NO,PYMT_BATCH_LNO,PYMT_BATCH_ILNO")
                grdARTPYMT3.DisplayLayout.Bands(0).SummaryFooterCaption = "Totals for Payment Ref " & .ActiveRow.Cells("CUST_PYMT_REF_NO").Text
            End If
        End With
    End Sub

    Private Sub chkPaymentsShowZero_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPaymentsShowZero.CheckedChanged
        Payments_Show_Zero()
    End Sub

    Sub Payments_Show_Zero()
        If chkPaymentsShowZero.Checked Then
            DirectCast(grdARTPYMT2.DataSource, DataTable).DefaultView.RowFilter = ""
        Else
            DirectCast(grdARTPYMT2.DataSource, DataTable).DefaultView.RowFilter = "CUST_PYMT_AMT <> 0"
        End If
        If grdARTPYMT2.Rows.Count <> 0 Then
            grdARTPYMT2.ActiveRow = grdARTPYMT2.Rows(0)
        Else
            grdARTPYMT3.Visible = False
        End If
    End Sub

    Sub Customer_Activity()
        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Loading List of Active Customers")
        Call Fill_Records("ARTCUST6")
        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        grdARTCUST6.Tag = "*"
        grdARTCUST6.Visible = True
        grdARTCUSTT_FUPS.Visible = False
    End Sub

    Private Sub btnCreditUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreditUpdate.Click

        If chkEditCredit.Checked Then
            If Val(Absx1.numFor("CUST_CREDIT_LIMIT").Value & "") <> 0 Then
                If Absx1.dteFor("CUST_CRED_LIMIT_EST").Value & "" = "" Then
                    MsgBox("Date Credit Limit Established is Required", MsgBoxStyle.OkOnly, "Cannot Update Credit Info")
                    Exit Sub
                End If
            End If

            BeginTrans()

            Dim X As CurrencyManager = Me.BindingContext(dst.Tables("ARTCUST1"))
            X.EndCurrentEdit()
            rowARTCUST1 = dst.Tables("ARTCUST1").Rows(0)

            Call Write_Audit_Trail(rowARTCUST1, rowARTCUST1_orig, "E")

            Dim rowARTCUST5 As DataRow = dst.Tables("ARTCUST5").NewRow()
            With rowARTCUST1

                dst.Tables("ARTCUST1").AcceptChanges()

                ASCMAIN1.sql = "Update ARTCUST1 Set" _
                & "  CUST_CREDIT_LIMIT = :PARM1" _
                & ", CUST_CRED_LIMIT_EST = :PARM2" _
                & ", CUST_CRED_LIMIT_REV = :PARM3" _
                & ", CUST_CREDIT_HOLD = :PARM4" _
                & ", CUST_FACTOR_IND = :PARM5" _
                & ", CUST_CREDIT_SCORE = :PARM6" _
                & ", CUST_CREDIT_SCORE_DATE = :PARM7" _
                & ", TERM_CODE = :PARM8" _
                & ", CUST_CREDIT_LIMIT_APPR_BY = :PARM9" _
                & ", CUST_CREDIT_LIMIT_NOTES = :PARM10" _
                & ", CUST_CREDIT_RATING = :PARM11" _
                & ", CUST_CREDIT_RATING_DATE = :PARM12" _
                & ", CUST_INS_AMT = :PARM13" _
                & ", CUST_INS_DATE = :PARM14" _
                & ", CUST_DUNS = :PARM15" _
                & ", CUST_PD_GRACE_DAYS = :PARM16" _
                & ", CUST_CREDIT_RELEASE = :PARM17" _
                & " where CUST_CODE = :PARM18"

                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NDDVVVDVVVVDNDVNVV", New Object() _
                {.Item("CUST_CREDIT_LIMIT"), _
                .Item("CUST_CRED_LIMIT_EST"), _
                .Item("CUST_CRED_LIMIT_REV"), _
                .Item("CUST_CREDIT_HOLD"), _
                .Item("CUST_FACTOR_IND"), _
                .Item("CUST_CREDIT_SCORE"), _
                .Item("CUST_CREDIT_SCORE_DATE"), _
                .Item("TERM_CODE"), _
                .Item("CUST_CREDIT_LIMIT_APPR_BY"), _
                .Item("CUST_CREDIT_LIMIT_NOTES"), _
                .Item("CUST_CREDIT_RATING"), _
                .Item("CUST_CREDIT_RATING_DATE"), _
                .Item("CUST_INS_AMT"), _
                .Item("CUST_INS_DATE"), _
                .Item("CUST_DUNS"), _
                .Item("CUST_PD_GRACE_DAYS"), _
                .Item("CUST_CREDIT_RELEASE"), _
                HFs("CUST_CODE")})

                rowARTCUST5.Item("CUST_CODE") = .Item("CUST_CODE")
                rowARTCUST5.Item("CUST_CREDIT_LIMIT") = .Item("CUST_CREDIT_LIMIT")
                rowARTCUST5.Item("CUST_CRED_LIMIT_EST") = .Item("CUST_CRED_LIMIT_EST")
                rowARTCUST5.Item("CUST_CRED_LIMIT_REV") = .Item("CUST_CRED_LIMIT_REV")
                rowARTCUST5.Item("CUST_CREDIT_HOLD") = .Item("CUST_CREDIT_HOLD")
                rowARTCUST5.Item("CUST_FACTOR_IND") = .Item("CUST_FACTOR_IND")
                rowARTCUST5.Item("CUST_CREDIT_SCORE") = .Item("CUST_CREDIT_SCORE")
                rowARTCUST5.Item("CUST_CREDIT_SCORE_DATE") = .Item("CUST_CREDIT_SCORE_DATE")
                rowARTCUST5.Item("TERM_CODE") = .Item("TERM_CODE")
                rowARTCUST5.Item("CUST_CREDIT_LIMIT_APPR_BY") = .Item("CUST_CREDIT_LIMIT_APPR_BY")
                rowARTCUST5.Item("CUST_CREDIT_LIMIT_NOTES") = .Item("CUST_CREDIT_LIMIT_NOTES")
                rowARTCUST5.Item("CUST_CREDIT_RATING") = .Item("CUST_CREDIT_RATING")
                rowARTCUST5.Item("CUST_CREDIT_RATING_DATE") = .Item("CUST_CREDIT_RATING_DATE")
                rowARTCUST5.Item("CUST_INS_AMT") = .Item("CUST_INS_AMT")
                rowARTCUST5.Item("CUST_INS_DATE") = .Item("CUST_INS_DATE")
                rowARTCUST5.Item("CUST_CREDIT_RELEASE") = .Item("CUST_CREDIT_RELEASE")
            End With

            rowARTCUST5.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowARTCUST5.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD

            dst.Tables("ARTCUST5").Rows.Add(rowARTCUST5)

            Sort_grdColumns(grdARTCUST5, "INIT_DATE".ToLower)

            Update_Record_TDA("ARTCUST5")
            ARCMAIN1.Record_Customer_Event(HFs("CUST_CODE"), "CR Limit Changed to " & Format$(Val(rowARTCUST1.Item("CUST_CREDIT_LIMIT") & ""), "$###,##0"), "C")

            CommitTrans("Credit Info Updated")

            chkEditCredit.Checked = False
            CalculateRemainingBalances()
        End If
    End Sub

    Private Sub btnCreditCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreditCancel.Click

        dst.Tables("ARTCUST1").Rows(0).EndEdit()
        dst.Tables("ARTCUST1").RejectChanges()

        chkEditCredit.Checked = False

    End Sub

    Private Sub chkEditCredit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEditCredit.CheckedChanged
        If chkEditCredit.Checked Then
            If Not ASCMAIN1.Logical_Lock("ARTCUST1", HFs("CUST_CODE")) Then
                chkEditCredit.Checked = False
                Exit Sub
            Else
                Load_Customer_and_Credit_History()
            End If
        Else
            ASCMAIN1.MultiTask_Release()
        End If
        ToggleCreditEditable()
    End Sub

    Sub ToggleCreditEditable()
        If chkEditCredit.Checked Then
            Set_Read_Only(grpARCredit, False)
            btnCreditUpdate.Enabled = True
            btnCreditCancel.Enabled = True
            rowARTCUST1_orig = dst.Tables("ARTCUST1").NewRow
            rowARTCUST1_orig.ItemArray = dst.Tables("ARTCUST1").Rows(0).ItemArray
        Else
            Set_Read_Only(grpARCredit, True)
            Set_Read_Only_for_ctl(chkEditCredit, False)
            btnCreditUpdate.Enabled = False
            btnCreditCancel.Enabled = False

        End If
    End Sub

    Sub CalculateRemainingBalances()
        grdARTCUST1_AR_TYPE.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("ARTCUST1_OPEN")
            .Rows.Clear()
            Dim CUST_CREDIT_LIMIT As Decimal = Val(Absx1.numFor("CUST_CREDIT_LIMIT").Value & "") ' Val(rowARTCUST1.Item("CUST_CREDIT_LIMIT") & "")

            .Rows.Add(New Object() {"CR Limit", CUST_CREDIT_LIMIT})
            .Rows.Add(New Object() {"Open AR", TOTAL_DUE})

            'ASCMAIN1.sql = "Select DECODE(NVL(SOTORDR1.ORDR_CREDIT_STATUS,'X'),'A','A','X') CRSTAT" _
            '& ", Sum (NVL(SOTORDR1.ORDR_AMT,0)) AMT from SOTORDR1" _
            '& "  where SOTORDR1.SO_STATUS_CODE IN ('O','H')" _
            '& "    and (SOTORDR1.CUST_CODE = '" & HFs("CUST_CODE") & "'" _
            '& "     or  SOTORDR1.CUST_BILL_TO_CUST IN " _
            '& " (Select CUST_CODE from ARTCUST1 where CUST_CREDIT_GROUP_CUST = '" & HFs("CUST_CODE") & "'))" _
            '& " group by DECODE(NVL(SOTORDR1.ORDR_CREDIT_STATUS,'X'),'A','A','X')"

            Dim OPEN_ORDERS_RELEASED As Decimal = 0 ' Val(ASCDATA1.GetDataValue)
            Dim OPEN_ORDERS_PENDING As Decimal = 0 ' Val(ASCDATA1.GetDataValue)

            'For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            '    Dim CRSTAT As String = row.Item("CRSTAT") & ""
            '    If CRSTAT = "X" Then
            '        OPEN_ORDERS_PENDING = Val(row.Item("AMT") & "")
            '    Else
            '        OPEN_ORDERS_RELEASED = Val(row.Item("AMT") & "")
            '    End If
            'Next

            .Rows.Add(New Object() {"Ords Reld", OPEN_ORDERS_RELEASED})
            .Rows.Add(New Object() {"Ords Pend", OPEN_ORDERS_PENDING})
            Dim CREDIT_AVAIL As Decimal = CUST_CREDIT_LIMIT - TOTAL_DUE - OPEN_ORDERS_RELEASED
            If CREDIT_AVAIL < 0 Then
                CREDIT_AVAIL = 0
            End If

            Dim CRP As String = ""
            If CUST_CREDIT_LIMIT <> 0 Then
                CRP = Format(100 * CREDIT_AVAIL / CUST_CREDIT_LIMIT, "##0") & "%"
            End If
            .Rows.Add(New Object() {"CR Avail " & CRP, CREDIT_AVAIL})
        End With
    End Sub

    Private Sub grdARTCUSTT_FUPS_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTCUSTT_FUPS.DoubleClickRow
        If grdARTCUSTT_FUPS.ActiveRow IsNot Nothing AndAlso grdARTCUSTT_FUPS.ActiveRow.IsDataRow Then
            Absx1.txtFor("CUST_CODE").Text = grdARTCUSTT_FUPS.ActiveRow.Cells("TABLE_KEY").Text
            Click_Command("Select Customer")
        End If
    End Sub

    Sub Refresh_FollowUps()
        Fill_Records("ARTCUSTT_FUPS")
        grdARTCUSTT_FUPS.Visible = True
        grdARTCUST6.Visible = False
        grdARTCUST6.Tag = ""

        grdARTCUSTT_FUPS.Text = "My Follow Ups (Action Required, or Completed Today)"
    End Sub

    Private Sub grdARTCUSTT_FUPS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUSTT_FUPS.InitializeRow
        If e.Row.Cells("CONV_FOLLOWUP_BY").Text <> "" Then
            e.Row.Cells("CONV_FOLLOWUP_BY").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.BackColor = Drawing.Color.Yellow
            If e.Row.Cells("CONV_FOLLOWUP_BY").Text = ASCMAIN1.USER_ID Then
                e.Row.Cells("CONV_FOLLOWUP_BY").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.ForeColor = Drawing.Color.Red
            End If
        Else
            e.Row.Cells("CONV_FOLLOWUP_BY").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("CONV_FOLLOWUP_BY").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.ForeColor = Drawing.Color.Empty
        End If

        With e.Row.Cells("CONV_STATUS")
            If .Value & "" = "1" Then
                .Appearance.BackColor = Drawing.Color.LightGreen
            ElseIf .Value & "" = "2" Then
                .Appearance.BackColor = Drawing.Color.PeachPuff
            Else
                .Appearance.BackColor = Drawing.Color.Empty
            End If
        End With

    End Sub

    'Private Sub grdARTCUSTT_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUSTT.InitializeRow
    '    If e.Row.Cells("CONV_FOLLOWUP_WITH").Text <> "" And e.Row.Cells("CONV_FOLLOWUP_ACK_BY").Text = "" Then
    '        e.Row.Cells("CONV_FOLLOWUP_WITH").Appearance.BackColor = Drawing.Color.Yellow
    '        e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.BackColor = Drawing.Color.Yellow
    '        If e.Row.Cells("CONV_FOLLOWUP_WITH").Text = ASCMAIN1.USER_ID Then
    '            e.Row.Cells("CONV_FOLLOWUP_WITH").Appearance.ForeColor = Drawing.Color.Red
    '            e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.ForeColor = Drawing.Color.Red
    '        End If
    '    Else
    '        e.Row.Cells("CONV_FOLLOWUP_WITH").Appearance.BackColor = Drawing.Color.Empty
    '        e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.BackColor = Drawing.Color.Empty
    '        e.Row.Cells("CONV_FOLLOWUP_WITH").Appearance.ForeColor = Drawing.Color.Empty
    '        e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.ForeColor = Drawing.Color.Empty
    '    End If
    'End Sub

    Private Sub UltraTabControl5_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl5.SelectedTabChanged
        Call Setup_Tab_AR()
    End Sub

    Sub Setup_Tab_AR()
        If SELECTION_NO = 0 Then Exit Sub
        If tabMain.SelectedTab.Key = "Accts Rec" Then
            If UltraTabControl5.SelectedTab.Key = "Payment History" Then
                If grdARTPYMT2.Tag = "*" Then
                    Load_ARTPYMT2()
                    grdARTPYMT2.Tag = ""
                End If
            End If
        End If
        SetControlPanel()
    End Sub

    Private Sub cmdOrderHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOrderHistory.Click
        grdSOTORDR0.Tag = "*"
        Setup_tabOrders()
    End Sub

    Private Sub grdSOTORDR1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTORDR0.AfterRowActivate
        Setup_SOTORDR2()
    End Sub

    Private Sub grdARTPYMT2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMT2.InitializeRow
        If e.Row.Cells("CCPA_NO_CREDIT").Text <> "" Then
            e.Row.Cells("STATUS").Value = "CC Credited"
        ElseIf e.Row.Cells("PYMT_STATUS").Text <> "2" Then
            e.Row.Cells("STATUS").Value = "Un-Applied"
            e.Row.Cells("STATUS").Appearance.BackColor = Drawing.Color.Yellow
        ElseIf e.Row.Cells("PYMT_DELETED").Value & "" = "1" Then
            e.Row.Cells("STATUS").Value = "Payment Deleted"
            e.Row.Cells("STATUS").Appearance.ForeColor = Drawing.Color.Red
        End If

        Dim PYMT_REVERSED As String = e.Row.Cells("PYMT_REVERSED").Text
        If PYMT_REVERSED = "1" Then
            e.Row.Cells("PYMT_BATCH_NO").Appearance.ForeColor = Drawing.Color.Red
        ElseIf PYMT_REVERSED = "2" Then
            e.Row.Cells("PYMT_BATCH_NO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("CUST_PYMT_AMT_CURR").Appearance.ForeColor = Drawing.Color.Red
        End If

    End Sub

    Private Sub UltraTabControl9_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl9.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
    End Sub

    Private Sub chkOpenOrdersOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOpenOrdersOnly.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTORDR1()
    End Sub

    Private Sub chkOrdersOnHold_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOrdersOnHold.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Orders_Filter()
    End Sub

    Sub Orders_Filter()
        Dim dvw As DataView
        Dim sql As String = ""
        splSOTORDR0.Panel2Collapsed = False
        If chkOrdersOnHold.Checked Then
            sql &= " and (ISNULL(ORDR_HOLD,'0') = '1')"
        End If

        dvw = DirectCast(grdSOTORDR0.DataSource, DataTable).DefaultView
        dvw.RowFilter = Mid(sql, 5)
        Setup_SOTORDR2()

    End Sub

    Private Sub grdSOTORDR1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR0.InitializeRow

        If e.Row.Band.Key = "SOTORDR0" Then
            'If e.Row.Cells("ORDR_HOLD_SALES").Value & "" = "1" Then
            '    e.Row.Cells("ORDR_HOLD_SALES").Appearance.BackColor = Drawing.Color.Yellow
            'End If
            'If e.Row.Cells("ORDR_HOLD_CREDIT").Value & "" = "1" Then
            '    e.Row.Cells("ORDR_HOLD_CREDIT").Appearance.BackColor = Drawing.Color.Red
            'End If
            If Val(e.Row.Cells("ORDR_CNT_OPEN").Value & "") > 0 Then
                e.Row.Cells("ORDR_QTY_OPEN").Appearance.BackColor = Drawing.Color.LightGreen
                e.Row.Cells("ORDR_AMT_OPEN").Appearance.BackColor = Drawing.Color.LightGreen
            End If
            If Val(e.Row.Cells("ORDR_CNT_PICK").Value & "") > 0 Then
                e.Row.Cells("ORDR_QTY_PICK").Appearance.BackColor = Drawing.Color.LightBlue
                e.Row.Cells("ORDR_AMT_PICK").Appearance.BackColor = Drawing.Color.LightBlue
            End If
        Else

        End If

    End Sub

    Private Sub cmdFetchSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetchSales.Click
        Me.SetUp12Month()
    End Sub

    Private Sub SetUp12Month(Optional ByVal initialize As Boolean = False)

        ASCMAIN1.sql = "Select ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
            & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH2.INV_TYPE" & vbCrLf

        Dim RYP As String = cbeSales.Value
        If initialize Then RYP = ASCMAIN1.CYP
        Dim RYPs(12) As String
        Dim RYP_legends(12) As String

        For I As Integer = 1 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(RYP, -12 + I)
            RYP_legends(I) = Mid(ASCMAIN1.Get_Legend(YP), 10, 6)
            RYPs(I) = YP
            Dim sqlw As String = ", SUM(DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',"
            ASCMAIN1.sql &= sqlw & "SOTINVH2.ORDR_QTY_SHIP,0)) QTY" & Format(I, "00") & vbCrLf
            ASCMAIN1.sql &= sqlw & "SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE,0)) SLS" & Format(I, "00") & vbCrLf
            ASCMAIN1.sql &= sqlw & "SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST,0)) CGS" & Format(I, "00") & vbCrLf
        Next
        ASCMAIN1.sql &= " from SOTINVH2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE"

        If initialize Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            ASCMAIN1.sql &= " group by ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
            & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH2.INV_TYPE"
            SATCUSTS = ASCMAIN1.Temp_Table
        Else
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Sales History")

            ASCMAIN1.sql &= " and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & RYPs(1) & "'"
            ASCMAIN1.sql &= " and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & RYPs(12) & "'"
            ASCMAIN1.sql &= " and SOTINVH2.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
            ASCMAIN1.sql &= IIf(optSalesIC.Value = "A", "", " and SOTINVH2.INV_TYPE = '" & optSalesIC.Value & "'")
            ASCMAIN1.sql &= " group by ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
            & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH2.INV_TYPE"

            ASCDATA1.ExecuteSQL("Truncate Table " & SATCUSTS)
            ASCDATA1.ExecuteSQL("Insert into " & SATCUSTS & " " & ASCMAIN1.sql)

            For i As Integer = 1 To 12
                grdSATCUSTS.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Header.Caption = RYP_legends(i)
                grdSATCUSTI.DisplayLayout.Bands(0).Columns("V" & Format(i, "00")).Header.Caption = RYP_legends(i)
            Next

            Fill_Records("SATCUSTS")

            Sort_grdColumns(grdSATCUSTS, "SALES_DIVISION_CODE,STYLE_CLASS_CODE")

            'Sort_grdColumns(grdSATCUSTS, "STYLE_CLASS_CODE")
            'With grdSATCUSTS.DisplayLayout.Bands(0)
            '    .SortedColumns.Add("SALES_DIVISION_CODE", False, True)
            'End With
            'grdSATCUSTS.Rows.ExpandAll(True)

            splSATCUSTS.Visible = True

            grdSATCUSTS.Text = "12 Months Sales Summary (" & optSales.Text & IIf(optSalesIC.Value = "A", "", ", " & optSalesIC.Text & " only") & ") for " & Absx1.txtFor("CUST_CODE").Text & ":" & Absx1.txtFor("CUST_NAME").Text
            grdSATCUSTS.Visible = True

            splSATCUSTS.Panel2Collapsed = True
            Setup_SATCUSTS()

            grdSATCUSTS.Tag = ""

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub

    Private Sub cmdFetchPayments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetchPayments.Click
        Load_ARTPYMT2()
    End Sub

    Private Sub grdSATCUSTS_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATCUSTS.AfterRowActivate
        Setup_SATCUSTS()
    End Sub

    Sub Setup_SATCUSTS()

        If SELECTION_NO = 0 Then Exit Sub
        If grdSATCUSTS.ActiveRow Is Nothing OrElse Not grdSATCUSTS.ActiveRow.IsDataRow Then
            grdSATCUSTI.Visible = False
        Else
            ' ASCMAIN1.sql = sqlSATCUSTI
            Dim SALES_DIVISION_CODE As String = grdSATCUSTS.ActiveRow.Cells("SALES_DIVISION_CODE").Value & ""
            Dim STYLE_CLASS_CODE As String = grdSATCUSTS.ActiveRow.Cells("STYLE_CLASS_CODE").Value & ""
            Fill_Records("SATCUSTI", New String() {SALES_DIVISION_CODE, STYLE_CLASS_CODE})
            grdSATCUSTI.Text = "12 Months Sales Summary (" & optSales.Text & IIf(optSalesIC.Value = "A", "", ", " & optSalesIC.Text & " only") & ") for " & Absx1.txtFor("CUST_CODE").Text & ", Division " & SALES_DIVISION_CODE & ", Class " & STYLE_CLASS_CODE
            grdSATCUSTI.Visible = True
        End If

        splSATCUSTS.Panel2Collapsed = False
    End Sub

    Private Sub cmdFindPymtAppl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFindPymtAppl.Click
        If txtINV_NO_PYMT.Text <> "" Then
            If txtINV_NO_PYMT.Text.Length <> 10 Then
                txtINV_NO_PYMT.Text = txtINV_NO_PYMT.Text.PadLeft(10, "0")
            End If

            Process_INV_NO_PYMT()
        End If
    End Sub

    Sub Process_INV_NO_PYMT(Optional ByVal CUST_CODE_PYMT As String = "")

        'If txtINV_NO_PYMT.Text.Length < 10 Then
        '    txtINV_NO_PYMT.Text = txtINV_NO_PYMT.Text.PadLeft(10, "0")
        'End If

        If txtINV_NO_PYMT.Tag = "" Or txtINV_NO_PYMT.Tag = "I" Or txtINV_NO_PYMT.Tag = "C" Then
            ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_INV_NO = '" & txtINV_NO_PYMT.Text & "'"
            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow
            If rowSOTINVH1 Is Nothing Then
                MsgBox("No Record of Invoice/Memo " & txtINV_NO_PYMT.Text, MsgBoxStyle.OkOnly, "Cannot Find Record")
                Exit Sub
            End If

            If rowSOTINVH1.Item("CUST_CODE") & "" <> HFs("CUST_CODE") Then
                If rowSOTINVH1.Item("CUST_CODE") & "" <> CUST_CODE_PYMT Then
                    MsgBox("Invoice " & txtINV_NO_PYMT.Text & " belongs to another customer (" & rowSOTINVH1.Item("CUST_CODE") & ")", MsgBoxStyle.OkOnly, "Wrong Customer")
                    txtINV_NO_PYMT.Text = ""
                    Exit Sub
                End If
            End If
        End If
        CUST_CODE_PYMT = HFs("CUST_CODE")
        Load_ARTPYMT2(txtINV_NO_PYMT.Text, CUST_CODE_PYMT)
        Load_ARTPYMT3()
        tabMain.SelectedTab = tabMain.Tabs("Accts Rec")
        UltraTabControl5.SelectedTab = UltraTabControl5.Tabs("Payment History")

    End Sub

    Private Sub grdARTPYMT3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMT3.InitializeRow
        Dim INV_NO As String = e.Row.Cells("INV_NUM").Text
        If INV_NO = txtINV_NO_PYMT.Text Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        End If
    End Sub

    Private Sub txtINV_NO_PYMT_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtINV_NO_PYMT.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            txtINV_NO_PYMT.Text = UCase(txtINV_NO_PYMT.Text)
            Process_INV_NO_PYMT()
        End If
    End Sub

    Private Sub cmdPrintSelected_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintSelected.Click

        If grdARTPYMT3.Selected.Rows.Count = 0 Then
            MessageBox.Show("There are no AR Batch detail Invoices selected.", "Print Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If MessageBox.Show("Do you want to print the " & grdARTPYMT3.Selected.Rows.Count & " selected documents?", _
                            "Print Selected", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim invoices As String = String.Empty
        Dim INV_NO As String = String.Empty
        Dim rowSOTINVH1 As DataRow = Nothing

        For Each gridRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdARTPYMT3.Selected.Rows
            INV_NO = (gridRow.Cells("INV_NUM").Value & String.Empty).ToString.Trim

            rowSOTINVH1 = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_NO = '" & INV_NO & "' AND CUST_CODE = '" & MyBase.Absx1.txtFor("CUST_CODE").Text & "'")
            If rowSOTINVH1 Is Nothing Then
                rowSOTINVH1 = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_REF = '" & INV_NO & "' AND CUST_CODE = '" & MyBase.Absx1.txtFor("CUST_CODE").Text & "'")
            End If

            If rowSOTINVH1 IsNot Nothing Then
                invoices &= ", '" & rowSOTINVH1.Item("INV_NO") & "'"
            End If
        Next

        invoices = invoices.Substring(1).Trim

        Try
            Initialize_Report("SORINVC1", String.Empty, "RPT", False)
            REPORTS("SORINVC1").Fill_Records_RPT(New String() {invoices, "B"})
            With REPORTS("SORINVC1")
                .Print_Report_Begin()
                .Print_Report()
                .Print_Report_End(False, False)
            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Print Selected", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub tabSales_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSales.SelectedTabChanged
        Setup_tabSales()
    End Sub

    Sub Setup_tabSales()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Sales").Visible = ScreenMode And (tabSales.SelectedTab.Key = "12 Mos Summary")
    End Sub

    Private Sub grdARTCUST1_CODES_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUST1_CODES.InitializeRow
        Dim C As System.Drawing.Color = System.Drawing.Color.Empty

        Select Case e.Row.Cells("CODE_CATEGORY").Value
            Case "SALES"
                C = Drawing.Color.Green
            Case "ACCTG"
                C = Drawing.Color.BlueViolet
        End Select

        e.Row.Cells("CODE_VALUE").Appearance.ForeColor = C
        e.Row.Cells("DESC_VALUE").Appearance.ForeColor = C

    End Sub

    Private Sub grdARTOPEN1_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdARTOPEN1.DoubleClickCell
        If e.Cell.Column.Key = "PYMT_BATCH_NO" Then
            Dim batch As String = e.Cell.Value
            Load_ARTPYMT2("", "", batch)
            grdARTPYMT2.Tag = ""
            UltraTabControl5.SelectedTab = UltraTabControl5.Tabs("Payment History")
        End If
    End Sub

    Private Sub grdARTOPEN1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTOPEN1.InitializeRow
        If e.Row.Band.Index = 0 Then
            Dim CURR_CODE As String = e.Row.Cells("CURR_CODE").Value & ""
            If CURR_CODE <> "USD" Then
                e.Row.Cells("CURR_CODE").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("INV_TOTAL_AMT").Appearance.ForeColor = Drawing.Color.Red
            End If

            Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & ""
            If CUST_CODE <> HFs("CUST_CODE") Then
                e.Row.Cells("CUST_CODE").Appearance.ForeColor = Drawing.Color.Red
            End If

            Dim CUST_CODE_SO As String = e.Row.Cells("CUST_CODE_SO").Value & ""
            If CUST_CODE_SO <> HFs("CUST_CODE") Then
                e.Row.Cells("CUST_CODE_SO").Appearance.ForeColor = Drawing.Color.Red
            End If

            Dim POST_CODE As String = e.Row.Cells("POST_CODE").Value & ""
            If POST_CODE = "REBCR" Then
                e.Row.Cells("POST_CODE").Appearance.BackColor = Drawing.Color.Yellow
            ElseIf POST_CODE <> rowARTCUST1.Item("POST_CODE") & "" And POST_CODE <> "" Then
                e.Row.Cells("POST_CODE").Appearance.ForeColor = Drawing.Color.Red
            End If

            Dim INV_TOTAL_AMOUNT As Decimal = Val(e.Row.Cells("INV_TOTAL_AMOUNT").Value & "")
            Dim INV_BALANCE As Decimal = Val(e.Row.Cells("INV_BALANCE").Value & "")
            If INV_BALANCE = 0 Or (INV_TOTAL_AMOUNT <> 0 AndAlso System.Math.Abs(INV_BALANCE) / System.Math.Abs(INV_TOTAL_AMOUNT) < 0.05) Then
                e.Row.Cells("DAYS").Appearance.BackColor = Drawing.Color.WhiteSmoke
            End If
        End If
    End Sub

#Region "grdARTCUSTD"

    Private Sub grdARTCUSTD_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTD.AfterCellUpdate

        'Select Case e.Cell.Column.Key
        '    Case "CUST_BILL_TO_CUST"
        '        grdCodeDesc(grdARTCUSTD, "ARTCUST1", "CUST_BILL_TO_CUST", "CUST_NAME")
        'End Select
    End Sub

    Private Sub grdARTCUSTD_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUSTD.AfterRowActivate
        'With grdARTCUSTD.DisplayLayout.Bands(0)
        '    If grdARTCUSTD.ActiveRow.IsAddRow Then
        '        .Columns("CONTACT_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
        '        grdARTCUSTD.ActiveCell = grdARTCUSTD.ActiveRow.Cells("CUST_CODE")
        '        grdARTCUSTD.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
        '    Else
        '        .Columns("CONTACT_NO").CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdARTCUSTD_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdARTCUSTD.BeforeExitEditMode
        'grdFieldFormat(grdARTCUSTD)
    End Sub

    Private Sub grdARTCUSTD_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUSTD.BeforeRowUpdate
        With grdARTCUSTD
            'If LookUp("ARTCUST1", e.Row.Cells("CUST_BILL_TO_CUST").Text) Is Nothing Then
            '    MsgBox("Invalid Value entered for Customer (" & e.Row.Cells("CUST_BILL_TO_CUST").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            '    Exit Sub
            'End If

            If e.Row.IsAddRow Then
                .ActiveRow.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
                .ActiveRow.Cells("CONTACT_NO").Value = Val(dst.Tables("ARTCUSTD").Compute("MAX(CONTACT_NO)", "") & "") + 1
            End If
        End With
    End Sub

    Private Sub grdARTCUSTD_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUSTD.Error
        grdARTCUSTD.ActiveRow.CancelUpdate()
    End Sub
#End Region

    Private Sub cmdAEB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAEB.Click
        Load_Order_Guide()
    End Sub

    Private Sub grdSOTAEBP1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTAEBP1.AfterRowActivate
        If grdSOTAEBP1.ActiveRow.IsDataRow Then
            Setup_SOTAEBP2()
        End If
    End Sub

    Private Sub tabOrders_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabOrders.SelectedTabChanged
        Setup_tabOrders()
    End Sub

    Sub Setup_tabOrders()
        If SELECTION_NO = 0 Then Exit Sub
        If tabOrders.SelectedTab.Key = "Summary by Style" Then
            'Load_Order_Guide()
        ElseIf tabOrders.SelectedTab.Key = "Order History" Then
            If grdSOTORDR0.Tag = "*" Then
                Load_SOTORDR1()
                grdSOTORDR0.Tag = ""
            End If
        ElseIf tabOrders.SelectedTab.Key = "Invoice History" Then
            If grdSOTINVH1.Tag = "*" Then
                Load_SOTINVH1()
                grdSOTINVH1.Tag = ""
            End If
        End If
        SetControlPanel()
    End Sub

    Private Sub optAEBDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAEBDate.ValueChanged
        dteAEBFrom.ReadOnly = (optAEBDate.Value = "A")
        dteAEBTo.ReadOnly = (optAEBDate.Value = "A")
    End Sub

    Sub Load_Order_Guide()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading All Ever Bought")

        Dim sql1 As String = sqlSOTAEBP1

        If optAEBDate.Value = "A" Then
            sql1 = Replace(sql1, " and SOTINVH1.INV_DATE >= :PARM2", "")
            sql1 = Replace(sql1, " and SOTINVH1.INV_DATE <= :PARM3", "")
            grdSOTAEBP1.Text = "All Ever Bought"
        Else
            sql1 = Replace(sql1, ":PARM2", "'" & Format(dteAEBFrom.Value, "dd-MMM-yyyy") & "'")
            sql1 = Replace(sql1, ":PARM3", "'" & Format(dteAEBTo.Value, "dd-MMM-yyyy") & "'")
            grdSOTAEBP1.Text = "All Ever Bought from " & Format(dteAEBFrom.Value, "MM/dd/yyyy") & " thru " & Format(dteAEBTo.Value, "MM/dd/yyyy")
        End If
        sql1 = Replace(sql1, ":PARM1", "'" & Absx1.txtFor("CUST_CODE").Text & "'")

        Fill_Records("SOTAEBP1", "", True, sql1)
        Sort_grdColumns(grdSOTAEBP1, "STYLE_CODE")

        splSOTAEBP1.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_SOTAEBP2()

        If grdSOTAEBP1.ActiveRow Is Nothing Then
            grdSOTAEBP2.Visible = False
        Else
            grdSOTAEBP2.Visible = True

            Dim sql As String = Replace(sqlSOTAEBP2, ":PARM1", "'" & Absx1.txtFor("CUST_CODE").Text & "'")

            If optAEBDate.Value = "A" Then
                sql = Replace(sql, " and SOTINVH1.INV_DATE >= :PARM2", "")
                sql = Replace(sql, " and SOTINVH1.INV_DATE <= :PARM3", "")
            Else
                sql = Replace(sql, ":PARM2", "'" & Format(dteAEBFrom.Value, "dd-MMM-yyyy") & "'")
                sql = Replace(sql, ":PARM3", "'" & Format(dteAEBTo.Value, "dd-MMM-yyyy") & "'")
            End If

            Dim STYLE_CODE As String = grdSOTAEBP1.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdSOTAEBP1.ActiveRow.Cells("COLOR_CODE").Value
            sql = Replace(sql, ":PARM4", "'" & STYLE_CODE & "'")
            sql = Replace(sql, ":PARM5", "'" & COLOR_CODE & "'")

            Fill_Records("SOTAEBP2", "", True, sql)
            Sort_grdColumns(grdSOTAEBP2, "INV_DATE".ToLower)
            grdSOTAEBP2.Text = "Invoice Details for Style/Color " & STYLE_CODE & "/" & COLOR_CODE
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Public Class srtComparerSIZE_CODE
        Implements IComparer

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

            Dim xCell As UltraWinGrid.UltraGridCell = DirectCast(x, UltraWinGrid.UltraGridCell)
            Dim yCell As UltraWinGrid.UltraGridCell = DirectCast(y, UltraWinGrid.UltraGridCell)

            Return System.Math.Sign((Val(xCell.Row.Cells("RANK1").Value & "") * 100000 + Val(xCell.Row.Cells("RANK2").Value & "")) _
                                  - (Val(yCell.Row.Cells("RANK1").Value & "") * 100000 + Val(yCell.Row.Cells("RANK2").Value & "")))

        End Function

    End Class

    Public Overrides Function CustomSummary_End( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal rows As UltraWinGrid.RowsCollection, _
    ByVal CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid) As Double
        Select Case grd.Name
            Case "grdSOTINVH1_FRT_WHSE"
                Dim KEY As String = summarySettings.Key
                Dim FRT_AMT As Decimal = 0
                Dim UNITS As Decimal = 0
                grdSOTINVH1_FRT_WHSE_Calculate_Totals(rows, FRT_AMT, UNITS, KEY)
                CustomValue = 0
                If UNITS <> 0 Then CustomValue = FRT_AMT / UNITS
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub grdSOTINVH1_FRT_WHSE_Calculate_Totals( _
    ByVal rows As UltraWinGrid.RowsCollection, _
    ByRef FRT_AMT As Decimal, _
    ByRef UNITS As Decimal, _
    ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                grdSOTINVH1_FRT_WHSE_Calculate_Totals(gbrow.Rows, FRT_AMT, UNITS, KEY)
            Else
                If KEY = "FRT_RATE" Then
                    FRT_AMT += Val(grow2.Cells("FRT_AMT_ACCRUED_WHSE").Value & "")
                    UNITS += Val(grow2.Cells("UNITS").Value & "")
                ElseIf KEY = "FRT_RATE_ACTUAL" Then
                    If grow2.Cells("PPD_INV_IND").Value <> "0" Then
                        FRT_AMT += Val(grow2.Cells("FRT_AMOUNT_WHSE").Value & "")
                        UNITS += Val(grow2.Cells("UNITS").Value & "")
                    End If
                End If
            End If
        Next
    End Sub

    Sub Load_Customer_and_Credit_History()
        ' Fill_Records("ARTCUST1", HFs("CUST_CODE"))
        ' rowARTCUST1 = dst.Tables("ARTCUST1").Rows(0)
        rowARTCUST1 = Fill_Record("ARTCUST1", HFs("CUST_CODE"))
        If rowARTCUST1.Item("CUST_CREDIT_RELEASE") & "" = "" Then
            rowARTCUST1.Item("CUST_CREDIT_RELEASE") = "0"
        End If

        If rowARTCUST1.Item("CUST_CREDIT_GROUP_CUST") & "" <> "" Then
            Fill_Records("ARTCUST1_CREDIT", rowARTCUST1.Item("CUST_CREDIT_GROUP_CUST"))
        End If

        Fill_Records("ARTCUST5", HFs("CUST_CODE"))
        Sort_grdColumns(grdARTCUST5, "INIT_DATE".ToLower)
        'For Each rowARTCUST5 As DataRow In dst.Tables("ARTCUST5").Select("CUST_CREDIT_RELEASE IS NULL")
        '    rowARTCUST5.Item("CUST_CREDIT_RELEASE") = "0"
        'Next

        Fill_Records("TATEVNT1", HFs("CUST_CODE"))
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

    End Sub

    Private Sub optSalesIC_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSalesIC.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        SetUp12Month()
    End Sub

    Sub Setup_optSales()
        If SELECTION_NO = 0 Then Exit Sub
        For i As Integer = 1 To 13
            dst.Tables("SATCUSTS").Columns("V" & Format(i, "00")).Expression = optSales.Value & Format(i, "00")
            dst.Tables("SATCUSTI").Columns("V" & Format(i, "00")).Expression = optSales.Value & Format(i, "00")
        Next

        grdSATCUSTS.Text = "12 Months Sales Summary (" & optSales.Text & IIf(optSalesIC.Value = "A", "", ", " & optSalesIC.Text & " only") & ") for " & Absx1.txtFor("CUST_CODE").Text & ":" & Absx1.txtFor("CUST_NAME").Text
        If grdSATCUSTI.Visible AndAlso grdSATCUSTS.ActiveRow IsNot Nothing AndAlso grdSATCUSTS.ActiveRow.IsDataRow Then
            Dim SALES_DIVISION_CODE As String = grdSATCUSTS.ActiveRow.Cells("SALES_DIVISION_CODE").Value & ""
            Dim STYLE_CLASS_CODE As String = grdSATCUSTS.ActiveRow.Cells("STYLE_CLASS_CODE").Value & ""
            grdSATCUSTI.Text = "12 Months Sales Summary (" & optSales.Text & IIf(optSalesIC.Value = "A", "", ", " & optSalesIC.Text & " only") & ") for " & Absx1.txtFor("CUST_CODE").Text & ", Division " & SALES_DIVISION_CODE & ", Class " & STYLE_CLASS_CODE
        End If

        grdSATCUSTS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdSATCUSTI.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdSATCUSTS.DataBind()
        grdSATCUSTI.DataBind()
    End Sub

    Private Sub optSales_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSales.ValueChanged
        Setup_optSales()
    End Sub

    Private Sub cmdInvoiceHistory_Click(sender As System.Object, e As System.EventArgs) Handles cmdInvoiceHistory.Click
        grdSOTINVH1.Tag = "*"
        Setup_tabOrders()
    End Sub

    Private Sub grdSOTINVH1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTINVH1.AfterRowActivate
        Setup_SOTINVH2()
    End Sub

    Private Sub chkShowGP_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowGP.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Toggle_AEB_GP()
    End Sub

    Sub Toggle_AEB_GP()
        With grdSOTAEBP1.DisplayLayout.Bands(0)
            .Columns("GP").Hidden = Not chkShowGP.Checked
            .Columns("GPPCT").Hidden = Not chkShowGP.Checked
        End With
        With grdSOTAEBP2.DisplayLayout.Bands(0)
            .Columns("CGS").Hidden = Not chkShowGP.Checked
            .Columns("GP").Hidden = Not chkShowGP.Checked
            .Columns("GPPCT").Hidden = Not chkShowGP.Checked
        End With
    End Sub
End Class