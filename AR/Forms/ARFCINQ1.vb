Imports System.Collections.Specialized
Public Class ARFCINQ1

    Dim AGED_TOTALS() As Decimal
    Dim DUE_TOTALS() As Decimal
    Dim AGE_DAYS(4) As Integer
    Dim DUE_DAYS(4) As Integer
    Dim DUE_DATE(4) As String
    Dim AGE_DATE(4) As String

    Dim rowARTCUST1 As DataRow
    Dim rowARTCUST1_orig As DataRow
    Dim TOTAL_DUE As Double
    Dim TOTAL_DUE_INV As Decimal
    Dim LogEditingEnabled As Boolean = False


    Dim sqlSOTORDR1 As String = ""
    Dim SATCUSTS As String
    Dim sqlSATCUSTS As String = ""
    Dim sqlSATCUSTI As String = ""
    Dim sqlARTOPEN1 As String = ""

    Dim ARTCSUMC As String

    Dim showZBI As Boolean = True

    Dim ARTCUST0 As String = ""
    Dim ARTCUSTS As String = ""
    Dim ARTCUSTX As String = ""

    Dim sqlSOTAEBP1 As String = ""
    Dim sqlSOTAEBP2 As String = ""

    Dim collections_mode As Boolean = (MENU_ITEM_OBJECT = "ARFCINQC")
    Dim collections_write_off_started As Boolean = False

    Dim REASON_CODEs As New List(Of String)

    Dim PYMT_BATCH_NO As String = ""
    Dim PYMT_BATCH_LNO As Integer = 0

    Dim order_guide_loaded As Boolean

    Dim INV_NUM_column As String = "INV_NUM"

    Dim ARTSTMTR As String

    Dim sqlARTCUSTT_FUPS As New Text.StringBuilder With {.Length = 0}

    Public Structure strTOTALS
        Public APPL_TOTAL As Decimal
        Public DISC_TOTAL As Decimal
        Public WOFF_TOTAL As Decimal
        Public DED_TOTAL As Decimal
        Public CHB_TOTAL As Decimal
        Public OA_TOTAL As Decimal
        Public GL_TOTAL As Decimal
        Public NET_AR As Decimal
        Public UNAPPLIED As Decimal
    End Structure
    Dim TOTALS As strTOTALS

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
        '    INV_NUM_column = "INV_NO"
        'End If


        'If ASCMAIN1.USER_ID = "wjz" Then
        '    collections_mode = True
        'End If

        AUDIT.Add("ARTOPEN1", "*")

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

            If ASCMAIN1.CLIENT = "RGI" Then
                Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
                ''PRD_END_DATE = rowGLTPARM2.Item("PRD_END_DATE")
                Dim Report_date As Date = Now
                ' Report_date = rowGLTPARM2.Item("PRD_END_DATE")


                TAC.ARCMAIN2.Get_Aging_Data_RGI(
                ROWs("ARTPARM1"),
                Report_date, True)

                ASCMAIN1.sql = "Select ARTOPEN1.* " & TAC.ARCMAIN2.DAYS_AND_BUCKETS _
                & ", DECODE (ARTOPEN1.INV_TYPE,'B',ARTOPEN1.INV_BALANCE,0) CHARGEBACKS " & vbCrLf _
                & ", CASE WHEN ARTOPEN1.INV_TYPE = 'C' OR ARTOPEN1.INV_TYPE = 'O' THEN ARTOPEN1.INV_BALANCE ELSE 0 END CREDITS" & vbCrLf _
                & " from ARTOPEN1, ARTCUST1 ARTCUSTX, TATTERM1 " & vbCrLf _
                & " where ARTOPEN1.INV_BALANCE <> 0" & vbCrLf _
                & " and ARTCUSTX.CUST_CODE = ARTOPEN1.CUST_CODE" & vbCrLf _
                & " and TATTERM1.TERM_CODE = ARTOPEN1.TERM_CODE" & vbCrLf _
                & " and ARTCUSTX.CUST_CODE = :PARM1"
                Create_TDA(dst.Tables.Add, "ARTSTMTR", "**", 0, False, "V", 3)
                Dim ADD_SQL As String = "(" & ASCMAIN1.sql & ")"

                ASCMAIN1.sql = "Select CUST_CODE" & vbCrLf _
                & ", SUM (DECODE(AGE_BUCKET,0,INV_BALANCE,0)) AGE_0" & vbCrLf _
                & ", SUM (DECODE(AGE_BUCKET,1,INV_BALANCE,0)) AGE_1" & vbCrLf _
                & ", SUM (DECODE(AGE_BUCKET,2,INV_BALANCE,0)) AGE_2" & vbCrLf _
                & ", SUM (DECODE(AGE_BUCKET,3,INV_BALANCE,0)) AGE_3" & vbCrLf _
                & ", SUM (DECODE(AGE_BUCKET,4,INV_BALANCE,0)) AGE_4" & vbCrLf _
                & ", SUM (DECODE(DUE_BUCKET,0,INV_BALANCE,0)) DUE_0" & vbCrLf _
                & ", SUM (DECODE(DUE_BUCKET,1,INV_BALANCE,0)) DUE_1" & vbCrLf _
                & ", SUM (DECODE(DUE_BUCKET,2,INV_BALANCE,0)) DUE_2" & vbCrLf _
                & ", SUM (DECODE(DUE_BUCKET,3,INV_BALANCE,0)) DUE_3" & vbCrLf _
                & ", SUM (DECODE(DUE_BUCKET,4,INV_BALANCE,0)) DUE_4" & vbCrLf _
                & " from " & ADD_SQL & "  group by CUST_CODE"
                Create_TDA(dst.Tables.Add, "ARTCUSTA", "**", 0, False, "V", 1)

            End If


            sqlARTOPEN1 = "SELECT ARTOPEN1.*" & vbCrLf _
            & ", DECODE(ARTOPEN1.OPS_YYYYPP_PAID,NULL,ARTOPEN1.DATE_PAID,TRUNC(SYSDATE)) - ARTOPEN1.INV_DATE DAYS" & vbCrLf _
            & ", (CASE WHEN ARTOPEN1.INV_DATE >= " & AGE_DATE_ORA(2) & " AND ARTOPEN1.INV_DATE < " & AGE_DATE_ORA(1) & " THEN '1' ELSE" & vbCrLf _
            & "  CASE WHEN ARTOPEN1.INV_DATE >= " & AGE_DATE_ORA(3) & " AND ARTOPEN1.INV_DATE < " & AGE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
            & "  CASE WHEN ARTOPEN1.INV_DATE >= " & AGE_DATE_ORA(4) & " AND ARTOPEN1.INV_DATE < " & AGE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
            & "  '4' END END END) AGE_BUCKET" & vbCrLf _
            & ", TRUNC(SYSDATE) - ARTOPEN1.INV_DATE AGE" & vbCrLf _
            & ", (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(2) & " THEN '1' ELSE" & vbCrLf _
            & "  CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(3) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
            & "  CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(4) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
            & "  '4' END END END) DUE_BUCKET" & vbCrLf _
            & ", TRUNC(SYSDATE) - ARTOPEN1.INV_DUE_DATE DUE" & vbCrLf _
            & ", SOTSHIP1.BILL_OF_LADING_NO" & vbCrLf _
            & ", SOTINVH1.WHSE_CODE" & vbCrLf
            ASCMAIN1.sql = sqlARTOPEN1 _
            & " from ARTOPEN1," & ARTCUST0 & " ARTCUST0, SOTSHIP1, SOTINVH1 where ARTOPEN1.CUST_CODE = ARTCUST0.CUST_CODE" & vbCrLf _
            & " and SOTINVH1.INV_TYPE (+) = 'I' and SOTINVH1.INV_NO (+) = ARTOPEN1.INV_NUM" & vbCrLf _
            & " and SOTSHIP1.SHIP_BOL_NO (+) = SOTINVH1.SHIP_BOL_NO" & vbCrLf _
            & IIf(collections_mode, vbCrLf & " and (INV_TYPE = 'I' or (INV_TYPE = 'B' and REASON_CODE in (Select REASON_CODE from ARTREAS1 where SHIPPING_VIOLATION = '1')))", "")
            Create_TDA(.Tables.Add, "ARTOPEN1", "**", 0, True, "", 3, "INV_NOTES,TERM_CODE,INV_DUE_DATE")
            '.Tables("ARTOPEN1").Columns("AGE").DataType = GetType(System.Int32)

            '& ", DECODE(DTP,NULL,TRUNC(SYSDATE) - ARTOPEN1.INV_DATE,NULL) AGE" _
            With .Tables("ARTOPEN1")
                .Columns.Add("WOFF_BTN")
                .Columns.Add("WOFF", GetType(System.Decimal))
                .Columns.Add("INV_BALANCE_NEW", GetType(System.Decimal), "ISNULL(INV_BALANCE,0)-ISNULL(WOFF,0)")
                .Columns.Add("AGE_1", GetType(System.Decimal), "IIF(AGE_BUCKET='1',ISNULL(INV_BALANCE,0),0)")
                .Columns.Add("AGE_2", GetType(System.Decimal), "IIF(AGE_BUCKET='2',ISNULL(INV_BALANCE,0),0)")
                .Columns.Add("AGE_3", GetType(System.Decimal), "IIF(AGE_BUCKET='3',ISNULL(INV_BALANCE,0),0)")
                .Columns.Add("AGE_4", GetType(System.Decimal), "IIF(AGE_BUCKET='4',ISNULL(INV_BALANCE,0),0)")
            End With

            ASCMAIN1.sql = "Select ARTPYMT2.*, ARTPYMT1.PYMT_BATCH_DATE" _
            & ", ARTPYMT1.BANK_CODE, ARTPYMT1.PYMT_APPL_ONLY" _
            & ", ARTPYMT1.PYMT_SOURCE, ARTPYMT1.OPS_YYYYPP" _
            & ", ARTCCPA1.CUST_CREDIT_CARD_LAST4" _
            & " from ARTPYMT2,ARTPYMT1," & ARTCUST0 & " ARTCUST0, ARTCCPA1 " _
            & " where ARTPYMT2.CUST_CODE = ARTCUST0.CUST_CODE" _
            & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & "   and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
            & "   and ARTCCPA1.CCPA_NO (+) = ARTPYMT2.CCPA_NO" _
            & "   and ARTPYMT1.OPS_YYYYPP >= :PARM1"
            Create_TDA(.Tables.Add, "ARTPYMTY", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select ARTPYMT3.PYMT_BATCH_NO, ARTPYMT3.PYMT_BATCH_LNO, ARTPYMT3.PYMT_BATCH_ILNO " & vbCrLf _
            & ", ARTPYMT3.INV_TYPE, ARTPYMT3.INV_NUM, ARTPYMT3.INV_DATE, ARTPYMT3.INV_DUE_DATE" & vbCrLf _
            & ", ARTPYMT3.INV_BALANCE, ARTPYMT3.INV_PMT PMT" & vbCrLf _
            & ", NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0) DED, ARTPYMT3.INV_BALANCE_NEW" & vbCrLf _
            & ", ARTPYMT3.INV_CUST_PO CUST_REF, ARTPYMT3.REASON_CODE, ARTREAS1.REASON_DESC" & vbCrLf _
            & ", ARTPYMT3.SEG2_CODE, ARTPYMT3.SEG3_CODE, ARTPYMT3.SEG4_CODE, '0' CHARGEBACK_IND, ARTPYMT3.PARTNER_ORDR_NO OUR_REF" & vbCrLf _
            & ", DECODE(ARTPYMT3.INV_TYPE,'I',NVL(ARTOPEN1.DTP,NVL(ARTOPENX.DTP,0)),NULL) DTP" & vbCrLf _
            & ", DECODE(ARTPYMT3.INV_TYPE,'I',NVL(ARTOPEN1.DTP,NVL(ARTOPENX.DTP,0)) - (ARTPYMT3.INV_DUE_DATE - ARTPYMT3.INV_DATE),NULL) DPD" & vbCrLf _
            & " from ARTPYMT3, ARTREAS1, ARTOPEN1, ARTOPENX" & vbCrLf _
            & " where ARTREAS1.REASON_CODE (+) = ARTPYMT3.REASON_CODE " & vbCrLf _
            & "   and ARTPYMT3.PYMT_BATCH_NO = :PARM1 AND ARTPYMT3.PYMT_BATCH_LNO = :PARM2" & vbCrLf _
            & "   and ARTOPEN1.INV_TYPE (+) = ARTPYMT3.INV_TYPE and ARTOPEN1.INV_NUM (+) = ARTPYMT3.INV_NUM" & vbCrLf _
            & "   and ARTOPENX.INV_TYPE (+) = ARTPYMT3.INV_TYPE and ARTOPENX.INV_NUM (+) = ARTPYMT3.INV_NUM" & vbCrLf _
            & " union" & vbCrLf _
            & " Select ARTPYMT4.PYMT_BATCH_NO, ARTPYMT4.PYMT_BATCH_LNO, ARTPYMT4.PYMT_BATCH_GLNO PYMT_BATCH_ILNO" & vbCrLf _
            & ", 'G' INV_TYPE, 'GL W/Off' INV_NUM, NULL INV_DATE, NULL INV_DUE_DATE" & vbCrLf _
            & ", NULL INV_BALANCE, NULL PMT" & vbCrLf _
            & ", NVL(ARTPYMT4.GL_DIST_AMT,0) DED, NULL INV_BALANCE_NEW" & vbCrLf _
            & ", ARTPYMT4.GL_DIST_REF CUST_REF, ARTPYMT4.ACCT_CODE REASON_CODE, GLTACCT1.ACCT_DESC REASON_DESC" & vbCrLf _
            & ", ARTPYMT4.SEG2_CODE, ARTPYMT4.SEG3_CODE, ARTPYMT4.SEG4_CODE, '0' CHARGEBACK_IND, NULL OUR_REF, NULL DTP, NULL DPD" & vbCrLf _
            & " from ARTPYMT4, GLTACCT1" & vbCrLf _
            & " where GLTACCT1.ACCT_CODE (+) = ARTPYMT4.ACCT_CODE" & vbCrLf _
            & "   and ARTPYMT4.PYMT_BATCH_NO = :PARM1 AND ARTPYMT4.PYMT_BATCH_LNO = :PARM2" & vbCrLf _
            & " union" & vbCrLf _
            & " Select ARTPYMT5.PYMT_BATCH_NO, ARTPYMT5.PYMT_BATCH_LNO, ARTPYMT5.PYMT_BATCH_DLNO PYMT_BATCH_ILNO " & vbCrLf _
            & ", DECODE (ARTPYMT5.CHARGEBACK_IND,'1',ARTPYMT5.INV_TYPE_CB,'X') INV_TYPE" & vbCrLf _
            & ", DECODE (ARTPYMT5.CHARGEBACK_IND,'1',ARTPYMT5.CHARGEBACK_NO,CASE WHEN ARTPYMT5.GL_DIST_AMT < 0 THEN 'CR (Income)' ELSE 'DR (Expense)' END) INV_NUM" & vbCrLf _
            & ", DECODE(NVL(ARTPYMT5.CHARGEBACK_IND,'0'),'1',ARTPYMT2.CUST_PYMT_REF_DATE,NULL) INV_DATE" & vbCrLf _
            & ", NULL INV_DUE_DATE" & vbCrLf _
            & ", NULL INV_BALANCE, NULL PMT" & vbCrLf _
            & ", NVL(ARTPYMT5.GL_DIST_AMT,0) DED" & vbCrLf _
            & ", DECODE(NVL(ARTPYMT5.CHARGEBACK_IND,'0'),'1',NVL(ARTPYMT5.GL_DIST_AMT,0),NULL) INV_BALANCE_NEW" & vbCrLf _
            & ", ARTPYMT5.CUST_REFERENCE CUST_REF, ARTPYMT5.REASON_CODE, ARTREAS1.REASON_DESC" & vbCrLf _
            & ", ARTPYMT5.SEG2_CODE, ARTPYMT5.SEG3_CODE, ARTPYMT5.SEG4_CODE, ARTPYMT5.CHARGEBACK_IND, OUR_REFERENCE OUR_REF, NULL DTP, NULL DPD" & vbCrLf _
            & " from ARTPYMT5, ARTPYMT2, ARTREAS1" & vbCrLf _
            & " where ARTREAS1.REASON_CODE (+) = ARTPYMT5.REASON_CODE" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT5.PYMT_BATCH_LNO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_NO = :PARM1 and ARTPYMT5.PYMT_BATCH_LNO = :PARM2"
            Create_TDA(.Tables.Add, "ARTPYMTX", "**", 0, False, "VI", 0)

            With .Tables("ARTPYMTX")
                .Columns.Add("DTPINV", GetType(System.Decimal), "IIF(INV_TYPE='I',PMT,0)")
                .Columns.Add("DTPWGT", GetType(System.Decimal), "IIF(INV_TYPE='I',PMT*ISNULL(DTP,0),0)")
                .Columns.Add("DPDINV", GetType(System.Decimal), "IIF(INV_TYPE='I',PMT,0)")
                .Columns.Add("DPDWGT", GetType(System.Decimal), "IIF(INV_TYPE='I',PMT*ISNULL(DPD,0),0)")
                .Columns.Add("CURR_CODE", GetType(System.String))
                .Columns.Add("CURR_EXCH_RATE", GetType(System.Decimal))
                .Columns.Add("CURR_GAIN_LOSS", GetType(System.Decimal))
            End With

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

            Create_TDA(.Tables.Add, "EDTTRPM1", "*")

            Create_TDA(.Tables.Add, "ARTCCPA1", "*")
            .Tables("ARTCCPA1").Columns.Add("CUST_NAME")
            Create_TDA(.Tables.Add, "ARTCCPA2", "*")

            Create_TDA(.Tables.Add, "ARTCUST1", "*", , False)
            ASCMAIN1.sql = "SELECT * FROM ARTCUST1"
            Create_TDA(.Tables.Add, "ARTCUST1_CREDIT", "**", 1, False)

            'If ASCMAIN1.CLIENT_CODE = "VDI" Then
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


            'ASCMAIN1.sql = "SELECT TATCONV1.*, ARTCUST1.CUST_NAME " _
            '& " from TATCONV1,ARTCUST1 " _
            '& " where ARTCUST1.CUST_CODE = TATCONV1.TABLE_KEY " _
            '& "   and TATCONV1.CONV_STATUS = '1'" _
            '& "   and TATCONV1.TABLE_NAME = 'ARTCUST1'" _
            '& "   and TATCONV1.CONV_FOLLOWUP_BY = '" & ASCMAIN1.USER_ID & "'"
            sqlARTCUSTT_FUPS.Length = 0
            sqlARTCUSTT_FUPS.AppendLine("SELECT TATCONV1.*,")
            sqlARTCUSTT_FUPS.AppendLine("ARTCUST1.CUST_NAME,")
            sqlARTCUSTT_FUPS.AppendLine("ARTCUST1.CUST_SALES_HOLD,")
            sqlARTCUSTT_FUPS.AppendLine("ARTCUST1.CUST_CREDIT_HOLD,")
            sqlARTCUSTT_FUPS.AppendLine("NVL(BL.INV_BALANCE_CURR,0) INV_BALANCE_CURR")
            sqlARTCUSTT_FUPS.AppendLine("from TATCONV1, ARTCUST1,")
            sqlARTCUSTT_FUPS.AppendLine("(")
            sqlARTCUSTT_FUPS.AppendLine("    SELECT")
            sqlARTCUSTT_FUPS.AppendLine("    CUST_CODE,")
            sqlARTCUSTT_FUPS.AppendLine("    SUM(NVL(INV_BALANCE,0)) AS INV_BALANCE_CURR") 'Of Course there are 2 open items with null _CURR.  Enough of this maddness.
            sqlARTCUSTT_FUPS.AppendLine("    FROM ARTOPEN1")
            sqlARTCUSTT_FUPS.AppendLine("    GROUP BY CUST_CODE")
            sqlARTCUSTT_FUPS.AppendLine(") BL")
            sqlARTCUSTT_FUPS.AppendLine("where ARTCUST1.CUST_CODE = TATCONV1.TABLE_KEY")
            sqlARTCUSTT_FUPS.AppendLine("AND ARTCUST1.CUST_CODE = BL.CUST_CODE (+)")
            sqlARTCUSTT_FUPS.AppendLine("and TATCONV1.CONV_STATUS = '1'")
            sqlARTCUSTT_FUPS.AppendLine("and TATCONV1.TABLE_NAME = 'ARTCUST1'")
            ASCMAIN1.sql = sqlARTCUSTT_FUPS.ToString & " and TATCONV1.CONV_FOLLOWUP_BY = '" & ASCMAIN1.USER_ID & "'"
            Create_TDA(.Tables.Add, "ARTCUSTT_FUPS", "**", 0, False, "", 1)


            chkALLFU.Visible = ASCMAIN1.CLIENT = "RGI" And ASCMAIN1.USER_ID = "andy"

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
            With .Tables("ARTCUST6").Columns
                .Add("AGE_1", GetType(System.Decimal))
                .Add("AGE_2", GetType(System.Decimal))
                .Add("AGE_3", GetType(System.Decimal))
                .Add("AGE_4", GetType(System.Decimal))
            End With

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
                .PrimaryKey = New DataColumn() { .Columns("AGE_DATE")}
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

            ASCMAIN1.sql = "Select SOTINVH1.*, SOTSHIP1.BILL_OF_LADING_NO from SOTINVH1,SOTSHIP1 WHERE SOTSHIP1.SHIP_BOL_NO = SOTINVH1.SHIP_BOL_NO"
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
            .Tables("SOTORDR1").Columns("CUST_STORE_NO").AllowDBNull = True
            .Tables("SOTORDR1").Columns("FRT_TERMS").AllowDBNull = True

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
                .Add("ORDR_HOLD", GetType(System.String), "MAX(CHILD.ORDR_HOLD)")
                .Add("ORDR_AMT_ALLO_CUR", GetType(System.Decimal))
                .Add("PCT_ALLO_CUR", GetType(System.Decimal), "IIF(ORDR_AMT=0, 0, 100 * ORDR_AMT_ALLO_CUR / ORDR_AMT)")
            End With

            ASCMAIN1.sql = "Select SOTORDR2.* from SOTORDR2 where SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "V", 0)
            .Tables("SOTORDR2").Columns("ORDR_STATUS").AllowDBNull = True
            If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then ' If ASCMAIN1.SOLUTION = "VDI" Then
            Else
                .Tables("SOTORDR2").Columns("CUST_CODE").AllowDBNull = True
                .Tables("SOTORDR2").Columns("CUST_STORE_NO").AllowDBNull = True
                .Tables("SOTORDR2").Columns("WHSE_CODE").AllowDBNull = True
            End If

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

            If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then ' If ASCMAIN1.SOLUTION = "VDI" Then
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
            Else
                ASCMAIN1.sql = "SELECT X.*, SOTINVH1.INV_DATE, SOTINVH1.ORDR_CUST_PO, ICTITEM1.ITEM_DESC " & vbCrLf _
                    & " from (" & vbCrLf _
                    & "SELECT SOTINVH2.ITEM_CODE " & vbCrLf _
                    & ", COUNT (*) LINES" & vbCrLf _
                    & ", SUM (SOTINVH2.ORDR_QTY_SHIP) QTY" & vbCrLf _
                    & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) SLS" & vbCrLf _
                    & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ITEM_UNIT_COST) CGS" & vbCrLf _
                    & ", MAX (DECODE(SOTINVH2.INV_TYPE,'I',SOTINVH2.INV_NO,NULL)) INV_NO" & vbCrLf _
                    & " FROM SOTINVH2,SOTINVH1" & vbCrLf _
                    & "WHERE SOTINVH1.CUST_CODE  = :PARM1" & vbCrLf _
                    & "  AND SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
                    & "  AND SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
                    & "GROUP BY SOTINVH2.ITEM_CODE " & vbCrLf _
                    & ") X, SOTINVH1, ICTITEM1" & vbCrLf _
                    & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE and SOTINVH1.INV_TYPE = 'I' AND SOTINVH1.INV_NO = X.INV_NO"
            End If

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

            ARTCSUMC = TAC.SOCMAIN1.Create_ARTCSUMC(Me)

            Dim sqlREASON_CODEs As String = ""
            If collections_mode Then
                ASCMAIN1.sql = "Select REASON_CODE from ARTREAS1 where SHIPPING_VIOLATION = '1'"
            Else
                ASCMAIN1.sql = "Select REASON_CODE from ARTREAS1" & vbCrLf _
                    & " where REASON_CODE in (Select REASON_CODE from ARTOPEN1 where INV_TYPE = 'B')"
            End If

            Dim sqlB As String = ""
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "REASON_CODE")
                Dim REASON_CODE As String = row.Item("REASON_CODE")
                sqlB &= ", Sum (Decode(REASON_CODE,'" & REASON_CODE & "',INV_BALANCE,0)) R" & REASON_CODE
                REASON_CODEs.Add(REASON_CODE)
                sqlREASON_CODEs &= ",'" & REASON_CODE & "'"
            Next

            ASCMAIN1.sql = "Select CUST_CODE, COUNT (*) RECORDS, SUM(INV_BALANCE) TOTAL" _
                & sqlB & " from ARTOPEN1 where INV_TYPE = 'B'" _
                & IIf(collections_mode, " and REASON_CODE in (" & Mid(sqlREASON_CODEs, 2) & ")", "") _
                & " group by CUST_CODE"
            Create_TDA(.Tables.Add, "ARTOPENB", "**", 0, False, "", 1)
            .Tables("ARTOPENB").Columns("RECORDS").DataType = GetType(System.Int64)


            ASCMAIN1.sql = "Select ARTPYMT4.*, GLTACCT1.ACCT_DESC" _
            & " from ARTPYMT4,GLTACCT1" _
            & " where GLTACCT1.ACCT_CODE = ARTPYMT4.ACCT_CODE"
            Create_TDA(.Tables.Add, "ARTPYMT4", "**", 2)


            .Tables.Add("ARTPYMTT")
            With .Tables("ARTPYMTT").Columns
                .Add("PYMT_TOTAL_CODE")
                .Add("PYMT_TOTAL_CAPTION")
                .Add("PYMT_TOTAL_AMT", GetType(System.Double))
            End With

            With .Tables("ARTPYMTT")
                .PrimaryKey = New DataColumn() { .Columns("PYMT_TOTAL_CODE")}
                .Rows.Add(New Object() {"1", "Amt Applied", 0})
                '.Rows.Add(New Object() {"2", "Discounts", 0})
                '.Rows.Add(New Object() {"3", "Write-Off", 0})
                .Rows.Add(New Object() {"4", "Deductions", 0})
                .Rows.Add(New Object() {"5", "ChargeBack", 0})
                .Rows.Add(New Object() {"6", "GL Dist", 0})
                .Rows.Add(New Object() {"7", "On Account", 0})
                .Rows.Add(New Object() {"8", "Net AR", 0})
                .Rows.Add(New Object() {"9", "Out of Bal", 0})
            End With


            Create_TDA(.Tables.Add("ARTOPEN1WOFF"), "ARTOPEN1", "*")

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")
            Create_TDA(.Tables.Add, "ARTPYMT3", "*")

        End With

        cbeYP_PYMTs.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -60) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP_PYMTs.SelectedItem = cbeYP_PYMTs.Items(3)

        cbeSales.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -60) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeSales.SelectedItem = cbeSales.Items(0)

        cbeYPSFrom.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -120) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYPSFrom.SelectedItem = cbeYPSFrom.Items(Val(Mid(ASCMAIN1.CYP, 5, 2)) - 1)
        cbeYPSTo.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -120) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYPSTo.SelectedItem = cbeYPSTo.Items(0)


        cbeSALES_DIVISION.DataSource = ASCDATA1.GetDataTable("Select SALES_DIVISION_CODE, SALES_DIVISION_NAME from SOTSDIV1 order by SALES_DIVISION_CODE")
        cbeSALES_DIVISION.SelectedItem = cbeSALES_DIVISION.Items(0)


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
        grdARTPYMTY.DataSource = dst.Tables("ARTPYMTY")
        grdARTPYMTX.DataSource = dst.Tables("ARTPYMTX")

        grdARTSTMT1.DataSource = dst.Tables("ARTSTMT1")

        grdARTCUST6.DataSource = dst.Tables("ARTCUST6")

        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")

        grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")
        grdSOTINVH2.DataSource = dst.Tables("SOTINVH2")

        grdSOTORDR0.DisplayLayout.GroupByBox.Hidden = True ' False
        grdARTCUST6.DisplayLayout.GroupByBox.Hidden = False


        dst.Tables("ARTCSUMA").Columns.Add("PCT", GetType(System.Decimal))
        dst.Tables("ARTCSUMB").Columns.Add("PCT", GetType(System.Decimal))
        grdARTCSUMA.DataSource = dst.Tables("ARTCSUMA")
        grdARTCSUMB.DataSource = dst.Tables("ARTCSUMB")
        grdARTPYMTT.DataSource = dst.Tables("ARTPYMTT")

        grdARTPYMT4.DataSource = dst.Tables("ARTPYMT4")
        Set_SEGS(grdARTPYMT4, "ARTPYMT4")
        Create_Summary(grdARTPYMT4, "PYMT_BATCH_GLNO", "Count")
        Create_Summary(grdARTPYMT4, "GL_DIST_AMT_CURR")

        grdARTOPENB.DataSource = dst.Tables("ARTOPENB")

        grdARTOPENB.DisplayLayout.GroupByBox.Hidden = True
        With grdARTOPENB.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Drawing.Color.White
                If GCOL.Key = "CUST_CODE" Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf GCOL.Key = "CUST_CODE" Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
            Next

            With .Columns("CUST_CODE")
                .Header.Caption = "Customer"
                .Width = 100
                .Header.Fixed = True
                Create_Summary(grdARTOPENB, "CUST_CODE", "Count")
            End With

            For Each REASON_CODE As String In REASON_CODEs
                With .Columns("R" & REASON_CODE)
                    .Header.Caption = REASON_CODE
                    .Width = 100
                    Dim rowARTEAS1 As DataRow = LookUp("ARTREAS1", REASON_CODE)
                    .Header.ToolTipText = rowARTEAS1.Item("REASON_DESC") & ""
                    Create_Summary(grdARTOPENB, "R" & REASON_CODE)
                End With
            Next
            With .Columns("RECORDS")
                .Header.Caption = "Count"
                .Width = 60
                .Header.Fixed = True
                Create_Summary(grdARTOPENB, "RECORDS")
            End With
            With .Columns("TOTAL")
                .Header.Caption = "Total"
                .Width = 100
                .Header.Fixed = True
                Create_Summary(grdARTOPENB, "TOTAL")
            End With
        End With

        'Show_Filter(grdSOTORDR1, True)
        Show_Filter(grdARTCUST6, True)

        With grdARTPYMTY.DisplayLayout.Bands(0)
            .Columns("PYMT_BATCH_NO").Header.Fixed = True
        End With

        With grdARTCUST6.DisplayLayout.Bands("ARTCUST6")
            .Groups("Customer").Header.Fixed = True
            .Columns("CUST_BILL_TO_CUST").Hidden = True
        End With

        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDR0, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})
        Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})
        Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT_ALLO_CUR"})

        Create_Summary(grdSOTINVH1, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1, New String() {"INV_SALES", "INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT"})

        Create_Summary(grdSOTINVH2, "INV_LNO", "Count")
        Create_Summary(grdSOTINVH2, New String() {"ORDR_QTY_SHIP", "ORDR_AMT_SHIP"})

        Create_Summary(grdARTPYMTY, "PYMT_BATCH_NO", "Count")
        Create_Summary(grdARTPYMTY, "CUST_PYMT_AMT")

        Create_Summary(grdARTPYMTX, "INV_NUM", "Count")
        Create_Summary(grdARTPYMTX, "PMT")
        Create_Summary(grdARTPYMTX, "DED")

        If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then ' If ASCMAIN1.SOLUTION = "VDI" Then
            Create_Summary(grdSOTAEBP1, "STYLE_CODE", "Count")
        Else
            Create_Summary(grdSOTAEBP1, "ITEM_CODE", "Count")
        End If
        Create_Summary(grdSOTAEBP1, New String() {"LINES", "QTY", "SLS", "CGS", "GP"})

        Create_Summary(grdSOTAEBP2, "INV_NO", "Count")
        Create_Summary(grdSOTAEBP2, New String() {"ORDR_QTY_SHIP", "SLS", "CGS", "GP"})

        If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then ' If ASCMAIN1.SOLUTION = "VDI" Then
        Else
            Create_Summary(grdARTCUST2, "CUST_STORE_NO", "Count")
            grdARTCUST2.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Header.Fixed = True

        End If
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
                    If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                        GCOL.Format = "MM/dd/yyyy"
                        GCOL.Width = 120
                    Else
                        GCOL.Format = "MM/dd"
                        GCOL.Width = 60
                    End If

                ElseIf New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO_CUR", "PCT_ALLO_CUR", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(GCOL.Key) Then
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
            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                With .Columns("ORDR_AMT_ALLO_CUR")
                    .Hidden = False
                    .Format = "#,##0.00"
                End With
                With .Columns("PCT_ALLO_CUR")
                    .Hidden = False
                    .Format = "#0.0"
                End With
            Else
                .Columns("ORDR_AMT_ALLO_CUR").Hidden = True
                .Columns("PCT_ALLO_CUR").Hidden = True
            End If
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
                    If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                        GCOL.Format = "MM/dd/yyyy"
                        GCOL.Width = 120
                    Else
                        GCOL.Format = "MM/dd"
                    End If
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
            'With .Columns("PROD_DESC")
            '    .Header.Fixed = True
            '    .Header.Caption = "Description"
            '    .Width = 100
            'End With
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
        Create_Summary(grdARTCUST6, New String() {"CUST_SALES_MTD", "CUST_SALES_YTD", "CUST_SALES_LYR", "CUST_CASH_MTD", "CUST_CASH_YTD", "CUST_CASH_LYR", "CUST_CRED_MTD", "CUST_CRED_YTD", "CUST_CRED_LYR", "CUST_NUM_INV_MTD", "CUST_NUM_INV_YTD", "CUST_NUM_INV_LYR", "INV_BALANCE"})

        Create_Summary(grdARTOPEN1, INV_NUM_column, "Count")
        Create_Summary(grdARTOPEN1, New String() {"INV_BALANCE", "INV_BALANCE_CURR", "INV_TOTAL_AMOUNT", "INV_TOTAL_AMOUNT_CURR", "INV_PMT", "INV_DISC_TAKEN", "INV_WRITE_OFF", "WOFF", "INV_BALANCE_NEW"})

        grdARTCUSTT_FUPS.DisplayLayout.UseFixedHeaders = True
        With grdARTCUSTT_FUPS.DisplayLayout.Bands(0)
            .Columns("TABLE_KEY").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
        End With

        With grdARTOPEN1.DisplayLayout.Bands("ARTOPEN1_ARTPYMTD")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With

        With grdARTOPEN1.DisplayLayout.Bands("ARTOPEN1")
            .Override.HeaderPlacement = UltraWinGrid.HeaderPlacement.FixedOnTop
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next

            .Columns("INV_TYPE").Header.Fixed = True

            If collections_mode Then
                .Columns("REASON_CODE").Header.Fixed = True
            End If

            .Columns(INV_NUM_column).Header.Fixed = True
            .Columns("AGE_BUCKET").GroupByMode = UltraWinGrid.GroupByMode.Value

            .Columns("AGE_1").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1") & ""
            .Columns("AGE_2").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2") & ""
            .Columns("AGE_3").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3") & ""
            .Columns("AGE_4").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4") & ""
            Create_Summary(grdARTOPEN1, "AGE_1")
            Create_Summary(grdARTOPEN1, "AGE_2")
            Create_Summary(grdARTOPEN1, "AGE_3")
            Create_Summary(grdARTOPEN1, "AGE_4")
            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                .Columns("WHSE_CODE").Hidden = False
                .Columns("INV_DUE_DATE").Width = .Columns("INV_DATE").Width
                .Columns("INV_DUE_DATE").Format = "MM/dd/yyyy"
            Else
                .Columns("WHSE_CODE").Hidden = True
            End If
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
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    .Columns("AGE_" & CStr(i)).Hidden = True
                Else
                    .Columns("AGE_" & CStr(i)).Hidden = False
                End If
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

        With grdARTCUSTT_FUPS.DisplayLayout.Bands(0)
            .Columns("INV_BALANCE_CURR").Format = "###,###,##0.00"
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

        ASCMAIN1.Add_Value_List(grdARTOPEN1, "AGE_BUCKET", ,
        New String() {":" _
                , "0:Zero Balance" _
                , "1:" & ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1") _
                , "2:" & ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2") _
                , "3:" & ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3") _
                , "4:" & ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4")})
        ASCMAIN1.Add_Value_List(grdARTOPEN1, "DUE_BUCKET", ,
        New String() {":" _
                , "0:Zero Balance" _
                , "1:" & ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_1") _
                , "2:" & ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_2") _
                , "3:" & ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_3") _
                , "4:" & ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_4")})

        With grdARTOPEN1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
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
            For Each COLUMN_NAME As String In New String() {"WOFF_BTN", "WOFF", "INV_BALANCE_NEW"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
                If Not collections_mode Then
                    .Columns(COLUMN_NAME).Hidden = True
                End If
            Next

        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCSUMA, grdARTCSUMB}
            With grd.DisplayLayout.Bands(0)
                If grd.Name = "grdARTCSUMA" Then
                    With .Columns("LINE_DESC")
                        .Header.Caption = "Description"
                        .Width = 300
                        .Header.Fixed = True
                    End With
                    .Columns("LINE").Hidden = True
                    .Columns("LINE_ABBR").Hidden = True
                    'Create_Summary(grd, "LINE_DESC", "Count")
                Else
                    .Columns("CODE_VALUE").Header.Caption = "Code"
                    .Columns("CODE_VALUE").Width = 100
                    .Columns("CODE_VALUE").Header.Fixed = True
                    Create_Summary(grd, "CODE_VALUE", "Count")
                    .Columns("DESC_VALUE").Header.Caption = "Description"
                    .Columns("DESC_VALUE").Width = 200
                    .Columns("DESC_VALUE").Header.Fixed = True
                    .Columns("LINE").Hidden = True
                End If


                With .Columns("PCT")
                    If grd.Name = "grdARTCSUMA" Then
                        .Header.Caption = "%NSls"
                    Else
                        .Header.Caption = "%Ttl"
                    End If
                    .Width = 60
                    .Header.Fixed = True
                    .Format = "##.0%"
                End With

                For i As Integer = 0 To 13
                    Dim C As String = "AMT" & Format(i, "00")
                    If grd.Name = "grdARTCSUMB" Then
                        Create_Summary(grd, C, , , "#,##0")
                    End If
                    .Columns(C).Format = "#,##0"
                    .Columns(C).Width = 90
                Next

                For Each COLUMN_NAME As String In New String() {"AMT00", "AMT13"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next

                For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                    GCOL.Header.Appearance.BackColor = Drawing.Color.White
                    GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"LINE_DESC", "CODE_VALUE", "DESC_VALUE"}.Contains(GCOL.Key) Then
                        GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        GCOL.Width = 60
                    ElseIf New String() {"AMT00", "AMT13"}.Contains(GCOL.Key) Then
                        GCOL.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    Else
                        GCOL.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    End If
                Next
                .Columns("AMT00").Hidden = True
            End With
        Next

        Setup_grdARTCSUMC()

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

        For I As Integer = 1 To 4
            With grdARTCUST6.DisplayLayout.Bands(0).Columns("AGE_" & Format(I, "0"))
                .Width = 80
                .Format = "#,##0.00"
                .Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_" & Format(I, "0")) & ""
            End With
            Create_Summary(grdARTCUST6, "AGE_" & Format(I, "0"))
        Next

        ASCMAIN1.Add_Value_List(grdARTCUST5, "CUST_CREDIT_RELEASE")
        ASCMAIN1.Add_Value_List(grdARTCUST6, "CUST_STATUS")

        If collections_mode Then
            For Each COLUMN_NAME As String In New String() _
                {"CUST_STORE_NO", "SREP_CODE", "SALES_DIVISION_CODE", "ORDR_CREDIT_APPR_BY",
                 "ORDR_CREDIT_APPR_DATE", "DTP", "ORDR_NO", "CUST_CODE_SO", "POST_CODE",
                 "ORDR_TYPE_CODE", "CUST_CODE", "TERM_CODE", "OPS_YYYYPP", "OPS_YYYYPP_F",
                 "INV_DUE_DATE", "AGE_BUCKET", "DUE_BUCKET", "DUE"}
                grdARTOPEN1.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
            Next
            grdARTOPEN1.DisplayLayout.Bands(0).Columns("INV_CUST_PO").Header.Caption = "Claim No"
            '   grdARTOPEN1.DisplayLayout.Bands(0).Columns("INV_PMT").Header.Caption = "Applied"

            grdARTOPEN1.DisplayLayout.Bands(0).Columns("INV_DISC_TAKEN").Hidden = (ROWs("ARTPARM1").Item("AR_PARM_USE_DISC") & "" <> "1")
            grdARTOPEN1.DisplayLayout.Bands(0).Columns("INV_WRITE_OFF").Hidden = (ROWs("ARTPARM1").Item("AR_PARM_USE_WOFF") & "" <> "1")

            splARTOPEN1.Parent = tabMain.Parent

            splARTSTMT1.Panel1Collapsed = True
        Else
            splARTSTMT1.Panel2Collapsed = True
        End If

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            grdARTOPEN1.DisplayLayout.Bands(0).Columns("PARTNER_ORDR_NO").Header.Caption = "Clarins Inv"
        End If
        If ASCMAIN1.CLIENT = "VAN" Then
            grdARTOPEN1.DisplayLayout.Bands(0).Columns("PARTNER_ORDR_NO").Hidden = True
        End If

        chkEditNotes.Appearance.ForeColor = System.Drawing.Color.White
        chkEditNotes.Appearance.BackColor = System.Drawing.Color.FromArgb(98, 160, 232)
        chkEditNotes.Appearance.BackColor2 = System.Drawing.Color.FromArgb(83, 115, 191)
        chkEditNotes.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select Customer"
                Validate_Code("CUST_CODE")
                'If Not ASCMAIN1.Logical_Lock("ARTCUST1", "TEST", , , , 1) Then Exit Sub

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

            Case "Start Write-Off"


                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then
                        Exit Sub
                    End If

                    If Not ASCMAIN1.Logical_Lock("ARTOPEN1", Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If



                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    Mode_Settings(False)
                    Absx1.txtFor("CUST_CODE").Text = CUST_CODE



                    If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then
                        Exit Sub
                    End If

                    If Not ASCMAIN1.Logical_Lock("ARTOPEN1", Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If
                End If


            Case "Update Write-Off"

                Dim activeOPEN1 As Integer = Val(dst.Tables("ARTOPEN1").Compute("COUNT (INV_NUM)", "WOFF <> 0") & "")
                Dim activePYMT4 As Integer = Val(dst.Tables("ARTPYMT4").Compute("COUNT (PYMT_BATCH_NO)", "GL_DIST_AMT_CURR <> 0") & "")
                ' Dim activePYMT5 As Integer = Val(dst.Tables("ARTPYMT5").Compute("COUNT (PYMT_BATCH_NO)", "GL_DIST_AMT_CURR <> 0") & "")

                If activeOPEN1 + activePYMT4 = 0 Then ' + activePYMT5 = 0 Then
                    EMsg &= vbCr & "No Application Details Found"
                End If

                If Get_TOTAL_UNAPPLIED() <> 0 Then
                    EMsg &= vbCr & "Entry is Out of Balance"
                End If

                'For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("INV_PMT_CURR <> 0 OR INV_DISC_TAKEN_CURR <> 0 OR INV_WRITE_OFF_CURR <> 0", "", DataViewRowState.CurrentRows)
                '    With rowARTPYMT3
                '        Dim rowARTOPEN1 As DataRow = LookUp("ARTOPEN1", New String() {HFs("CUST_CODE"), .Item("INV_TYPE"), .Item("INV_NUM")})
                '        If rowARTOPEN1.Item("OPS_YYYYPP") & "" > ASCMAIN1.CYP Then
                '            EMsg &= vbCr & "You may not apply Items which have a future posting date" & vbCr & " (See AR Item " & rowARTPYMT3.Item("INV_NUM") & ")"
                '        End If
                '    End With
                'Next

                If Absx1.dteFor("CUST_PYMT_REF_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Transaction Reference Date Required"
                Else
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
                    Dim rowGLTPARM2_prior As DataRow = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1))
                    Dim CUST_PYMT_REF_DATE As String = Format(Absx1.dteFor("CUST_PYMT_REF_DATE").Value, "yyyyMMdd")
                    If CUST_PYMT_REF_DATE > Format(rowGLTPARM2.Item("PRD_END_DATE"), "yyyyMMdd") _
                    Or CUST_PYMT_REF_DATE <= Format(rowGLTPARM2_prior.Item("PRD_END_DATE"), "yyyyMMdd") Then
                        EMsg &= vbCr & "Transaction Reference Date should be between " _
                            & Format(CDate(rowGLTPARM2_prior.Item("PRD_END_DATE")).AddDays(1), "MM/dd/yyyy") _
                            & " and " _
                            & Format(rowGLTPARM2.Item("PRD_END_DATE"), "MM/dd/yyyy")
                    End If
                End If

                If Absx1.txtFor("CUST_PYMT_REF_NO").Text = "" Then
                    EMsg &= vbCr & "Transaction Reference No Required"
                End If

                For Each rowARTPYMT4 As DataRow In dst.Tables("ARTPYMT4").Select("ISNULL(GL_DIST_AMT_CURR,0) = 0", "", DataViewRowState.CurrentRows)
                    EMsg &= vbCr & "Distribution Amount may not be 0 (See Line " & rowARTPYMT4.Item("PYMT_BATCH_GLNO") & ")"
                Next
                For Each row As DataRow In dst.Tables("ARTPYMT4").Select("")
                    Dim ACCT_CODE As String = row.Item("ACCT_CODE") & ""
                    If LookUp("GLTACCT1", ACCT_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Account Code " & ACCT_CODE
                    Else
                        If cdr.Item("ACCT_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Acct Code " & ACCT_CODE & " is not Active"
                        End If
                        If cdr.Item("ACCT_SUB_CTL") & "" = "1" Then
                            EMsg &= vbCr & "Acct Code " & ACCT_CODE & " is a Control Account - no Manual J/E permitted"
                        End If
                    End If
                Next

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

            Case "Chargebacks"
                Chargebacks_Summary()
            Case "Data Dump"
                Dim clsARCRGIDD As New ARCRGIDD(Me)
                If clsARCRGIDD.eMsg.Length = 0 Then
                    clsARCRGIDD.makeExcel()
                Else
                    MsgBox(clsARCRGIDD.eMsg, vbCritical, "Excel Creation Cancelled")
                End If
                Me.Cursor = Cursors.Default
            Case "Refresh Follow Ups"
                Refresh_FollowUps()

            Case "Print"
                Print_Hard_Copy()

            Case "Start Write-Off"
                'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                'Mode_Settings(False)
                'Absx1.txtFor("CUST_CODE").Text = CUST_CODE

                collections_write_off_started = True
                PYMT_BATCH_NO = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
                PYMT_BATCH_LNO = 1
                ' Click_Command("Select Customer")
                EntryMode = "I"
                Load_Record()
                Mode_Settings(True)

                Absx1.txtFor("CUST_PYMT_REF_NO").Text = "CUSA W-Off"
                Absx1.dteFor("CUST_PYMT_REF_DATE").Value = Now.Date

            Case "Update Write-Off"
                Update_Record()
                collections_write_off_started = False
                Mode_Settings(False)

            Case "Cancel Write-Off"
                collections_write_off_started = False
                Mode_Settings(True)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Select Customer").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Customer Activity").Visible = Not ScreenMode And Not collections_mode
                    .Items("Chargebacks").Visible = Not ScreenMode
                    .Items("Refresh Follow Ups").Visible = Not ScreenMode

                    .Items("Start Write-Off").Visible = collections_mode And ScreenMode
                    .Items("Chargebacks").Visible = Not ScreenMode
                    .Items("Chargebacks").Visible = Not ScreenMode

                    ' .Items("Print").Visible = False
                    .Items("Data Dump").Visible = False
                    If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                        If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                            .Items("Data Dump").Visible = True
                        End If
                    End If
                End With
                .Groups("Screen Control").Visible = Not collections_write_off_started
                .Groups("Write-Off").Visible = collections_write_off_started
                .Groups("Control Totals").Visible = collections_write_off_started
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = ScreenMode
        grdARTCUST6.Visible = Not ScreenMode And (grdARTCUST6.Tag = "*")
        ' grdARTOPENB.Visible = Not ScreenMode And (grdARTOPENB.Tag = "*")
        tabChargebacks.Visible = Not ScreenMode And (grdARTOPENB.Tag = "*")

        grdSATCUSTS.Visible = False
        splSATCUSTS.Panel2Collapsed = True

        chkEditCredit.Checked = False
        grdARTCUSTT_FUPS.Visible = Not ScreenMode And (grdARTCUST6.Tag <> "*")
        chkALLFU.Visible = Not ScreenMode And (grdARTCUST6.Tag <> "*") And ASCMAIN1.CLIENT = "RGI" And ASCMAIN1.USER_ID = "andy"
        chkALLFU.Checked = False
        chkSALES_HOLD.Visible = ASCMAIN1.CLIENT = "RGI"

        SetControlPanel()

        grpContact.Visible = ScreenMode
        lblSREP.Visible = ScreenMode

        If ScreenMode Then

            grdARTPYMTY.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
            grdARTPYMTY.Text = ""
            SplitContainer5.Parent = UltraTabControl5.Tabs("Payment History").TabPage
            grdARTOPEN1.Parent = splARTOPEN1.Panel1
            grdARTOPEN1.Text = "Open AR Items"
            If collections_mode Then

                Dim EE As Log_Entity = Log_Context()

                If EE.TABLE_NAME <> "" And EE.enabled Then
                    Dim F As New ASFCONV1(ASCMAIN1.ActiveForm, EE)
                    F.TopLevel = False
                    tabCollections.Tabs("Conversation Log").TabPage.Controls.Add(F)
                    F.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
                    F.Dock = DockStyle.Fill
                    F.specific_tab = "G"
                    F.single_row_grid = True
                    F.sqlTATCONV1_where = " and INIT_OPER in (Select USER_ID from ASTUSER2 where SECURITY_CODE = 'CF')"
                    F.Show()
                End If

                grdARTOPEN1.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
                tabCollections.Tabs("GL Distribution").Visible = collections_write_off_started
                If collections_write_off_started Then
                    chkEditNotes.Checked = False
                    chkEditNotes.Visible = False

                    tabCollections.SelectedTab = tabCollections.Tabs("GL Distribution")
                    With grdARTOPEN1.DisplayLayout.Bands(0)
                        .Columns("WOFF_BTN").Hidden = False
                        .Columns("WOFF").Hidden = False
                        .Columns("WOFF").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("INV_BALANCE_NEW").Hidden = False
                    End With

                    grdARTOPEN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                Else
                    chkEditNotes.Visible = True

                    With grdARTOPEN1.DisplayLayout.Bands(0)
                        .Columns("WOFF_BTN").Hidden = True
                        .Columns("WOFF").Hidden = True
                        .Columns("INV_BALANCE_NEW").Hidden = True
                    End With

                    grdARTOPEN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                End If

            Else

                Dim EE As Log_Entity = Log_Context()

                If EE.TABLE_NAME <> "" And EE.enabled Then
                    Dim F As New ASFCONV1(ASCMAIN1.ActiveForm, EE)
                    F.TopLevel = False
                    tabMain.Tabs("Log").TabPage.Controls.Add(F)
                    F.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
                    F.Dock = DockStyle.Fill
                    F.specific_tab = "G"
                    F.single_row_grid = True
                    F.Show()
                End If
            End If


            tabMain.ActiveTab = tabMain.Tabs("General")
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show $0 Balance Items"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = True
            Filter_ARTOPEN1()

            If collections_mode Then
                tabMain.Visible = False
                splARTOPEN1.Visible = True
            End If
        Else
            Clear_Record()

            grdARTPYMTY.Text = "Application History"
            grdARTPYMTY.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = False
            SplitContainer5.Parent = tabChargebacks.Tabs("Applications").TabPage
            grdARTOPEN1.Parent = tabChargebacks.Tabs("Chargebacks").TabPage
            grdARTOPEN1.Text = "Open Chargebacks"
            If collections_mode Then

                If tabCollections.Tabs("Conversation Log").TabPage.Controls.Count > 0 Then
                    Dim F As ASFCONV1 = tabCollections.Tabs("Conversation Log").TabPage.Controls(0)
                    tabCollections.Tabs("Conversation Log").TabPage.Controls.Clear()
                    F.Dispose()
                End If

                grdARTOPEN1.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = False
                grdARTOPEN1.DisplayLayout.Bands(0).Columns("CUST_CODE").Header.Fixed = True
                grdARTOPEN1.DisplayLayout.Bands(0).Columns("CUST_CODE").Header.VisiblePosition = 0
                With grdARTOPEN1.DisplayLayout.Bands(0)
                    .Columns("WOFF_BTN").Hidden = True
                    .Columns("WOFF").Hidden = True
                    .Columns("INV_BALANCE_NEW").Hidden = True
                End With

                grdARTOPEN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Else

                If tabMain.Tabs("Log").TabPage.Controls.Count > 0 Then
                    Dim F As ASFCONV1 = tabMain.Tabs("Log").TabPage.Controls(0)
                    tabMain.Tabs("Log").TabPage.Controls.Clear()
                    F.Dispose()
                End If
            End If

            splSATCUSTS.Visible = False
            splSummary.Visible = False
            If collections_mode Then
                splARTOPEN1.Visible = False
            End If
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            tabMain.Tabs("Outbound").Visible = True
        Else
            tabMain.Tabs("Outbound").Visible = False
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("CSZ").Text = ""

        txtINV_NO_PYMT.Text = ""
        chkEditNotes.Checked = False

        Absx1.dteFor("CUST_PYMT_REF_DATE").Value = DBNull.Value
        Absx1.txtFor("CUST_PYMT_REF_NO").Text = ""

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() _
        {"ARTPYMTY", "ARTPYMTX", "ARTCUST1", "ARTCUST1_CREDIT", "ARTCUST0", "ARTCUSTS",
         "SOTAEBP1", "SOTAEBP2", "SOTORDR0", "SOTORDR1", "SOTORDR2", "SATCUSTS", "SATCUSTI",
         "ARTCUST2", "ARTCUST5", "ARTOPEN1", "ARTSTMT1", "ARTCUST1_AGEDAR",
         "SOTSREP1", "TATCONV1", "TATTERM1", "ARTPYMTD", "ARTSTMT1_DESC", "ARTPYMT1", "ARTPYMT2", "ARTPYMT3", "ARTPYMT4"}
            dst.Tables(TABLE_NAME).Rows.Clear()

        Next
        dst.Tables("EDTTRPM1").Clear()

        If grdARTCUST6.Tag = "" Then
            Refresh_FollowUps()
        End If

        EnforceConstraints(True)

        grdARTPYMTX.DisplayLayout.Bands(0).SummaryFooterCaption = ""

        dst.AcceptChanges()
        If collections_mode Then
        Else
            tabMain.SelectedTab = tabMain.Tabs(0)
        End If

        SetControlPanel()

        splSOTAEBP1.Visible = False
        grdARTOPEN1.DisplayLayout.Bands(0).Columns("BILL_OF_LADING_NO").Hidden = True

        For Each grdname As String In GRDs.Keys
            Dim grd As UltraWinGrid.UltraGrid = GRDs(grdname)
            grd.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        Next

        ClearOBData()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading " & HFs("CUST_CODE"))

        EnforceConstraints(False)

        grdARTPYMTY.Tag = "*"
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

        Load_ARTOPEN1()

        ASCMAIN1.Progress("Now Loading " & HFs("CUST_CODE"))

        If collections_mode Then
            Fill_Records("TATCONV1", Absx1.txtFor("CUST_CODE").Text)

            dst.Tables("ARTPYMTD").Rows.Clear()
            For Each rowARTOPEN1 As DataRow In dst.Tables("ARTOPEN1").Select("")
                Dim INV_NUM As String = rowARTOPEN1.Item("INV_NUM")
                Dim INV_TYPE As String = rowARTOPEN1.Item("INV_TYPE")
                Fill_Records("ARTPYMTD", New String() {INV_TYPE, INV_NUM}, False)
            Next

        Else
            Fill_Records("ARTCUSTD", HFs("CUST_CODE"))
            Sort_grdColumns(grdARTCUSTD, "CUST_CODE")
            Fill_Records("ARTCUSTS")
            Sort_grdColumns(grdARTCUSTS, "CUST_CODE")

            Fill_Records("ARTCUST2", HFs("CUST_CODE"))

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

            order_guide_loaded = False
            'Load_Order_Guide()

            If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then ' If ASCMAIN1.SOLUTION = "VDI" Then
                Sort_grdColumns(grdSOTAEBP1, "STYLE_CODE,COLOR_CODE")
            Else
                Sort_grdColumns(grdSOTAEBP1, "ITEM_CODE")
            End If
            splSOTAEBP1.Visible = True
        End If

        EnforceConstraints(True)


        ASCMAIN1.Progress("Now Loading " & HFs("CUST_CODE"))

        Fill_Records("SOTSREP1")
        Fill_Records("TATTERM1")

        Dim SREP_CODE As String = rowARTCUST1.Item("SREP_CODE") & ""
        Dim SREP_NAME As String = LookUp("SOTSREP1", SREP_CODE, True).Item("SREP_NAME") & ""
        lblSREP.Text = "SRep: " & SREP_CODE & " - " & SREP_NAME

        Absx1.txtFor("CSZ").Text = Absx1.txtFor("CUST_CITY").Text & ", " & Absx1.txtFor("CUST_STATE").Text

        Setup_ARTCUST1_tables()
        If collections_mode Then
        Else
            tabMain.SelectedTab = tabMain.Tabs("Info")
        End If

        grdARTOPEN1.DisplayLayout.Bands(0).Columns("INV_NO_CONS").Hidden = (rowARTCUST1.Item("CUST_CONS_INV") & "" <> "1")

        Toggle_Forex_Columns()

        'Absx1.txtFor("LAST_3612").Text = TAC.ARCMAIN1.Last_3612_SAvg(HFs("CUST_CODE"))

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Delete_Record()
        Stop
    End Sub

    Sub Update_Record()

        Write_Batch_Headers()

        Dim PYMT_BATCH_ILNO As Integer = 0
        For Each rowARTOPEN1 As DataRow In dst.Tables("ARTOPEN1").Select("WOFF <> 0", "INV_NUM")
            PYMT_BATCH_ILNO += 1
            Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow
            With rowARTPYMT3
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                .Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO

                For Each COLUMN_NAME As String In New String() _
                    {"INV_TYPE", "INV_NUM", "REASON_CODE", "INV_DATE", "INV_DUE_DATE",
                     "CUST_CODE_SO", "CUST_STORE_NO", "INV_CUST_PO",
                     "POST_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE",
                     "ORDR_TYPE_CODE", "CURR_CODE", "CURR_EXCH_RATE", "INV_NO_CONS"}
                    .Item(COLUMN_NAME) = rowARTOPEN1.Item(COLUMN_NAME)
                Next

                .Item("INV_DISC_TAKEN") = 0
                .Item("INV_DISC_TAKEN_CURR") = 0
                .Item("INV_WRITE_OFF") = 0
                .Item("INV_WRITE_OFF_CURR") = 0
                .Item("CURR_GAIN_LOSS") = 0

                .Item("INV_PMT") = rowARTOPEN1.Item("WOFF")
                .Item("INV_BALANCE") = rowARTOPEN1.Item("INV_BALANCE")
                .Item("INV_BALANCE_NEW") = rowARTOPEN1.Item("INV_BALANCE_NEW")
                .Item("INV_PMT_CURR") = rowARTOPEN1.Item("WOFF")
                .Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_BALANCE")
                .Item("INV_BALANCE_NEW_CURR") = rowARTOPEN1.Item("INV_BALANCE_NEW")
            End With
            dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
        Next

        dst.Tables("ARTOPEN1WOFF").Rows.Clear()

        For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("", "PYMT_BATCH_ILNO")
            With rowARTPYMT3
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                Dim INV_TYPE As String = .Item("INV_TYPE")
                Dim INV_NUM As String = .Item("INV_NUM")
                Dim INV_PMT As Decimal = Val(.Item("INV_PMT") & "")
                Dim rowARTOPEN1 As DataRow = Fill_Record("ARTOPEN1WOFF", New String() {CUST_CODE, INV_TYPE, INV_NUM}, False, False)
                Dim INV_BALANCE As Decimal = Val(rowARTOPEN1.Item("INV_BALANCE") & "")
                rowARTOPEN1.Item("INV_PMT") = Val(rowARTOPEN1.Item("INV_PMT") & "") + INV_PMT
                rowARTOPEN1.Item("INV_BALANCE") = INV_BALANCE - INV_PMT
                ' rowARTOPEN1.Item("INV_LAST_PMT") = DATETIME_STAMP.Date
                rowARTOPEN1.Item("INV_LAST_PMT_REF") = Absx1.txtFor("CUST_PYMT_REF_NO").Text
                rowARTOPEN1.Item("INV_LAST_PMT_REF_DT") = Absx1.dteFor("CUST_PYMT_REF_DATE").Value
                rowARTOPEN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowARTOPEN1.Item("LAST_DATE") = DATETIME_STAMP
                rowARTOPEN1.Item("INV_PMT_CURR") = rowARTOPEN1.Item("INV_PMT")
                rowARTOPEN1.Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_BALANCE")
            End With
        Next

        For Each rowARTPYMT4 As DataRow In dst.Tables("ARTPYMT4").Select("")
            rowARTPYMT4.Item("GL_DIST_AMT") = rowARTPYMT4.Item("GL_DIST_AMT_CURR")
        Next

        Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").Rows.Find(New Object() {PYMT_BATCH_NO, PYMT_BATCH_LNO})
        rowARTPYMT2.Item("CUST_PYMT_REF_NO") = Absx1.txtFor("CUST_PYMT_REF_NO").Text
        rowARTPYMT2.Item("CUST_PYMT_REF_DATE") = Absx1.dteFor("CUST_PYMT_REF_DATE").Value

        BeginTrans()

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")
        Update_Record_TDA("ARTPYMT3")
        Update_Record_TDA("ARTPYMT4")
        Update_Record_TDA("ARTOPEN1WOFF")

        CommitTrans("Update Complete")
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
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
        If MENU_ITEM_OBJECT = "ARFCINQC" Then
            collections_mode = True
        End If
        Load_Popup_Menu(grdARTSTMT1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        If collections_mode Then
            Load_Popup_Menu(grdARTOPEN1, "SSSSBBBBB", "Show Filter", "Show GroupBox", "Show $0 Balance Items", "Show BOL", "Show", "Show Pymt Applications", "Retrieve Paid Items", "Create Log", "Total Balance")
        Else
            Load_Popup_Menu(grdARTOPEN1, "SSSSBBBBBBBBBBBBBBBB", "Show Filter", "Show GroupBox", "Show $0 Balance Items", "Show BOL",
                        "email", "Fax", "Show", "Sales Order Inquiry", "Customer Returns Inquiry", "Show Pymt Applications",
                        "Retrieve Paid Invoices", "Create Log", "Total Balance", "Change Terms", "Credit Card", "Sales Order Entry", "Show Aged AR", "email Aged AR", "email Cust Statement", "Show Cust Statement")

        End If
        Load_Popup_Menu(grdSOTORDR0, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Order Inquiry", "Sales Order Entry", "Print Selected", "Print Pro-Forma")
        Load_Popup_Menu(grdSOTORDR2, "B", "Style Status Inquiry")
        Load_Popup_Menu(grdSATCUSTS, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdARTCUST6, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTAEBP1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
        Load_Popup_Menu(grdSOTAEBP2, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice")
        Load_Popup_Menu(grdARTCUSTS, "B", "Customer Inquiry")
        If (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") Then
            Load_Popup_Menu(grdSOTINVH1, "SSSPBPBBBPB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice", "Resend EDI Invoice")
        ElseIf (ASCMAIN1.CLIENT = "RGI") Then 'AndAlso ASCMAIN1.DBS_SERVER = "RGI") Then
            Load_Popup_Menu(grdSOTINVH1, "SSSPBPBBBPBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice", "Send Invoice to Web", "Resend EDI Invoice", "Copy Invoice to clipboard")
        Else
            Load_Popup_Menu(grdSOTINVH1, "SSSPBPBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "email Invoice", "Fax Invoice", "Show Invoice")
        End If

        Load_Popup_Menu(grdSOTINVH2, "B", "Style Status Inquiry")

        If ASCMAIN1.USER_SECURITY_CODEs.Contains("AR") Then
            Load_Popup_Menu(grdARTPYMTY, "BSSB", "Issue CC Credit", "Show Filter", "Show GroupBox", "Show Raw EDI")
            Load_Popup_Menu(grdARTPYMTX, "SBBB", "Show Filter", "email Invoice", "Fax Invoice", "Show Invoice")

        Else
            Load_Popup_Menu(grdARTPYMTY, "SS", "Show Filter", "Show GroupBox")
        End If
        Load_Popup_Menu(grdTATEVNT1, "B", "Show email")
        'grdARTCUSTT_FUPS

        If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
            Load_Popup_Menu(grdARTCUSTT_FUPS, "SS", "Show Filter", "Show GroupBox")
        End If
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

        If tlb_pop.Tools.Exists("Show BOL") Then
            ' tlb_sbt.Tag = "x"
            tlb_sbt = DirectCast(tlb_pop.Tools("Show BOL"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grdARTOPEN1.DisplayLayout.Bands(0).Columns("BILL_OF_LADING_NO").Hidden
            'tlb_sbt.Tag = ""
        End If

        If tlb_pop.Tools.Exists("Show Outline") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Show Outline"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = ScreenMode
        End If


        If tlb_pop.Tools.Exists("Change Terms") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Change Terms"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = ASCMAIN1.USER_SECURITY_CODEs.Contains("CR")
        End If

        If tlb_pop.Tools.Exists("Email Cust Statement") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Email Cust Statement"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = ASCMAIN1.DBS_COMPANY = "RGI"

        End If

        If tlb_pop.Tools.Exists("Show Cust Statement") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Show Cust Statement"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = ASCMAIN1.DBS_COMPANY = "RGI"

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

                Case "grdARTPYMTY"
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
                        'tlb_pop.Tools("Sales Order Entry").SharedProps.Visible = False
                        If collections_mode Then
                            '"Show Pymt Applications",
                            For Each tool As String In New String() _
                             {"Show $0 Balance Items", "Retrieve Paid Items", "Create Log"}
                                tlb_pop.Tools(tool).SharedProps.Visible = ScreenMode
                            Next
                        Else
                            tlb_pop.Tools("Sales Order Inquiry").SharedProps.Visible = (grd.ActiveRow.IsDataRow AndAlso grd.ActiveRow.Cells("INV_TYPE").Text = "I")
                            tlb_pop.Tools("Customer Returns Inquiry").SharedProps.Visible = (grd.ActiveRow.IsDataRow AndAlso grd.ActiveRow.Cells("INV_TYPE").Text = "R")

                        End If
                    End If

                Case "grdSOTORDR0"
                    If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
                        tlb_pop.Tools("Sales Order Entry").SharedProps.Visible = True
                        tlb_pop.Tools("Print Selected").SharedProps.Visible = True
                    Else
                        tlb_pop.Tools("Sales Order Entry").SharedProps.Visible = False
                        tlb_pop.Tools("Print Selected").SharedProps.Visible = False
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Print Pro-Forma"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ASCMAIN1.CLIENT = "RGI"

                Case "grdSOTINVH1"
                    If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                    Else
                        'If grd.ActiveRow.Cells("INV_TOTAL_AMOUNT").Value <= 0 OrElse grd.ActiveRow.Cells("INV_TYPE").Value <> "I" OrElse grd.Selected.Rows.Count <> 1 Then
                        '    tlb_pop.Tools("Charge Credit Card").SharedProps.Visible = False
                        'Else
                        '    tlb_pop.Tools("Charge Credit Card").SharedProps.Visible = True
                        'End If
                    End If

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

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Print Pro-Forma"
                If grd.Selected.Rows.Count = 0 Then
                    MessageBox.Show("You must select at least 1 sales order", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim lstCustCodes As New List(Of String)
                Dim lstOrderNos As New List(Of String)
                For Each grdRow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    'Dim CUST_CODE As String = grdRow.Cells("CUST_CODE").Value
                    'If Not lstCustCodes.Contains(CUST_CODE) Then
                    '    lstCustCodes.Add(CUST_CODE)
                    'End If

                    'If lstCustCodes.Count > 1 Then
                    '    MessageBox.Show("All selected sales orders must be for the same customer.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    '    Exit Sub
                    'End If
                    Dim ORDR_NO As String = String.Empty

                    If ASCMAIN1.CLIENT = "RGI" Then
                        ORDR_NO = grdRow.Cells("ORDR_GROUP_NO").Value
                        lstOrderNos.Add(ORDR_NO)
                    Else
                        Dim ORDR_GROUP_NO = grdRow.Cells("ORDR_GROUP_NO").Value
                        ORDR_NO = ASCDATA1.GetDataValue("Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                        Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                        If rowSOTORDR1 IsNot Nothing Then
                            lstOrderNos.Add(ORDR_NO)
                        End If
                    End If
                Next

                If lstOrderNos.Count = 0 Then
                    MessageBox.Show("Could not locate all Selected Orders.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Using F As New ASFMSGBF
                    Dim ExportInfo As Boolean = False
                    If MessageBox.Show("Do you want to show Export Information?", e.Tool.Key, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                        ExportInfo = True
                    End If

                    If ASCMAIN1.CLIENT = "VAN" And InquiryMode Then
                        Print_ProForma("ORDR_QTY", lstOrderNos, ExportInfo)
                    Else
                        Dim ORDR_QTY_fields() As String = New String() {"Qty Ordered", "Qty Open", "Qty Allocated", "Qty In Pick", "Qty Available"}
                        'Dim ORDR_QTY_fields() As String = New String() {"Qty Ordered", "Qty Open", "Qty Allocated", "Qty Allocated Current", "Qty In Pick"}
                        Dim i As Integer = F.Get_opt_from_User("Which Qty Field should be Used", ORDR_QTY_fields, 0, "Pro-Forma Qty Option")
                        If i <> -1 Then
                            Dim ORDR_QTY_field As String = New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_PICK", "ORDR_QTY_ALLO_X"}(i)
                            Print_ProForma(ORDR_QTY_field, lstOrderNos, ExportInfo)
                        End If
                    End If
                End Using


            Case "Show $0 Balance Items"
                Filter_ARTOPEN1()
            Case "Retrieve Paid Invoices"
                Dim numDays As Double = 0
                Using FRM As New ASFMSGBF
                    numDays = FRM.Get_numint_from_User("Days to Retrieve", "Retrieve Paid Invoices")
                    If FRM.user_option <> -1 Then
                        Retrieve_Paid_Invoices(numDays)
                    End If
                End Using

            Case "Retrieve Paid Items"
                Dim numDays As Int64 = 280
                Retrieve_Paid_Items(280)

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

            Case "Show Aged AR"
                Print_Hard_Copy(True)
                Exit Sub

            Case "Show BOL"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grdARTOPEN1.DisplayLayout.Bands(0).Columns("BILL_OF_LADING_NO").Hidden = Not tlb_sbt.Checked

            Case "email Aged AR"
                Dim FILENAME As String = Print_Hard_Copy(True, True)
                '  Show_Document(ASCMAIN1.Folders("Temp") & FILENAME & ".PDF")

                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                Dim CUST_NAME As String = Absx1.txtFor("CUST_NAME").Text
                Dim CUST_EMAIL As String = Absx1.txtFor("CUST_EMAIL").Text
                Dim CUST_CONTACT As String = Absx1.txtFor("CUST_CONTACT").Text
                If CUST_EMAIL <> "" Then
                    EMAIL_ADDRESSs.Add(CUST_EMAIL, IIf(CUST_CONTACT = "", CUST_EMAIL, CUST_CONTACT))
                End If

                Dim ATTACHMENTs As New Dictionary(Of String, String)
                ATTACHMENTs.Add(FILENAME & ".pdf", ASCMAIN1.Folders("Temp") & FILENAME & ".PDF")

                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                "Aged Accounts Receivable", "AR", False, True, CUST_CODE, CUST_NAME, "Customer")

                If SEND_NO <> "" Then
                    TAC.TACMAIN1.Record_Event("ARTOPEN1", CUST_CODE, DATETIME_STAMP, ASCMAIN1.USER_ID, "EML", "email Aged AR " & Format(Now, "MM/dd/yyyy"), SEND_NO)
                    Dim ATTACHMENT_NO As String = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")
                    Dim CONV_NO As String = ASCMAIN1.Next_Control_No("TATCONV1.CONV_NO")

                    If Not dst.Tables.Contains("ASTATTA2") Then
                        Create_TDA(dst.Tables.Add, "ASTATTA2", "*")
                    End If
                    Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").NewRow
                    With rowASTATTA2
                        .Item("TABLE_NAME") = "TATCONV1"
                        .Item("COLUMN_NAME") = "CONV_NO"
                        .Item("CODE_VALUE") = CONV_NO
                        .Item("ATTACHMENT_NO") = ATTACHMENT_NO
                        .Item("ATTACHMENT_DESC") = FILENAME & ".pdf"
                        .Item("ATTACHMENT_FILENAME") = ASCMAIN1.Folders("Temp") & FILENAME & ".PDF"
                        .Item("ATTACHMENT_EXT") = "PDF"
                        .Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
                        .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_DATE") = DATETIME_STAMP
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("ATTACHMENT_DATETIME") = DATETIME_STAMP
                    End With
                    dst.Tables("ASTATTA2").Rows.Add(rowASTATTA2)
                    Update_Record_TDA("ASTATTA2")
                    dst.Tables("ASTATTA2").Rows.Clear()

                    Dim rowTATSEND1 As DataRow = LookUp("TATSEND1", SEND_NO)

                    dst.Tables("TATCONV1").AcceptChanges()
                    Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").NewRow
                    With rowTATCONV1
                        .Item("CONV_NO") = CONV_NO
                        .Item("CONV_DATE") = DATETIME_STAMP.Date
                        .Item("CONV_SUBJECT") = rowTATSEND1.Item("SEND_SUBJECT")
                        .Item("CONV_NOTES") = rowTATSEND1.Item("SEND_BODY")
                        .Item("CONV_STATUS") = "0"
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("TABLE_NAME") = "ARTCUST1"
                        .Item("TABLE_KEY") = CUST_CODE
                        .Item("SEND_NO") = SEND_NO
                    End With
                    dst.Tables("TATCONV1").Rows.Add(rowTATCONV1)
                    Update_Record_TDA("TATCONV1")

                    Dim F2 As ASFCONV1 = DirectCast(tabMain.Tabs("Log").TabPage.Controls(0), ASFCONV1)
                    If F2.tblTATCONV1 IsNot Nothing Then
                        F2.tblTATCONV1.Rows.Add(rowTATCONV1.ItemArray)
                    End If

                    dst.Tables("TATCONV1").Rows.Clear()
                End If

                Exit Sub

            Case "email Cust Statement"
                If grd.ActiveRow Is Nothing Then
                    Exit Sub
                End If

                Call AR_STATEMENT()
                Dim FILENAME As String = Print_Hard_Copy_Statement(True, True)

                dst.Tables("ARTSTMTZ").Rows.Clear()


                '  Show_Document(ASCMAIN1.Folders("Temp") & FILENAME & ".PDF")

                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                Dim CUST_NAME As String = Absx1.txtFor("CUST_NAME").Text
                Dim CUST_EMAIL As String = Absx1.txtFor("CUST_EMAIL").Text
                Dim CUST_CONTACT As String = Absx1.txtFor("CUST_CONTACT").Text
                If CUST_EMAIL <> "" Then
                    EMAIL_ADDRESSs.Add(CUST_EMAIL, IIf(CUST_CONTACT = "", CUST_EMAIL, CUST_CONTACT))
                End If

                Dim ATTACHMENTs As New Dictionary(Of String, String)
                ATTACHMENTs.Add(FILENAME & ".pdf", ASCMAIN1.Folders("Temp") & FILENAME & ".PDF")

                Dim SUBJECT As String = "Statement " & " - Customer No " & grd.ActiveRow.Cells("CUST_CODE").Value
                Dim BODY As String = "Attached is your statement.  Thank you for the opportunity to do business with you."


                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, "AR", False, True, CUST_CODE, CUST_NAME, "Customer", BODY)

                If SEND_NO <> "" Then
                    TAC.TACMAIN1.Record_Event("ARTOPEN1", CUST_CODE, DATETIME_STAMP, ASCMAIN1.USER_ID, "EML", "email Cust Statement " & Format(Now, "MM/dd/yyyy"), SEND_NO)
                    Dim ATTACHMENT_NO As String = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")
                    Dim CONV_NO As String = ASCMAIN1.Next_Control_No("TATCONV1.CONV_NO")

                    If Not dst.Tables.Contains("ASTATTA2") Then
                        Create_TDA(dst.Tables.Add, "ASTATTA2", "*")
                    End If
                    Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").NewRow
                    With rowASTATTA2
                        .Item("TABLE_NAME") = "TATCONV1"
                        .Item("COLUMN_NAME") = "CONV_NO"
                        .Item("CODE_VALUE") = CONV_NO
                        .Item("ATTACHMENT_NO") = ATTACHMENT_NO
                        .Item("ATTACHMENT_DESC") = FILENAME & ".pdf"
                        .Item("ATTACHMENT_FILENAME") = ASCMAIN1.Folders("Temp") & FILENAME & ".PDF"
                        .Item("ATTACHMENT_EXT") = "PDF"
                        .Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
                        .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_DATE") = DATETIME_STAMP
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("ATTACHMENT_DATETIME") = DATETIME_STAMP
                    End With
                    dst.Tables("ASTATTA2").Rows.Add(rowASTATTA2)
                    Update_Record_TDA("ASTATTA2")
                    dst.Tables("ASTATTA2").Rows.Clear()

                    Dim rowTATSEND1 As DataRow = LookUp("TATSEND1", SEND_NO)

                    dst.Tables("TATCONV1").AcceptChanges()
                    Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").NewRow
                    With rowTATCONV1
                        .Item("CONV_NO") = CONV_NO
                        .Item("CONV_DATE") = DATETIME_STAMP.Date
                        .Item("CONV_SUBJECT") = rowTATSEND1.Item("SEND_SUBJECT")
                        .Item("CONV_NOTES") = rowTATSEND1.Item("SEND_BODY")
                        .Item("CONV_STATUS") = "0"
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("TABLE_NAME") = "ARTCUST1"
                        .Item("TABLE_KEY") = CUST_CODE
                        .Item("SEND_NO") = SEND_NO
                    End With
                    dst.Tables("TATCONV1").Rows.Add(rowTATCONV1)
                    Update_Record_TDA("TATCONV1")

                    Dim F2 As ASFCONV1 = DirectCast(tabMain.Tabs("Log").TabPage.Controls(0), ASFCONV1)
                    If F2.tblTATCONV1 IsNot Nothing Then
                        F2.tblTATCONV1.Rows.Add(rowTATCONV1.ItemArray)
                    End If

                    dst.Tables("TATCONV1").Rows.Clear()


                End If

                Exit Sub

            Case "Show Cust Statement"
                Call AR_STATEMENT()
                Dim FILENAME As String = Print_Hard_Copy_Statement(True, False)
                'Show_Document(FILENAME)
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

                    REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {" And SOTORDR1.ORDR_NO In ('" & String.Join("', '", ordrGroupNoList) & "')"})
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

            Case "email", "email Invoice", "email Statement", "Fax", "Fax Invoice", "Fax Statement", "Copy Invoice to clipboard"

                Dim FILENAME As String = ""
                Dim ATTACHMENT As String = ""
                Dim SUBJECT As String = ""
                Dim INV_NO As String = ""
                Dim INVNOs As String = ""

                If e.Tool.OwningMenu.Key = "grdARTSTMT1" Then
                    Dim RYP As String = grd.ActiveRow.Cells("OPS_YYYYPP").Value
                    Dim STMT_NO As String = grd.ActiveRow.Cells("STMT_NO").Value & ""
                    If ASCMAIN1.useUNCPath Then
                        FILENAME = $"{ASCMAIN1.Folders("SharedRoot")}\OSG\" & RYP & "\PDF\" & STMT_NO & ".PDF"
                    Else
                        FILENAME = "S:\OSG\" & RYP & "\PDF\" & STMT_NO & ".PDF"
                    End If

                    ATTACHMENT = ASCMAIN1.Folders("Temp") & STMT_NO & "." & "PDF"
                    SUBJECT = "Statement for " & Mid(RYP, 5, 2) & "/" & Mid(RYP, 1, 4) &
                            " (Acct# " & grd.ActiveRow.Cells("CUST_CODE").Value & " " & rowARTCUST1.Item("CUST_NAME") & ")"

                ElseIf e.Tool.OwningMenu.Key = "grdARTOPEN1" Or e.Tool.OwningMenu.Key = "grdARTPYMTX" Then
                    'INVNOs = GetInvoiceList(grd, e.Tool.OwningMenu.Key, INV_NO)
                    'If INVNOs.Length = 0 Then
                    '    Exit Sub
                    'End If
                    'FILENAME = Create_Invoice(INVNOs)
                    'SUBJECT = "Invoice " & INV_NO
                    'ATTACHMENT = FILENAME 'INV_NO & ".pdf"

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
                    If (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") Then
                        If e.Tool.OwningMenu.Key = "grdARTPYMTX" Then
                            SUBJECT = SUBJECT
                        Else
                            SUBJECT = SUBJECT & " - Customer No " & grd.ActiveRow.Cells("CUST_CODE").Value
                        End If
                    End If

                ElseIf e.Tool.OwningMenu.Key = "grdSOTINVH1" Then
                    'INVNOs = GetInvoiceList(grd, e.Tool.OwningMenu.Key, INV_NO)
                    'If INVNOs.Length = 0 Then
                    '    Exit Sub
                    'End If
                    'FILENAME = Create_Invoice(INVNOs)
                    'SUBJECT = "Invoice " & INV_NO
                    'ATTACHMENT = FILENAME 'INV_NO & ".pdf"

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

                Dim CUST_EMAIL_AR As String = Absx1.txtFor("CUST_EMAIL").Text

                If ASCMAIN1.CLIENT = "VAN" Then
                    If rowARTCUST1.Item("CUST_XMIT_INV_VIA") & "" = "E" And rowARTCUST1.Item("CUST_INV_EMAIL") & "" <> "" Then
                        CUST_EMAIL_AR = rowARTCUST1.Item("CUST_INV_EMAIL") & ""
                    End If
                End If
                If e.Tool.Key Like "email*" Then
                    TAC.SOCMAIN1.email_Invoice(Me,
                        Absx1.txtFor("CUST_CODE").Text,
                        Absx1.txtFor("CUST_NAME").Text,
                        CUST_EMAIL_AR,
                        CUST_EMAIL_AR,
                        FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT, INV_NO)

                ElseIf e.Tool.Key Like "Fax*" Then
                    Send_fax(FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT)
                ElseIf e.Tool.Key Like "*clipboard" Then
                    Dim paths As New StringCollection()
                    paths.Add(FILENAME)
                    If My.Computer.FileSystem.FileExists(FILENAME) Then
                        Clipboard.SetFileDropList(paths)
                        MsgBox(SUBJECT & " copied to clipboard", vbOKOnly, "File Copied")
                    Else
                        MsgBox("File not Found", MsgBoxStyle.OkOnly, "Error Attempting to Copy File ")
                    End If
                End If

            Case "Show", "Show Invoice", "Show Statement", "Show Credit"
                Dim FILENAME As String = ""
                Dim InvNos As String = ""

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

                ElseIf e.Tool.OwningMenu.Key = "grdARTPYMTX" Then

                    If grd.ActiveRow IsNot Nothing Then
                        Dim INV_NUM_TO_SHOW As String = grd.ActiveRow.Cells("INV_NUM").Value
                        Dim INV_TYPE_TO_SHOW As String = grd.ActiveRow.Cells("INV_TYPE").Value
                        FILENAME = Create_Invoice("('" & INV_TYPE_TO_SHOW & "','" & INV_NUM_TO_SHOW & "')")
                        Show_Document(FILENAME)
                    End If

                ElseIf e.Tool.OwningMenu.Key = "grdARTOPEN1" Or e.Tool.OwningMenu.Key = "grdSOTORDR1" Or e.Tool.OwningMenu.Key = "grdSOTINVH1" Or e.Tool.OwningMenu.Key = "grdSOTAEBP2" Then
                    Dim C As String = "INV_NO"
                    If e.Tool.OwningMenu.Key = "grdARTOPEN1" Then C = "INV_NUM"

                    Dim INV_NO As String = grd.ActiveRow.Cells(C).Value
                    InvNos = GetInvoiceList(grd, e.Tool.OwningMenu.Key, INV_NO)
                    If InvNos.Length = 0 Then
                        Exit Sub
                    End If

                    FILENAME = Create_Invoice(InvNos)
                    Show_Document(FILENAME)
                End If

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

                If collections_mode And Not ScreenMode Then
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        If grow.IsDataRow Then
                            INV_NO = grow.Cells(INV_NUM_column).Text
                            INV_TYPE = grow.Cells("INV_TYPE").Text
                            If dst.Tables("ARTPYMTD").Select("INV_TYPE = '" & INV_TYPE & "' AND " & INV_NUM_column & " = '" & INV_NO & "'").Length = 0 Then
                                Fill_Records("ARTPYMTD", New String() {INV_TYPE, INV_NO}, False)
                            End If
                        End If
                    Next
                    grd.DisplayLayout.Bands(1).Hidden = False
                    grd.Rows.ExpandAll(True)
                    Exit Sub
                End If

                If grd.Name = "grdARTOPEN1" Then
                    INV_NO = grd.ActiveRow.Cells(INV_NUM_column).Text
                    INV_TYPE = grd.ActiveRow.Cells("INV_TYPE").Text
                    If dst.Tables("ARTPYMTD").Select("INV_TYPE = '" & INV_TYPE & "' AND " & INV_NUM_column & " = '" & INV_NO & "'").Length = 0 Then
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
                           & " exceeds the maximum size (1000). Please reduce the number of selected items.",
                            MsgBoxStyle.OkOnly, "Cannot Proceed")

                Else
                    Dim F As New ASFCONV2(Me, "ARTCUST1", HFs("CUST_CODE"), logText)
                    F.EntryMode = "N"
                    F.ShowDialog()
                    If F.result = "U" Then
                        dst.Tables("TATCONV1").Rows.Add(F.rowTATCONV1.ItemArray)
                        Update_Record_TDA("TATCONV1")
                        If collections_mode Then
                            If tabCollections.Tabs("Conversation Log").TabPage.Controls.Count > 0 Then
                                Dim F2 As ASFCONV1 = DirectCast(tabCollections.Tabs("Conversation Log").TabPage.Controls(0), ASFCONV1)
                                F2.tblTATCONV1.Rows.Clear()
                                F2.tblTATCONV1.Merge(dst.Tables("TATCONV1"))
                            End If
                        End If
                    End If
                    F.Dispose()

                End If

            Case "Total Balance"

                If grd.Selected.Rows.Count <> 0 Then
                    Dim invTotal As Decimal = 0
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        invTotal = invTotal + grow.Cells("INV_BALANCE").Value
                    Next
                    MsgBox("Total Balance: " & String.Format("{0:c}", invTotal),
                       MsgBoxStyle.OkOnly,
                       String.Format("Total Balance for {0} Item(s) Selected", grd.Selected.Rows.Count))
                Else
                    MsgBox("You must select the rows that you want totaled",
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

            Case "Charge Credit Card"
                Charge_Credit_Card(grd.ActiveRow.Cells("INV_NO").Value)

            Case "Show Raw EDI"

                If grd.Name = "grdARTPYMTY" And grdARTPYMTY.ActiveRow IsNot Nothing AndAlso grdARTPYMTY.ActiveRow.IsDataRow Then
                    Dim EDI_DOC_SEQ_NO As String = grdARTPYMTY.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                    If EDI_DOC_SEQ_NO = "" Then
                        MsgBox("This is not an EDI 820 Payment,", MsgBoxStyle.OkOnly, "Cannot Show Raw EDI")
                        Exit Sub
                    End If
                    Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, , "820")
                    Using frm As New ASFTEXT1
                        frm.t = RAW_EDI
                        frm.Text = "Raw EDI for " & grdARTPYMTY.ActiveRow.Cells("CUST_CODE").Value & " Check No " & grdARTPYMTY.ActiveRow.Cells("CUST_PYMT_REF_NO").Value
                        frm.ShowDialog()
                    End Using
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
                                    & vbCrLf & vbCrLf & "An Audit Trail Record will be Recorded.",
                                    MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If

                        BeginTrans()

                        ASCMAIN1.sql = "Update ARTOPEN1 Set TERM_CODE = :PARM1, INV_DUE_DATE = :PARM2 where INV_TYPE = :PARM3 and INV_NUM = :PARM4 and CUST_CODE = :PARM5"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VDVVV", New Object() {TERM_CODE, INV_DUE_DATE, INV_TYPE, INV_NUM, CUST_CODE})

                        TAC.TACMAIN1.Record_Event("SOTINVH1", CUST_CODE & ":" & INV_TYPE & ":" & INV_NUM,
                                                  DATETIME_STAMP, ASCMAIN1.USER_ID, "CHGTERMS",
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

    Private Function GetInvoiceList(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid,
                                    ByVal OwningMenuKey As String,
                                    ByVal INV_NO As String) As String

        Dim INV_TYPE As String = String.Empty
        Dim INV_NOs As String = String.Empty

        If grd.ActiveRow IsNot Nothing Then
            If Not grd.ActiveRow.Selected Then
                grd.Selected.Rows.Clear()
                grd.ActiveRow.Selected = True
            End If
        End If

        If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.Selected Then
            Return String.Empty
        End If

        If grd.Selected.Rows.Count = 0 Then
            Return String.Empty
        End If

        For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
            INV_TYPE = grow.Cells("INV_TYPE").Value & String.Empty
            If (ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA") And INV_TYPE = "R" Then
                INV_TYPE = "C"
            End If

            If OwningMenuKey = "grdARTOPEN1" Then
                INV_NO = grow.Cells(INV_NUM_column).Text
            Else
                INV_NO = grow.Cells("INV_NO").Text
            End If

            If INV_NO <> "" AndAlso ",I,C,R,".Contains(INV_TYPE) Then
                INV_NOs &= ", " & "('" & INV_TYPE & "', '" & INV_NO & "')"
            End If

        Next

        If INV_NOs.Length > 2 Then
            INV_NOs = Mid(INV_NOs, 2)
        End If

        Return INV_NOs

    End Function

    Sub Print_ProForma(ByVal ORDR_QTY_field As String,
                       ByVal lstOrderNos As List(Of String),
                       ByVal ExportInfo As Boolean)
        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Preparing Invoice")
            Dim pfComment As String = ""
            Dim REPORT_NAME As String = "SORINVP1"

            If Not REPORTS.ContainsKey(REPORT_NAME) Then
                REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
                REPORTS(REPORT_NAME).Prepare_dst(False, "")
            End If

            Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
            If RPT = "" Then
                RPT = REPORT_NAME
            End If

            Dim proFormaQuery As New List(Of String)
            Dim salesOrders As Boolean = False
            Dim invoices As Boolean = False
            Dim INV_TYPE_requested As String = String.Empty

            Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTORDR1 WHERE ORDR_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))", "SOTORDR1", "V", {String.Join(",", lstOrderNos.ToArray)})
            For Each dr As DataRow In tbl.Select("")
                Select Case dr.Item("ORDR_STATUS") & String.Empty
                    Case "F"
                        proFormaQuery.Add($"(SOTINVH1.ORDR_NO = '{dr.Item("ORDR_NO")}')")
                        invoices = True
                    Case Else
                        proFormaQuery.Add($"(SOTORDR1.ORDR_NO = '{dr.Item("ORDR_NO")}')")
                        salesOrders = True
                End Select
            Next

            If invoices AndAlso salesOrders Then
                INV_TYPE_requested = "B"
            ElseIf salesOrders Then
                INV_TYPE_requested = "O"
            Else
                ' Nothing at this time
            End If

            Dim rptQuery As String = " AND (" & String.Join(" OR ", proFormaQuery.ToArray) & ")"

            If ASCMAIN1.CLIENT = "VAN" Then
                REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {rptQuery, "1", pfComment})
            Else
                REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {rptQuery, "1", INV_TYPE_requested, ORDR_QTY_field})
            End If

            With REPORTS(REPORT_NAME).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "0")
                .CR_params.Add("EXPORT_INFO", IIf(ExportInfo, "1", "0"))
                .Generate_Report(RPT, "Sales Invoice", , True, , , , , False)
                .Print_Report_End()
            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Print ProForma", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try

    End Sub



#End Region

#Region "Popup Menu Functions"

    Function Create_Invoice(ByVal INV_NO As String) As String
        Me.Cursor = Cursors.WaitCursor

        ASCMAIN1.Progress("Now Preparing Invoice for Printing")

        Dim REPORT_NAME As String = "SORINVP1"
        Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If RPT = "" Then RPT = REPORT_NAME

        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        Dim sql As String = " and (SOTINVH1.INV_TYPE, SOTINVH1.INV_NO) in (" & INV_NO & ")"
        Dim tempFileName As String = "INV" & DateTime.Now.ToString("yyyyMMddHHmmss")

        ASCMAIN1.sql = "Select INV_NO from SOTINVH1 where INV_NO_CONS is Not Null " & sql
        Dim row() As DataRow = ASCDATA1.GetDataTable.Select
        Dim CONS_INV As String = "0"
        If row.Length <> 0 Then
            If MsgBox("Show Consolidated Invoices as Consolidated?", MsgBoxStyle.YesNo,
                      "Option to Print Consolidated Invoice") = MsgBoxResult.Yes Then
                CONS_INV = "1"
            End If
        End If

        Dim EXPORT_INFO As String = "0"
        If ASCMAIN1.CLIENT = "RGI" Then
            If MsgBox("Include Export Info?", MsgBoxStyle.YesNo,
                      "Option to Include Export Info") = MsgBoxResult.Yes Then
                EXPORT_INFO = "1"
            End If
        End If

        REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {sql, "", "", "", CONS_INV})
        Dim FILENAME As String = ""
        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()
            If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjzZ" Then
                '.CR_params.Add("SUBT", "")
                '.CR_params.Add("CONS_INV", CONS_INV)
                '.Generate_Report(REPORT_NAME, , , , , , , , False)
                'FILENAME = "" ' .F.REPORT_FILENAMES(REPORT_NO)
                '.Print_Report_End(True, False, , 1, "192.168.135.77:9100")
            Else
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", CONS_INV)
                '.CR_params.Add("EXPORT_INFO", "0")
                .CR_params.Add("EXPORT_INFO", EXPORT_INFO)

                Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
                FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
                .Print_Report_End(, True)
            End If
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return FILENAME
    End Function

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

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            End Using

        Catch ex As Exception
            MessageBox.Show($"Processing error: {ex.Message}", "Process Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

#End Region

#Region "grdARTPYMTY"

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

            'If rowARTCUST1.Item("COLLECTOR_CODE") & "" <> "" Then
            '    .Rows.Add(New Object() {"ACCTG", "Collector", rowARTCUST1.Item("COLLECTOR_CODE") & "", LookUp("ARTCOLL1", rowARTCUST1.Item("COLLECTOR_CODE") & "", True).Item("COLLECTOR_NAME")})
            'End If
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
        TOTAL_DUE_INV = Val(rowOPENAR.Item("OPEN_I") & "")

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

    Sub Setup_ARTPYMTY()

    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
        Me.SetControlPanel()
    End Sub

    Sub Load_ARTPYMTY(Optional ByVal INV_NO As String = "", Optional ByVal CUST_CODE_P As String = "",
                      Optional ByVal PYMT_BATCH_NO As String = "")

        Call ASCMAIN1.Progress("Now Loading Payment History")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Dim CUST_CODE_PYMT As String = ""

        grdARTPYMTX.Visible = False
        If CUST_CODE_P <> "" Then
            CUST_CODE_PYMT = CUST_CODE_P
        Else
            CUST_CODE_PYMT = HFs("CUST_CODE")
        End If
        If INV_NO <> "" And PYMT_BATCH_NO = "" Then
            ASCMAIN1.sql = "Select ARTPYMT2.*, ARTPYMT1.PYMT_BATCH_DATE" _
            & " from ARTPYMT2,ARTPYMT1 " _
            & " where (ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO) in " _
            & " (Select PYMT_BATCH_NO, PYMT_BATCH_LNO from ARTPYMT3 " _
            & "  where INV_NUM = '" & txtINV_NO_PYMT.Text & "'" & IIf(txtINV_NO_PYMT.Tag <> "", " and INV_TYPE = '" & txtINV_NO_PYMT.Tag & "'", "") & ")" _
            & " and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
            & " and ARTPYMT2.CUST_CODE = '" & CUST_CODE_PYMT & "'"

            Fill_Records("ARTPYMTY", "", True, ASCMAIN1.sql)
            grdARTPYMTY.Text = "Payment Applications involving Invoice No " & INV_NO
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

            Fill_Records("ARTPYMTY", "", True, ASCMAIN1.sql)
            grdARTPYMTY.Text = "Payment Applications involving Batch " & PYMT_BATCH_NO
            chkPaymentsShowZero.Checked = True

        Else
            Fill_Records("ARTPYMTY", New String() {cbeYP_PYMTs.Value})
            grdARTPYMTY.Text = "Payments Received from Customer " & CUST_CODE_PYMT & " since " & cbeYP_PYMTs.Text
        End If

        Sort_grdColumns(grdARTPYMTY, "PYMT_BATCH_NO,PYMT_BATCH_LNO".ToLower)

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
        & " from ARTOPENX ARTOPEN1," & ARTCUST0 & " ARTCUST0, SOTSHIP1, SOTINVH1 " _
        & " where ARTOPEN1.CUST_CODE = ARTCUST0.CUST_CODE" _
        & " and SOTINVH1.INV_TYPE (+) = 'I' and SOTINVH1.INV_NO (+) = ARTOPEN1.INV_NUM" & vbCrLf _
        & " and SOTSHIP1.SHIP_BOL_NO (+) = SOTINVH1.SHIP_BOL_NO" & vbCrLf _
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

    Sub Retrieve_Paid_Items(ByVal numDays As Integer)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Retrieving Paid Items")

        ASCMAIN1.sql = sqlARTOPEN1 _
        & " from ARTOPENX ARTOPEN1," & ARTCUST0 & " ARTCUST0, SOTSHIP1, SOTINVH1 " _
        & " where ARTOPEN1.CUST_CODE = ARTCUST0.CUST_CODE" _
        & " and SOTINVH1.INV_TYPE (+) = 'I' and SOTINVH1.INV_NO (+) = ARTOPEN1.INV_NUM" & vbCrLf _
        & " and SOTSHIP1.SHIP_BOL_NO (+) = SOTINVH1.SHIP_BOL_NO" & vbCrLf _
        & "   and ARTOPEN1.DATE_PAID > TRUNC(SYSDATE) - " & CStr(numDays) _
        & "   and (ARTOPEN1.INV_TYPE = 'I' or (ARTOPEN1.INV_TYPE = 'B' and ARTOPEN1.REASON_CODE in (Select REASON_CODE from ARTREAS1 where SHIPPING_VIOLATION = '1')))"
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

    Function Print_Hard_Copy(Optional aged_ar As Boolean = False, Optional email As Boolean = False) As String

        Dim RPT_TITLE As String = "Customer Inquiry"
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

        Dim rowARTSTMT1 As DataRow = dst.Tables("ARTSTMT1").Rows.Find("999999")
        Dim rowARTSTMT1_save As DataRow = dst.Tables("ARTSTMT1").NewRow
        rowARTSTMT1_save.ItemArray = rowARTSTMT1.ItemArray

        Dim D(4) As Decimal
        For I As Integer = 1 To 4
            rowARTSTMT1.Item("DUE_" & CStr(I)) = Val(dst.Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "DUE_BUCKET = " & CStr(I) & " AND (INV_TYPE = 'I' OR INV_TYPE = 'B' OR INV_TYPE = 'D')") & "")
        Next

        If aged_ar Then
            CR_params.Add("SKIP_MIDDLE", "1")
            If ASCMAIN1.CLIENT = "VAN" Then
                CR_params.Add("SKIP_CREDITS", "1")

            Else
                CR_params.Add("SKIP_CREDITS", "0")
            End If
        Else
            CR_params.Add("SKIP_MIDDLE", "0")
            CR_params.Add("SKIP_CREDITS", "0")
        End If

        If email Then
            FILENAME = ASCMAIN1.Next_Control_No("ARFCINQ1.STMT")
            Dim REPORT_NO As String = Generate_Report(reportFile, IIf(aged_ar, "Aged Accounts Receivable", RPT_TITLE), "As of " & Format(Now, "MM/dd/yyyy"), "", "PDF", FILENAME, False)
            Print_Report_End(, True)
        Else
            Generate_Report(reportFile, IIf(aged_ar, "Aged Accounts Receivable", RPT_TITLE))
            Print_Report_End()
        End If

        rowARTSTMT1.Item("DUE_1") = rowARTSTMT1_save.Item("DUE_1")
        rowARTSTMT1.Item("DUE_2") = rowARTSTMT1_save.Item("DUE_2")
        rowARTSTMT1.Item("DUE_3") = rowARTSTMT1_save.Item("DUE_3")
        rowARTSTMT1.Item("DUE_4") = rowARTSTMT1_save.Item("DUE_4")
        grdARTSTMT1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        'Show_Document(FILENAME)
        Return FILENAME

    End Function
    Function Print_Hard_Copy_Statement(Optional aged_ar As Boolean = False, Optional email As Boolean = False) As String

        Dim RPT_TITLE As String = "Customer Statement Printing"
        Dim reportFile As String = "ARRSTMTR"
        Dim FILENAME As String = ""

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing " & RPT_TITLE)

        Print_Report_Begin()
        CR_params.Add("SUBT", "")

        If email Then
            FILENAME = ASCMAIN1.Next_Control_No("ARFCINQ1.STMT")
            Dim REPORT_NO As String = Generate_Report(reportFile, IIf(aged_ar, "Customer Statement Printing", RPT_TITLE), "As of " & Format(Now, "MM/dd/yyyy"), "", "PDF", FILENAME, False)
            Print_Report_End(, True)
        Else
            Generate_Report(reportFile, IIf(aged_ar, "Customer Statement Printing", RPT_TITLE))
            Print_Report_End()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        'Show_Document(FILENAME)
        Return FILENAME

    End Function
    Sub Load_SOTORDR1()

        Me.Cursor = Cursors.WaitCursor
        Dim CUST_CODE As String = HFs("CUST_CODE")
        Dim CUST_NAME As String = HFs("CUST_NAME")
        splSOTORDR0.Panel2Collapsed = False

        dst.Tables("SOTORDR2").Rows.Clear()
        dst.Tables("SOTORDR1").Rows.Clear()
        dst.Tables("SOTORDR0").Rows.Clear()

        Dim subQuery As String = String.Empty

        If chkOpenOrdersOnly.Checked Then
            ASCMAIN1.Progress("Now Loading Open Order Data")
            grdSOTORDR0.Text = "Open Orders for " & CUST_CODE & ":" & CUST_NAME & " (by Order)"

            subQuery = "Select SOTORDR0.* from SOTORDR0 where CUST_CODE = '" & CUST_CODE & "' and (ORDR_CNT_OPEN <> 0 or ORDR_CNT_PICK <> 0)"

            ASCMAIN1.sql = "Select SOTORDR0.* from SOTORDR0 where CUST_CODE = '" & CUST_CODE & "' and (ORDR_CNT_OPEN <> 0 or ORDR_CNT_PICK <> 0)"
            Fill_Records("SOTORDR0", "", , ASCMAIN1.sql)

        Else
            ASCMAIN1.Progress("Now Loading Sales Order History")
            grdSOTORDR0.Text = "Order History for " & CUST_CODE & ":" & CUST_NAME & " since " & Format(dteOrderHistory.DateTime, "MM/dd/yyyy")

            subQuery = "Select SOTORDR0.* from SOTORDR0 where CUST_CODE = '" & CUST_CODE & "' and ORDR_DATE > '" & Format(dteOrderHistory.DateTime, "dd-MMM-yyyy") & "'"

            ASCMAIN1.sql = "Select SOTORDR0.* from SOTORDR0 where CUST_CODE = '" & CUST_CODE & "' and ORDR_DATE > '" & Format(dteOrderHistory.DateTime, "dd-MMM-yyyy") & "'"
            Fill_Records("SOTORDR0", "", , ASCMAIN1.sql)
        End If

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select()
                Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO").ToString & String.Empty
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine($"SELECT SUM(NVL(ORDR_AMT_ALLO_CUR,0)) ORDR_AMT_ALLO_CUR FROM SOTORDRS WHERE ORDR_GROUP_NO = '{ORDR_GROUP_NO}'")
                ASCMAIN1.sql = SQLS.ToString()
                Dim ORDR_AMT_ALLO_CUR As Int32 = Val(ASCDATA1.GetDataValue)
                rowSOTORDR0.Item("ORDR_AMT_ALLO_CUR") = ORDR_AMT_ALLO_CUR
            Next
        End If

        ASCMAIN1.sql = "Select SOTORDR1.* from SOTORDR1 where ORDR_GROUP_NO in (" & Replace(subQuery, ".*", ".ORDR_GROUP_NO") & ")"
        Fill_Records("SOTORDR1", "", , ASCMAIN1.sql)
        Sort_grdColumns(grdSOTORDR0, "ORDR_GROUP_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Orders_Filter()
        ' Setup_SOTORDR2()
    End Sub

    Sub Load_SOTINVH1()

        Me.Cursor = Cursors.WaitCursor
        Dim CUST_CODE As String = HFs("CUST_CODE")
        Dim CUST_NAME As String = HFs("CUST_NAME")
        splSOTINVH1.Panel2Collapsed = False

        ASCMAIN1.Progress("Now Loading Sales Invoice History")
        grdSOTORDR0.Text = "Invoice History for " & CUST_CODE & ":" & CUST_NAME & " since " & Format(dteInvoiceHistory.DateTime, "MM/dd/yyyy")

        ASCMAIN1.sql = "Select SOTINVH1.*, SOTSHIP1.BILL_OF_LADING_NO" _
            & " from SOTINVH1,SOTSHIP1 where SOTINVH1.CUST_CODE = '" & CUST_CODE & "'" _
            & "  and SOTSHIP1.SHIP_BOL_NO (+) = SOTINVH1.SHIP_BOL_NO" _
            & "  and SOTINVH1.INV_DATE > '" & Format(dteInvoiceHistory.DateTime, "dd-MMM-yyyy") & "'"
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
        UltraExplorerBar1.Groups("Open Balances").Visible = Not collections_mode And ScreenMode And (tabMain.SelectedTab.Key = "General" Or tabMain.SelectedTab.Key = "Info" Or (tabMain.SelectedTab.Key = "Accts Rec" And UltraTabControl5.SelectedTab.Key = "Open AR && Aging History"))
        UltraExplorerBar1.Groups("Open AR by Type").Visible = Not collections_mode And ScreenMode And (tabMain.SelectedTab.Key = "General" Or tabMain.SelectedTab.Key = "Info" Or (tabMain.SelectedTab.Key = "Accts Rec" And UltraTabControl5.SelectedTab.Key = "Open AR && Aging History"))
        'UltraExplorerBar1.Groups("Aged Open AR").Visible = ScreenMode And (UltraTabControl1.SelectedTab.Key = "General" Or UltraTabControl1.SelectedTab.Key = "Info" Or (UltraTabControl1.SelectedTab.Key = "Accts Rec" And UltraTabControl5.SelectedTab.Key = "Open AR && Aging History"))
        UltraExplorerBar1.Groups("Freight").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Freight")
        UltraExplorerBar1.Groups("Summary").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Summary")
        UltraExplorerBar1.Groups("Update Inv Due Date").Visible = Not collections_mode And ASCMAIN1.CLIENT = "RGI" And InStr(ASCMAIN1.USER_SECURITY_CODEs, "CX") And ScreenMode And (tabMain.SelectedTab.Key = "Accts Rec")
    End Sub

    Sub Setup_tabMain()

        If tabMain.SelectedTab Is Nothing Then Exit Sub
        'UltraExplorerBar1.Groups("Customer Log").Visible = _
        '(tabMain.SelectedTab.Key = "Log")

        Select Case tabMain.SelectedTab.Key
            Case "Name && Address"
            Case "Info"
            Case "Accts Rec"
                Call Setup_Tab_AR()
            Case "Orders"
                Setup_tabOrders()
            Case "Sales"
                If Not order_guide_loaded Then
                    Load_Order_Guide()
                End If

                If grdARTSLSMA.Tag = "" Then
                    Fill_Records("ARTSLSMA", HFs("CUST_CODE"))
                    Sort_grdColumns(grdARTSLSMA, "OPS_YYYYPP")
                    grdARTSLSMA.Tag = HFs("CUST_CODE")
                End If
                If grdSATCUSTS.Tag = "*" Then
                    SetUp12Month()
                End If
            Case "Outbound"
                If txtOBTerms.Text & String.Empty = "" Then
                    fillOBData()
                End If
        End Select

    End Sub

    Private Sub grdARTPYMTY_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMTY.AfterRowActivate
        Load_ARTPYMTX()
    End Sub

    Sub Load_ARTPYMTX()

        With grdARTPYMTY
            If .ActiveRow Is Nothing OrElse .ActiveRow.IsGroupByRow OrElse Not .ActiveRow.IsDataRow Then
                grdARTPYMTX.Visible = False
            Else
                grdARTPYMTX.Visible = True

                Dim PYMT_BATCH_NO As String = .ActiveRow.Cells("PYMT_BATCH_NO").Text
                Dim PYMT_BATCH_LNO As Integer = Val(.ActiveRow.Cells("PYMT_BATCH_LNO").Text)
                grdARTPYMTX.Text = "Payment Details for Batch-Lno " & PYMT_BATCH_NO & "-" & CStr(PYMT_BATCH_LNO) & "; " & .ActiveRow.Cells("CUST_CODE").Text & ":" & .ActiveRow.Cells("CUST_NAME").Text
                Call Fill_Records("ARTPYMTX", New Object() {PYMT_BATCH_NO, PYMT_BATCH_LNO})
                Call Sort_grdColumns(grdARTPYMTX, "PYMT_BATCH_NO,PYMT_BATCH_LNO,PYMT_BATCH_ILNO")
                grdARTPYMTX.DisplayLayout.Bands(0).SummaryFooterCaption = "Totals for Payment Ref " & .ActiveRow.Cells("CUST_PYMT_REF_NO").Text

                Dim DTPINV As Decimal = Val(dst.Tables("ARTPYMTX").Compute("SUM(DTPINV)", "") & "")
                Dim DTPWGT As Decimal = Val(dst.Tables("ARTPYMTX").Compute("SUM(DTPWGT)", "") & "")
                Dim DPDINV As Decimal = Val(dst.Tables("ARTPYMTX").Compute("SUM(DPDINV)", "") & "")
                Dim DPDWGT As Decimal = Val(dst.Tables("ARTPYMTX").Compute("SUM(DPDWGT)", "") & "")

                Dim DTP As Decimal = 0
                If DTPINV <> 0 Then DTP = DTPWGT / DTPINV
                Dim DPD As Decimal = 0
                If DPDINV <> 0 Then DPD = DPDWGT / DPDINV
                grdARTPYMTX.Text &= "; Avg DTP = " & Format(DTP, "##0.0")
                grdARTPYMTX.Text &= "; Avg DPD = " & Format(DPD, "##0.0")

            End If
        End With
    End Sub

    Private Sub chkPaymentsShowZero_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPaymentsShowZero.CheckedChanged
        Payments_Show_Zero()
    End Sub

    Sub Payments_Show_Zero()
        If chkPaymentsShowZero.Checked Then
            DirectCast(grdARTPYMTY.DataSource, DataTable).DefaultView.RowFilter = ""
        Else
            DirectCast(grdARTPYMTY.DataSource, DataTable).DefaultView.RowFilter = "CUST_PYMT_AMT <> 0"
        End If
        If grdARTPYMTY.Rows.Count <> 0 Then
            grdARTPYMTY.ActiveRow = grdARTPYMTY.Rows(0)
        Else
            grdARTPYMTX.Visible = False
        End If
    End Sub

    Sub Customer_Activity()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading List of Active Customers")
        Fill_Records("ARTCUST6")
        Sort_grdColumns(grdARTCUST6, "CUST_CODE")
        TAC.ARCMAIN1.Get_Aging_Data(ROWs.Item("ARTPARM1"), Now.Date)
        ASCMAIN1.sql = "Select CUST_CODE" & TAC.ARCMAIN1.AGED_TOTALS & " from ARTOPEN1 group by CUST_CODE"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim rowARTCUST6 As DataRow = dst.Tables("ARTCUST6").Rows.Find(CUST_CODE)
            If rowARTCUST6 IsNot Nothing Then
                rowARTCUST6.Item("AGE_1") = row.Item("AGE_1")
                rowARTCUST6.Item("AGE_2") = row.Item("AGE_2")
                rowARTCUST6.Item("AGE_3") = row.Item("AGE_3")
                rowARTCUST6.Item("AGE_4") = row.Item("AGE_4")
            End If
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        grdARTCUST6.Tag = "*"
        grdARTCUST6.Visible = True
        grdARTOPENB.Tag = ""
        ' grdARTOPENB.Visible = False
        tabChargebacks.Visible = False
        grdARTCUSTT_FUPS.Visible = False
        chkALLFU.Visible = False
    End Sub

    Sub Chargebacks_Summary()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Chargebacks Summary")

        Fill_Records("ARTOPENB")
        Sort_grdColumns(grdARTOPENB, "CUST_CODE")

        ASCMAIN1.sql = sqlARTOPEN1 _
            & " from ARTOPEN1, SOTSHIP1, SOTINVH1 where ARTOPEN1.INV_TYPE = 'B'" _
            & " and SOTINVH1.INV_TYPE (+) = 'I' and SOTINVH1.INV_NO (+) = ARTOPEN1.INV_NUM" & vbCrLf _
            & " and SOTSHIP1.SHIP_BOL_NO (+) = SOTINVH1.SHIP_BOL_NO" & vbCrLf _
            & IIf(collections_mode, vbCrLf & " and ARTOPEN1.REASON_CODE in (Select REASON_CODE from ARTREAS1 where SHIPPING_VIOLATION = '1')", "")

        dst.Tables("ARTPYMTD").Rows.Clear()
        Fill_Records("ARTOPEN1", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdARTOPEN1, "CUST_CODE,INV_NUM")
        '  grdARTOPEN1.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)

        ASCMAIN1.sql = "Select ARTPYMT2.*, ARTPYMT1.PYMT_BATCH_DATE" & vbCrLf _
            & " from ARTPYMT2,ARTPYMT1 " & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'" & vbCrLf _
            & "   and ARTPYMT1.PYMT_SOURCE = 'CWO'"
        Fill_Records("ARTPYMTY", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdARTPYMTY, "PYMT_BATCH_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        grdARTCUST6.Tag = ""
        grdARTCUST6.Visible = False
        grdARTOPENB.Tag = "*"
        ' grdARTOPENB.Visible = True
        tabChargebacks.Visible = True
        grdARTCUSTT_FUPS.Visible = False
        chkALLFU.Visible = False
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

                If ASCMAIN1.CLIENT = "RGI" Then
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
                    & ", CUST_PD_GRACE_PCT = :PARM17" _
                    & ", CUST_CREDIT_RELEASE = :PARM18" _
                    & ", CUST_SALES_HOLD = :PARM19" _
                    & " where CUST_CODE = :PARM20"

                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NDDVVVDVVVVDNDVNNVVV", New Object() _
                        { .Item("CUST_CREDIT_LIMIT"),
                        .Item("CUST_CRED_LIMIT_EST"),
                        .Item("CUST_CRED_LIMIT_REV"),
                        .Item("CUST_CREDIT_HOLD"),
                        .Item("CUST_FACTOR_IND"),
                        .Item("CUST_CREDIT_SCORE"),
                        .Item("CUST_CREDIT_SCORE_DATE"),
                        .Item("TERM_CODE"),
                        .Item("CUST_CREDIT_LIMIT_APPR_BY"),
                        .Item("CUST_CREDIT_LIMIT_NOTES"),
                        .Item("CUST_CREDIT_RATING"),
                        .Item("CUST_CREDIT_RATING_DATE"),
                        .Item("CUST_INS_AMT"),
                        .Item("CUST_INS_DATE"),
                        .Item("CUST_DUNS"),
                        .Item("CUST_PD_GRACE_DAYS"),
                        .Item("CUST_PD_GRACE_PCT"),
                        .Item("CUST_CREDIT_RELEASE"),
                        .Item("CUST_SALES_HOLD"),
                        HFs("CUST_CODE")})
                Else
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
                        & ", CUST_PD_GRACE_PCT = :PARM17" _
                        & ", CUST_CREDIT_RELEASE = :PARM18" _
                        & " where CUST_CODE = :PARM19"

                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NDDVVVDVVVVDNDVNNVV", New Object() _
                        { .Item("CUST_CREDIT_LIMIT"),
                        .Item("CUST_CRED_LIMIT_EST"),
                        .Item("CUST_CRED_LIMIT_REV"),
                        .Item("CUST_CREDIT_HOLD"),
                        .Item("CUST_FACTOR_IND"),
                        .Item("CUST_CREDIT_SCORE"),
                        .Item("CUST_CREDIT_SCORE_DATE"),
                        .Item("TERM_CODE"),
                        .Item("CUST_CREDIT_LIMIT_APPR_BY"),
                        .Item("CUST_CREDIT_LIMIT_NOTES"),
                        .Item("CUST_CREDIT_RATING"),
                        .Item("CUST_CREDIT_RATING_DATE"),
                        .Item("CUST_INS_AMT"),
                        .Item("CUST_INS_DATE"),
                        .Item("CUST_DUNS"),
                        .Item("CUST_PD_GRACE_DAYS"),
                        .Item("CUST_PD_GRACE_PCT"),
                        .Item("CUST_CREDIT_RELEASE"),
                        HFs("CUST_CODE")})
                End If

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
                rowARTCUST5.Item("CUST_PD_GRACE_DAYS") = .Item("CUST_PD_GRACE_DAYS")
                rowARTCUST5.Item("CUST_PD_GRACE_PCT") = .Item("CUST_PD_GRACE_PCT")
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
            ASCMAIN1.sql = "Select " & vbCrLf _
                & "  SUM (NVL(SOTORDR2.ORDR_QTY_OPEN,0)*NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_OPEN" & vbCrLf _
                & ", SUM (NVL(SOTORDR2.ORDR_QTY_PICK,0)*NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_PICK" & vbCrLf _
                & " from SOTORDR1,SOTORDR2 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO and SOTORDR1.CUST_CODE = '" & HFs("CUST_CODE") & "' AND SOTORDR1.ORDR_STATUS IN ('O','P')"
            Dim rowOP As DataRow = ASCDATA1.GetDataRow
            Dim OPEN_ORDERS_RELEASED As Decimal = Val(rowOP.Item("ORDR_PICK") & "") ' 0 ' Val(ASCDATA1.GetDataValue)
            Dim OPEN_ORDERS_PENDING As Decimal = Val(rowOP.Item("ORDR_OPEN") & "") ' 0 ' Val(ASCDATA1.GetDataValue)

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
            If ASCMAIN1.CLIENT = "INT" Then
                CREDIT_AVAIL = CUST_CREDIT_LIMIT - TOTAL_DUE_INV - OPEN_ORDERS_RELEASED
            End If
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
        If Not ScreenMode Then
            If grdARTCUSTT_FUPS.ActiveRow IsNot Nothing AndAlso grdARTCUSTT_FUPS.ActiveRow.IsDataRow Then
                Absx1.txtFor("CUST_CODE").Text = grdARTCUSTT_FUPS.ActiveRow.Cells("TABLE_KEY").Text
                Click_Command("Select Customer")
            End If
        End If
    End Sub

    Sub Refresh_FollowUps()
        Fill_Records("ARTCUSTT_FUPS")
        grdARTCUSTT_FUPS.Visible = True
        chkALLFU.Visible = ASCMAIN1.CLIENT = "RGI" And (ASCMAIN1.USER_ID = "andy" Or ASCMAIN1.USER_ID = "wayne")
        grdARTCUST6.Visible = False
        ' grdARTOPENB.Visible = False
        tabChargebacks.Visible = False
        grdARTCUST6.Tag = ""
        grdARTOPENB.Tag = ""

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
                If grdARTPYMTY.Tag = "*" Then
                    Load_ARTPYMTY()
                    grdARTPYMTY.Tag = ""
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

    Private Sub grdARTPYMTY_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMTY.InitializeRow
        If e.Row.Cells("CCPA_NO_CREDIT").Text <> "" Then
            e.Row.Cells("STATUS").Value = "CC Credited"
        ElseIf e.Row.Cells("PYMT_STATUS").Text <> "2" Then
            e.Row.Cells("STATUS").Value = "Un-Applied"
            e.Row.Cells("STATUS").Appearance.BackColor = Drawing.Color.Yellow
        ElseIf e.Row.Cells("PYMT_DELETED").Value & "" = "1" Then
            e.Row.Cells("STATUS").Value = "Payment Deleted"
            e.Row.Cells("STATUS").Appearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("PYMT_STATUS").Value & "" = "1" Then
            e.Row.Cells("STATUS").Value = "Not Applied"
        ElseIf e.Row.Cells("PYMT_STATUS").Value & "" = "2" Then
            e.Row.Cells("STATUS").Value = "Applied"
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
        Load_ARTPYMTY()
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

        Dim PYMT_BATCH_NO_found As String = ""

        If txtINV_NO_PYMT.Tag = "" Or txtINV_NO_PYMT.Tag = "I" Or txtINV_NO_PYMT.Tag = "C" Then
            Dim INV_NO As String = txtINV_NO_PYMT.Text
            INV_NO = INV_NO.PadLeft(10, "0")
            ASCMAIN1.sql = "Select * from SOTINVH1 where INV_NO = '" & INV_NO & "'"
            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow
            If rowSOTINVH1 Is Nothing Then
                ASCMAIN1.sql = "Select * from ARTPYMT3 where INV_NUM = '" & INV_NO & "'"
                Dim row2 As DataRow = ASCDATA1.GetDataRow
                If row2 Is Nothing Then
                    ASCMAIN1.sql = "Select * from ARTPYMT5 where CHARGEBACK_NO = '" & INV_NO & "'"
                    row2 = ASCDATA1.GetDataRow
                    If row2 IsNot Nothing Then
                        PYMT_BATCH_NO_found = row2.Item("PYMT_BATCH_NO")
                    End If
                End If
                If row2 Is Nothing Then
                    MsgBox("No Record of Invoice/Memo " & txtINV_NO_PYMT.Text, MsgBoxStyle.OkOnly, "Cannot Find Record")
                    Exit Sub
                End If
            Else
                txtINV_NO_PYMT.Text = INV_NO

                If rowSOTINVH1.Item("CUST_CODE") & "" <> HFs("CUST_CODE") Then
                    If rowSOTINVH1.Item("CUST_CODE") & "" <> CUST_CODE_PYMT Then
                        MsgBox("Invoice " & txtINV_NO_PYMT.Text & " belongs to another customer (" & rowSOTINVH1.Item("CUST_CODE") & ")", MsgBoxStyle.OkOnly, "Wrong Customer")
                        txtINV_NO_PYMT.Text = ""
                        Exit Sub
                    End If
                End If
            End If
        End If
        CUST_CODE_PYMT = HFs("CUST_CODE")
        Load_ARTPYMTY(txtINV_NO_PYMT.Text, CUST_CODE_PYMT, PYMT_BATCH_NO_found)
        Load_ARTPYMTX()
        If collections_mode Then
            tabMain.SelectedTab = tabMain.Tabs("Accts Rec")
        Else
            tabMain.SelectedTab = tabMain.Tabs("Accts Rec")
        End If
        UltraTabControl5.SelectedTab = UltraTabControl5.Tabs("Payment History")

    End Sub

    Private Sub grdARTPYMTX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMTX.InitializeRow
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

        If grdARTPYMTX.Selected.Rows.Count = 0 Then
            MessageBox.Show("There are no AR Batch detail Invoices selected.", "Print Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If MessageBox.Show("Do you want to print the " & grdARTPYMTX.Selected.Rows.Count & " selected documents?",
                            "Print Selected", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim invoices As String = String.Empty
        Dim INV_NO As String = String.Empty
        Dim rowSOTINVH1 As DataRow = Nothing

        For Each gridRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdARTPYMTX.Selected.Rows
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

    Private Sub grdARTOPEN1_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdARTOPEN1.AfterRowUpdate
        If chkEditNotes.Checked Then
            Update_Record_TDA("ARTOPEN1")
        End If
        If collections_write_off_started Then
            Display_Application_Totals()
        End If
    End Sub

    Private Sub grdARTOPEN1_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdARTOPEN1.ClickCellButton
        If Val(e.Cell.Row.Cells("WOFF").Value & "") = 0 Then
            e.Cell.Row.Cells("WOFF").Value = e.Cell.Row.Cells("INV_BALANCE").Value
        Else
            e.Cell.Row.Cells("WOFF").Value = 0
        End If

        e.Cell.Row.Update()
    End Sub

    Private Sub grdARTOPEN1_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdARTOPEN1.DoubleClickCell
        If e.Cell.Column.Key = "PYMT_BATCH_NO" Then
            Dim batch As String = e.Cell.Value
            Load_ARTPYMTY("", "", batch)
            grdARTPYMTY.Tag = ""
            UltraTabControl5.SelectedTab = UltraTabControl5.Tabs("Payment History")
        End If
    End Sub

    Private Sub grdARTOPEN1_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTOPEN1.DoubleClickRow
        If Not ScreenMode Then
            Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & ""
            Absx1.txtFor("CUST_CODE").Text = CUST_CODE
            Click_Command("Select Customer")
        End If
    End Sub

    Private Sub grdARTOPEN1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTOPEN1.InitializeRow
        If e.Row.Band.Index = 0 Then

            Select Case e.Row.Cells("INV_TYPE").Value & ""
                Case "C"
                    If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                        e.Row.Appearance.BackColor = Drawing.Color.LightGreen
                    Else
                        e.Row.Cells("INV_TYPE").Appearance.BackColor = Drawing.Color.LightGreen
                    End If

                Case "B"
                    If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                        e.Row.Appearance.BackColor = Drawing.Color.LightSteelBlue
                    Else
                        e.Row.Cells("INV_TYPE").Appearance.BackColor = Drawing.Color.LightSteelBlue
                    End If

            End Select

            If ScreenMode Then
                Dim CURR_CODE As String = e.Row.Cells("CURR_CODE").Value & ""
                If CURR_CODE <> "USD" Then
                    e.Row.Cells("CURR_CODE").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("INV_TOTAL_AMT").Appearance.ForeColor = Drawing.Color.Red
                End If

                Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & ""
                If CUST_CODE <> Absx1.txtFor("CUST_CODE").Text Then
                    e.Row.Cells("CUST_CODE").Appearance.ForeColor = Drawing.Color.Red
                End If

                Dim CUST_CODE_SO As String = e.Row.Cells("CUST_CODE_SO").Value & ""
                If CUST_CODE_SO <> Absx1.txtFor("CUST_CODE").Text Then
                    e.Row.Cells("CUST_CODE_SO").Appearance.ForeColor = Drawing.Color.Red
                End If

                Dim POST_CODE As String = e.Row.Cells("POST_CODE").Value & ""
                If POST_CODE = "REBCR" Then
                    e.Row.Cells("POST_CODE").Appearance.BackColor = Drawing.Color.Yellow
                ElseIf POST_CODE <> rowARTCUST1.Item("POST_CODE") & "" And POST_CODE <> "" Then
                    e.Row.Cells("POST_CODE").Appearance.ForeColor = Drawing.Color.Red
                End If
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

    Public Overrides Function CustomSummary_End(
    ByVal summarySettings As UltraWinGrid.SummarySettings,
    ByVal rows As UltraWinGrid.RowsCollection,
    ByVal CustomValue As Double,
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

    Sub grdSOTINVH1_FRT_WHSE_Calculate_Totals(
    ByVal rows As UltraWinGrid.RowsCollection,
    ByRef FRT_AMT As Decimal,
    ByRef UNITS As Decimal,
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

        If ASCMAIN1.CLIENT = "RGI" Then
            Fill_Records("ARTSTMTR", HFs("CUST_CODE"))
            Fill_Records("ARTCUSTA", HFs("CUST_CODE"))
        End If


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

            If MessageBox.Show("Do you want to make a Credit Card Charge in the amount of " & Format(inv_total_amount, "$#,##0.00"), "Charge Credit Card",
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

    Private Sub chkShowGP_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowGP.CheckedChanged

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

    Private Sub cmdFetchSummary_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetchSummary.Click
        Dim YP0 As String = cbeYPSFrom.Value
        Dim YP1 As String = cbeYPSTo.Value
        Dim P As Integer = ASCMAIN1.Period_Diff(YP0, YP1) + 1
        If P > 12 Then
            MsgBox("Period Range cannot span more than 12 Months", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        ElseIf P < 1 Then
            MsgBox("Period Range must be at least 1 month", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Summary Data")

        Setup_grdARTCSUMC()

        Dim SALES_DIVISION_CODE As String = ""
        If Not chkAllDivisions.Checked Then
            SALES_DIVISION_CODE = cbeSALES_DIVISION.Value
        End If

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        TAC.SOCMAIN1.Build_ARTCSUMC(Me, CUST_CODE, YP0, YP1, ARTCSUMC, , SALES_DIVISION_CODE)

        For Each C As String In New String() {"GRS", "NET", "DED"}
            Dim row2() As DataRow = dst.Tables("ARTCSUMA").Select("LINE_ABBR = '" & C & "'")
            Dim LINE As Integer = Val(row2(0).Item("LINE"))
            Dim T As Decimal = Val(row2(0).Item("AMT00") & "")
            For Each rowARTCSUMB As DataRow In dst.Tables("ARTCSUMB").Select("LINE = " & LINE)
                Dim V As Decimal = (rowARTCSUMB.Item("AMT00") & "")
                Dim PCT As Decimal = 0
                If T <> 0 Then PCT = V / T
                rowARTCSUMB.Item("PCT") = PCT
            Next
        Next



        Setup_grdARTCSUMB()
        splSummary.Visible = True
        grdARTCSUMA.Text = "Customer P&L for " & CUST_CODE & " for the " & CStr(P) & " Months " & grdARTCSUMA.DisplayLayout.Bands(0).Columns("AMT01").Header.Caption & " thru " & grdARTCSUMA.DisplayLayout.Bands(0).Columns("AMT" & Format(P, "00")).Header.Caption

        If SALES_DIVISION_CODE <> "" Then
            grdARTCSUMA.Text &= ", Division " & SALES_DIVISION_CODE & " - Sales & CGS only, Deductions shown are Total Customer"
        End If

        Dim row() As DataRow
        row = dst.Tables("ARTCSUMA").Select("LINE_ABBR = 'GRS'")
        Dim GRS As Decimal = Val(row(0).Item("AMT00") & "")
        row = dst.Tables("ARTCSUMA").Select("LINE_ABBR = 'NET'")
        Dim NET As Decimal = Val(row(0).Item("AMT00") & "")


        row = dst.Tables("ARTCSUMA").Select("LINE_ABBR = 'GP'")
        Dim GP As Decimal = Val(row(0).Item("AMT00") & "")
        row = dst.Tables("ARTCSUMA").Select("LINE_ABBR = 'DED'")
        Dim DED As Decimal = Val(row(0).Item("AMT00") & "")

        row = dst.Tables("ARTCSUMA").Select("LINE_ABBR = 'NP'")
        Dim NP As Decimal = Val(row(0).Item("AMT00") & "")

        Dim NPP As Decimal = 0
        If NET <> 0 Then NPP = System.Math.Abs(NP / NET) * System.Math.Sign(NP)
        If ASCMAIN1.CLIENT = "VAN" Then
            If NET <> 0 Then NPP = System.Math.Abs(NP / (NET - DED)) * System.Math.Sign(NP)
        End If
        'row(0).Item("LINE_DESC") = "Net Profit " & Format(NPP, "##.0%")
        row(0).Item("PCT") = NPP

        row = dst.Tables("ARTCSUMA").Select("LINE_ABBR = 'GP'")
        Dim GPP As Decimal = 0
        If NET <> 0 Then GPP = System.Math.Abs(GP / NET) * System.Math.Sign(GP)
        'row(0).Item("LINE_DESC") = "$GP on Net Sales " & Format(GPP, "##.0%")
        row(0).Item("PCT") = GPP

        row = dst.Tables("ARTCSUMA").Select("LINE_ABBR = 'DED'")
        Dim DP As Decimal = 0
        If NET <> 0 Then DP = System.Math.Abs(DED / NET) * System.Math.Sign(DED)
        '  row(0).Item("LINE_DESC") = "Deductions " & Format(DP, "##.0%")
        row(0).Item("PCT") = DP


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_grdARTCSUMC()
        Dim YP0 As String = cbeYPSFrom.Value
        Dim YP1 As String = cbeYPSTo.Value
        Dim P As Integer = ASCMAIN1.Period_Diff(YP0, YP1) + 1
        If P > 12 Then
            MsgBox("Period Range cannot span more than 12 Months", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        ElseIf P < 1 Then
            MsgBox("Period Range must be at least 1 month", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCSUMA, grdARTCSUMB}
            For i As Integer = 1 To 12
                Dim C As String = "AMT" & Format(i, "00")
                With grd.DisplayLayout.Bands(0).Columns(C)
                    If i > P Then
                        .Hidden = True
                    Else
                        .Hidden = False
                        Dim LEGEND As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(YP0, i - 1), False, True)
                        .Header.Caption = LEGEND
                    End If
                End With
                grd.DisplayLayout.Bands(0).Columns("AMT00").Header.Caption = "Period"
                grd.DisplayLayout.Bands(0).Columns("AMT13").Header.Caption = "Total"
            Next
            grd.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        Next

        grdARTCSUMA.DisplayLayout.Bands(0).Columns("LINE_DESC").Width = 300
        grdARTCSUMB.DisplayLayout.Bands(0).Columns("CODE_VALUE").Width = 100
        grdARTCSUMB.DisplayLayout.Bands(0).Columns("DESC_VALUE").Width = 200

        EnforceConstraints(False)

        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            dst.Tables("ARTCSUMA").Rows.Clear()
            TAC.SOCMAIN1.Create_ARTCSUMA(Me, Absx1.txtFor("CUST_CODE").Text)
            dst.Tables("ARTCSUMB").Rows.Clear()
        End If

        EnforceConstraints(True)

        Sort_grdColumns(grdARTCSUMA, "LINE", True)

    End Sub

    Private Sub grdARTCSUMA_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdARTCSUMA.AfterRowActivate
        Setup_grdARTCSUMB()
    End Sub

    Sub Setup_grdARTCSUMB()
        If grdARTCSUMA.ActiveRow Is Nothing Then
            grdARTCSUMB.Visible = False
        Else

            Dim LINE_ABBR As String = grdARTCSUMA.ActiveRow.Cells("LINE_ABBR").Value & ""
            If LINE_ABBR = "NET" Or LINE_ABBR = "GP" Or LINE_ABBR = "NP" Then
                grdARTCSUMB.Visible = False
            Else
                grdARTCSUMB.Visible = True
                grdARTCSUMB.Text = "Customer P&L Detail of " & grdARTCSUMA.ActiveRow.Cells("LINE_DESC").Value & " for " & Absx1.txtFor("CUST_CODE").Text
                With grdARTCSUMB.DisplayLayout.Bands(0)
                    Select Case grdARTCSUMA.ActiveRow.Cells("LINE_ABBR").Value & ""
                        Case "GRS", "RTN", "ST", "PP", "GWP", "DSP", "CGS"
                            .Columns("CODE_VALUE").Header.Caption = "Item Code"
                            .Columns("DESC_VALUE").Header.Caption = "Description"
                        Case "DED", "CR", "DR"
                            .Columns("CODE_VALUE").Header.Caption = "Reason Code"
                            .Columns("DESC_VALUE").Header.Caption = "Description"
                        Case "GL"
                            .Columns("CODE_VALUE").Header.Caption = "Acct Code"
                            .Columns("DESC_VALUE").Header.Caption = "Description"
                        Case Else
                            .Columns("CODE_VALUE").Header.Caption = "?"
                            .Columns("DESC_VALUE").Header.Caption = "?"
                    End Select
                End With

                Dim dvw As DataView = DirectCast(grdARTCSUMB.DataSource, DataTable).DefaultView
                dvw.RowFilter = "LINE = " & grdARTCSUMA.ActiveRow.Cells("LINE").Value
                Sort_grdColumns(grdARTCSUMB, "CODE_VALUE")
            End If
        End If
    End Sub

    Private Sub grdARTCSUMA_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCSUMA.InitializeRow
        Select Case e.Row.Cells("LINE_ABBR").Value
            Case "GRS"
                e.Row.Cells("LINE_DESC").Appearance.BackColor = Drawing.Color.LightGreen
            Case "RTN"
                e.Row.Cells("LINE_DESC").Appearance.ForeColor = Drawing.Color.Red
            Case "NET"
                e.Row.Appearance.BackColor = Drawing.Color.LightGray
            Case "CGS"
                e.Row.Cells("LINE_DESC").Appearance.ForeColor = Drawing.Color.Blue
            Case "GP"
                e.Row.Appearance.BackColor = Drawing.Color.LightGray
            Case "ST", "GWP", "DSP", "PP"
                e.Row.Cells("LINE_DESC").Appearance.BackColor = Drawing.Color.LightBlue
            Case "CR"
                e.Row.Cells("LINE_DESC").Appearance.BackColor = Drawing.Color.Yellow
            Case "DR"
                e.Row.Cells("LINE_DESC").Appearance.BackColor = Drawing.Color.Yellow
            Case "DED"
                e.Row.Cells("LINE_DESC").Appearance.BackColor = Drawing.Color.Orange
            Case "GL"
                e.Row.Cells("LINE_DESC").Appearance.BackColor = Drawing.Color.Orange
            Case "NP"
                e.Row.Appearance.BackColor = Drawing.Color.LightGray
        End Select
    End Sub

    Private Sub grdARTOPENB_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTOPENB.DoubleClickRow
        If e.Row.IsDataRow Then
            Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value
            Absx1.txtFor("CUST_CODE").Text = CUST_CODE
            Click_Command("Select Customer")
        End If
    End Sub

    Private Sub chkEditNotes_CheckedChanged(sender As Object, e As EventArgs) Handles chkEditNotes.CheckedChanged
        If chkEditNotes.Checked Then
            grdARTOPEN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            With grdARTOPEN1.DisplayLayout.Bands(0).Columns("INV_NOTES")
                .CellAppearance.BackColor = Drawing.Color.Yellow
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            End With
        Else
            grdARTOPEN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            With grdARTOPEN1.DisplayLayout.Bands(0).Columns("INV_NOTES")
                .CellAppearance.BackColor = Drawing.Color.Empty
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        End If
    End Sub


#Region "grdARTPYMT4"

    Private Sub grdARTPYMT4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT4.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdARTPYMT4, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
        End Select
    End Sub

    Private Sub grdARTPYMT4_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT4.AfterExitEditMode
        With grdARTPYMT4
            Select Case .ActiveCell.Column.Key
                Case "ACCT_CODE"
                    Dim ACCT_CODE As String = .ActiveCell.Text
                    If ACCT_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdARTPYMT4_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT4.AfterRowActivate
        With grdARTPYMT4
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdARTPYMT4.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdARTPYMT4_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT4.AfterRowsDeleted
        Display_Application_Totals()
    End Sub

    Private Sub grdARTPYMT4_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdARTPYMT4.AfterRowUpdate
        Display_Application_Totals()
    End Sub

    Private Sub grdARTPYMT4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTPYMT4.BeforeRowUpdate
        With grdARTPYMT4
            If e.Row.Cells("ACCT_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                    If cdr.Item("ACCT_STATUS") & "" <> "A" Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is not Active", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                    If cdr.Item("ACCT_SUB_CTL") & "" = "1" Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is a Control Account - no Manual J/E permitted", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If

                    Dim DIST_APP_CODE As String = "AR"
                    If collections_mode Then DIST_APP_CODE = "CF"
                    If LookUp("GLTDSTR1", DIST_APP_CODE) IsNot Nothing AndAlso cdr.Item("DIST_APP_STATUS") & "" = "A" Then
                        If LookUp("GLTDSTR2", New String() {DIST_APP_CODE, e.Row.Cells("ACCT_CODE").Text}) Is Nothing Then
                            MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is not permitted for Posting in this Application (" & DIST_APP_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If

                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        e.Cancel = True
                    Else
                        LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("PYMT_BATCH_NO").Text = "" Then
                    .ActiveRow.Cells("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
                    .ActiveRow.Cells("PYMT_BATCH_LNO").Value = PYMT_BATCH_LNO
                    .ActiveRow.Cells("PYMT_BATCH_GLNO").Value = Val(dst.Tables("ARTPYMT4").Compute("Max(PYMT_BATCH_GLNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdARTPYMT4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT4.ClickCellButton
        Dim sql_where As String = ""

        Select Case e.Cell.Column.Key

            Case "ACCT_CODE"
                Dim DIST_APP_CODE As String = "AR"
                If collections_mode Then DIST_APP_CODE = "CF"

                If LookUp("GLTDSTR1", DIST_APP_CODE) IsNot Nothing AndAlso cdr.Item("DIST_APP_STATUS") & "" = "A" Then
                    sql_where = "ACCT_CODE in (Select ACCT_CODE from GLTDSTR2 where DIST_APP_CODE = '" & DIST_APP_CODE & "')"
                End If

            Case "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"

        End Select

        grdClickCellButton(grdARTPYMT4, sql_where, sql_where <> "")
    End Sub

    Private Sub grdARTPYMT4_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdARTPYMT4.DoubleClickCell
        If e.Cell.Column.Key = "GL_DIST_AMT_CURR" Then
            If grdARTPYMT4.ActiveCell Is Nothing Then
            Else
                'grdARTPYMT4.ActiveCell.Value = Val(grdARTPYMT4.ActiveCell.Value & "") + Get_TOTAL_UNAPPLIED()
                'grdARTPYMT4.UpdateData()

                grdARTPYMT4.ActiveCell.Value = 0
                grdARTPYMT4.ActiveRow.Update()
                grdARTPYMT4.ActiveCell.Value = -1 * Get_TOTAL_UNAPPLIED()
                grdARTPYMT4.ActiveRow.Update()

            End If
        End If

    End Sub

    Private Sub grdARTPYMT4_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTPYMT4.Error
        grdARTPYMT4.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub Calculate_Totals()

        Dim CUST_PYMT_AMT_CURR As Decimal = 0
        TOTALS.APPL_TOTAL = Val(dst.Tables("ARTOPEN1").Compute("SUM (WOFF)", "") & "")
        'TOTALS.APPL_TOTAL = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_PMT_CURR)", "") & "")
        'TOTALS.DISC_TOTAL = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_DISC_TAKEN_CURR)", "") & "")
        'TOTALS.WOFF_TOTAL = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_WRITE_OFF_CURR)", "") & "")

        'TOTALS.DED_TOTAL = Val(dst.Tables("ARTPYMT5").Compute("SUM (GL_DIST_AMT_CURR)", "CHARGEBACK_IND IS NULL OR CHARGEBACK_IND = '0'") & "")
        'TOTALS.CHB_TOTAL = Val(dst.Tables("ARTPYMT5").Compute("SUM (GL_DIST_AMT_CURR)", "GL_DIST_AMT_CURR > 0 AND CHARGEBACK_IND = '1'") & "")
        'TOTALS.OA_TOTAL = Val(dst.Tables("ARTPYMT5").Compute("SUM (GL_DIST_AMT_CURR)", "GL_DIST_AMT_CURR < 0 AND CHARGEBACK_IND = '1'") & "")
        TOTALS.GL_TOTAL = Val(dst.Tables("ARTPYMT4").Compute("SUM (GL_DIST_AMT_CURR)", "") & "")
        'TOTALS.NET_AR = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE_NEW_CURR)", "") & "") - Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE_CURR)", "") & "") + TOTALS.CHB_TOTAL + TOTALS.OA_TOTAL ' + TOTALS.WOFF_TOTAL
        TOTALS.NET_AR = Val(dst.Tables("ARTOPEN1").Compute("SUM (INV_BALANCE_NEW)", "") & "") - Val(dst.Tables("ARTOPEN1").Compute("SUM (INV_BALANCE)", "") & "") + TOTALS.CHB_TOTAL + TOTALS.OA_TOTAL ' + TOTALS.WOFF_TOTAL
        TOTALS.UNAPPLIED = System.Math.Round(CUST_PYMT_AMT_CURR - (TOTALS.APPL_TOTAL - TOTALS.DED_TOTAL - TOTALS.CHB_TOTAL - TOTALS.OA_TOTAL - TOTALS.GL_TOTAL), 2)

    End Sub

    Sub Display_Application_Totals()
        Calculate_Totals()
        With dst.Tables("ARTPYMTT").Rows
            .Find("1").Item("PYMT_TOTAL_AMT") = TOTALS.APPL_TOTAL
            '.Find("2").Item("PYMT_TOTAL_AMT") = TOTALS.DISC_TOTAL
            '.Find("3").Item("PYMT_TOTAL_AMT") = TOTALS.WOFF_TOTAL
            .Find("4").Item("PYMT_TOTAL_AMT") = TOTALS.DED_TOTAL
            .Find("5").Item("PYMT_TOTAL_AMT") = TOTALS.CHB_TOTAL
            .Find("6").Item("PYMT_TOTAL_AMT") = TOTALS.GL_TOTAL
            .Find("7").Item("PYMT_TOTAL_AMT") = TOTALS.OA_TOTAL * -1
            .Find("8").Item("PYMT_TOTAL_AMT") = TOTALS.NET_AR
            .Find("9").Item("PYMT_TOTAL_AMT") = TOTALS.UNAPPLIED
            ' cmdLeaveOA.Visible = (TOTALS.UNAPPLIED > 0)
        End With
        grdARTPYMTT.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        If ScreenMode Then
            If TOTALS.UNAPPLIED < 0 Then
                ' MsgBox("You are now out of Funds.", MsgBoxStyle.OkOnly, "Notification")
            End If
        End If
    End Sub


    Function Get_TOTAL_UNAPPLIED() As Decimal
        Dim CUST_PYMT_AMT_CURR As Decimal = 0

        If Absx1.txtFor("CUST_CODE").Text = "" Then
            Dim TOTAL4 As Decimal = Val(dst.Tables("ARTPYMT4").Compute("SUM(GL_DIST_AMT_CURR)", "") & "")
            Return System.Math.Round(CUST_PYMT_AMT_CURR + TOTAL4, 2)
            'Return System.Math.Round(CUST_PYMT_AMT_CURR + TOTAL4, 2)
        Else
            Return System.Math.Round(Val(dst.Tables("ARTPYMTT").Rows.Find("9").Item("PYMT_TOTAL_AMT") & ""), 2)
        End If
        'Return System.Math.Round(Val(dst.Tables("ARTPYMTT").Rows.Find("9").Item("PYMT_TOTAL_AMT") & ""), 2)
    End Function

    Sub Write_Batch_Headers()

        dst.Tables("ARTPYMT1").Rows.Clear()
        Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
        With rowARTPYMT1
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO ' PYMT_BATCH_NO_application_only
            .Item("PYMT_BATCH_DATE") = DATETIME_STAMP.Date
            .Item("STATUS") = "1"
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("PYMT_APPL_ONLY") = "1"
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1
            .Item("PYMT_SOURCE") = "CWO"
        End With
        dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim CUST_NAME As String = Absx1.txtFor("CUST_NAME").Text

        dst.Tables("ARTPYMT2").Rows.Clear()

        Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
        With rowARTPYMT2
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO ' PYMT_BATCH_NO_application_only
            .Item("PYMT_BATCH_LNO") = 1 'PYMT_BATCH_LNO_application_only
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_NAME") = CUST_NAME
            .Item("CUST_PYMT_REF_NO") = ""
            .Item("CUST_PYMT_REF_DATE") = DATETIME_STAMP.Date
            .Item("CUST_PYMT_AMT") = 0
            .Item("PYMT_STATUS") = "2"
            .Item("CUST_PYMT_AMT_CURR") = 0
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("CURR_CODE") = DBNull.Value
            .Item("CURR_EXCH_RATE") = DBNull.Value
        End With
        dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)

        INIT_LAST("ARTPYMT1")

    End Sub

    Sub Format_grdARTOPEN1(tf As Boolean)
        Dim showForex As Boolean = tf
        With grdARTOPEN1.DisplayLayout.Bands(0)
            .Columns("INV_BALANCE").Hidden = showForex
            .Columns("INV_BALANCE_CURR").Hidden = Not showForex
            .Columns("CURR_CODE").Hidden = Not showForex
            .Columns("CURR_EXCH_RATE").Hidden = Not showForex
            Dim VP As Integer = .Columns("INV_BALANCE").Header.VisiblePosition
            .Columns("INV_BALANCE_CURR").Header.VisiblePosition = VP + 1
            .Columns("CURR_CODE").Header.VisiblePosition = VP + 2
            .Columns("CURR_EXCH_RATE").Header.VisiblePosition = VP + 3
            .Columns("INV_BALANCE_CURR").Header.Appearance.BackColor2 = Drawing.Color.Lime
            .Columns("INV_BALANCE_CURR").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Columns("CURR_CODE").Header.Appearance.BackColor2 = Drawing.Color.Lime
            .Columns("CURR_CODE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Columns("CURR_EXCH_RATE").Header.Appearance.BackColor2 = Drawing.Color.Lime
            .Columns("CURR_EXCH_RATE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        End With

    End Sub
    Sub Format_grdARTPYMT2(tf As Boolean)
        Dim showForex As Boolean = tf
        With grdARTPYMTY.DisplayLayout.Bands(0)
            .Columns("CURR_CODE").Hidden = Not showForex
            .Columns("CURR_EXCH_RATE").Hidden = Not showForex
            Dim VP As Integer = .Columns("CURR_CODE").Header.VisiblePosition
            .Columns("CURR_EXCH_RATE").Header.VisiblePosition = VP + 1
            .Columns("CURR_CODE").Header.Appearance.BackColor2 = Drawing.Color.Lime
            .Columns("CURR_CODE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Columns("CURR_EXCH_RATE").Header.Appearance.BackColor2 = Drawing.Color.Lime
            .Columns("CURR_EXCH_RATE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        End With

    End Sub
    Sub Format_grdARTPYMT3(tf As Boolean)
        Dim showForex As Boolean = tf
        If showForex Then
            For Each row3 As DataRow In dst.Tables("ARTPYMT3").Select()
                Dim PYMT_BATCH_NO As String = row3.Item("PYMT_BATCH_NO")
                Dim PYMT_BATCH_LNO As Integer = row3.Item("PYMT_BATCH_LNO")
                Dim PYMT_BATCH_ILNO As Integer = row3.Item("PYMT_BATCH_ILNO")
                ASCMAIN1.sql = "SELECT CURR_CODE, CURR_EXCH_RATE, CURR_GAIN_LOSS" _
                    & " from ARTYPMT3 where PYMT_BATCH_NO = :PARM1 and PYMT_BATCH_LNO = :PARM2 and PYMT_BATCH_ILNO = :PARM3"
                Dim rowARTPYMT3 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VNN", New String() {PYMT_BATCH_NO, PYMT_BATCH_LNO, PYMT_BATCH_ILNO})
                If rowARTPYMT3 IsNot Nothing Then
                    row3.Item("CURR_CODE") = rowARTPYMT3.Item("CURR_CODE") & ""
                    row3.Item("CURR_EXCH_RATE") = Val(rowARTPYMT3.Item("CURR_CODE") & "")
                    row3.Item("CURR_GAIN_LOSS") = Val(rowARTPYMT3.Item("CURR_CODE") & "")
                End If
            Next


        End If
        With grdARTPYMTX.DisplayLayout.Bands(0)
            .Columns("CURR_CODE").Hidden = Not showForex
            .Columns("CURR_EXCH_RATE").Hidden = Not showForex
            .Columns("CURR_GAIN_LOSS").Hidden = Not showForex
            Dim VP As Integer = .Columns("CURR_CODE").Header.VisiblePosition
            .Columns("CURR_EXCH_RATE").Header.VisiblePosition = VP + 1
            .Columns("CURR_GAIN_LOSS").Header.VisiblePosition = VP + 2
            .Columns("CURR_CODE").Header.Appearance.BackColor2 = Drawing.Color.Lime
            .Columns("CURR_CODE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Columns("CURR_EXCH_RATE").Header.Appearance.BackColor2 = Drawing.Color.Lime
            .Columns("CURR_EXCH_RATE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Columns("CURR_GAIN_LOSS").Header.Appearance.BackColor2 = Drawing.Color.Lime
            .Columns("CURR_GAIN_LOSS").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        End With
    End Sub

    Sub Toggle_Forex_Columns()
        If rowARTCUST1 Is Nothing Then Exit Sub

        Dim CURR_CODE As String = rowARTCUST1.Item("CURR_CODE") & ""

        Dim showForex As Boolean = (CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE"))

        lblCurr_Code.Text = IIf(CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE"), "", CURR_CODE)
        lblCurr_Code.Visible = (lblCurr_Code.Text <> "")

        Format_grdARTOPEN1(showForex)
        Format_grdARTPYMT2(showForex)
        Format_grdARTPYMT3(showForex)
    End Sub

    Private Sub cmdFetchSummary_HandleDestroyed(sender As Object, e As EventArgs) Handles cmdFetchSummary.HandleDestroyed

    End Sub

    Private Sub chkAllDivisions_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllDivisions.CheckedChanged
        cbeSALES_DIVISION.Visible = Not chkAllDivisions.Checked
    End Sub

    Private Sub UltraButton1_Click(sender As Object, e As EventArgs) Handles cmdUpdate_Due_Date.Click

        ' dgj
        If grdARTOPEN1.Selected.Rows.Count = 0 Then
            MsgBox("You must select the A/R Item row to change the Invoice Due Date",
                            MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If
        Dim TERM_CODE As String = "CDD" ' Custom Due Date Terms Code Need to not re-calc Due Date if terms 'CDD'
        Dim INV_DUE_DATE As Date = dteDue_Date.Value

        If MsgBox("Click Yes to Update the Due Dates to  " & Format(INV_DUE_DATE, "MM/dd/yy") & " and the Terms Codes to 'CDD' for all Open A/R Items Selected" _
                          & vbCrLf & vbCrLf & "An Audit Trail Record will be Recorded.",
            MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        BeginTrans()

        dst.Tables("ARTOPEN1").AcceptChanges()


        For Each grow As UltraWinGrid.UltraGridRow In grdARTOPEN1.Selected.Rows
            Dim INV_DUE_DATE_orig As Date = grow.Cells("INV_DUE_DATE").Value & ""


            grow.Cells("TERM_CODE").Value = TERM_CODE
            grow.Cells("INV_DUE_DATE").Value = INV_DUE_DATE

            Dim TERM_CODE_orig As String = grdARTOPEN1.ActiveRow.Cells("TERM_CODE").Value & ""
            Dim INV_TYPE As String = grow.Cells("INV_TYPE").Value & ""
            Dim INV_NUM As String = grow.Cells("INV_NUM").Value & ""
            Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value & ""
            Dim INV_DATE As Date = grow.Cells("INV_DATE").Value & ""

            TAC.TACMAIN1.Record_Event("ARTOPEN1", CUST_CODE & ":" & INV_TYPE & ":" & INV_NUM,
                                                      DATETIME_STAMP, ASCMAIN1.USER_ID, "CHGDUEDATE",
                                                      "AR Due Date Changed from " & Format(INV_DUE_DATE_orig, "MM/dd/yy") & " to " & Format(INV_DUE_DATE, "MM/dd/yy"), "", Me.Name)
            grow.Update()
        Next

        Update_Record_TDA("ARTOPEN1")
        CommitTrans("Update Due Date Complete")



    End Sub

    Private Sub chkALLFU_CheckedChanged(sender As Object, e As EventArgs) Handles chkALLFU.CheckedChanged
        '    grdARTCUSTT_FUPS.Visible = True
        If chkALLFU.Checked Then
            ASCMAIN1.sql = sqlARTCUSTT_FUPS.ToString & " and TATCONV1.INIT_OPER <> 'ana'"
            'ASCMAIN1.sql = "SELECT TATCONV1.*, ARTCUST1.CUST_NAME " _
            '    & " from TATCONV1,ARTCUST1 " _
            '    & " where ARTCUST1.CUST_CODE = TATCONV1.TABLE_KEY " _
            '    & "   and TATCONV1.CONV_STATUS = '1'" _
            '    & "   and TATCONV1.INIT_OPER <> 'ana'" _
            '    & "   and TATCONV1.TABLE_NAME = 'ARTCUST1'"
        Else
            ASCMAIN1.sql = sqlARTCUSTT_FUPS.ToString & " and TATCONV1.CONV_FOLLOWUP_BY = '" & ASCMAIN1.USER_ID & "'"
            'ASCMAIN1.sql = "SELECT TATCONV1.*, ARTCUST1.CUST_NAME " _
            '& " from TATCONV1,ARTCUST1 " _
            '& " where ARTCUST1.CUST_CODE = TATCONV1.TABLE_KEY " _
            '& "   and TATCONV1.CONV_STATUS = '1'" _
            '& "   and TATCONV1.TABLE_NAME = 'ARTCUST1'" _
            '& "   and TATCONV1.CONV_FOLLOWUP_BY = '" & ASCMAIN1.USER_ID & "'"
        End If
        Fill_Records("ARTCUSTT_FUPS", "", True, ASCMAIN1.sql)

    End Sub
    Sub AR_STATEMENT()
        Dim Report_date As Date = Now
        If dst.Tables.Contains("ARTSTMTZ") Then
            dst.Tables("ARTSTMTZ").Rows.Clear()
        Else
            With dst.Tables.Add("ARTSTMTZ")
                .Columns.Add("AR_PARM_KEY")
                .Columns.Add("REMIT0")
                .Columns.Add("REMIT1")
                .Columns.Add("REMIT2")
                .Columns.Add("REMIT3")
                .Columns.Add("AR_PARM_REMIT_MESSAGE")
                .Columns.Add("AR_PARM_DUNS_NO")
                .Columns.Add("ADDRESS_LINE")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .Columns.Add("AR_PARM_FIN_CHG_RATE", GetType(System.Decimal))
                .Columns.Add("STMT_DATE", GetType(System.DateTime))
                .PrimaryKey = New DataColumn() { .Columns("AR_PARM_KEY")}
            End With
        End If

        Dim rowARTSTMTZ As DataRow = dst.Tables("ARTSTMTZ").NewRow
        With ROWs("ARTPARM1")
            rowARTSTMTZ.Item("AR_PARM_KEY") = "Z"
            rowARTSTMTZ.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            rowARTSTMTZ.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            rowARTSTMTZ.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
                    & .Item("AR_PARM_REMIT_STATE") & " " _
                    & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                    & .Item("AR_PARM_REMIT_COUNTRY")
            If .Item("AR_PARM_REMIT_PHONE") & "" <> "" And .Item("AR_PARM_REMIT_FAX") & "" <> "" Then
                rowARTSTMTZ.Item("REMIT3") = "" _
                    & " Tel " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_PHONE")) _
                    & ", Fax " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_FAX"))
            End If
            rowARTSTMTZ.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
            If 1 = 1 Then
                rowARTSTMTZ.Item("AR_PARM_REMIT_MESSAGE") = rowARTSTMTZ.Item("AR_PARM_REMIT_MESSAGE") & vbCrLf & .Item("AR_PARM_REMIT_MESSAGE_EXPORT")
            End If
            rowARTSTMTZ.Item("AR_PARM_DUNS_NO") = .Item("AR_PARM_DUNS_NO") & ""
            rowARTSTMTZ.Item("AR_PARM_FIN_CHG_RATE") = .Item("AR_PARM_FIN_CHG_RATE") & ""
            rowARTSTMTZ.Item("ADDRESS_LINE") = "" _
                & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ADDR1") _
                & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_CITY") _
                & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_STATE") _
                & " " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ZIP_CODE") _
                & IIf(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "" <> "" _
                  And ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & "" <> "", "" _
                      & ", Tel " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "") _
                      & ", Fax " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & ""), "")
        End With
        rowARTSTMTZ.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        rowARTSTMTZ.Item("STMT_DATE") = Report_date
        dst.Tables("ARTSTMTZ").Rows.Add(rowARTSTMTZ)

    End Sub

#Region "outbound Requests"
    Private Sub btnEmailOutbound_Click(sender As Object, e As EventArgs) Handles btnEmailOutbound.Click
        Dim emsg As New Text.StringBuilder With {.Length = 0}
        If txtOBSendName.Text.Length = 0 Then
            emsg.AppendLine("Missing Sender Name.")
        End If
        If txtOBSendEmail.Text.Length = 0 Then
            emsg.AppendLine("Missing Sender Email.")
        End If
        If emsg.Length > 0 Then
            Dim iTitle As String = "Errors"
            MsgBox(emsg.ToString(), MsgBoxStyle.OkOnly, iTitle)
        Else
            Dim content As String = MakeHTMLBody()
            Dim fileName As String = ASCMAIN1.Folders("Temp") & "Outbound.html"
            If System.IO.File.Exists(fileName) Then
                System.IO.File.Delete(fileName)
            End If
            System.IO.File.WriteAllText(fileName, content)
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            Dim ATTACHMENTs As New Dictionary(Of String, String)
            EMAIL_ADDRESSs.Add(txtOBSendEmail.Text, txtOBSendName.Text)

            Dim TEMPLATE_NAME As String = "CREDIT"
            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                 "Credit Reference Request", TEMPLATE_NAME, True, False, TEMPLATE_NAME, TEMPLATE_NAME, "Credit Reference Request", content)
            If chkBCCAR.Checked Then
                Dim EMAIL_ADDRESSs_AR As New Dictionary(Of String, String)
                EMAIL_ADDRESSs_AR.Add("ar@regency-rib.com", "Accounts Receivable")
                SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs_AR, ATTACHMENTs,
                    "Credit Reference Request (Copy)", TEMPLATE_NAME, True, False, TEMPLATE_NAME, TEMPLATE_NAME, "Credit Reference Request (Copy)", content)
            End If
            MsgBox("Mail Sent", vbOKOnly, "Done")
        End If
    End Sub

    Private Sub btnPrintOutbound_Click(sender As Object, e As EventArgs) Handles btnPrintOutbound.Click
        Dim emsg As New Text.StringBuilder With {.Length = 0}
        If txtOBSendName.Text.Length = 0 Then
            emsg.AppendLine("Missing Sender Name.")
        End If
        If txtOBSendEmail.Text.Length = 0 Then
            emsg.AppendLine("Missing Sender Email.")
        End If
        If emsg.Length > 0 Then
            Dim iTitle As String = "Errors"
            MsgBox(emsg.ToString(), MsgBoxStyle.OkOnly, iTitle)
        Else
            Dim content As String = MakeHTMLBody()
            Dim fileName As String = ASCMAIN1.Folders("Temp") & "Outbound.html"
            If System.IO.File.Exists(fileName) Then
                System.IO.File.Delete(fileName)
            End If
            System.IO.File.WriteAllText(fileName, content)
            Show_Document(fileName)
        End If
    End Sub
    Private Sub fillOBData()

        Dim ADDR As New Text.StringBuilder With {.Length = 0}
        Dim CUST_ADDR1 As String = Absx1.txtFor("CUST_ADDR1").Text.ToString & String.Empty
        Dim CUST_ADDR2 As String = Absx1.txtFor("CUST_ADDR2").Text.ToString & String.Empty
        Dim CUST_ADDR3 As String = Absx1.txtFor("CUST_ADDR3").Text.ToString & String.Empty
        Dim CUST_CITY As String = Absx1.txtFor("CUST_CITY").Text.ToString & String.Empty
        Dim CUST_STATE As String = Absx1.txtFor("CUST_STATE").Text.ToString & String.Empty
        Dim CUST_ZIP_CODE As String = Absx1.txtFor("CUST_ZIP_CODE").Text.ToString & String.Empty

        ADDR.AppendLine(CUST_ADDR1)
        If CUST_ADDR2.Length > 0 Then
            ADDR.AppendLine(CUST_ADDR2)
        End If
        If CUST_ADDR3.Length > 0 Then
            ADDR.AppendLine(CUST_ADDR3)
        End If
        ADDR.AppendLine($"{CUST_CITY}, {CUST_STATE} {CUST_ZIP_CODE}")
        txtOBAddress.Text = ADDR.ToString

        txtOBTerms.Text = Absx1.txtFor("TERM_DESC").Text.ToString & String.Empty

        Dim rowARTCUST6 As DataRow = LookUp("ARTCUST6", HFs("CUST_CODE"), True)
        If Not IsNothing(rowARTCUST6) Then
            dteOBFirstOrder.Value = rowARTCUST6.Item("CUST_FIRST_PURCH")
            dteOBLastOrder.Value = rowARTCUST6.Item("CUST_LAST_INV_DATE")
        End If

        numOBCreditLimit.Value = Val(Absx1.numFor("CUST_CREDIT_LIMIT").Value & String.Empty)

        Dim rCnt As Int64 = 0
        Dim rcntMax As Int64 = 36
        Dim OBHightCredit As Double = 0
        Dim OBCurrBal As Double = 0
        Dim OBPastDue As Double = 0
        For Each rowARTSTMT1 As DataRow In dst.Tables("ARTSTMT1").Select("", "OPS_YYYYPP DESC")
            rCnt += 1
            Dim TOTAL_OPEN_AMT As Decimal = Val(rowARTSTMT1.Item("TOTAL_OPEN_AMT").ToString & String.Empty)
            If rowARTSTMT1.Item("OPS_YYYYPP") = "999999" Then
                Dim AGE_2 As Decimal = Val(rowARTSTMT1.Item("AGE_2").ToString & String.Empty)
                Dim AGE_3 As Decimal = Val(rowARTSTMT1.Item("AGE_3").ToString & String.Empty)
                Dim AGE_4 As Decimal = Val(rowARTSTMT1.Item("AGE_4").ToString & String.Empty)
                OBPastDue = AGE_2 + AGE_3 + AGE_4
                OBCurrBal = TOTAL_OPEN_AMT
                If AGE_4 > 0 Then
                    rdoOBRatingUnsatisfactory.Checked = True
                Else
                    If AGE_3 > 0 Then
                        rdoOBRatingSatisfactory.Checked = True
                    Else
                        rdoOBRatingPrompt.Checked = True
                    End If
                End If
            Else
                If rCnt >= rcntMax Then
                    Exit For
                Else
                    If TOTAL_OPEN_AMT > OBHightCredit Then
                        OBHightCredit = TOTAL_OPEN_AMT
                    End If
                End If
            End If
        Next
        numOBHightCredit.Value = OBHightCredit
        numOBCurrBal.Value = OBCurrBal
        numOBPastDue.Value = OBPastDue

        txtOBNotes.Text = ""

        txtOBSendName.Text = ""
        txtOBSendEmail.Text = ""
    End Sub

    Private Sub ClearOBData()
        txtOBAddress.Text = ""
        txtOBTerms.Text = ""
        dteOBFirstOrder.Value = Null
        dteOBLastOrder.Value = Null
        numOBCreditLimit.Value = Null
        rdoOBRatingSatisfactory.Checked = True
        numOBHightCredit.Value = Null
        numOBCurrBal.Value = Null
        numOBPastDue.Value = Null
        txtOBNotes.Text = ""
        txtOBSendName.Text = ""
        txtOBSendEmail.Text = ""
    End Sub

    Private Function MakeHTMLBody() As String
        Dim RetVal As String
        Dim TEMPLATE As String = $"{If(ASCMAIN1.useUNCPath, ASCMAIN1.Folders("SharedRoot"), "S:")}\Archive\templates\OutboundCredit.html"
        If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            TEMPLATE = "C:\Users\Wayne\Dropbox\Regency International\Shopsite Integration\Customers\OutboundCredit.html"
        End If
        Dim WEB_FIELDS As New Dictionary(Of String, String)
        WEB_FIELDS.Add("{AccountNumber}", Absx1.txtFor("CUST_CODE").Text.ToString & String.Empty)
        WEB_FIELDS.Add("{AccountName}", Absx1.txtFor("CUST_NAME").Text.ToString & String.Empty)
        WEB_FIELDS.Add("{SendName}", txtOBSendName.Text.ToString & String.Empty)
        WEB_FIELDS.Add("{SendEmail}", txtOBSendEmail.Text.ToString & String.Empty)
        WEB_FIELDS.Add("{Address}", txtOBAddress.Text.ToString & String.Empty)
        If rdoOBFAX.Checked Then
            WEB_FIELDS.Add("{SendType}", "Fax:")
        Else
            WEB_FIELDS.Add("{SendType}", "E-Mail:")
        End If
        '
        If rdoOBRatingNone.Checked Then
            WEB_FIELDS.Add("{PayRating}", "")
        Else
            If rdoOBRatingUnsatisfactory.Checked Then
                WEB_FIELDS.Add("{PayRating}", "<li><bold>Pay Rating:</bold> Unsatisfactory")
            Else
                If rdoOBRatingSatisfactory.Checked Then
                    WEB_FIELDS.Add("{PayRating}", "<li><bold>Pay Rating:</bold> Satisfactory")
                Else
                    WEB_FIELDS.Add("{PayRating}", "<li><bold>Pay Rating:</bold> Prompt")
                End If
            End If
        End If

        If (txtOBTerms.Text.ToString & String.Empty).Length > 0 Then
            WEB_FIELDS.Add("{Terms}", $"<li><bold>Terms:</bold> {txtOBTerms.Text.ToString & String.Empty}</li> ")
        Else
            WEB_FIELDS.Add("{Terms}", "")
        End If

        If Not IsNothing(dteOBFirstOrder.Value) Then
            WEB_FIELDS.Add("{FirstOrder}", $"<li><bold>Date Opened:</bold> {Format(CDate(dteOBFirstOrder.Value.ToString & String.Empty), "MM/dd/yyyy")}</li>")
        Else
            WEB_FIELDS.Add("{FirstOrder}", "")
        End If

        If Not IsNothing(dteOBLastOrder.Value) Then
            WEB_FIELDS.Add("{LastOrder}", $"<li><bold>Last Sale:</bold> {Format(CDate(dteOBLastOrder.Value.ToString & String.Empty), "MM/dd/yyyy")}</li>")
        Else
            WEB_FIELDS.Add("{LastOrder}", "")
        End If

        If IsNumeric(numOBCreditLimit.Value.ToString & String.Empty) Then
            If Val(numOBCreditLimit.Value.ToString & String.Empty) > 0 Then
                WEB_FIELDS.Add("{CreditLimit}", $"<li><bold>Credit Limit:</bold> {Format(Val(numOBCreditLimit.Value.ToString & String.Empty), "###,###,##0")}</li>")
            Else
                WEB_FIELDS.Add("{CreditLimit}", "")
            End If
        Else
            WEB_FIELDS.Add("{CreditLimit}", "")
        End If

        If IsNumeric(numOBHightCredit.Value.ToString & String.Empty) Then
            If Val(numOBHightCredit.Value.ToString & String.Empty) > 0 Then
                WEB_FIELDS.Add("{HightCredit}", $"<li><bold>High Credit:</bold> {Format(Val(numOBHightCredit.Value.ToString & String.Empty), "###,###,##0")}</li>")
            Else
                WEB_FIELDS.Add("{HightCredit}", "")
            End If
        Else
            WEB_FIELDS.Add("{HightCredit}", "")
        End If

        If IsNumeric(numOBCurrBal.Value.ToString & String.Empty) Then
            If Val(numOBCurrBal.Value.ToString & String.Empty) > 0 Then
                WEB_FIELDS.Add("{CurrBal}", $"<li><bold>Current Balance:</bold> {Format(Val(numOBCurrBal.Value.ToString & String.Empty), "###,###,##0")}</li>")
            Else
                WEB_FIELDS.Add("{CurrBal}", "")
            End If
        Else
            WEB_FIELDS.Add("{CurrBal}", "")
        End If

        If IsNumeric(numOBPastDue.Value.ToString & String.Empty) Then
            If Val(numOBPastDue.Value.ToString & String.Empty) > 0 Then
                WEB_FIELDS.Add("{PastDue}", $"<li><bold>Past Due:</bold> {Format(Val(numOBPastDue.Value.ToString & String.Empty), "###,###,##0")}</li>")
            Else
                WEB_FIELDS.Add("{PastDue}", "")
            End If
        Else
            WEB_FIELDS.Add("{PastDue}", "")
        End If

        If (txtOBNotes.Text.ToString & String.Empty).Length > 0 Then
            WEB_FIELDS.Add("{Notes}", $"<bold>Other Notes:</bold> {txtOBNotes.Text.ToString & String.Empty}")
        Else
            WEB_FIELDS.Add("{Notes}", "")
        End If

        Dim BodyContent As String = System.IO.File.ReadAllText(TEMPLATE)
        BodyContent = BodyContent.Replace(vbCrLf, "")
        For Each WEB_FIELD As KeyValuePair(Of String, String) In WEB_FIELDS
            BodyContent = BodyContent.Replace(WEB_FIELD.Key, WEB_FIELD.Value)
        Next
        RetVal = BodyContent

        Return RetVal
    End Function

    Private Sub rdoOBEMAIL_CheckedChanged(sender As Object, e As EventArgs) Handles rdoOBEMAIL.CheckedChanged
        setOBOptions()
    End Sub

    Private Sub rdoOBFAX_CheckedChanged(sender As Object, e As EventArgs) Handles rdoOBFAX.CheckedChanged
        setOBOptions()
    End Sub

    Private Sub setOBOptions()
        If rdoOBEMAIL.Checked Then
            btnPrintOutbound.Enabled = True
            btnEmailOutbound.Enabled = True
        End If

        If rdoOBFAX.Checked Then
            btnPrintOutbound.Enabled = True
            btnEmailOutbound.Enabled = False
        End If
    End Sub
#End Region

End Class