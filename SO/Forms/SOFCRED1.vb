Public Class SOFCRED1

    Dim CUST_CODE As String
    Dim CUST_BILL_TO_CUST As String
    Dim CUST_CREDIT_GROUP_CUST As String

    Dim ORDR_NO As String
    Dim ORDR_TOTAL_AMT As Decimal
    Dim ORDR_GROUP_NO As String

    Dim rowARTCUST1 As DataRow
    Dim rowSOTORDR1 As DataRow
    Dim rowSOTORDRA As DataRow

    Dim SOTORDRA As String
    Dim SOTORDRAsql As String
 
    Dim ARTSTMTX As String
    Dim sqlARTSTMTX As String

    Dim REVIEW_RESULTS As String
    Dim REVIEW_RESULTS_DESC As String

    Dim AGED_TOTALS() As Decimal
    Dim DUE_TOTALS() As Decimal

    'Dim sqlARTCUSTX As String

    ' aging history shown in customer tree form
    ' batches = these also updated ORDR_REL, and had a report
    ' credit reports need to be credit group aware
    ' recent payments need to show customer. and artcust6 for payments needs to join to bill to

    Dim time_to_wait As Integer = 0

#Region "ABS Standard Routines"
    ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ARTPARM1")
        Get_PARM("SOTPARM1")

        TAC.ARCMAIN1.Get_Aging_Data(ROWs("ARTPARM1"), Now.Date)

        With dst
            SOTORDRAsql = "Select SOTORDR0.*" & vbCrLf _
                & ",SOTAUTH1.ORDR_CRED_HOLD_CODES" & vbCrLf _
                & ",SOTAUTH1.ORDR_REL_HOLD_CODES" & vbCrLf _
                & ",SOTAUTH1.ORDR_CREDIT_NOTE" & vbCrLf _
                & ",SOTAUTH1.ORDR_CRED_CLR_BY" & vbCrLf _
                & ",SOTAUTH1.ORDR_CRED_CLR_AUTH" & vbCrLf _
                & ",SOTAUTH1.ORDR_CRED_CLR_DATE" & vbCrLf _
                & ",SOTAUTH1.ORDR_CRED_CLR_AUTH_TYPE" & vbCrLf _
                & ",SOTAUTH1.ORDR_CRED_CLR_AUTH_NO" & vbCrLf _
                & ",SOTAUTH1.ORDR_CRED_CLR_AUTH_CTL_NO" & vbCrLf _
                & ",SOTAUTH1.ORDR_CRED_CLR_AUTH_AMT" & vbCrLf _
                & ",ARTCUST1.CUST_NAME" & vbCrLf _
                & " from SOTORDR0,SOTAUTH1,ARTCUST1" & vbCrLf & vbCrLf _
                & " where SOTAUTH1.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
                & "   and SOTORDR0.ORDR_CNT_OPEN <> 0" & vbCrLf _
                & "   and SOTORDR0.ORDR_TYPE_CODE <> 'XFR'"

            SOTORDRA = ASCMAIN1.Temp_Table(SOTORDRAsql & " and ROWNUM < 1")
            ASCMAIN1.sql = "Select * from " & SOTORDRA
            Create_TDA(.Tables.Add, "SOTORDRA", "**", 0, False, "", 1)

            sqlARTSTMTX = "" _
            & "Select ARTSTMT1.CUST_CODE, ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.CUST_CRED_LIMIT_REV" & vbCrLf _
            & ", MAX (NVL(CUST_HIGH_BAL_AMT,0)) HIGH_BAL" & vbCrLf _
            & ", SUM (ARTSTMT1.TOTAL_OPEN_AMT) OPEN_AMT" & vbCrLf _
            & ", SUM (ARTSTMT1.TOTAL_OPEN_DDS) OPEN_AMT_DAYS" & vbCrLf _
            & ", SUM (ARTSTMT1.TOTAL_CLSD_AMT) PAID_AMT" & vbCrLf _
            & ", SUM (ARTSTMT1.TOTAL_CLSD_DDS) PAID_AMT_DAYS" & vbCrLf _
            & ", SUM (CASE WHEN NVL(ARTSTMT1.DUE_2,0)+NVL(ARTSTMT1.DUE_3,0)+NVL(ARTSTMT1.DUE_4,0)>0 THEN 1 ELSE 0 END) TIMES_PAST_DUE" & vbCrLf _
            & " from ARTSTMT1,ARTCUST1" & vbCrLf _
            & " where ARTSTMT1.OPS_YYYYPP >= :PARM1" & vbCrLf _
            & "   and ARTSTMT1.OPS_YYYYPP <= :PARM2" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = ARTSTMT1.CUST_CODE" & vbCrLf _
            & "   and NVL(ARTCUST1.CUST_CREDIT_LIMIT,0) > 0" & vbCrLf _
            & " group by ARTSTMT1.CUST_CODE, ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.CUST_CRED_LIMIT_REV"
            ARTSTMTX = ASCMAIN1.Temp_Table(Replace(Replace(sqlARTSTMTX, ":PARM1", "'000000'"), ":PARM2", "'000000'"))
            ASCDATA1.ExecuteSQL("Alter Table " & ARTSTMTX & " Add Primary Key (CUST_CODE)")

            Dim sqlARTOPENA As String = "" _
            & " Select CUST_CODE" & vbCrLf _
            & TAC.ARCMAIN1.AGED_TOTALS & vbCrLf _
            & ", SUM (INV_BALANCE) TOTAL_OPEN_AMT" & vbCrLf _
            & ", MAX (CASE WHEN INV_BALANCE > 0 THEN TRUNC(SYSDATE) - INV_DATE ELSE 0 END) OLDEST_DR_ITEM" _
            & " from ARTOPEN1 where INV_BALANCE <> 0" & vbCrLf _
            & " and CUST_CODE in (Select CUST_CODE from " & ARTSTMTX & ")" & vbCrLf _
            & " group by CUST_CODE"
            'ASCMAIN1.sql = sqlARTOPENA
            'Create_TDA(.Tables.Add, "ARTOPENA", "**", 0, False, "V", 1)
            ASCMAIN1.sql = "Select Sum (AGE_1) AGE_1, Sum (AGE_2) AGE_2, Sum (AGE_3) AGE_3, Sum (AGE_4) AGE_4, Sum (DUE_1) DUE_1, Sum (DUE_2) DUE_2, Sum (DUE_3) DUE_3, Sum (DUE_4) DUE_4, Sum (TOTAL_OPEN_AMT) TOTAL_OPEN_AMT, Max (OLDEST_DR_ITEM) OLDEST_DR_ITEM from (" & sqlARTOPENA & ")"
            Create_TDA(.Tables.Add, "ARTOPENA", "**", 0, False)

            Dim sqlSOTORDRX As String = "" _
            & " Select SOTORDR0.CUST_CODE" & vbCrLf _
            & ", SUM (SOTORDR0.ORDR_AMT_PICK) PEND" & vbCrLf _
            & ", SUM (SOTORDR0.ORDR_AMT_OPEN) APPR" & vbCrLf _
            & " from SOTORDR0 " & vbCrLf _
            & " where SOTORDR0.CUST_CODE in (Select CUST_CODE from " & ARTSTMTX & ")" & vbCrLf _
            & " group by CUST_CODE"

            ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.CUST_CONTACT" & vbCrLf _
            & ", ARTCUST1.CUST_DBA_NAME, ARTCUST1.CUST_PHONE, ARTCUST1.CUST_URL, ARTCUST1.CUST_EMAIL, ARTCUST1.CUST_DUNS" & vbCrLf _
            & ", ARTCUST1.CUST_STATUS, ARTCUST1.CUST_STATUS_DATE, ARTCUST1.CUST_STATUS_COMMENT, ARTCUST1.CUST_GROUP_CODE" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.CUST_CRED_LIMIT_EST, ARTCUST1.CUST_CRED_LIMIT_REV" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_LIMIT_NOTES ,ARTCUST1.CUST_CREDIT_LIMIT_APPR_BY" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_HOLD, ARTCUST1.CUST_CREDIT_RELEASE, ARTCUST1.CUST_FACTOR_IND" & vbCrLf _
            & ", ARTCUST1.CUST_BUS_ESTAB, ARTCUST1.TERM_CODE, ARTCUST1.CUST_TERMS_NOTE" & vbCrLf _
            & ", ARTCUST1.SREP_CODE, ARTCUST1.CUST_BILL_TO_CUST, ARTCUST1.CUST_CREDIT_GROUP_CUST" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_SCORE, ARTCUST1.CUST_CREDIT_SCORE_DATE, ARTCUST1.CUST_PD_GRACE_DAYS" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_RATING, ARTCUST1.CUST_CREDIT_RATING_DATE" & vbCrLf _
            & ", ARTCUST1.CUST_INS_AMT, ARTCUST1.CUST_INS_DATE" & vbCrLf _
            & ", SOTORDRX.PEND, SOTORDRX.APPR, ARTSTMTX.HIGH_BAL" & vbCrLf _
            & ", ARTOPENA.TOTAL_OPEN_AMT" & vbCrLf _
            & ", ARTOPENA.AGE_1, ARTOPENA.AGE_2, ARTOPENA.AGE_3, ARTOPENA.AGE_4" & vbCrLf _
            & ", ARTOPENA.DUE_1, ARTOPENA.DUE_2, ARTOPENA.DUE_3, ARTOPENA.DUE_4" & vbCrLf _
            & ", ARTCUST6.CUST_LAST_INV_DATE, ARTCUST6.CUST_LAST_PMT_DATE" & vbCrLf _
            & ", CASE WHEN ABS(NVL(OPEN_AMT,0)) < 1 THEN 0 ELSE TRUNC(100 * OPEN_AMT_DAYS/OPEN_AMT) / 100 END AVG_AGE" & vbCrLf _
            & ", CASE WHEN ABS(NVL(PAID_AMT,0)) < 1 THEN 0 ELSE TRUNC(100 * PAID_AMT_DAYS/PAID_AMT) / 100 END AVG_DTP" & vbCrLf _
            & ", ARTSTMTX.TIMES_PAST_DUE, SOTSREP1.SREP_NAME" & vbCrLf _
            & " from ARTCUST1, ARTCUST6, SOTSREP1, " & ARTSTMTX & " ARTSTMTX" & vbCrLf _
            & ", (" & sqlARTOPENA & ") ARTOPENA" & vbCrLf _
            & ", (" & sqlSOTORDRX & ") SOTORDRX" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = ARTSTMTX.CUST_CODE" & vbCrLf _
            & " and ARTCUST6.CUST_CODE (+) = ARTSTMTX.CUST_CODE" & vbCrLf _
            & " and ARTOPENA.CUST_CODE (+) = ARTSTMTX.CUST_CODE" & vbCrLf _
            & " and SOTORDRX.CUST_CODE (+) = ARTSTMTX.CUST_CODE" & vbCrLf _
            & " and SOTSREP1.SREP_CODE (+) = ARTCUST1.SREP_CODE"
            Create_TDA(.Tables.Add, "ARTCREDX", "**", 0, False, "N")
            .Tables("ARTCREDX").Columns("AVG_AGE").DataType = GetType(System.Int32)
            .Tables("ARTCREDX").Columns("AVG_DTP").DataType = GetType(System.Int32)
            .Tables("ARTCREDX").Columns("TIMES_PAST_DUE").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_GROUP_NO = :PARM1 and ORDR_STATUS = 'O'"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V", 1)
            '"ORDR_CRED_HOLD_CODES, ORDR_CRED_CLR_BY, ORDR_CRED_CLR_AUTH, ORDR_CRED_CLR_DATE, ORDR_REL_HOLD_CODES, ORDR_CRED_CLEARED"

            .Tables("SOTORDR1").Columns.Add("ORDR_AMT", GetType(System.Decimal))
            .Tables("SOTORDR1").Columns.Add("ORDR_TOTAL_AMT", GetType(System.Decimal), "ISNULL(ORDR_AMT,0)")

            ASCMAIN1.sql = "Select ARTCUST1.*" _
            & " from ARTCUST1 where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)

            .Tables.Add("ARTCREDA")
            With .Tables("ARTCREDA")
                .Columns.Add("LINE", GetType(System.Int32))
                .Columns.Add("DESCRIPTION", GetType(System.String))
                .Columns.Add("AMOUNT", GetType(System.Decimal))
            End With

            .Tables.Add("ARTCREDL")
            With .Tables("ARTCREDL")
                .Columns.Add("CODE", GetType(System.String))
                .Columns.Add("SEL", GetType(System.String))
                .Columns.Add("DESCRIPTION", GetType(System.String))
                .PrimaryKey = New DataColumn() {.Columns("CODE")}
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
             & "   and ARTSTMT1.CUST_CODE in (Select CUST_CODE from " & ARTSTMTX & ")" & vbCrLf _
             & " group by ARTSTMT1.OPS_YYYYPP, GLTPARM2.LEGEND"

            'ASCMAIN1.sql = "SELECT ARTSTMT1.*, GLTPARM2.LEGEND" _
            '& " from ARTSTMT1, GLTPARM2 " _
            '& " where GLTPARM2.OPS_YYYYPP (+) = ARTSTMT1.OPS_YYYYPP" _
            '& "   and ARTSTMT1.CUST_CODE in (Select CUST_CODE from " & ARTSTMTX & ")"
            Create_TDA(.Tables.Add, "ARTSTMT1", "**", 0, False, "", 1)
            .Tables("ARTSTMT1").Columns.Add("DAYS_OPEN", GetType(System.Decimal), "IIF(ISNULL(TOTAL_OPEN_AMT,0)=0,0,ISNULL(TOTAL_OPEN_DDS,0) / ISNULL(TOTAL_OPEN_AMT,0))")
            .Tables("ARTSTMT1").Columns.Add("DAYS_PAID", GetType(System.Decimal), "IIF(ISNULL(INV_CLSD_AMT,0)=0,0,ISNULL(INV_CLSD_DDS,0) / ISNULL(INV_CLSD_AMT,0))")
            .Tables("ARTSTMT1").Columns.Add("DAYS_PAID_SAVG", GetType(System.Decimal), "IIF(ISNULL(INV_CLSD_CNT,0)=0,0,ISNULL(INV_CLSD_DYS,0) / ISNULL(INV_CLSD_CNT,0))")

            ASCMAIN1.sql = "Select ARTPYMT2.*, ARTPYMT1.PYMT_BATCH_DATE" _
            & " from ARTPYMT2,ARTPYMT1 " _
            & " where ARTPYMT2.CUST_CODE in (Select CUST_CODE from " & ARTSTMTX & ")" & vbCrLf _
            & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & "   and ARTPYMT1.OPS_YYYYPP >= :PARM1" _
            & "   and NVL(ARTPYMT1.PYMT_APPL_ONLY,'0') <> '1'"
            Create_TDA(.Tables.Add, "ARTPYMT2", "**", 0, False, "V", 2)

            Create_TDA(.Tables.Add, "SOTAUTH1", "*")

            Create_TDA(.Tables.Add, "SOTORDRF", "*")

            ASCMAIN1.sql = "Select TATCONV1.*" _
              & " from TATCONV1" _
              & " where TABLE_NAME = 'ARTCUST1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATCONV1", "**", 0, , "V", 1)
            .Tables("TATCONV1").Columns.Add("CONV_ATTACHMENTS", GetType(System.Int64))

        End With

        grdSOTORDRA.DataSource = dst.Tables("SOTORDRA")
        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdARTCREDA.DataSource = dst.Tables("ARTCREDA")
        grdARTCREDL.DataSource = dst.Tables("ARTCREDL")
        grdARTCREDX.DataSource = dst.Tables("ARTCREDX")
        grdARTSTMT1.DataSource = dst.Tables("ARTSTMT1")
        grdARTPYMT2.DataSource = dst.Tables("ARTPYMT2")
        grdTATCONV1.DataSource = dst.Tables("TATCONV1")
      
        Create_Summary(grdSOTORDRA, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDRA, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP"})

        Show_Filter(grdSOTORDRA)

        With grdSOTORDRA.DisplayLayout.Bands("SOTORDRA")
            .Columns("ORDR_GROUP_NO").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
            .Columns("ORDR_AMT_OPEN").Header.Fixed = True
        End With

        With grdSOTORDR1.DisplayLayout.Bands("SOTORDR1")
            .Columns("ORDR_NO").Header.Fixed = True
            .Columns("CUST_STORE_NO").Header.Fixed = True
            .Columns("ORDR_AMT").Header.Fixed = True
        End With

        For i As Integer = 1 To 4
            With grdARTSTMT1.DisplayLayout.Bands(0)
                .Columns("AGE_" & CStr(i)).Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_" & CStr(i)) & ""
                .Columns("DUE_" & CStr(i)).Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_" & CStr(i)) & ""
            End With
        Next

        With grdARTCREDX.DisplayLayout.Bands("ARTCREDX")
            .Columns("CUST_CODE").Header.Fixed = True
        End With

        With grdARTCREDX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
            {"CUST_CREDIT_LIMIT", "CUST_CRED_LIMIT_EST", "CUST_CRED_LIMIT_REV", _
             "CUST_CREDIT_LIMIT_NOTES", "CUST_CREDIT_LIMIT_APPR_BY"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
            For Each COLUMN_NAME As String In New String() _
            {"CUST_STATUS", "CUST_STATUS_DATE", "CUST_STATUS_COMMENT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightPink
            Next
            For Each COLUMN_NAME As String In New String() _
            {"TOTAL_OPEN_AMT", "AGE_1", "AGE_2", "AGE_3", "AGE_4", "DUE_1", "DUE_2", "DUE_3", "DUE_4", "AVG_AGE", "AVG_DTP", "TIMES_PAST_DUE", "CUST_PD_GRACE_DAYS"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightSalmon
            Next
            For Each COLUMN_NAME As String In New String() _
            {"CUST_LAST_INV_DATE", "CUST_LAST_PMT_DATE", "HIGH_BAL"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGoldenrodYellow
            Next
            For Each COLUMN_NAME As String In New String() _
            {"CUST_DUNS", "CUST_CREDIT_SCORE", "CUST_CREDIT_SCORE_DATE", _
             "CUST_CREDIT_RATING", "CUST_CREDIT_RATING_DATE", "CUST_INS_AMT", "CUST_INS_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            Next
        End With

        grdARTSTMT1.DisplayLayout.UseFixedHeaders = True
        With grdARTSTMT1.DisplayLayout.Bands(0)
            .Columns("LEGEND").Header.Fixed = True
        End With
        With grdARTPYMT2.DisplayLayout.Bands("ARTPYMT2")
            .Columns("PYMT_BATCH_NO").Header.Fixed = True
        End With

        Create_Summary(grdARTCREDX, "CUST_CODE", "Count")

        Create_Summary(grdARTPYMT2, "PYMT_BATCH_NO", "Count")
        Create_Summary(grdARTPYMT2, "CUST_PYMT_AMT")


        Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDR1, "ORDR_AMT")

        Bind_Controls(grpHeader, "SOTORDR1")
        Bind_Controls(SplitContainer3, "SOTORDR1")
        Bind_Controls(grpCustCredit, "ARTCUST1")

        grdARTCREDX.Visible = False

        Add_ARTCREDL("CRH", "Credit Hold")
        Add_ARTCREDL("INA", "Inactive > " & ROWs("SOTPARM1").Item("SO_PARM_CRH_DAYS_INACTIVE") & " Days")
        Add_ARTCREDL("REV", "CR Limit Review")
        Add_ARTCREDL("LIM", "Over CR Limit")
        Add_ARTCREDL("P/D", "Past Due > " & ROWs("SOTPARM1").Item("SO_PARM_CRH_DAYS_PAST_DUE"))
        Add_ARTCREDL("SLH", "Sales Hold")
        Add_ARTCREDL("TRM", "Std Terms")

        'ASCMAIN1.Add_Value_List(grdSOTORDRA, "ORDR_CREDIT_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTORDRA, "ORDR_TYPE_CODE")
        ASCMAIN1.Add_Value_List(grdARTCREDX, "CUST_STATUS")

        ASCMAIN1.Add_Value_List(grdSOTORDRA, "ORDR_CRED_CLR_AUTH", , New String() {":", "A:Appr", "H:Hold", "D:Decl", "W:Work"})

        cmbCRED_CODE.DataSource = ASCDATA1.GetDataTable("Select CRED_CODE, CRED_DESC from ARTCRED1 order by CRED_CODE")
        SplitContainer4.Panel2Collapsed = True

        Absx1.numFor("CUST_CREDIT_LIMIT").MaskInput = "nnn,nnn,nnn" ' this keeps decimal places from showing up when you click into the control.  Actually need to provide as many n's as the max value warrants
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select Order"

                ORDR_NO = ""
                ORDR_GROUP_NO = ""

                If grdSOTORDRA.Selected.Rows.Count = 0 Then
                    If grdSOTORDRA.ActiveRow IsNot Nothing Then
                        grdSOTORDRA.Selected.Rows.Add(grdSOTORDRA.ActiveRow)
                    End If
                End If

                If grdSOTORDRA.Selected.Rows.Count > 0 Then
                    If grdSOTORDRA.Selected.Rows.Count > 1 Then
                        EMsg &= "You Must Select only one order to be Reviewed"
                    Else

                        ORDR_GROUP_NO = grdSOTORDRA.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                        rowSOTORDRA = dst.Tables("SOTORDRA").Rows.Find(ORDR_GROUP_NO)

                        CUST_CODE = grdSOTORDRA.ActiveRow.Cells("CUST_CODE").Value & ""
                        ORDR_TOTAL_AMT = Val(grdSOTORDRA.ActiveRow.Cells("ORDR_AMT").Value & "")

                        ASCMAIN1.sql = "Select ORDR_NO, ORDR_STATUS, CUST_BILL_TO_CUST from SOTORDR1" _
                            & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS = 'O'"

                        ASCMAIN1.sql = "Select * from SOTORDR1" _
                            & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS = 'O'"

                        rowSOTORDR1 = ASCDATA1.GetDataRow
                        If rowSOTORDR1 Is Nothing Then
                            EMsg &= "Order " & ORDR_NO & " is no longer Open"
                        Else
                            If rowSOTORDR1.Item("ORDR_STATUS") & "" <> "O" Then
                                EMsg &= "Order " & ORDR_NO & " is no longer Open"
                            End If

                            If EMsg = "" Then

                                CUST_BILL_TO_CUST = rowSOTORDR1.Item("CUST_BILL_TO_CUST") & ""
                                'CUST_CREDIT_GROUP_CUST = grdSOTORDRA.ActiveRow.Cells("CUST_CREDIT_GROUP_CUST").Value & ""
                                ORDR_NO = rowSOTORDR1.Item("ORDR_NO") & ""

                                Dim rowSOTAUTH1 As DataRow = LookUp("SOTAUTH1", ORDR_GROUP_NO)
                                If rowSOTAUTH1 IsNot Nothing Then
                                    If rowSOTAUTH1.Item("ORDR_CRED_CLR_AUTH") & "" = "" Then
                                        'If cdr.Item("ORDR_CREDIT_STATUS") & "" = "" Or cdr.Item("ORDR_CREDIT_STATUS") & "" = "H" Then
                                        If tabOrders.SelectedTab.Key <> "Pending" Then
                                            EMsg &= "Order " & ORDR_NO & " has already been Reviewed (Please Refresh)"
                                        End If
                                    Else
                                        If tabOrders.SelectedTab.Key <> "Reviewed" Then
                                            EMsg &= "Order " & ORDR_NO & " is Pending Review (Please Refresh)"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    EMsg &= "You Must Select an Order to be Reviewed by Double-Clicking a Row"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ARTCUST1", CUST_BILL_TO_CUST) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                End If

            Case "Refresh"

            Case "Cancel"

            Case "Update"
                'If txtORDR_CREDIT_NOTE.Visible Then
                '    If txtORDR_CREDIT_NOTE.Text = "" Then
                '        EMsg &= vbCr & "You Must Provide a Note"
                '    End If
                'End If
                If cmbCRED_CODE.Visible Then
                    If cmbCRED_CODE.Value & "" = "" Then
                        If optAction.Value & "" = "A" Then
                            ' NOTHING ELSE REQUIRED
                        ElseIf optAction.Value & "" = "C" Then
                            EMsg &= vbCr & "You Must Select a Conditional Approval Code"
                        ElseIf optAction.Value & "" = "H" Then
                            EMsg &= vbCr & "You Must Select an Approval Pending Code"
                        ElseIf optAction.Value & "" = "D" Then
                            EMsg &= vbCr & "You Must Select a Credit Rejection Code"
                        Else
                            EMsg &= vbCr & "You Must Select a Credit Rejection Code"
                        End If

                    End If
                End If
                If optAction.CheckedIndex = -1 Then
                    EMsg &= vbCr & "You Must Select an Action"
                End If
                If EMsg = "" Then
                    Dim rowARTCREDL As DataRow = dst.Tables("ARTCREDL").Rows.Find("CRH")
                    If rowARTCREDL.Item("SEL") & "" = "1" Then
                        If MsgBox("The Customer is on Credit Hold." _
                                        & vbCrLf & vbCrLf & "Are you sure that you want to Continue with this Action?", _
                                        MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If

                    'If ORDR_NOsPending.Count = 1 Then
                    '    rowSOTORDR1 = LookUp("SOTORDR1", ORDR_NOsPending.Item(0))
                    'End If
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

            Case "Select Order"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                Clear_Record()

            Case "Cancel"
                Mode_Settings(False)

            Case "Update", "Approve", "Reject", "Still Pending"
                Update_Record()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
                '.Groups("Screen Control").Items("Approve").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Reject").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Still Pending").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Select Order").Settings.Enabled = not_iScreenMode
                .Groups("Approve/Reject").Visible = ScreenMode

                '.Groups("Screen Control").Items("Update").Visible = False
            End With
        End If

        grpHeader.Visible = ScreenMode
        grpReports.Visible = Not ScreenMode

        tabMain.Tabs("Orders").Visible = Not ScreenMode
        tabMain.Tabs("Approve/Reject").Visible = ScreenMode
        tabMain.Tabs("Credit Limit Reports").Visible = False ' Not ScreenMode ' not ready for prime time

        Set_Read_Only(grpHeader, True)
        Set_Read_Only(grpApprove, Not ScreenMode)
        Set_Read_Only(grpCustCredit, True)

        If ScreenMode Then
            If Absx1.txtFor("CUST_CODE").ButtonsRight.Count = 0 Then
                Absx1.txtFor("CUST_CODE").ButtonsRight.Add(New Infragistics.Win.UltraWinEditors.EditorButton)
                Absx1.txtFor("CUST_BILL_TO_CUST").ButtonsRight.Add(New Infragistics.Win.UltraWinEditors.EditorButton)
                ' Absx1.txtFor("CUST_CREDIT_GROUP_CUST").ButtonsRight.Add(New Infragistics.Win.UltraWinEditors.EditorButton)

                Absx1.txtFor("CUST_CODE").ButtonsRight(0).Key = "ARFCINQ1"
                Absx1.txtFor("CUST_BILL_TO_CUST").ButtonsRight(0).Key = "ARFCINQ1"
                '  Absx1.txtFor("CUST_CREDIT_GROUP_CUST").ButtonsRight(0).Key = "ARFCINQ1"

                Absx1.txtFor("CUST_CODE").ButtonsRight(0).Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "FORM_GREEN")
                Absx1.txtFor("CUST_BILL_TO_CUST").ButtonsRight(0).Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "FORM_RED")
                ' Absx1.txtFor("CUST_CREDIT_GROUP_CUST").ButtonsRight(0).Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "FORM_YELLOW")

                AddHandler Absx1.txtFor("CUST_CODE").EditorButtonClick, AddressOf txtCUST_CODE_EditorButtonClick
                AddHandler Absx1.txtFor("CUST_BILL_TO_CUST").EditorButtonClick, AddressOf txtCUST_CODE_EditorButtonClick
                ' AddHandler Absx1.txtFor("CUST_CREDIT_GROUP_CUST").EditorButtonClick, AddressOf txtCUST_CODE_EditorButtonClick

                Absx1.txtFor("CUST_CODE").ButtonsRight.Add(New Infragistics.Win.UltraWinEditors.EditorButton)
                Absx1.txtFor("CUST_BILL_TO_CUST").ButtonsRight.Add(New Infragistics.Win.UltraWinEditors.EditorButton)
                '  Absx1.txtFor("CUST_CREDIT_GROUP_CUST").ButtonsRight.Add(New Infragistics.Win.UltraWinEditors.EditorButton)

                Absx1.txtFor("CUST_CODE").ButtonsRight(1).Key = "ARTCUST1"
                Absx1.txtFor("CUST_BILL_TO_CUST").ButtonsRight(1).Key = "ARTCUST1"
                ' Absx1.txtFor("CUST_CREDIT_GROUP_CUST").ButtonsRight(1).Key = "ARTCUST1"

                Absx1.txtFor("CUST_CODE").ButtonsRight(1).Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "FORM_GREEN_EDIT")
                Absx1.txtFor("CUST_BILL_TO_CUST").ButtonsRight(1).Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "FORM_RED_EDIT")
                ' Absx1.txtFor("CUST_CREDIT_GROUP_CUST").ButtonsRight(1).Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "FORM_YELLOW_EDIT")

                'AddHandler Absx1.txtFor("CUST_CODE").EditorButtonClick, AddressOf txtCUST_CODE_EditorButtonClick
                'AddHandler Absx1.txtFor("CUST_BILL_TO_CUST").EditorButtonClick, AddressOf txtCUST_CODE_EditorButtonClick
                'AddHandler Absx1.txtFor("CUST_CREDIT_GROUP_CUST").EditorButtonClick, AddressOf txtCUST_CODE_EditorButtonClick
            End If
            Absx1.txtFor("CUST_CODE").ButtonsRight(0).Enabled = True
            Absx1.txtFor("CUST_BILL_TO_CUST").ButtonsRight(0).Enabled = True
            '   Absx1.txtFor("CUST_CREDIT_GROUP_CUST").ButtonsRight(0).Enabled = True

            Setup_grdSOTORDRA()

        Else
            Clear_Record()
            Show_Filter(grdSOTORDRA, True)
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() _
        {"SOTORDR1", "ARTCUST1", "ARTCREDX", "ARTCREDA", "ARTSTMT1", "ARTPYMT2", _
         "SOTAUTH1", "TATCONV1", "SOTORDRF"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        '"ARTCUSTX_C", "ARTCUSTX_B", "ARTCUSTX_S",

        EnforceConstraints(True)


        ORDR_NO = ""
        ORDR_GROUP_NO = ""

        txtORDR_CREDIT_NOTE.Text = ""
        cmbCRED_CODE.Value = ""
        txtORDR_CREDIT_NOTE.Visible = False
        lblORDR_CREDIT_NOTE.Visible = False
        cmbCRED_CODE.Visible = False
        REVIEW_RESULTS = ""
        REVIEW_RESULTS_DESC = ""

        For Each rowARTCREDL As DataRow In dst.Tables("ARTCREDL").Rows
            rowARTCREDL.Item("SEL") = "0"
            If rowARTCREDL.Item("CODE") = "TRM" Then
                rowARTCREDL.Item("DESCRIPTION") = "Std Terms"
            End If
        Next

        SetupGrids()
        tabMain.SelectedTab = tabMain.Tabs(0)
        Setup_Tab()
        Setup_Exceed_Pct()
    End Sub

    Sub Load_Record()

        'Call Save_Header_Fields(UltraGroupBox1)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Selecting Order for Credit Review")

        EnforceConstraints(False)

        Dim ORDR_TOTAL_AMT As Decimal = Val(rowSOTORDRA.Item("ORDR_AMT") & "")
        Absx1.numFor("ORDR_TOTAL_AMT").Value = ORDR_TOTAL_AMT

        Fill_Records("SOTORDR1", ORDR_GROUP_NO)

        ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO" _
            & ", Sum (NVL(SOTORDR2.ORDR_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT" _
            & " from SOTORDR2,SOTORDR1" _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and SOTORDR1.ORDR_STATUS = 'O'" _
            & " group by SOTORDR2.ORDR_NO"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim O As String = row.Item("ORDR_NO") & ""
            Dim ORDR_AMT As Decimal = Val(row.Item("ORDR_AMT") & "")
            Dim rowO As DataRow = dst.Tables("SOTORDR1").Rows.Find(O)
            rowO.Item("ORDR_AMT") = ORDR_AMT
            'If rowO IsNot Nothing Then
            '    rowO.Item("ORDR_AMT") = ORDR_AMT
            'Else
            '    ORDR_AMT = 0
            'End If
        Next

        Dim dvw As DataView = DirectCast(grdSOTORDRA.DataSource, DataTable).DefaultView '  dst.Tables("SOTORDRA").DefaultView
        dvw.RowFilter = "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"

        Fill_Records("ARTCUST1", CUST_BILL_TO_CUST)

        ASCDATA1.ExecuteSQL("Truncate Table " & ARTSTMTX)
        ASCDATA1.ExecuteSQL("Insert into " & ARTSTMTX & " (CUST_CODE) Select '" & CUST_BILL_TO_CUST & "' CUST_CODE from DUAL union Select CUST_CODE from ARTCUST1 where NVL(CUST_CREDIT_GROUP_CUST,CUST_CODE) = '" & CUST_CREDIT_GROUP_CUST & "'  union Select CUST_CODE from ARTCUST1 where NVL(CUST_CREDIT_GROUP_CUST,CUST_BILL_TO_CUST) = '" & CUST_CREDIT_GROUP_CUST & "'")

        grpCreditGroupCustomer.Text = "Credit Information for (Credit Group Customer) " & CUST_CREDIT_GROUP_CUST

        Add_Customer(CUST_CODE, "Sold-To", txtSOLD_TO)
        Add_Customer(CUST_BILL_TO_CUST, "Bill-To", txtBILL_TO)

        'Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO")
        'Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO}, True)
        'txtSHIP_TO.Text = rowARTCUST2.Item("CUST_NAME") & ""

        Fill_Records("TATCONV1", CUST_CODE)
        If CUST_BILL_TO_CUST <> CUST_CODE Then
            Fill_Records("TATCONV1", CUST_BILL_TO_CUST, False)
            If CUST_CREDIT_GROUP_CUST <> CUST_BILL_TO_CUST Then
                Fill_Records("TATCONV1", CUST_CREDIT_GROUP_CUST, False)
            End If
        End If
        Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)

        Dim CUST_CREDIT_LIMIT As Decimal = Val(Absx1.numFor("CUST_CREDIT_LIMIT").Value & "")

        ASCMAIN1.sql = "Select Count (*) ORDS" & vbCrLf _
            & ", Sum (CASE WHEN ORDR_CNT_PICK <> 0 THEN ORDR_AMT_PICK ELSE 0 END) PICK" & vbCrLf _
            & ", Sum (CASE WHEN ORDR_CNT_OPEN <> 0 THEN DECODE(SOTAUTH1.ORDR_CRED_CLR_AUTH,'A',ORDR_AMT_OPEN,0) ELSE 0 END) APPR" & vbCrLf _
            & ", Sum (CASE WHEN ORDR_CNT_OPEN <> 0 THEN DECODE(SOTAUTH1.ORDR_CRED_CLR_AUTH,'A',0,ORDR_AMT_OPEN) ELSE 0 END) PEND" & vbCrLf _
            & " from SOTORDR0,SOTAUTH1,ARTCUST1" & vbCrLf _
            & " where SOTAUTH1.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
            & "   and (SOTORDR0.ORDR_CNT_OPEN <> 0 OR SOTORDR0.ORDR_CNT_PICK <> 0)" & vbCrLf _
            & "   and SOTORDR0.ORDR_TYPE_CODE <> 'XFR'" & vbCrLf _
            & "   and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'"
        Dim row2 As DataRow = ASCDATA1.GetDataRow

        Dim APPR As Decimal = Val(row2.Item("APPR") & "") + Val(row2.Item("PICK") & "")
        Dim PEND As Decimal = Val(row2.Item("PEND") & "")

        Fill_Records("ARTOPENA")

        Dim rowARTOPENA As DataRow
        If dst.Tables("ARTOPENA").Rows.Count = 0 Then
            rowARTOPENA = dst.Tables("ARTOPENA").NewRow
        Else
            rowARTOPENA = dst.Tables("ARTOPENA").Rows(0)
        End If
        Dim TOTAL_OPEN_AMT As Decimal = Val(rowARTOPENA.Item("TOTAL_OPEN_AMT") & "")

        Dim LINE As Int32 = 0
        Add_ARTCREDA(LINE, ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_1") & "", Val(rowARTOPENA.Item("DUE_1") & ""))
        Add_ARTCREDA(LINE, ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_2") & "", Val(rowARTOPENA.Item("DUE_2") & ""))
        Add_ARTCREDA(LINE, ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_3") & "", Val(rowARTOPENA.Item("DUE_3") & ""))
        Add_ARTCREDA(LINE, ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_DESC_4") & "", Val(rowARTOPENA.Item("DUE_4") & ""))
        Add_ARTCREDA(LINE, "Total Balance", TOTAL_OPEN_AMT)
        Add_ARTCREDA(LINE, "Orders Approved", APPR)
        Add_ARTCREDA(LINE, "Credit Remaining", CUST_CREDIT_LIMIT - TOTAL_OPEN_AMT - APPR)
        Add_ARTCREDA(LINE, "Orders Pending", PEND)
        Sort_grdColumns(grdARTCREDA, "LINE", True)

        If CUST_CREDIT_LIMIT - TOTAL_OPEN_AMT - APPR < PEND Then
            dst.Tables("ARTCREDL").Rows.Find("LIM").Item("SEL") = "1"
        End If
        If Val(rowARTOPENA.Item("OLDEST_DR_ITEM") & "") > Val(ROWs("SOTPARM1").Item("SO_PARM_CRH_DAYS_PAST_DUE") & "") Then
            dst.Tables("ARTCREDL").Rows.Find("P/D").Item("SEL") = "1"
        End If

        'If rowSOTORDRA.Item("CUST_LAST_INV_DATE") & "" = "" OrElse _
        '   Format(CDate(rowSOTORDRA.Item("CUST_LAST_INV_DATE")).AddDays(Val(ROWs("SOTPARM1").Item("SO_PARM_CRH_DAYS_INACTIVE") & "")), "yyyyMMdd") < Format(Now + ASCMAIN1.NowTSD, "yyyyMMdd") Then
        '    dst.Tables("ARTCREDL").Rows.Find("INA").Item("SEL") = "1"
        'End If

        grdARTSTMT1.DisplayLayout.Rows.FixedRows.Clear()

        Setup_grdARTSTMT1(CUST_CREDIT_GROUP_CUST)

        Load_Payment_History()

        EnforceConstraints(True)

        txtORDR_CREDIT_NOTE.Text = ""
        cmbCRED_CODE.Value = ""
        optAction.CheckedIndex = -1

        Absx1.txtFor("LAST_3612").Text = TAC.ARCMAIN1.Last_3612_SAvg(CUST_CREDIT_GROUP_CUST)

        tabOrder.Tabs("Orders Pending").Text = "Group No " & ORDR_GROUP_NO

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub SetupGrids()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ASCDATA1.ExecuteSQL("Truncate table " & SOTORDRA)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDRA & " (" & SOTORDRAsql & ")")

        EnforceConstraints(False)

        Fill_Records("SOTORDRA")
        Setup_grdSOTORDRA()
        Sort_grdColumns(grdSOTORDRA, "ORDR_GROUP_NO".ToLower)
        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Dim msg As String

        BeginTrans()

        Dim ORDR_NOs As Integer = dst.Tables("SOTORDR1").Rows.Count

        If ORDR_NOs > 1 Then
            msg = String.Format("All {0} Orders Pending Credit Review for Credit Group Customer {1} " & IIf(ORDR_GROUP_NO = "", "", " in Batch " & ORDR_GROUP_NO) & "have been {2}", _
                                ORDR_NOs, CUST_CODE, REVIEW_RESULTS_DESC)
        Else
            msg = String.Format("Order {0} has been {1}", _
                                ORDR_NO, REVIEW_RESULTS_DESC)
        End If

        Dim rowSOTAUTH1 As DataRow = Fill_Record("SOTAUTH1", ORDR_GROUP_NO)
        If rowSOTAUTH1 Is Nothing Then
            rowSOTAUTH1 = dst.Tables("SOTAUTH1").NewRow
            rowSOTAUTH1.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            dst.Tables("SOTAUTH1").Rows.Add(rowSOTAUTH1)
        End If

        With rowSOTAUTH1

            TAC.TACMAIN1.Record_Event("SOTORDR0", ORDR_GROUP_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "CR_" & REVIEW_RESULTS, txtORDR_CREDIT_NOTE.Text)

            .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            If REVIEW_RESULTS = "APR" Then
                .Item("ORDR_CRED_CLR_AUTH") = "A"
                .Item("ORDR_CRED_CLEARED") = .Item("ORDR_CRED_HOLD_CODES")
                rowSOTAUTH1.Item("CRED_CODE") = DBNull.Value
            Else
                If REVIEW_RESULTS = "CAP" Then
                    .Item("ORDR_CRED_CLR_AUTH") = "A"
                ElseIf REVIEW_RESULTS = "REJ" Then
                    .Item("ORDR_CRED_CLR_AUTH") = "D"
                ElseIf REVIEW_RESULTS = "PND" Then
                    .Item("ORDR_CRED_CLR_AUTH") = "H"
                End If
                rowSOTAUTH1.Item("CRED_CODE") = cmbCRED_CODE.Value
                rowSOTAUTH1.Item("ORDR_CREDIT_NOTE") = txtORDR_CREDIT_NOTE.Text
            End If

            rowSOTAUTH1.Item("ORDR_CRED_CLR_BY") = .Item("ORDR_CRED_CLR_BY")
            rowSOTAUTH1.Item("ORDR_CRED_CLR_DATE") = .Item("ORDR_CRED_CLR_DATE")
            rowSOTAUTH1.Item("ORDR_CRED_HOLD_CODES") = .Item("ORDR_CRED_HOLD_CODES")
            rowSOTAUTH1.Item("ORDR_REL_HOLD_CODES") = .Item("ORDR_REL_HOLD_CODES")
            rowSOTAUTH1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSOTAUTH1.Item("INIT_DATE") = DATETIME_STAMP

            Dim rowSOTORDRF As DataRow = dst.Tables("SOTORDRF").NewRow
            rowSOTORDRF.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO

            Dim rowARTOPENA As DataRow = dst.Tables("ARTOPENA").Rows(0)
            If rowARTOPENA IsNot Nothing Then
                rowSOTORDRF.Item("AGE_1") = rowARTOPENA.Item("AGE_1")
                rowSOTORDRF.Item("AGE_2") = rowARTOPENA.Item("AGE_2")
                rowSOTORDRF.Item("AGE_2") = rowARTOPENA.Item("AGE_2")
                rowSOTORDRF.Item("AGE_4") = rowARTOPENA.Item("AGE_4")
            Else
                rowSOTORDRF.Item("AGE_1") = 0
                rowSOTORDRF.Item("AGE_2") = 0
                rowSOTORDRF.Item("AGE_2") = 0
                rowSOTORDRF.Item("AGE_4") = 0
            End If

            Dim APPR As Decimal = 0 ' Val(row.Item("APPR") & "")
            Dim PEND As Decimal = 0 ' Val(row.Item("PEND") & "")
            rowSOTORDRF.Item("ORDERS_APPR") = APPR
            rowSOTORDRF.Item("ORDERS_PEND") = PEND
            rowSOTORDRF.Item("CRED_CODE") = cmbCRED_CODE.Value
            rowSOTORDRF.Item("ORDR_CREDIT_NOTE") = txtORDR_CREDIT_NOTE.Text

            rowSOTORDRF.Item("ORDR_CREDIT_STATUS") = .Item("ORDR_CRED_CLR_AUTH")
            rowSOTORDRF.Item("ORDR_CREDIT_APPR_BY") = .Item("ORDR_CRED_CLR_BY")
            rowSOTORDRF.Item("ORDR_CREDIT_APPR_DATE") = .Item("ORDR_CRED_CLR_DATE")

            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            rowSOTORDRF.Item("CUST_CODE") = CUST_CREDIT_GROUP_CUST
            rowSOTORDRF.Item("CUST_CREDIT_LIMIT") = rowARTCUST1.Item("CUST_CREDIT_LIMIT")
            rowSOTORDRF.Item("CUST_CRED_LIMIT_EST") = rowARTCUST1.Item("CUST_CRED_LIMIT_EST")
            rowSOTORDRF.Item("CUST_CRED_LIMIT_REV") = rowARTCUST1.Item("CUST_CRED_LIMIT_REV")
            rowSOTORDRF.Item("CUST_CREDIT_HOLD") = rowARTCUST1.Item("CUST_CREDIT_HOLD")
            rowSOTORDRF.Item("CUST_CREDIT_SCORE") = rowARTCUST1.Item("CUST_CREDIT_SCORE")
            rowSOTORDRF.Item("CUST_CREDIT_SCORE_DATE") = rowARTCUST1.Item("CUST_CREDIT_SCORE_DATE")
            rowSOTORDRF.Item("CUST_CREDIT_LIMIT_APPR_BY") = rowARTCUST1.Item("CUST_CREDIT_LIMIT_APPR_BY")
            rowSOTORDRF.Item("CUST_CREDIT_LIMIT_NOTES") = rowARTCUST1.Item("CUST_CREDIT_LIMIT_NOTES")
            rowSOTORDRF.Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
            rowSOTORDRF.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSOTORDRF.Item("INIT_DATE") = DATETIME_STAMP
            rowSOTORDRF.Item("CUST_CREDIT_RATING") = rowARTCUST1.Item("CUST_CREDIT_RATING")
            rowSOTORDRF.Item("CUST_CREDIT_RATING_DATE") = rowARTCUST1.Item("CUST_CREDIT_RATING_DATE")
            rowSOTORDRF.Item("CUST_INS_AMT") = rowARTCUST1.Item("CUST_INS_AMT")
            rowSOTORDRF.Item("CUST_INS_DATE") = rowARTCUST1.Item("CUST_INS_DATE")
            dst.Tables("SOTORDRF").Rows.Add(rowSOTORDRF)
        End With

        Update_Record_TDA("SOTAUTH1")
        Update_Record_TDA("SOTORDRF")
        Update_Record_TDA("TATCONV1")

        'If optAction.Value <> "A" Then
        '    Generate_emails()
        'End If

        ASCMAIN1.MultiTask_Release() ' UNLOCK CUSTOMER AND ORDER 

        CommitTrans(msg)

    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOTORDRA, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Inquiry")
        Call Load_Popup_Menu(grdARTCREDX, "SSB", "Show Filter", "Show GroupBox", "Customer Inquiry")
        Call Load_Popup_Menu(grdTATCONV1, "SSBBBB", "Show Filter", "Show GroupBox", "Add to Log", "Show Log", "Edit Log", "Follow-Up")
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

        Select Case e.SourceControl.Name
            Case "grdTATCONV1"
                'Dim btnEnabled_Show As Boolean = False
                'Dim btnEnabled_Follow_Up As Boolean = False
                'If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                '    Dim CONV_NO As String = grd.ActiveRow.Cells("CONV_NO").Text
                '    Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                '    If rowTATCONV1 IsNot Nothing Then
                '        btnEnabled_Show = True
                '        If (EntryMode = "E" Or EntryMode = "N") Then
                '            btnEnabled_Follow_Up = True
                '        End If
                '    End If
                'End If

                tlb_btn = DirectCast(tlb_pop.Tools("Show Log"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Enabled = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow ' btnEnabled_Show
                tlb_btn = DirectCast(tlb_pop.Tools("Edit Log"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Enabled = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow ' btnEnabled_Show And btnEnabled_Follow_Up
                tlb_btn = DirectCast(tlb_pop.Tools("Follow-Up"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Enabled = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow '  btnEnabled_Show And btnEnabled_Follow_Up

                Exit Sub
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            Case "grdSOTORDRA"
                e.Tool.ToolbarsManager.Tools("Sales Order Inquiry").SharedProps.Visible = True ' MAYBE NEED TO BE RESTRICTED
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Add to Log"
                Dim F As New TAFCONV2(Me, "ARTCUST1", Absx1.txtFor("CUST_CREDIT_GROUP_CUST").Text)
                F.EntryMode = "N"
                F.ShowDialog()
                If F.result = "U" Then
                    dst.Tables("TATCONV1").Rows.Add(F.rowTATCONV1.ItemArray)
                    Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)
                End If
                F.Dispose()

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Text
                'Dim row As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                ASCMAIN1.sql = "Select Min (ORDR_NO) ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                If row IsNot Nothing Then
                    Dim ORDR_NO As String = row.Item("ORDR_NO")
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI", "F", "SO")
                End If

            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")



            Case "Show Log"
                Dim CONV_NO As String = grd.ActiveRow.Cells("CONV_NO").Text
                Dim F As New TAFCONV2(Me, "ARTCUST1", Absx1.txtFor("CUST_CREDIT_GROUP_CUST").Text)
                F.EntryMode = "V"
                F.rowTATCONV1 = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                F.ShowDialog()
                F.Dispose()

            Case "Edit Log"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CREDIT_GROUP_CUST").Text
                Dim CONV_NO As String = grd.ActiveRow.Cells("CONV_NO").Text
                Dim F As New TAFCONV2(Me, "ARTCUST1", Absx1.txtFor("CUST_CREDIT_GROUP_CUST").Text)
                F.EntryMode = "E"
                F.rowTATCONV1 = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                F.ShowDialog()
                If F.result = "U" Then
                    Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)
                End If
                F.Dispose()

            Case "Follow-Up"
                Dim CONV_NO As String = grd.ActiveRow.Cells("CONV_NO").Text
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CREDIT_GROUP_CUST").Text
                Dim F As New TAFCONV2(Me, "ARTCUST1", Absx1.txtFor("CUST_CREDIT_GROUP_CUST").Text)
                F.EntryMode = "F"
                Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                F.rowTATCONV1_PREV = rowTATCONV1
                F.ShowDialog()
                If F.result = "U" Then
                    dst.Tables("TATCONV1").Rows.Add(F.rowTATCONV1.ItemArray)
                    Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)
                End If
                F.Dispose()

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub chk_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.chk_CheckedChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CREDIT_HOLD"
                With Absx1.chkFor("CUST_CREDIT_HOLD")
                    If .Checked Then
                        .Appearance.ForeColor = Drawing.Color.Red
                        .Appearance.ForeColorDisabled = Drawing.Color.Red
                    Else
                        .Appearance.ForeColor = Drawing.Color.Empty
                        .Appearance.ForeColorDisabled = Drawing.Color.Black
                    End If
                End With
        End Select
    End Sub
#End Region

    Sub txtCUST_CODE_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs)
        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
        Dim CUST_CODE As String = txtctl.Text
        If e.Button.Key = "ARFCINQ1" Then
            Context_Launch("Select Customer", CUST_CODE, "Customer Inquiry", "ARFCINQ1")
        Else
            Context_Launch("View", Column_Values("CUST_CODE", CUST_CODE), "Customer Master", "ARTCUST1")
        End If
    End Sub

    Private Sub grdSOTORDRA_DoubleClickRow(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDRA.DoubleClickRow
        If e.Row.IsDataRow Then
            Click_Command("Select Order")
        End If
    End Sub

    Private Sub grdARTCUST1_AGING_InitializeRow(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCREDA.InitializeRow
        If e.Row.Cells("DESCRIPTION").Text = "Total Due" Then
            e.Row.Appearance.BackColor = Drawing.Color.Beige
            e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
        End If
        If e.Row.Cells("DESCRIPTION").Text = "Credit Remaining" Then
            If Val(e.Row.Cells("AMOUNT").Value & "") > 0 Then
                e.Row.Appearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.Appearance.ForeColor = Drawing.Color.Red
            End If
            e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
        End If
        If e.Row.Cells("DESCRIPTION").Text = "Past Due DRs" Then
            If Val(e.Row.Cells("AMOUNT").Value & "") > 0 Then
                e.Row.Cells("AMOUNT").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
            End If
        End If


        If e.Row.Cells("DESCRIPTION").Text = "Over 90" Then
            If Val(e.Row.Cells("AMOUNT").Value & "") > 0 Then
                e.Row.Cells("AMOUNT").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
            End If
        End If

    End Sub

    Private Sub cmdCreditReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreditReport.Click
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building List of Customers")

        Dim sqlCRGroup As String = "SELECT NVL(CUST_CREDIT_GROUP_CUST,CUST_CODE) FROM ARTCUST1 WHERE CUST_CODE IN (SELECT NVL(CUST_BILL_TO_CUST,CUST_CODE) CUST_BILL_TO_CUST from ARTCUST1)"
        If optCRGroup.Value = "C" Then
            sqlCRGroup = " and ARTCUST1.CUST_CODE in (" & sqlCRGroup & ")"
        ElseIf optCRGroup.Value = "N" Then
            sqlCRGroup = " and ARTCUST1.CUST_CODE in (Select CUST_CODE from ARTCUST1 minus " & sqlCRGroup & ")"
        Else
            sqlCRGroup = ""
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & ARTSTMTX)
        ASCMAIN1.sql = "Insert into " & ARTSTMTX & " " & Replace(sqlARTSTMTX, "group by", sqlCRGroup & " " & "group by")

        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -2), ASCMAIN1.CYP})

        Dim PCT As String = CStr(Val(numExceed.Value) / 100)
        If optCreditReport.Value = "E" Then
            ASCMAIN1.sql = "Delete from " & ARTSTMTX _
            & " where HIGH_BAL <= CUST_CREDIT_LIMIT * " & PCT & ""
        ElseIf optCreditReport.Value = "N" Then
            ASCMAIN1.sql = "Delete from " & ARTSTMTX _
            & " where HIGH_BAL >= CUST_CREDIT_LIMIT * " & PCT & ""
        ElseIf optCreditReport.Value = "R" Then
            ASCMAIN1.sql = "Delete from " & ARTSTMTX _
            & " where CUST_CRED_LIMIT_REV is not Null and CUST_CRED_LIMIT_REV > SYSDATE + " & numExceed.Value
        End If
        ASCDATA1.ExecuteSQL()


        Dim LEGEND As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -2), False, True)

        Fill_Records("ARTCREDX")
        If optCreditReport.Value = "E" Then
            grdARTCREDX.Text = "Customers Exceeding " & numExceed.Value & "% of Credit Limit since " & LEGEND
        ElseIf optCreditReport.Value = "N" Then
            grdARTCREDX.Text = "Customers NOT Exceeding " & numExceed.Value & "% of Credit Limit since " & LEGEND
        ElseIf optCreditReport.Value = "R" Then
            grdARTCREDX.Text = "Customers with a CR Limit Review Date Expiring in the next " & numExceed.Value & " Days"
        End If
        If optCRGroup.Value = "C" Then
            grdARTCREDX.Text &= "; Credit Groups Only"
        ElseIf optCRGroup.Value = "N" Then
            grdARTCREDX.Text &= "; Non-Credit Groups Only"
        End If
        grdARTCREDX.Visible = True

        Sort_grdColumns(grdARTCREDX, "CUST_CODE")

        SplitContainer4.Panel2Collapsed = False

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Tab()
    End Sub

    Sub Setup_Tab()
        UltraExplorerBar1.Groups("Select Report Type").Visible = tabMain.SelectedTab IsNot Nothing AndAlso (tabMain.SelectedTab.Key = "Credit Limit Reports")
        UltraExplorerBar1.Groups("Screen Control").Visible = tabMain.SelectedTab Is Nothing OrElse (tabMain.SelectedTab.Key <> "Credit Limit Reports")

        If ScreenMode Then
            grdARTSTMT1.Text = ""
            grdARTSTMT1.Parent = tabOrder.Tabs("Aging History").TabPage
        Else
            grdARTSTMT1.Parent = SplitContainer4.Panel2
        End If
        'If tabMain.SelectedTab.Key = "Approve/Reject" Then
        '    grdARTSTMT1.Text = ""
        '    grdARTSTMT1.Parent = tabOrder.Tabs("Aging History").TabPage
        'ElseIf tabMain.SelectedTab.Key = "Credit Limit Reports" Then
        '    grdARTSTMT1.Parent = SplitContainer4.Panel2
        'End If
    End Sub

    Sub Load_Payment_History()
        Call ASCMAIN1.Progress("Now Loading Payment History")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6)
        Call Fill_Records("ARTPYMT2", New String() {YP})

        'Dim PYMT_BATCH_LNO As Int16
        'ASCMAIN1.sql = "Select * from ARTCCPA1 where ARTCCPA1.CCPA_STATUS = 'A' and ARTCCPA1.CCPA_TYPE IN ('S','C') and CUST_CODE = '" & CUST_CODE & "'"
        ''ASCMAIN1.sql = "Select * from ARTCCPA1 where ARTCCPA1.CCPA_STATUS = 'A' and ARTCCPA1.CCPA_TYPE IN ('S','C') and CCPA_NOTE IS NOT NULL and ROWNUM < 10"
        'For Each rowARTCCPA1 As DataRow In ASCDATA1.GetDataTable.Rows
        '    Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
        '    rowARTPYMT2.Item("PYMT_BATCH_NO") = "9999999999"
        '    PYMT_BATCH_LNO += 1
        '    rowARTPYMT2.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
        '    rowARTPYMT2.Item("CUST_CODE") = rowARTCCPA1.Item("CUST_CODE")
        '    rowARTPYMT2.Item("CUST_PYMT_AMT") = rowARTCCPA1.Item("CCPA_AMT")
        '    rowARTPYMT2.Item("PYMT_BATCH_DATE") = rowARTCCPA1.Item("CCPA_DATE_SALE")
        '    rowARTPYMT2.Item("CUST_PYMT_REF_NO") = rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") & ":" & rowARTCCPA1.Item("CUST_CREDIT_CARD_LAST4")
        '    rowARTPYMT2.Item("PYMT_NOTE") = rowARTCCPA1.Item("CCPA_NOTE")
        '    dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)
        'Next

        Sort_grdColumns(grdARTPYMT2, "PYMT_BATCH_NO,PYMT_BATCH_LNO".ToLower)
        tabOrder.Tabs("Recent Payments").Text = "Recent Payments (since " & ASCMAIN1.Get_Legend(YP, False, True) & ")"

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Add_ARTCREDA(ByRef LINE As Int32, ByVal DESCRIPTION As String, ByVal AMOUNT As Decimal)
        LINE += 1
        dst.Tables("ARTCREDA").Rows.Add(New Object() {LINE, DESCRIPTION, AMOUNT})
    End Sub

    Sub Add_ARTCREDL(ByVal CODE As String, ByVal DESCRIPTION As String)
        dst.Tables("ARTCREDL").Rows.Add(New Object() {CODE, "0", DESCRIPTION})
    End Sub

    Private Sub optExceed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCreditReport.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If optCreditReport.Value = "R" Then
            lblExceed.Text = "Days Away"
        Else
            lblExceed.Text = "% of their CR Limit in the past 3 mos"
        End If

        lblExceed.Visible = (optCreditReport.Value <> "CBS")
        numExceed.Visible = (optCreditReport.Value <> "CBS")

        If optCreditReport.Value <> "CBS" Then
            Setup_Exceed_Pct()
        End If

    End Sub

    Sub Setup_Exceed_Pct()
        Select Case optCreditReport.Value
            Case "E"
                numExceed.Value = 90
            Case "N"
                numExceed.Value = 50
            Case "R"
                numExceed.Value = 15
        End Select
    End Sub

    Private Sub grdARTCREDX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCREDX.AfterRowActivate
        If grdARTCREDX.ActiveRow IsNot Nothing AndAlso grdARTCREDX.ActiveRow.IsDataRow Then
            Dim CUST_CODE As String = grdARTCREDX.ActiveRow.Cells("CUST_CODE").Text
            Dim CUST_NAME As String = grdARTCREDX.ActiveRow.Cells("CUST_NAME").Text
            Setup_grdARTSTMT1(CUST_CODE)
            grdARTSTMT1.Text = "Aging History for " & CUST_CODE & ":" & CUST_NAME
            grdARTSTMT1.Visible = True
        Else
            grdARTSTMT1.Visible = False
        End If

    End Sub

    Private Sub grdARTCREDX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTCREDX.InitializeLayout

    End Sub

    Sub Setup_grdARTSTMT1(ByVal CUST_CODE As String)
        ASCMAIN1.sql = "SELECT ARTSTMT1.*, GLTPARM2.LEGEND" _
        & " from ARTSTMT1, GLTPARM2 " _
        & " where GLTPARM2.OPS_YYYYPP (+) = ARTSTMT1.OPS_YYYYPP" _
        & "   and ARTSTMT1.CUST_CODE = '" & CUST_CODE & "'"
        Fill_Records("ARTSTMT1", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdARTSTMT1, "OPS_YYYYPP".ToLower)
    End Sub

    Sub Setup_grdSOTORDRA()
        Dim dvw As DataView = DirectCast(grdSOTORDRA.DataSource, DataTable).DefaultView
        If ScreenMode Then
            'dvw.RowFilter = "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            'grdSOTORDRA.Text = ""
            'grdSOTORDRA.Selected.Rows.Clear()
            'grdSOTORDRA.Rows.ColumnFilters.ClearAllFilters()
            'Show_Filter(grdSOTORDRA, False)
            'grdSOTORDRA.DisplayLayout.GroupByBox.Hidden = True

            'grdSOTORDRA.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        Else
            If tabOrders.SelectedTab.Key = "Pending" Then
                dvw.RowFilter = "ISNULL(ORDR_CRED_CLR_AUTH,'?') = '?'"
                grdSOTORDRA.Text = "Orders to be Reviewed"
                grdSOTORDRA.Parent = tabOrders.SelectedTab.TabPage
            ElseIf tabOrders.SelectedTab.Key = "Reviewed" Then
                dvw.RowFilter = "ISNULL(ORDR_CRED_CLR_AUTH,'?') <> '?'"
                grdSOTORDRA.Text = "Orders already Reviewed"
                grdSOTORDRA.Parent = tabOrders.SelectedTab.TabPage
            End If
        End If
    End Sub

    Private Sub grdSOTORDRA_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRA.InitializeRow

        If tabOrders.SelectedTab.Key = "Pending" Then Exit Sub

        Dim ORDR_CRED_CLR_AUTH As String = e.Row.Cells("ORDR_CRED_CLR_AUTH").Value & ""
        If ORDR_CRED_CLR_AUTH = "" Then ORDR_CRED_CLR_AUTH = "H"

        If ORDR_CRED_CLR_AUTH = "H" Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            If ORDR_CRED_CLR_AUTH = "D" Then
                e.Row.Cells("ORDR_CRED_CLR_AUTH").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Red
            Else
                If ORDR_CRED_CLR_AUTH = "A" Then
                    If e.Row.Cells("ORDR_CREDIT_NOTE").Value & "" <> "" Then
                        e.Row.Cells("ORDR_CRED_CLR_AUTH").Appearance.ForeColor = Drawing.Color.DarkOrange
                        'e.Row.Cells("CRED_CODE").Appearance.ForeColor = Drawing.Color.DarkOrange
                        e.Row.Cells("ORDR_CREDIT_NOTE").Appearance.ForeColor = Drawing.Color.DarkOrange
                        e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.DarkOrange
                    Else
                        e.Row.Cells("ORDR_CRED_CLR_AUTH").Appearance.BackColor = Drawing.Color.LightGreen
                        e.Row.Cells("ORDR_GROUP_NO").Appearance.BackColor = Drawing.Color.LightGreen
                    End If
                Else
                    e.Row.Cells("ORDR_CRED_CLR_AUTH").Appearance.ForeColor = Drawing.Color.Blue
                    e.Row.Cells("ORDR_GROUP_NO").Appearance.ForeColor = Drawing.Color.Blue
                End If
            End If
        End If

        Dim ORDR_TYPE_CODE As String = e.Row.Cells("ORDR_TYPE_CODE").Value & ""
        If ORDR_TYPE_CODE <> "REG" Then
            e.Row.Cells("ORDR_TYPE_CODE").Appearance.BackColor = Drawing.Color.Pink
            e.Row.Cells("ORDR_GROUP_NO").Appearance.BackColor = Drawing.Color.Pink
        End If
    End Sub

    Private Sub tabOrders_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabOrders.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdSOTORDRA()
    End Sub

    Sub Add_Customer(ByVal CUST_CODE As String, ByVal RELATIONSHIP As String, txt As UltraWinEditors.UltraTextEditor)
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        txt.Text = rowARTCUST1.Item("CUST_NAME") & ", " & rowARTCUST1.Item("CUST_CITY") & ", " & rowARTCUST1.Item("CUST_STATE")

        If rowARTCUST1.Item("CUST_CREDIT_HOLD") & "" = "1" Then
            dst.Tables("ARTCREDL").Rows.Find("CRH").Item("SEL") = "1"
        End If
        If rowARTCUST1.Item("CUST_SALES_HOLD") & "" = "1" Then
            dst.Tables("ARTCREDL").Rows.Find("SLH").Item("SEL") = "1"
        End If

        If RELATIONSHIP = "Bill-To" Then
            If rowARTCUST1.Item("TERM_CODE") & "" <> rowSOTORDR1.Item("TERM_CODE") & "" Then
                dst.Tables("ARTCREDL").Rows.Find("TRM").Item("SEL") = "1"
                dst.Tables("ARTCREDL").Rows.Find("TRM").Item("DESCRIPTION") = "Std Terms: " & rowARTCUST1.Item("TERM_CODE")
            End If
        End If

    End Sub

    Private Sub optAction_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAction.ValueChanged
        Select Case optAction.Value
            Case "A"
                txtORDR_CREDIT_NOTE.Visible = False
                lblORDR_CREDIT_NOTE.Visible = False
                cmbCRED_CODE.Visible = False
                REVIEW_RESULTS = "APR"
                REVIEW_RESULTS_DESC = "Approved"
            Case "D"
                txtORDR_CREDIT_NOTE.Visible = True
                lblORDR_CREDIT_NOTE.Visible = True
                lblORDR_CREDIT_NOTE.Text = "Reason for Rejection"
                cmbCRED_CODE.Visible = True
                REVIEW_RESULTS = "REJ"
                REVIEW_RESULTS_DESC = "Rejected"
            Case "C"
                txtORDR_CREDIT_NOTE.Visible = True
                lblORDR_CREDIT_NOTE.Visible = True
                lblORDR_CREDIT_NOTE.Text = "Condition for Approval"
                cmbCRED_CODE.Visible = True
                REVIEW_RESULTS = "CAP"
                REVIEW_RESULTS_DESC = "Cond Appr"
            Case "H"
                txtORDR_CREDIT_NOTE.Visible = True
                lblORDR_CREDIT_NOTE.Visible = True
                lblORDR_CREDIT_NOTE.Text = "Reason Still Pending"
                cmbCRED_CODE.Visible = True
                REVIEW_RESULTS = "PND"
                REVIEW_RESULTS_DESC = "Pending"
        End Select
    End Sub

    Private Sub grdARTCREDL_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCREDL.InitializeRow
        If e.Row.Cells("SEL").Value & "" = "1" Then
            e.Row.Cells("SEL").Appearance.BackColor = Drawing.Color.Red
        Else
            e.Row.Cells("SEL").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub


#Region "grdTATCONV1"

    Function Get_CONV_ATTACHMENTS(ByVal CONV_NO As String) As String
        ASCMAIN1.sql = "Select Count (*) from ASTATTA2 " _
        & " where TABLE_NAME = 'TATCONV1' and COLUMN_NAME = 'CONV_NO' " _
        & " and CODE_VALUE = '" & CONV_NO & "'"
        Dim CONV_ATTACHMENTS As Int64 = Val(ASCDATA1.GetDataValue)
        If CONV_ATTACHMENTS = 0 Then
            Return ""
        Else
            Return CStr(CONV_ATTACHMENTS)
        End If
    End Function

    Private Sub grdTATCONV1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdTATCONV1.ClickCellButton
        If e.Cell.Column.Key = "CONV_ATTACHMENTS" Then
            Dim ENTITY As New Dropped_On_Entity
            ENTITY.TABLE_NAME = "TATCONV1"
            ENTITY.COLUMN_NAME = "CONV_NO"
            ENTITY.CODE_VALUE = e.Cell.Row.Cells("CONV_NO").Value
            Dim DESC_VALUE = "Log by " & e.Cell.Row.Cells("INIT_OPER").Value _
                & " " & Format(e.Cell.Row.Cells.Item("INIT_DATE").Value, "MM/dd/yyyy HH:mm") _
                & " (" & e.Cell.Row.Cells("CONV_SUBJECT").Value & ")"
            ENTITY.DESC_VALUE = DESC_VALUE
            ENTITY.ATTACHMENT_NOTES = ""

            Dim F As New ASFATTA1
            F.ENTITY = ENTITY
            F.ShowDialog()
            F.Dispose()

            grdTATCONV1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        End If
    End Sub

    Private Sub grdTATCONV1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdTATCONV1.DoubleClickRow
        If Not ScreenMode Then
            If e.Row.IsAddRow Or e.Row.IsFilterRow Then
                Exit Sub
            End If
            Dim OPPY_NO As String = e.Row.Cells("TABLE_KEY").Value & ""
            Dim CONV_NO As String = e.Row.Cells("CONV_NO").Value & ""
            Absx1.txtFor("OPPY_NO").Text = OPPY_NO
            Click_Command("View")

            Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Rows.Find(CONV_NO)

            Dim i As Int64 = dst.Tables("TATCONV1").Rows.IndexOf(rowTATCONV1)
            grdTATCONV1.ActiveRow = grdTATCONV1.Rows.GetRowWithListIndex(i)
            grdTATCONV1.Selected.Rows.Add(grdTATCONV1.ActiveRow)
        End If
    End Sub

    Private Sub grdTATCONV1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdTATCONV1.InitializeRow
        Dim CONV_STATUS As String = e.Row.Cells("CONV_STATUS").Value & ""
        If CONV_STATUS = "1" Then
            e.Row.Cells("CONV_FOLLOWUP_BY").Appearance.BackColor = Drawing.Color.LightGreen
            e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.BackColor = Drawing.Color.LightGreen
        Else
            e.Row.Cells("CONV_FOLLOWUP_BY").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.BackColor = Drawing.Color.Empty
        End If

        Dim CONV_NO As String = ""
        Dim CONV_ATTACHMENTS As String = Get_CONV_ATTACHMENTS(e.Row.Cells("CONV_NO").Value)
        If CONV_ATTACHMENTS <> e.Row.Cells("CONV_ATTACHMENTS").Value & "" Then
            e.Row.Cells("CONV_ATTACHMENTS").Value = CONV_ATTACHMENTS
            grdTATCONV1.UpdateData()
        End If

        If e.Row.Cells("TABLE_KEY").Value & "" = "" Then ' Absx1.txtFor("CUST_CREDIT_GROUP_CUST").Text Then
            e.Row.Cells("INIT_DATE").Appearance.BackColor = Drawing.Color.LightYellow
        ElseIf e.Row.Cells("TABLE_KEY").Value & "" = Absx1.txtFor("CUST_BILL_TO_CUST").Text Then
            e.Row.Cells("INIT_DATE").Appearance.BackColor = Drawing.Color.LightPink
        Else
            e.Row.Cells("INIT_DATE").Appearance.BackColor = Drawing.Color.LightGreen
        End If

    End Sub
#End Region

    Sub Generate_emails()

        ' IF MULTIPLE ORDERS WERE REVIEWED, THEN MULTIPLE ORDERS SHOULD BE LISTED
        ' NEED TO CHECK THAT THE EVENT LOG IN SOINQ SHOWS THE EMAIL EVENT
        ' MAKE SURE SO INQ ALLOWS RIGHT CLICK TO SHOW THE EMAIL, LIKE PO
        For Each rowSOTORDR1 In dst.Tables("SOTORDR1").Rows
            Using frmTAFSEND1 As New TAFSEND1(Me)
                With frmTAFSEND1
                    ORDR_NO = rowSOTORDR1.Item("ORDR_NO")
                    If rowSOTORDR1.Item("INIT_OPER") & "" <> "" Then
                        Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", rowSOTORDR1.Item("INIT_OPER"))
                        If rowASTUSER1 IsNot Nothing AndAlso rowASTUSER1.Item("USER_EMAIL") & "" <> "" Then
                            If Not .SEND_TOs.ContainsKey(rowASTUSER1.Item("USER_EMAIL") & "") Then
                                .SEND_TOs.Add(rowASTUSER1.Item("USER_EMAIL") & "", "")
                            End If
                        End If
                    End If
                    If rowSOTORDR1.Item("LAST_OPER") & "" <> "" Then
                        Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", rowSOTORDR1.Item("LAST_OPER"))
                        If rowASTUSER1 IsNot Nothing AndAlso rowASTUSER1.Item("USER_EMAIL") & "" <> "" Then
                            If Not .SEND_TOs.ContainsKey(rowASTUSER1.Item("USER_EMAIL") & "") Then
                                .SEND_TOs.Add(rowASTUSER1.Item("USER_EMAIL") & "", "")
                            End If
                        End If
                    End If
                    If rowSOTORDR1.Item("SREP_CODE") & "" <> "" Then
                        ASCMAIN1.sql = "Select ASTUSER1.* from ASTUSER1,TATUSER1 " _
                        & " where TATUSER1.SREP_CODE = :PARM1 and ASTUSER1.USER_ID = TATUSER1.USER_ID"
                        Dim rowASTUSER1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {rowSOTORDR1.Item("SREP_CODE")})
                        If rowASTUSER1 IsNot Nothing AndAlso rowASTUSER1.Item("USER_EMAIL") & "" <> "" Then
                            If Not .SEND_TOs.ContainsKey(rowASTUSER1.Item("USER_EMAIL") & "") Then
                                .SEND_TOs.Add(rowASTUSER1.Item("USER_EMAIL") & "", "")
                            End If
                        End If
                    End If

                    .EMAIL_KEY = "CR_REJ"
                    .SEND_FROM = "donotreply" & "@" & ASCMAIN1.rowASTPARM1.Item("AS_PARM_DEFAULT_EMAIL_DOMAIN")
                    .SEND_FROM_NAME = ASCMAIN1.USER_NAME
                    .SEND_TO = ""
                    .SEND_TO_NAME = ""
                    .SEND_SUBJECT = "Credit Review Result for Sales Order " & ORDR_NO & ":" & optAction.Text
                    .SEND_BODY = "Sales Order " & ORDR_NO _
                        & " was Reviewed by " & ASCMAIN1.USER_NAME & " " & Format(Now, "MM/dd/yy HH:mm") _
                        & vbCrLf _
                        & "Reason Given for Credit Review: " & cmbCRED_CODE.Text & vbCrLf & txtORDR_CREDIT_NOTE.Text
                    '.SEND_ATTACHMENT = PO_Discrepancy_Report
                    .SEND_METHOD = "E"
                    .SEND_ENTITY_CAPTION = "Customer"
                    .SEND_ENTITY_TABLE = "ARTCUST1"
                    .SEND_ENTITY_KEY = Absx1.txtFor("CUST_CODE").Text
                    .SEND_ENTITY_NAME = rowSOTORDR1.Item("CUST_NAME") & ""
                    '.ShowDialog()
                    '.Send_email()
                    .Send_email_automatically(False)

                    If .SEND_STATUS <> "C" Then
                        TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO, _
                                       Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, _
                                       "EML", _
                                        "Credit Exception", _
                                       .SEND_NO)
                    End If
                End With
            End Using
        Next
    End Sub
End Class