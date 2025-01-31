Public Class SORSORD1
    ' hardcoding below for walmart, kmart, sears

    Dim REPORT_DATE0 As Date
    Dim REPORT_DATE1 As Date
    Dim dtSOTRSRV1 As New DataTable
    Dim ICTSTATDSQL As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Absx1.dteFor("DTE0").Value = Now.Date
        Absx1.dteFor("DTE1").Value = Now.Date
        ASCMAIN1.sql = "Select RSRV_NO from SOTRSRV1 WHERE ROWNUM < 0"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTRSRV1", 1))
        dtSOTRSRV1 = dst.Tables("SOTRSRV1").Clone
        grdSOTRSRV1.DataSource = dtSOTRSRV1
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Absx1.chkFor("CHKTRANONLY").Visible = True
            Absx1.chkFor("CHKSTYLESTATS").Visible = True
        Else
            Absx1.chkFor("CHKTRANONLY").Visible = False
            Absx1.chkFor("CHKSTYLESTATS").Visible = False

        End If
        Absx1.chkFor("CHKTRANONLY").Checked = False
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        REPORT_DATE0 = Absx1.dteFor("DTE0").Value
        REPORT_DATE1 = Absx1.dteFor("DTE1").Value

        Dim CONDITION As String = ""
        If optSR.Value = "A" Then
            CONDITION = " 1=1"
        Else
            CONDITION = " SOTORDR1." _
                & IIf(optSR.Value = "R", "ORDR_DATE_RECD", "ORDR_SHIP_DATE") _
                & IIf(Format(REPORT_DATE0, "dd-MMM-yyyy") = Format(REPORT_DATE1, "dd-MMM-yyyy"),
                  " = '" & Format(REPORT_DATE0, "dd-MMM-yyyy") & "'",
                  " between '" & Format(REPORT_DATE0, "dd-MMM-yyyy") & "' and '" & Format(REPORT_DATE1, "dd-MMM-yyyy") & "'")
        End If

        Dim sqlw As String = ASCMAIN1.SQL_Add_WHERE(" AND " & CONDITION _
             & " and (SOTORDR1.ORDR_STATUS = 'O' OR SOTORDR1.ORDR_STATUS = 'P' OR SOTORDR1.ORDR_STATUS = 'F')")

        If Absx1.chkFor("CHKEDI_ONLY").Checked Then
            sqlw &= " and SOTORDR1.ORDR_SOURCE = 'E' "
        End If

        'sql &= SQL_in("CUST_CODE", "SOTORDR1.CUST_CODE")
        MyBase.Get_SQL("*")
        sqlw &= sql_WHERE & sql_JOIN
        If Absx1.optFor("OPTORB").Value = "R" Then
            sqlw &= " and ROWNUM < 1"
        End If

        ASCMAIN1.Progress("Order Summary", "")

        ASCMAIN1.sql = "Select 'O' ORDR_GROUP_TYPE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", NVL(SOTORDR1.CUST_DC_NO,'XXXXXX') CUST_DC_NO" & vbCrLf _
            & ", Min (SOTORDR1.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", Min (SOTORDR1.ORDR_DATE) ORDR_DATE" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_SHIP_DATE) ORDR_SHIP_DATE" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_CANCEL_DATE) ORDR_CANCEL_DATE" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_CUST_PO) ORDR_CUST_PO" & vbCrLf _
            & ", Count (*) ORDR_CNT" & vbCrLf _
            & ", MIN (SOTORDR1.SALES_DIVISION_CODE) SALES_DIVISION_CODE" & vbCrLf _
            & ", MIN (EDI_APPOINTMENT) EDI_APPOINTMENT" & vbCrLf _
            & ", MIN (ORDR_DEPT) ORDR_DEPT" & vbCrLf _
            & " from SOTORDR1" & vbCrLf _
            & sqlw & vbCrLf _
            & " GROUP BY SOTORDR1.ORDR_GROUP_NO, NVL(SOTORDR1.CUST_DC_NO,'XXXXXX')"
        Dim SOTORDR0 As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_TYPE,ORDR_GROUP_NO,CUST_DC_NO)")



        Dim CONDITION_R As String = ""
        Dim sqlw_r As String = ""
        If Absx1.optFor("OPTORB").Value = "R" Or Absx1.optFor("OPTORB").Value = "B" Then
            If optSR.Value = "A" Then
                CONDITION_R = " 1=1"
            Else
                CONDITION_R = " SOTRSRV1." _
                    & IIf(optSR.Value = "R", "INIT_DATE", "ORDR_SHIP_DATE") _
                    & IIf(Format(REPORT_DATE0, "dd-MMM-yyyy") = Format(REPORT_DATE1, "dd-MMM-yyyy"),
                      " = '" & Format(REPORT_DATE0, "dd-MMM-yyyy") & "'",
                      " between '" & Format(REPORT_DATE0, "dd-MMM-yyyy") & "' and '" & Format(REPORT_DATE1, "dd-MMM-yyyy") & "'")
            End If

            sqlw_r = ASCMAIN1.SQL_Add_WHERE(" AND " & CONDITION_R _
                 & " and (SOTRSRV1.RSRV_STATUS = 'O')")

            'If Absx1.chkFor("CHKEDI_ONLY").Checked Then
            '    sqlw &= " and SOTORDR1.ORDR_SOURCE = 'E' "
            'End If

            sqlw_r &= Replace(sql_WHERE & sql_JOIN, "SOTORDR1", "SOTRSRV1")

            If dtSOTRSRV1.Rows.Count > 0 Then
                Dim fltrList As String = ""
                For Each rowSOTRSRV1 As DataRow In dtSOTRSRV1.Select()
                    fltrList = fltrList & String.Format("'{0}',", rowSOTRSRV1.Item("RSRV_NO").ToString & String.Empty)
                Next
                sqlw_r &= " AND SOTRSRV1.RSRV_NO IN (" & fltrList.Substring(0, fltrList.Length - 1) & ")"
            End If


            ASCMAIN1.Progress("Reservations", "")

            ASCMAIN1.sql = "Select 'R' ORDR_GROUP_TYPE, SOTRSRV1.RSRV_NO ORDR_GROUP_NO" & vbCrLf _
                & ", 'XXXXXX' CUST_DC_NO" & vbCrLf _
                & ", Min (SOTRSRV1.CUST_CODE) CUST_CODE" & vbCrLf _
                & ", Min (TRUNC(SOTRSRV1.INIT_DATE)) ORDR_DATE" & vbCrLf _
                & ", MIN (SOTRSRV1.ORDR_SHIP_DATE) ORDR_SHIP_DATE" & vbCrLf _
                & ", MIN (SOTRSRV1.ORDR_CANCEL_DATE) ORDR_CANCEL_DATE" & vbCrLf _
                & ", MIN (SOTRSRV1.ORDR_CUST_PO) ORDR_CUST_PO" & vbCrLf _
                & ", Count (*) ORDR_CNT" & vbCrLf _
                & ", MIN (SOTRSRV1.SALES_DIVISION_CODE) SALES_DIVISION_CODE" & vbCrLf _
                & ", 'R' EDI_APPOINTMENT" & vbCrLf _
                & ", MIN (SOTRSRV1.ORDR_DEPT) ORDR_DEPT" & vbCrLf _
                & " from SOTRSRV1" & vbCrLf _
                & sqlw_r & vbCrLf _
                & " group by SOTRSRV1.RSRV_NO"
            ASCMAIN1.sql = "Insert into " & SOTORDR0 & " " & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()
        End If


        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_QTY NUMBER (6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_AMT NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add GROUP_SEQ NUMBER (6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add MABD VARCHAR2(10)")






        If Absx1.optFor("OPTCONS_GROUPS").Value = "C" Then
            ASCDATA1.ExecuteSQL("Update " & SOTORDR0 & " Set GROUP_SEQ = 0")
        ElseIf Absx1.optFor("OPTCONS_GROUPS").Value = "D" Then
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is Select X.*, ROWNUM GROUP_SEQ from " & vbCrLf _
                & "  (Select Distinct CUST_CODE, NVL(ORDR_DEPT,'X') ORDR_DEPT from " & SOTORDR0 & ") X;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDR0 & " Set GROUP_SEQ = R1.GROUP_SEQ" & vbCrLf _
                & "    where CUST_CODE = R1.CUST_CODE and NVL(ORDR_DEPT,'X') = R1.ORDR_DEPT;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        ElseIf Absx1.optFor("OPTCONS_GROUPS").Value = "DC" Then
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is Select X.*, ROWNUM GROUP_SEQ from " & vbCrLf _
                & "  (Select Distinct CUST_CODE, NVL(CUST_DC_NO,'X') CUST_DC_NO from " & SOTORDR0 & ") X;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDR0 & " Set GROUP_SEQ = R1.GROUP_SEQ" & vbCrLf _
                & "    where CUST_CODE = R1.CUST_CODE and NVL(CUST_DC_NO,'X') = R1.CUST_DC_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        ElseIf Absx1.optFor("OPTCONS_GROUPS").Value = "N" Then
            ASCDATA1.ExecuteSQL("Update " & SOTORDR0 & " Set GROUP_SEQ = ROWNUM")
        End If
        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO" _
            & ", NVL(SOTORDR1.CUST_DC_NO,'XXXXXX') CUST_DC_NO" _
            & ", SUM (ORDR_QTY) ORDR_QTY, SUM (ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT" _
            & " from SOTORDR1,SOTORDR2" _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & " and (CUST_CODE NOT IN ('BEALLS','ROSS') OR SOTORDR2.ORDR_QTY <> SOTORDR2.ORDR_QTY_CANC)" _
            & " and SOTORDR1.ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR0 & " where ORDR_GROUP_TYPE = 'O')" _
            & " GROUP BY SOTORDR1.ORDR_GROUP_NO, NVL(SOTORDR1.CUST_DC_NO,'XXXXXX');"
        Else
            ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO" _
            & ", NVL(SOTORDR1.CUST_DC_NO,'XXXXXX') CUST_DC_NO" _
            & ", SUM (ORDR_QTY) ORDR_QTY, SUM (ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT" _
            & " from SOTORDR1,SOTORDR2" _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & " and SOTORDR1.ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR0 & " where ORDR_GROUP_TYPE = 'O')" _
            & " GROUP BY SOTORDR1.ORDR_GROUP_NO, NVL(SOTORDR1.CUST_DC_NO,'XXXXXX');"

        End If

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is " & ASCMAIN1.sql _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update " & SOTORDR0 & " Set " _
            & "     ORDR_QTY = R1.ORDR_QTY" _
            & "    ,ORDR_AMT = R1.ORDR_AMT" _
            & "    where ORDR_GROUP_TYPE = 'O' and ORDR_GROUP_NO = R1.ORDR_GROUP_NO" _
            & "      and CUST_DC_NO = R1.CUST_DC_NO;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()


        If Absx1.optFor("OPTORB").Value = "R" Or Absx1.optFor("OPTORB").Value = "B" Then
            ASCMAIN1.sql = "Select SOTRSRV1.RSRV_NO ORDR_GROUP_NO" _
              & ", 'XXXXXX' CUST_DC_NO" _
              & ", SUM (RSRV_QTY) ORDR_QTY, SUM (RSRV_QTY * ORDR_UNIT_PRICE) ORDR_AMT" _
              & " from SOTRSRV1,SOTRSRV2" _
              & " where SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO" _
              & "   and SOTRSRV1.RSRV_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR0 & " where ORDR_GROUP_TYPE = 'R')" _
              & " GROUP BY SOTRSRV1.RSRV_NO;"

            ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is " & ASCMAIN1.sql _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Update " & SOTORDR0 & " Set " _
                & "     ORDR_QTY = R1.ORDR_QTY" _
                & "    ,ORDR_AMT = R1.ORDR_AMT" _
                & "    where ORDR_GROUP_TYPE = 'R' and ORDR_GROUP_NO = R1.ORDR_GROUP_NO" _
                & "      and CUST_DC_NO = R1.CUST_DC_NO;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & SOTORDR0, "SOTORDR0", 3))

        dst.Tables.Add(ASCDATA1.GetDataTable("Select Distinct CUST_CODE, GROUP_SEQ from " & SOTORDR0, "SOTSORD0", 2))

        ASCMAIN1.Progress("Order Information", "")
        Dim sqlO As String
        If ASCMAIN1.CLIENT = "VAN" Then
            sqlO = "Select SOTORDR0.CUST_CODE, SOTORDR0.GROUP_SEQ, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'WALMART', NULL" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'KMART', NULL" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'SEARS', NULL, NVL(SOTORDR2.RANGE_STYLE_CODE,EDT850T2.EDI_LBL_CODE)))) RANGE_STYLE_CODE" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY ELSE 0 END) ORDR_QTY_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_OPEN ELSE 0 END) ORDR_QTY_OPEN_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_PICK ELSE 0 END) ORDR_QTY_PICK_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_SHIP ELSE 0 END) ORDR_QTY_SHIP_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_CANC ELSE 0 END) ORDR_QTY_CANC_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_OPEN_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_PICK_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_SHIP_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_CANC_X" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & " from SOTORDR1, SOTORDR2, " & SOTORDR0 & " SOTORDR0, EDT850T2" & vbCrLf _
            & sql_TABLE_NAMEs & vbCrLf _
            & sqlw & vbCrLf _
            & " and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " and SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & " and NVL(SOTORDR0.CUST_DC_NO,'XXXXXX') = NVL(SOTORDR1.CUST_DC_NO,'XXXXXX')" & vbCrLf _
            & " and SOTORDR2.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO(+)" & vbCrLf _
            & " and SOTORDR2.EDI_DTL_SEQ = EDT850T2.EDI_DTL_SEQ(+)" & vbCrLf _
            & " group by SOTORDR0.CUST_CODE, SOTORDR0.GROUP_SEQ, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'WALMART', NULL" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'KMART', NULL" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'SEARS', NULL, NVL(SOTORDR2.RANGE_STYLE_CODE,EDT850T2.EDI_LBL_CODE))))" & vbCrLf _
            & " having SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY ELSE 0 END) <> 0"
        Else
            sqlO = "Select SOTORDR0.CUST_CODE, SOTORDR0.GROUP_SEQ, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'WALMART', NULL" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'KMART', NULL" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'SEARS', NULL, nvl(SOTORDR2.RANGE_STYLE_CODE,EDT850T2.EDI_LBL_CODE)))) RANGE_STYLE_CODE" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY ELSE 0 END) ORDR_QTY_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_OPEN ELSE 0 END) ORDR_QTY_OPEN_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_PICK ELSE 0 END) ORDR_QTY_PICK_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_SHIP ELSE 0 END) ORDR_QTY_SHIP_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_CANC ELSE 0 END) ORDR_QTY_CANC_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_OPEN_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_PICK_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_SHIP_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_CANC_X" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & " from SOTORDR1, SOTORDR2, EDT850T2, " & SOTORDR0 & " SOTORDR0" & vbCrLf _
            & sql_TABLE_NAMEs & vbCrLf _
            & sqlw & vbCrLf _
            & " and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " and SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & " and NVL(SOTORDR0.CUST_DC_NO,'XXXXXX') = NVL(SOTORDR1.CUST_DC_NO,'XXXXXX')" & vbCrLf _
            & " and SOTORDR2.EDI_DOC_SEQ_NO =  EDT850T2.EDI_DOC_SEQ_NO(+) and SOTORDR2.EDI_DTL_SEQ =  EDT850T2.EDI_DTL_SEQ(+)" & vbCrLf _
            & " group by SOTORDR0.CUST_CODE, SOTORDR0.GROUP_SEQ, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'WALMART', NULL" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'KMART', NULL" & vbCrLf _
            & ", DECODE (SOTORDR1.CUST_CODE, 'SEARS', NULL, nvl(SOTORDR2.RANGE_STYLE_CODE,EDT850T2.EDI_LBL_CODE))))" & vbCrLf _
            & " having SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY ELSE 0 END) <> 0"
        End If


        Dim sqlR = "Select SOTORDR0.CUST_CODE, SOTORDR0.GROUP_SEQ, SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE" & vbCrLf _
            & ", NULL RANGE_STYLE_CODE" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN SOTRSRV2.RSRV_QTY ELSE 0 END) RSRV_QTY_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN SOTRSRV2.RSRV_QTY_OPEN ELSE 0 END) RSRV_QTY_OPEN_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN 0 ELSE 0 END) RSRV_QTY_PICK_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN SOTRSRV2.RSRV_QTY_USED ELSE 0 END) RSRV_QTY_SHIP_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN SOTRSRV2.RSRV_QTY_CANC ELSE 0 END) RSRV_QTY_CANC_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN SOTRSRV2.RSRV_QTY * SOTRSRV2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN SOTRSRV2.RSRV_QTY_OPEN * SOTRSRV2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_OPEN_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN 0 * SOTRSRV2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_PICK_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN SOTRSRV2.RSRV_QTY_USED * SOTRSRV2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_SHIP_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION_R & " THEN SOTRSRV2.RSRV_QTY_CANC * SOTRSRV2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_CANC_X" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY) RSRV_QTY" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY_OPEN) RSRV_QTY_OPEN" & vbCrLf _
            & ", SUM (0) RSRV_QTY_PICK" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY_USED) RSRV_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTRSRV2.RSRV_QTY_CANC) RSRV_QTY_CANC" & vbCrLf _
            & " from SOTRSRV1, SOTRSRV2, " & SOTORDR0 & " SOTORDR0" & vbCrLf _
            & sql_TABLE_NAMEs & vbCrLf _
            & sqlw_r & vbCrLf _
            & " and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
            & " and SOTORDR0.ORDR_GROUP_NO = SOTRSRV1.RSRV_NO" & vbCrLf _
            & " group by SOTORDR0.CUST_CODE, SOTORDR0.GROUP_SEQ, SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE"

        ASCMAIN1.sql = "Select CUST_CODE, GROUP_SEQ, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", RANGE_STYLE_CODE" & vbCrLf _
            & ", SUM (ORDR_QTY_X) ORDR_QTY_X" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN_X) ORDR_QTY_OPEN_X" & vbCrLf _
            & ", SUM (ORDR_QTY_PICK_X) ORDR_QTY_PICK_X" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP_X) ORDR_QTY_SHIP_X" & vbCrLf _
            & ", SUM (ORDR_QTY_CANC_X) ORDR_QTY_CANC_X" & vbCrLf _
            & ", SUM (ORDR_AMT_X) ORDR_AMT_X" & vbCrLf _
            & ", SUM (ORDR_AMT_OPEN_X) ORDR_AMT_OPEN_X" & vbCrLf _
            & ", SUM (ORDR_AMT_PICK_X) ORDR_AMT_PICK_X" & vbCrLf _
            & ", SUM (ORDR_AMT_SHIP_X) ORDR_AMT_SHIP_X" & vbCrLf _
            & ", SUM (ORDR_AMT_CANC_X) ORDR_AMT_CANC_X" & vbCrLf _
            & ", SUM (ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & " from (" & vbCrLf _
            & sqlO

        If Absx1.optFor("OPTORB").Value = "R" Or Absx1.optFor("OPTORB").Value = "B" Then
            ASCMAIN1.sql &= vbCrLf & " union " & vbCrLf & sqlR & vbCrLf
        End If

        ASCMAIN1.sql &= ") group by CUST_CODE, GROUP_SEQ, STYLE_CODE, COLOR_CODE, RANGE_STYLE_CODE"

        Dim SOTSORD1 As String = ASCMAIN1.Temp_Table

        If ASCMAIN1.CLIENT = "VAN" Then
            sql = "Delete From " & SOTSORD1 & " WHERE CUST_CODE IN ('BEALLS','ROSS') AND ORDR_QTY_X = ORDR_QTY_CANC_X"
            ASCDATA1.ExecuteSQL(sql)
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            sql = "Update " & SOTSORD1 & " Set ORDR_QTY_X = ORDR_QTY_OPEN_X  + ORDR_QTY_PICK_X, ORDR_AMT_X = ORDR_AMT_X - ORDR_AMT_CANC_X  WHERE CUST_CODE IN ('BEALLS','ROSS') AND ORDR_QTY_X <> ORDR_QTY_CANC_X and ORDR_QTY_CANC_X <> 0 and (ORDR_QTY_OPEN_X <> 0 OR ORDR_QTY_PICK_X <> 0)"
            ASCDATA1.ExecuteSQL(sql)
        End If

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & SOTSORD1, "SOTSORD1", 0))

        ASCMAIN1.Progress("WIP && In Transit", "")
        ASCMAIN1.sql = "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_ETA, POTSHIP3.PO_QTY_SHP" & vbCrLf _
            & " From POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2" & vbCrLf _
            & " Where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTORDR2.STYLE_CODE in " & vbCrLf _
            & " (Select Distinct STYLE_CODE from " & SOTSORD1 & ")" & vbCrLf
        dst.Tables.Add(ASCDATA1.GetDataTable("", "POTSHIPX", 0))


        ASCMAIN1.Progress("Master Files", "")
        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from SOTSDIV1", "SOTSDIV1", 1))


        ASCMAIN1.sql = "Select * from ARTCUST1" _
            & " where CUST_CODE in (Select DISTINCT CUST_CODE from " & SOTSORD1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTCUST1", 1))


        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC " & vbCrLf _
            & " from ICTSTYL1 " & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE in " & vbCrLf _
            & " (Select Distinct STYLE_CODE from " & SOTSORD1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTSTYL1", 1))


        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then

            'ASCMAIN1.sql = "Select ICTBODY2.SUB_BODY_CODE,ICTBODY2.SUB_BODY_DESC,ICTBODY2.MASTER_BODY_CODE" & vbCrLf _
            '& " FROM ICTSTYL1,ICTBODY2 " & vbCrLf _
            '& " where ICTBODY2.SUB_BODY_CODE = ICTSTYL1.SUB_BODY_CODE" & vbCrLf _
            '& " AND ICTSTYL1.STYLE_CODE in " & vbCrLf _
            '& " (Select Distinct STYLE_CODE from " & SOTSORD1 & ")" & vbCrLf _
            '& " GROUP BY ICTBODY2.SUB_BODY_CODE,ICTBODY2.SUB_BODY_DESC,ICTBODY2.MASTER_BODY_CODE HAVING MAX(NVL(STANDARD_CUBE_PER_UNIT,0)) = 0"
            'dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTBODYX", 1))
            If chkStyleStats.Checked Then
                ASCMAIN1.sql = "Select * from (" & vbCrLf _
            & " Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.INIT_DATE, POTSHIP1.WHSE_CODE, POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & ", POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
            & ", POTORDR1.FACTORY_CODE, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
            & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO" & vbCrLf _
            & ", POTSHIP2.PO_DATE_RECEIVED" & vbCrLf _
            & ", POTSHIP3.PO_QTY_SHP, POTSHIP3.PO_QTY_REC" & vbCrLf _
            & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
            & ", POTORDR2.PO_QTY_ORD, 0 PO_QTY_OPN" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
            & ", POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY" & vbCrLf _
            & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
            & "From POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2, ICTATOP2" & vbCrLf _
            & "Where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & " And POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & " And POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & " And POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "  And POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & " And POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "  And ICTATOP2.PS_CODE (+) = 'S'" & vbCrLf _
            & " And ICTATOP2.PS_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & " And POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & " And (POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE) IN  (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & SOTSORD1 & ")" & vbCrLf _
            & " ) union (" & vbCrLf _
            & "Select  POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE,POTORDR1.INIT_DATE, POTORDR1.WHSE_CODE, POTORDR2.PO_ORDER_NO" & vbCrLf _
            & ", POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
            & ", POTORDR1.FACTORY_CODE, POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & ", Null PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO" & vbCrLf _
            & ", Decode(nvl(POTORDR2.PO_QTY_OPN,0),0,'ClosedPO','OpenPO') PO_SHIP_VESSEL" & vbCrLf _
            & ", POTORDR2.PO_DATE_SHIP_BY, POTORDR2.PO_DATE_ETA" & vbCrLf _
            & ", 10 PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
            & ", Null PO_SHIP_REF_NO, Null CONTAINER_NO" & vbCrLf _
            & ", NULL PO_DATE_RECEIVED" & vbCrLf _
            & ", 0 PO_QTY_SHP, 0 PO_QTY_REC" & vbCrLf _
            & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
            & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN" & vbCrLf _
            & ", POTORDR2.PO_DATE_ETA + 10 PO_ARRIVAL_DATE" & vbCrLf _
            & ", POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY" & vbCrLf _
            & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
            & " From POTORDR1, POTORDR2, ICTATOP2" & vbCrLf _
            & "Where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
            & "  And ICTATOP2.PS_CODE (+) = 'P'" & vbCrLf _
            & "   And ICTATOP2.PS_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   And POTORDR2.PO_QTY_OPN <> 0" & vbCrLf _
            & " And (POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE) IN  (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & SOTSORD1 & ")" & vbCrLf _
            & ")"

                ICTSTATDSQL = ASCMAIN1.sql
                Create_TDA(dst.Tables.Add, "ICTSTATD", "**", 0, False)

                Fill_Records("ICTSTATD", "", True, ICTSTATDSQL)


            End If
        End If



        ASCMAIN1.Progress("Inventory Status", "")

        'ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
        '    & ", 0 WHSE_QTY_ON_HAND, 0 WHSE_QTY_PICK" & vbCrLf _
        '    & ", 0 WHSE_QTY_ON_ORDER, 0 WHSE_QTY_TRAN" & vbCrLf _
        '    & ", 0 SOON_SHIP_QTY, SYSDATE SOON_SHIP_DATE, 0 REC_LW" & vbCrLf _
        '    & " from ICTSTAT2 " & vbCrLf _
        '    & " WHERE ROWNUM < 0"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSORD2", 2))

        ASCMAIN1.sql = "Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE" & vbCrLf _
            & ", SUM(ICTSTAT2.WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND, SUM(ICTSTAT2.WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
            & ", SUM(ICTSTAT2.WHSE_QTY_ON_ORDER) WHSE_QTY_ON_ORDER, SUM(ICTSTAT2.WHSE_QTY_TRAN) WHSE_QTY_TRAN" & vbCrLf _
            & ", SYSDATE SOON_SHIP_DATE" & vbCrLf _
            & " from ICTSTAT2 " & vbCrLf _
            & " where ICTSTAT2.STYLE_CODE in " & vbCrLf _
            & " (Select Distinct STYLE_CODE from " & SOTSORD1 & ")" & vbCrLf _
            & " GROUP BY ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", 0 WHSE_QTY_ON_HAND, 0 WHSE_QTY_PICK, 0 WHSE_QTY_ON_ORDER, 0 WHSE_QTY_TRAN, NULL SOON_SHIP_DATE from (" & vbCrLf _
            & "Select Distinct STYLE_CODE, COLOR_CODE from " & SOTSORD1 & vbCrLf _
            & " minus" & vbCrLf _
            & "Select Distinct STYLE_CODE, COLOR_CODE from ICTSTAT2)"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSORD2", 2))
        With dst.Tables("SOTSORD2")
            .Columns("WHSE_QTY_PICK").ReadOnly = False
            .Columns.Add("SOON_SHIP_QTY", GetType(System.Int64))
            .Columns.Add("REC_LW", GetType(System.Int64))
        End With


        ASCMAIN1.sql = "Select ICTTRAN2.STYLE_CODE, ICTTRAN2.COLOR_CODE" & vbCrLf _
            & ", SUM (ICTTRAN2.TRAN_QTY) TRAN_QTY" & vbCrLf _
            & " from ICTTRAN2" & vbCrLf _
            & " where (ICTTRAN2.STYLE_CODE,ICTTRAN2.COLOR_CODE) in (Select Distinct STYLE_CODE, COLOR_CODE from " & SOTSORD1 & ")" & vbCrLf _
            & "   and TRAN_TYPE = 'R'" & vbCrLf _
            & "   and TRAN_NO in (" & vbCrLf _
            & "Select TRAN_NO from ICTTRAN1 where ICTTRAN1.TRAN_TYPE = 'R'    " & vbCrLf _
            & "   and ICTTRAN1.TRAN_DATE +7 >= '" & Format(REPORT_DATE1, "dd-MMM-yyyy") & "'" & vbCrLf _
            & "   and ICTTRAN1.TRAN_DATE <= '" & Format(REPORT_DATE0, "dd-MMM-yyyy") & "'" & vbCrLf _
            & "   and ICTTRAN1.TRAN_STATUS_UPD = 'U'" & vbCrLf _
            & ") GROUP BY STYLE_CODE, COLOR_CODE"
        'Replaced this line above because it was brining in too many colors and hit an error with no records in the loop below. WR-12/5/16
        '& " where ICTTRAN2.STYLE_CODE in (Select Distinct STYLE_CODE from " & SOTSORD1 & ")" & vbCrLf _

        For Each rowICTTRANX As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim rowSOTSORD2 As DataRow = dst.Tables("SOTSORD2").Rows.Find(New Object() {rowICTTRANX.Item("STYLE_CODE"), rowICTTRANX.Item("COLOR_CODE")})
            rowSOTSORD2.Item("REC_LW") = Val(rowICTTRANX.Item("TRAN_QTY") & "")
        Next

        ASCMAIN1.sql = "Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE" & vbCrLf _
            & ", SUM (ICTSTAT2.WHSE_QTY_OPEN) WHSE_QTY_OPEN from ICTSTAT2 " & vbCrLf _
            & " where ICTSTAT2.STYLE_CODE in " & vbCrLf _
            & " (Select Distinct STYLE_CODE from " & SOTSORD1 & ")" & vbCrLf _
            & " group by ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", 0 WHSE_QTY_OPEN  from (" & vbCrLf _
            & "Select Distinct STYLE_CODE, COLOR_CODE from " & SOTSORD1 & vbCrLf _
            & " minus" & vbCrLf _
            & "Select Distinct STYLE_CODE, COLOR_CODE from ICTSTAT2)"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTSTATX", 2))
        With dst.Tables("ICTSTATX")
            .Columns("WHSE_QTY_OPEN").ReadOnly = False
        End With

        ' Put This Order's Pick Qty & Amt into Open

        Create_Relation("SOTSORD2", "SOTSORD1", "STYLE_CODE,COLOR_CODE")
        dst.Tables("SOTSORD2").Columns.Add("ORDR_QTY_PICK_X", GetType(System.Int64), "SUM(CHILD(SOTSORD2_SOTSORD1).ORDR_QTY_PICK_X)")

        If "" = "gabe saks ok" Then
            For Each rowSOTSORD2 As DataRow In dst.Tables("SOTSORD2").Select("ORDR_QTY_PICK_X <> 0")
                rowSOTSORD2.Item("WHSE_QTY_PICK") = Val(rowSOTSORD2.Item("WHSE_QTY_PICK") & "") _
                                                  - Val(rowSOTSORD2.Item("ORDR_QTY_PICK_X") & "")
            Next
        End If

        'For Each rowSOTSORD1 As DataRow In dst.Tables("SOTSORD1").Select("ORDR_QTY_PICK_X <> 0")
        '    ' Dim rowSOTSORD2 As DataRow = dst.Tables("SOTSORD2").Rows.Find(New Object() {rowSOTSORD1.Item("STYLE_CODE"), rowSOTSORD1.Item("COLOR_CODE")})
        '    Dim rowSOTSORD2 As DataRow = rowSOTSORD1.GetParentRow("SOTSORD2_SOTSORD1")
        '    rowSOTSORD2.Item("WHSE_QTY_PICK") = Val(rowSOTSORD2.Item("WHSE_QTY_PICK") & "") _
        '                                      - Val(rowSOTSORD1.Item("ORDR_QTY_PICK_X") & "")
        'Next

        If "" = "gabe saks ok" Then
            For Each rowSOTSORD1 As DataRow In dst.Tables("SOTSORD1").Select("ORDR_QTY_PICK_X <> 0")
                With rowSOTSORD1
                    .Item("ORDR_QTY_OPEN_X") = Val(.Item("ORDR_QTY_OPEN_X") & "") _
                                             + Val(.Item("ORDR_QTY_PICK_X") & "")
                    .Item("ORDR_AMT_OPEN_X") = Val(.Item("ORDR_AMT_OPEN_X") & "") _
                                             + Val(.Item("ORDR_AMT_PICK_X") & "")
                    .Item("ORDR_QTY_OPEN") = Val(.Item("ORDR_QTY_OPEN") & "") _
                                             + Val(.Item("ORDR_QTY_PICK_X") & "")
                End With
            Next
        End If



        Create_Relation("ICTSTATX", "SOTSORD1", "STYLE_CODE,COLOR_CODE")
        dst.Tables("ICTSTATX").Columns.Add("ORDR_QTY_PICK_X", GetType(System.Int64), "SUM(CHILD(ICTSTATX_SOTSORD1).ORDR_QTY_PICK_X)")

        If "" = "gabe saks ok" Then
            ' qty open is screwed up even when we enable this code
            ' a change needed to be made to the report.
            For Each rowICTSTATX As DataRow In dst.Tables("ICTSTATX").Select("ORDR_QTY_PICK_X <> 0")
                rowICTSTATX.Item("WHSE_QTY_OPEN") = Val(rowICTSTATX.Item("WHSE_QTY_OPEN") & "") _
                                                  - Val(rowICTSTATX.Item("ORDR_QTY_PICK_X") & "")
            Next
        End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            If Absx1.chkFor("CHKTRANONLY").Checked Then
                For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select()
                    If rowSOTORDR0.Item("CUST_CODE").ToString & String.Empty = "WALMART" Then
                        rowSOTORDR0.Item("MABD") = GET_MABD(rowSOTORDR0.Item("ORDR_GROUP_NO").ToString & String.Empty)
                    End If
                Next
            End If
        End If
    End Sub

    Public Overrides Sub Print_Report()
        Dim RPT_NAME As String = RPT
        If Absx1.chkFor("CHKTRANONLY").Checked Then
            RPT_NAME = "SORSORDW"
        End If
        SUBT = ""

        Dim z As String = ""
        If Format(REPORT_DATE0, "yyyyMMdd") = Format(REPORT_DATE1, "yyyyMMdd") Then
            z = " on " & Format(REPORT_DATE0, "MM/dd/yyyy")
        Else
            z = " between " & Format(REPORT_DATE0, "MM/dd/yyyy") & " and " & Format(REPORT_DATE1, "MM/dd/yyyy")
        End If
        If optSR.Value = "R" Then
            SUBT &= "Orders Received" & z
        Else
            SUBT &= "Orders to Ship" & z
        End If

        If Absx1.chkFor("CHKEDI_ONLY").Checked Then
            SUBT &= " (EDI Orders Only)"
        End If

        CR_params.Add("NO_PRICING", IIf(Absx1.chkFor("CHKNO_PRICING").Checked, "1", "0"))
        'CR_params.Add("TRANONLY", IIf(Absx1.chkFor("CHKTRANONLY").Checked, "1", "0"))
        'Generate_Report(RPT, , SUBT)
        Generate_Report(RPT_NAME, , SUBT)


        If (ASCMAIN1.CLIENT = "VAN") Then

            'Dim rowICTBODYX() As DataRow = dst.Tables("ICTBODYX").Select("")
            'If rowICTBODYX.Count <> 0 Then
            '    MsgBox("Please Review Missing Standard Cube Per Unit Report", MsgBoxStyle.OkOnly, "Need to Update the Standard Cube Per Unit field in Body Types (Sub) FM")
            '    RPT = "SORSORDX"
            '    RPT_TITLE = "Missing Standard Cube Per Unit Report"
            '    SUBT = "Enter Standard Cube Per Unit for these Sub Body Codes in Body Types (Sub) Maint"
            '    Generate_Report(RPT, RPT_TITLE, SUBT)
            'End If
            If chkStyleStats.Checked Then


                SUBT &= " (w/In-Transit Details)"
                CR_params.Add("NO_PRICING", IIf(Absx1.chkFor("CHKNO_PRICING").Checked, "1", "0"))
                RPT_NAME = "SORSORDZ"

                Generate_Report(RPT_NAME, , SUBT)
            End If


        End If



    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If optSR.Value = "A" Then
                If tblASTDSQLA.Select("CODE_VALUES IS NOT NULL").Length = 0 Then
                    EMsg &= vbCr & "You must select a filter when Selecting Any Date"
                End If
            End If

            If Absx1.optFor("OPTORB").Value <> "O" Then
                Dim rows() As DataRow = tblASTDSQLA.Select("COLUMN_NAME <> 'CUST_CODE' and COLUMN_NAME <> 'SALES_DIVISION_CODE' and ISNULL(CODE_VALUES,'') <> ''")
                If rows.Length <> 0 Then
                     MsgBox("Please call ABS", MsgBoxStyle.OkOnly, "Need to check SQL for Performance")
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Private Sub btnFilterReservations_Click(sender As Object, e As EventArgs) Handles btnFilterReservations.Click
        dtSOTRSRV1.Clear()
        Dim S As New Text.StringBuilder With {.Length = 0}
        S.AppendLine("SELECT RSRV_NO, CUST_CODE, ORDR_CUST_PO, ORDR_SHIP_DATE, ORDR_CANCEL_DATE")
        S.AppendLine("FROM SOTRSRV1")
        S.AppendLine("WHERE RSRV_STATUS = 'O'")
        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Select Reservations To Filter"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable
        End With
        Dim F As New ASFCODE1
        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            For Each DR As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                Dim rowSOTRSRV1 As DataRow = dtSOTRSRV1.NewRow
                rowSOTRSRV1.Item("RSRV_NO") = DR.Item("RSRV_NO") & String.Empty
                dtSOTRSRV1.Rows.Add(rowSOTRSRV1)
            Next
        End If
        grdSOTRSRV1.Refresh()
    End Sub

    Private Sub UltraOptionSet2_ValueChanged(sender As Object, e As EventArgs) Handles UltraOptionSet2.ValueChanged
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            grpFilterReservations.Visible = True
        Else
            grpFilterReservations.Visible = False
        End If
    End Sub

    Private Function GET_MABD(ByVal ORDR_GROUP_NO As String) As String
        Dim RETVAL As String = ""
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT DISTINCT E1.EDI_SHIP_DATE")
        SQLS.AppendLine("FROM SOTORDR1 O1, EDT850T1 E1")
        SQLS.AppendLine("WHERE O1.EDI_DOC_SEQ_NO = E1.EDI_DOC_SEQ_NO")
        SQLS.AppendLine("AND O1.EDI_JRNL_NO = E1.EDI_JRNL_NO")
        SQLS.AppendLine(String.Format("AND O1.ORDR_GROUP_NO = '{0}'", ORDR_GROUP_NO))
        ASCMAIN1.sql = SQLS.ToString()
        Dim EDI_SHIP_DATE As String = ASCDATA1.GetDataValue
        If IsDate(EDI_SHIP_DATE) Then
            RETVAL = Format(CDate(EDI_SHIP_DATE), "MM/dd/yy")
        End If
        Return RETVAL
    End Function

    Private Sub txtDescription_ValueChanged(sender As Object, e As EventArgs) Handles txtDescription.ValueChanged

    End Sub
End Class