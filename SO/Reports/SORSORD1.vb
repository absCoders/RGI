Imports System.Text
Imports SpreadsheetGear

Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports Infragistics.Win.UltraWinGrid


Public Class SORSORD1
    ' hardcoding below for walmart, kmart, sears

    Dim REPORT_DATE0 As Date
    Dim REPORT_DATE1 As Date
    Dim dtSOTRSRV1 As New DataTable
    Dim ICTSTATDSQL As String
    Dim ICTSTAT2SQL As String
    Dim SOTORDRXSQL As String = ""
    Dim IMG_Error_Reported As Boolean = False

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

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

        txtPOREF.Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")
        With UltraExplorerBar1.Groups("Special Functions")
            .Visible = False
        End With
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

        ' new 
        Dim POREF As String = ""
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            If txtPOREF.Text <> "" Then
                Dim i As Integer
                POREF = ""
                Dim datarec() As String = Split(txtPOREF.Text, vbCrLf)
                For i = 0 To UBound(datarec)
                    If datarec(i).Length <> 0 And POREF = "" Then
                        POREF = POREF & "("
                    End If
                    If datarec(i).Length <> 0 Then
                        POREF = POREF & "'" & datarec(i) & "',"
                    End If
                    '     MessageBox.Show(datarec(i))
                Next i

                If POREF <> "" Then
                    POREF = POREF.TrimEnd(CChar(","))
                    POREF = POREF & ")"
                    sqlw &= "" _
                              & " and SOTORDR1.ORDR_CUST_PO IN " & POREF & vbCrLf
                End If
            End If
        End If

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

            If POREF <> "" Then
                sqlw_r &= "" _
                              & " and SOTRSRV1.ORDR_CUST_PO IN " & POREF & vbCrLf
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
            & " and (CUST_CODE NOT IN ('BEALLS','ROSS','BEALLSDS','MARSHAL') OR SOTORDR2.ORDR_QTY <> SOTORDR2.ORDR_QTY_CANC)" _
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
            & ", MAX (SOTORDR1.ORDR_CANCEL_DATE) CANCEL_DATE" & vbCrLf _
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
            & ", MAX (SOTORDR1.ORDR_CANCEL_DATE) CANCEL_DATE" & vbCrLf _
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
            & ", MAX (SOTRSRV1.ORDR_CANCEL_DATE) CANCEL_DATE" & vbCrLf _
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
            & ", MAX (CANCEL_DATE) CANCEL_DATE" & vbCrLf _
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
            sql = "Delete From " & SOTSORD1 & " WHERE CUST_CODE IN ('BEALLS','ROSS','BEALLSDS','MARSHAL') AND ORDR_QTY_X = ORDR_QTY_CANC_X"
            ASCDATA1.ExecuteSQL(sql)
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            sql = "Update " & SOTSORD1 & " Set ORDR_QTY_X = ORDR_QTY_OPEN_X  + ORDR_QTY_PICK_X, ORDR_AMT_X = ORDR_AMT_X - ORDR_AMT_CANC_X  WHERE CUST_CODE IN ('BEALLS','ROSS','BEALLSDS','MARSHAL') AND ORDR_QTY_X <> ORDR_QTY_CANC_X and ORDR_QTY_CANC_X <> 0 and (ORDR_QTY_OPEN_X <> 0 OR ORDR_QTY_PICK_X <> 0)"
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


        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC,ICTSTYL1.IMAGE_NAME " & vbCrLf _
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


                ASCMAIN1.sql = "select *  from ictstat2 WHERE (STYLE_CODE,COLOR_CODE) IN (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & SOTSORD1 & ")"
                ICTSTAT2SQL = ASCMAIN1.sql
                Create_TDA(dst.Tables.Add, "ICTSTAT2", "**", 2, False)

                Fill_Records("ICTSTAT2", "", True, ICTSTAT2SQL)

                ASCMAIN1.sql = "Select * from (" & vbCrLf _
                & " Select SOTORDR2.STYLE_CODE,SOTORDR2.COLOR_CODE,'O' ORDR_TYPE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ",  SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                & ", MIN(SOTORDR1.SREP_CODE) SREP_CODE, MIN(SOTORDR1.WHSE_CODE) WHSE_CODE, SOTORDR0.ORDR_TYPE_CODE" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR, SUM (SOTORDR2.ORDR_QTY_OPEN) OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) PICK, SUM (SOTORDR2.ORDR_QTY_ALLO) ALLO" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP) SHIP, SUM (SOTORDR2.ORDR_QTY_CANC) CANC, MAX (SOTORDR2.ORDR_UNIT_PRICE) PRICE" & vbCrLf _
                & ", COUNT (DISTINCT SOTORDR1.ORDR_NO) ORDERS" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY      * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_CANC" & vbCrLf _
                & ", ARTCUST1.CUST_NAME" & vbCrLf _
                & ", MIN (SOTORDR1.ORDR_DATE_RECD) ORDR_DATE_RECD, MIN (SOTORDR1.INIT_DATE) INIT_DATE" & vbCrLf _
                & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
                & " From SOTORDR2, SOTORDR1, SOTORDR0, ARTCUST1, ICTATOP1" & vbCrLf _
                & " Where (SOTORDR2.ORDR_STATUS = 'O' OR SOTORDR2.ORDR_STATUS = 'P')" & vbCrLf _
                & " And SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & " And ICTATOP1.ORDR_TYPE(+) = 'O'" & vbCrLf _
                & " And ICTATOP1.ORDR_NO (+) = SOTORDR2.ORDR_NO" & vbCrLf _
                & " And SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & " And ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE " & vbCrLf _
                & " And (SOTORDR2.STYLE_CODE,SOTORDR2.COLOR_CODE) IN  (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & SOTSORD1 & ")" & vbCrLf _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, ARTCUST1.CUST_NAME, SOTORDR0.ORDR_TYPE_CODE" & vbCrLf _
                & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
                & " ) union (" & vbCrLf _
                & " Select SOTRSRV2.STYLE_CODE,SOTRSRV2.COLOR_CODE,'R' ORDR_TYPE, SOTRSRV2.RSRV_NO ORDR_GROUP_NO, SOTRSRV1.CUST_CODE, SOTRSRV1.ORDR_CUST_PO ORDR_CUST_PO" & vbCrLf _
                & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", MIN(SOTRSRV1.SREP_CODE) SREP_CODE, MIN(SOTRSRV1.WHSE_CODE) WHSE_CODE, NULL ORDR_TYPE_CODE" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY) ORDR, SUM (SOTRSRV2.RSRV_QTY_OPEN) OPEN" & vbCrLf _
                & ", SUM (0) PICK, SUM (SOTRSRV2.RSRV_QTY_ALLO) ALLO" & vbCrLf _
                & ", 0 SHIP, 0 CANC,MAX (SOTRSRV2.ORDR_UNIT_PRICE) PRICE" & vbCrLf _
                & ", 0 ORDERS" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY      * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY_OPEN * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN" & vbCrLf _
                & ", SUM (0                      * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
                & ", SUM (0                      * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _
                & ", SUM (SOTRSRV2.RSRV_QTY_CANC * SOTRSRV2.ORDR_UNIT_PRICE) ORDR_AMT_CANC" & vbCrLf _
                & ", ARTCUST1.CUST_NAME" & vbCrLf _
                & ", SOTRSRV1.INIT_DATE AS ORDR_DATE_RECD, SOTRSRV1.INIT_DATE" & vbCrLf _
                & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
                & " From SOTRSRV2, SOTRSRV1, ARTCUST1, ICTATOP1" & vbCrLf _
                & " Where SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
                & " And SOTRSRV2.RSRV_QTY_OPEN <> 0" & vbCrLf _
                & " And SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & " And ICTATOP1.ORDR_TYPE (+) = 'R'" & vbCrLf _
                & " And ICTATOP1.ORDR_NO (+) = SOTRSRV2.RSRV_NO" & vbCrLf _
                & " And ARTCUST1.CUST_CODE = SOTRSRV1.CUST_CODE" & vbCrLf _
                & " And (SOTRSRV2.STYLE_CODE,SOTRSRV2.COLOR_CODE) IN  (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & SOTSORD1 & ")" & vbCrLf _
                & " group by SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, SOTRSRV2.RSRV_NO, SOTRSRV1.CUST_CODE, SOTRSRV1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTRSRV1.ORDR_SHIP_DATE, SOTRSRV1.ORDR_CANCEL_DATE, ARTCUST1.CUST_NAME, SOTRSRV1.INIT_DATE" & vbCrLf _
                & ", ICTATOP1.STYLE_SHIP_WINDOW_DAYS, ICTATOP1.ORDR_SHIP_DATE_PLUS, ICTATOP1.STYLE_AT_ONCE_UNTIL, ICTATOP1.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
                & ")"

                SOTORDRXSQL = ASCMAIN1.sql
                Create_TDA(dst.Tables.Add, "SOTORDRX", "**", 0, False)

                Fill_Records("SOTORDRX", "", True, SOTORDRXSQL)

                With dst.Tables.Add("SOTCADSZ")
                    .Columns.Add("SEQ").DataType = GetType(System.Int32)
                    .Columns.Add("STYLE_CODE")
                    .Columns.Add("IMAGE", GetType(System.Byte()))
                    .Columns.Add("IMAGE_NAME")
                    .Columns.Add("SELECTED")
                    .Columns("SELECTED").DefaultValue = "0"
                End With


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


        ASCMAIN1.sql = "Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE" & vbCrLf _
            & ", SUM(ICTSTAT2.WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND, SUM(ICTSTAT2.WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
            & ", SUM(ICTSTAT2.WHSE_QTY_ON_ORDER) WHSE_QTY_ON_ORDER, SUM(ICTSTAT2.WHSE_QTY_TRAN) WHSE_QTY_TRAN,SUM(ICTSTAT2.WHSE_QTY_OPEN) WHSE_QTY_OPEN" & vbCrLf _
            & ", SUM(ICTSTAT2.WHSE_QTY_ON_HAND) - SUM(ICTSTAT2.WHSE_QTY_PICK) - SUM(ICTSTAT2.WHSE_QTY_OPEN) +  SUM(ICTSTAT2.WHSE_QTY_ON_ORDER) + SUM(ICTSTAT2.WHSE_QTY_TRAN) NET_POS" & vbCrLf _
            & ", SYSDATE SOON_SHIP_DATE" & vbCrLf _
            & " from ICTSTAT2 " & vbCrLf _
            & " where ICTSTAT2.STYLE_CODE in " & vbCrLf _
            & " (Select Distinct STYLE_CODE from " & SOTSORD1 & ")" & vbCrLf _
            & " GROUP BY ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", 0 WHSE_QTY_ON_HAND, 0 WHSE_QTY_PICK, 0 WHSE_QTY_ON_ORDER, 0 WHSE_QTY_TRAN,0 WHSE_QTY_OPEN, 0 NET_POS, NULL SOON_SHIP_DATE from (" & vbCrLf _
            & "Select Distinct STYLE_CODE, COLOR_CODE from " & SOTSORD1 & vbCrLf _
            & " minus" & vbCrLf _
            & "Select Distinct STYLE_CODE, COLOR_CODE from ICTSTAT2)"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSORDX", 2))


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

        With UltraExplorerBar1.Groups("Special Functions")
            .Visible = True
        End With
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        'Sticking a stupid row in the table to keep the standards from being an ass.
        Dim newASTSRPT1 As DataRow = dst.Tables("ASTSRPT1").NewRow
        newASTSRPT1.Item("G1") = "XX"
        dst.Tables("ASTSRPT1").Rows.Add(newASTSRPT1)

        UpdateReportRows()
    End Sub
    Private Sub UpdateReportRows()

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

        If eItemKey = "Done" Then
            '     Build_Init_Sel()
            With UltraExplorerBar1.Groups("Special Functions")
                .Visible = False
            End With
        End If

        If eItemKey = "Print Full CADs" Then
            Print_Full_CAD_Print(eItemKey)
            'With UltraExplorerBar1.Groups("Special Functions")
            '    .Visible = False
            'End With
            Exit Sub
        End If


        If eItemKey = "Net Position" Then

            ASCMAIN1.Progress("Creating Net Position Report", "")
            Dim XLS_FILENAME1 As String = MakeExcelWorkbook()
            Dim XLS_FILENAME2 As String = ""
            Show_Document(XLS_FILENAME1)


            ASCMAIN1.Progress("", "")


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


    Private Function MakeExcelWorkbook() As String
        Dim XLS_FILENAME As String = ""

        Dim StyleList As New List(Of String)
        For Each rowSOTSORDX As DataRow In dst.Tables("SOTSORDX").Select("NET_POS <> 0")
            Dim STYLE_CODE As String = rowSOTSORDX.Item("STYLE_CODE").ToString & String.Empty
            If dst.Tables("SOTSORD1").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & rowSOTSORDX.Item("COLOR_CODE") & String.Empty & "'").Length <> 0 <> 0 Then
                If Not StyleList.Contains(STYLE_CODE) Then
                    StyleList.Add(STYLE_CODE)
                Else
                    sql = sql
                End If

            End If

        Next


        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        worksheet.Name = "Style Info
"
        Create_Excel_WorkSheet(worksheet, StyleList)


        If ASCMAIN1.Folders("Temp").EndsWith("\") Then
            XLS_FILENAME = ASCMAIN1.Folders("Temp") & String.Format("{0}.XLSX", "Net Pos Report")
        Else
            XLS_FILENAME = ASCMAIN1.Folders("Temp") & "\" & String.Format("{0}.XLSX", "Net Pos Report")
        End If
        Dim success As Boolean = False

        ASCMAIN1.Progress("Now Saving Workbook")

        Do Until success
            Try
                If System.IO.File.Exists(XLS_FILENAME) Then
                    System.IO.File.Delete(XLS_FILENAME)
                End If
                workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                success = True
            Catch ex As Exception

            End Try
        Loop
        Return XLS_FILENAME



    End Function


    Sub Create_Excel_WorkSheet(worksheet As SpreadsheetGear.IWorksheet,
                               ByVal StyleList As List(Of String), Optional sqlWB As String = "")

        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        If (ASCMAIN1.Running_in_VS) Then
            If Not System.IO.Directory.Exists(IMAGE_FOLDER) Then
                Stop 'You Need to Set up Image Folder.
            End If
        End If

        Dim CX As Integer = 0
        Dim RX As Integer = 0

        Dim I As Integer = 0
        I += 4

        Dim COL0 As Integer = 12

        Dim COL As Integer = COL0

        Excel_DefaultColumns(worksheet, COL)

        With worksheet.Cells(I, 0, I, COL)
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        Dim I0 As Integer = 0
        Dim IA As Integer = 0
        Dim RT(11) As String
        Dim ROW0 As Integer = I
        Dim style_count As Integer = 0
        Dim pages As Integer = 0

        For Each STYLE_CODE As String In StyleList
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            ASCMAIN1.Progress("-", STYLE_CODE)
            I += 1
            I0 = I
            COL = COL0

            Excel_StyleHeader(worksheet, COL, I, COL0)

            I += 1

            Dim ImageRows = 0
            Dim ImageRowsBig = 0
            Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""
            Excel_ImageInsert(worksheet, IMAGE_NAME, IMAGE_FOLDER, ImageRows, ImageRowsBig, I)

            CX = 1

            Excel_StyleMasterfile(worksheet, I, CX, rowICTSTYL1, STYLE_CODE)

            Dim CI As Integer = 0
            Excel_ColorDetails(worksheet, STYLE_CODE, I, COL, COL0, CI)

            For iCOL As Integer = 1 To 1
                COL += 1
                Select Case iCOL
                    Case 1
                        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        'Case 2
                        '    If chkShip2.Checked Then
                        '        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        '    Else
                        '        COL -= 1
                        '    End If
                    Case 3
                        '       COL -= 1
                End Select

                RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
            Next

            COL += 1

            Dim colsLess As Int16 = 1

            If chkStyleStats.Checked Then
                COL = COL - colsLess
                For iCOL As Integer = 2 To 8
                    COL += 1
                    worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                Next
                COL += 0
            End If

            If chkStyleStats.Checked Then
                worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray
            Else
                worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL - colsLess).Interior.Color = SpreadsheetGear.Colors.LightGray
            End If


            With worksheet.Cells(I, COL0 - 1, I + CI - 1, COL - colsLess)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
            End With

            I += ImageRowsBig

            Dim CJ As Integer = ImageRows

            If CJ < 6 Then CJ = 6

            If CI > CJ Then
                I += CI
            Else
                I += CJ
            End If

            style_count += 1

            If (((I - 5) Mod 80) < ((I0 - 5) Mod 80)) Or (style_count >= 5) Or style_count >= 9 Then
                Dim R As SpreadsheetGear.IRange = worksheet.Cells(I0, 0).EntireRow
                worksheet.HPageBreaks.Add(R)
                style_count = 1
                pages += 1
            End If

            If chkStyleStats.Checked Then
                Dim interior As SpreadsheetGear.IInterior
                Dim range As SpreadsheetGear.IRange
                '  I += 1
                COL = COL0
                Dim chkcnt As Int64 = 0
                Dim NEWSTYLE As Boolean = True

                For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE, ORDR_SHIP_DATE")
                    If NEWSTYLE = True Then
                        worksheet.Cells(I - 1, COL - 1).Value = "Ord/Res Details"
                        I += 1
                        ' Headinds and headingsFOrmat
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Col"
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Ord Typ"
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Customer"
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Customer PO"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        End With
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Ord Shp Dt"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        '    range = worksheet.Cells(I - 1, COL - 1, I - 1, COL + 6)
                        ' interior = range.Interior
                        'interior.Color = SpreadsheetGear.Colors.Aquamarine
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Ord Can Dt"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        range = worksheet.Cells(I - 1, COL - 1, I - 1, COL + 8)
                        interior = range.Interior
                        interior.Color = SpreadsheetGear.Colors.Aquamarine
                        chkcnt += 1


                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Qty Ord"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Price"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Open"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        chkcnt += 1

                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Pick"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        chkcnt += 1

                        NEWSTYLE = False
                    End If



                    I += 1
                    chkcnt = 1
                    If sql = sql Then
                        ' avoid printing if no records in SOTORDRX
                        ' worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & "***"

                    End If



                    '  worksheet.Cells(I - 1, COL - 2 + chkcnt).Value = Val(rowSOTORDRXItem(1) & String.Empty)


                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = Format(Val(rowSOTORDRX.Item("COLOR_CODE") & String.Empty), "000")
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = rowSOTORDRX.Item("ORDR_TYPE") & String.Empty
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If
                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = rowSOTORDRX.Item("CUST_CODE") & String.Empty
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = rowSOTORDRX.Item("ORDR_CUST_PO") & String.Empty
                        .NumberFormat = ""
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = rowSOTORDRX.Item("ORDR_SHIP_DATE") & String.Empty
                        .NumberFormat = "MM/dd/yy"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = rowSOTORDRX.Item("ORDR_CANCEL_DATE") & String.Empty
                        .NumberFormat = "MM/dd/yy"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowSOTORDRX.Item("ORDR") & String.Empty)
                        .NumberFormat = "#,##0"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1
                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowSOTORDRX.Item("PRICE") & String.Empty)
                        .NumberFormat = "#,##0.00"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowSOTORDRX.Item("OPEN") & String.Empty)
                        .NumberFormat = "#,##0"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowSOTORDRX.Item("PICK") & String.Empty)
                        .NumberFormat = "#,##0"
                        .Font.Size = 14
                        If rowSOTORDRX.Item("ORDR_TYPE") & "" = "R" Then
                            .Font.Color = SpreadsheetGear.Colors.Red
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Green
                        End If

                    End With
                    chkcnt += 1


                Next
                COL = COL0
                chkcnt = 0
                NEWSTYLE = True
                I += 2

                For Each rowICTSTATD As DataRow In dst.Tables("ICTSTATD").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE, PO_DATE_SHIP_BY")
                    If NEWSTYLE = True Then
                        worksheet.Cells(I - 1, COL - 1).Value = "In-Transit Details"
                        I += 1
                        ' Headinds and headingsFOrmat
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Color"
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Factory"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        End With
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Qty Ord"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Qty Shp"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15

                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Rev Ship Dt"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "ETA"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Vessel"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 20
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        End With

                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Shp Dt Rev"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        range = worksheet.Cells(I - 1, COL - 1, I - 1, COL + 6)
                        interior = range.Interior
                        interior.Color = SpreadsheetGear.Colors.Aquamarine

                        NEWSTYLE = False
                    End If



                    I += 1
                    chkcnt = 1
                    If sql = sql Then
                        ' avoid printing if no records in ICTSTATD
                        ' worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & "***"

                    End If



                    '  worksheet.Cells(I - 1, COL - 2 + chkcnt).Value = Val(rowICTSTATD.Item(1) & String.Empty)

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = Format(Val(rowICTSTATD.Item("COLOR_CODE") & String.Empty), "000")
                        .Font.Size = 14
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                    End With
                    chkcnt += 1

                    '   worksheet.Cells(I - 1, COL - 2 + chkcnt).Value = Val(rowICTSTATD.Item(4) & String.Empty)

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = rowICTSTATD.Item("FACTORY_CODE") & String.Empty
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                        .Font.Size = 14
                    End With
                    chkcnt += 1


                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowICTSTATD.Item("PO_QTY_ORD") & String.Empty)
                        .NumberFormat = "#,##0"
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                        .Font.Size = 14
                    End With
                    chkcnt += 1
                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowICTSTATD.Item("PO_QTY_SHP") & String.Empty)
                        .NumberFormat = "#,##0"
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If


                        .Font.Size = 14
                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = rowICTSTATD.Item("PO_DATE_SHIP_BY") & String.Empty
                        .NumberFormat = "MM/dd/yy"
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                        .Font.Size = 14
                    End With
                    chkcnt += 1
                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .NumberFormat = "MM/dd/yy"
                        .Value = rowICTSTATD.Item("PO_SHIP_ETA") & String.Empty
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                        .Font.Size = 14
                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Left
                        .Value = rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If
                        .Font.Size = 14
                    End With
                    chkcnt += 1
                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .NumberFormat = "MM/dd/yy"
                        .Value = rowICTSTATD.Item("LAST_DATE_SHIP_BY") & String.Empty
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If
                        .Font.Size = 14
                    End With
                    chkcnt += 1

                Next
                'T = ""
                'COL += 1



            End If




            With worksheet.Cells(I0, 0, I + 1 - 1, COL)
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With
        Next

        I += 2
        COL = COL0

        'Trying to get away without totals here :)
        worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"

        Dim GT = ""
        For iCOL As Integer = 1 To 1
            COL += 1
            Select Case iCOL
                Case 1
                    worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                Case 3

                    COL -= 2

            End Select
        Next

        '    COL += 1

        If chkStyleStats.Checked Then
            For iCOL As Integer = 2 To 8
                COL += 1
                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")



            Next
            COL += 0
        End If

        worksheet.Cells(I - 1, COL0 - 1, I - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

        Excel_Header(worksheet)

        Excel_PageSetup(worksheet)
    End Sub

    Private Sub Excel_ColorDetails(ByRef worksheet As IWorksheet,
                                   ByVal STYLE_CODE As String,
                                   ByRef i As Integer,
                                   ByRef COL As Integer,
                                   ByRef COL0 As Integer,
                                   ByRef CI As Integer)
        Dim SZMAX As Integer = 0
        Dim SZTOT As Integer = 0
        Dim T As String = ""
        Dim styleTotal As Int64 = 0
        Dim LAST_COLOR As String = ""
        For Each rowSOTSORDX As DataRow In dst.Tables("SOTSORDX").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE")

            '   Dim rowSOTSORD1 As DataRow = dst.Tables("SOTSORD1").Rows.Find(New String() {rowSOTSORDX.Item("STYLE_CODE") & String.Empty, rowSOTSORDX.Item("COLOR_CODE") & String.Empty})

            ''Dim rowSOTSORD1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTSORD1 WHERE STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & rowSOTSORDX.Item("COLOR_CODE") & String.Empty & "'")

            ''If rowSOTSORD1 IsNot Nothing And Val(rowSOTSORDX.Item("NET_POS") & String.Empty) <> 0 Then

            If dst.Tables("SOTSORD1").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & rowSOTSORDX.Item("COLOR_CODE") & String.Empty & "'").Length <> 0 And Val(rowSOTSORDX.Item("NET_POS") & String.Empty) <> 0 Then
                ' And Val(rowSOTSORDX.Item("NET_POS") & String.Empty) <> 0
                CI += 1
                COL = COL0 + 1
                'COL = COL0
                Dim chkcnt As Int64 = 1
                Dim COLOR_DESC As String = ""
                If LAST_COLOR <> rowSOTSORDX.Item("COLOR_CODE") & String.Empty Then
                    worksheet.Cells(i + CI - 1, COL - 2).Value = "'" & rowSOTSORDX.Item("COLOR_CODE") & String.Empty
                    ' worksheet.Cells(i + CI - 1, COL - 1).Value = rowSOTSORDX.Item("COLOR_DESC") & String.Empty
                    worksheet.Cells(i + CI - 1, COL - 1).Value = GetAltColorCode(STYLE_CODE, rowSOTSORDX.Item("COLOR_CODE") & String.Empty, COLOR_DESC)
                    '     worksheet.Cells(i + CI - 1, COL - 1).Value = COLOR_DESC

                    LAST_COLOR = rowSOTSORDX.Item("COLOR_CODE") & String.Empty
                End If
                'worksheet.Cells(i + CI - 1, COL + 1).Value = "ORDR_CUST_PO" 'rowSOTCUSTQ.Item("ORDR_CUST_PO") & String.Empty
                'worksheet.Cells(i + CI - 1, COL + 2).Value = "ORDR_SHIP_DATE" 'rowSOTCUSTQ.Item("ORDR_SHIP_DATE") & String.Empty
                'worksheet.Cells(i + CI - 1, COL + 3).Value = "ORDR_CANCEL_DATE" 'rowSOTCUSTQ.Item("ORDR_CANCEL_DATE") & String.Empty
                worksheet.Cells(i + CI - 1, COL).Value = rowSOTSORDX.Item("NET_POS") & String.Empty
                'chkcnt += 1


                'If chkShip2.Checked Then
                '    worksheet.Cells(i + CI - 1, COL + 1).Value = rowSOTCUSTS.Item("QTY_SHP_02") & String.Empty
                '    'chkcnt += 1
                'End If
                'If chkShip3.Checked Then
                '    worksheet.Cells(i + CI - 1, COL + 2).Value = rowSOTCUSTS.Item("QTY_SHP_03") & String.Empty
                '    'chkcnt += 1
                'End If
                'T = ""
                'COL += 1

                If chkStyleStats.Checked Then
                    chkcnt = 1
                    ' COL = 10

                    ' ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty & "'"
                    '       For Each rowICTSTAT2 As DataRow In ASCDATA1.GetDataTable.Select("")
                    For Each rowICTSTAT2 As DataRow In dst.Tables("ICTSTAT2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowSOTSORDX.Item("COLOR_CODE") & String.Empty & "'")


                        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & String.Empty)
                        chkcnt += 1
                        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & String.Empty)
                        chkcnt += 1
                        Dim OTS As Integer = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "") - Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & "")
                        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = OTS
                        chkcnt += 1

                        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty)
                        chkcnt += 1

                        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty)
                        chkcnt += 1

                        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
                        chkcnt += 1

                        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = OTS + Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty) + Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty) - Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
                        chkcnt += 1

                    Next
                    T = ""
                    COL += 1
                End If

            Else

            End If



        Next




        CI += 2
        COL = COL0

        worksheet.Cells(i - 1, COL - 1, i + CI - 1, COL - 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & "***"
        worksheet.Cells(i + CI - 1, COL - 0).Value = "'" & "Total"
    End Sub

    Private Sub Excel_StyleMasterfile(ByRef worksheet As IWorksheet, ByRef i As Integer, ByRef cx As Integer, ByRef rowICTSTYL1 As DataRow, ByVal STYLE_CODE As String)
        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange

        With worksheet.Cells(i - 1, 3)
            .Value = "'" & STYLE_CODE
            .Font.Color = SpreadsheetGear.Colors.Purple
            .Font.Size = 24
            .Font.Bold = True
        End With

        cx = 3

        worksheet.Cells(i + 2, cx).Value = "Net Position"

        range = worksheet.Cells(i + 1, 3, i + 2, 4)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.LightGray

        range = worksheet.Cells(i + 1, 3 + 4, i + 2, 4 + 4)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.LightGray

        cx = 5
        worksheet.Cells(i, cx - 2).Value = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
        worksheet.Cells(i + 2, cx).Value = rowICTSTYL1.Item("CARTON_PACK_QTY")
    End Sub

    Private Sub Excel_ImageInsert(ByRef worksheet As IWorksheet,
                                  ByVal iMAGE_NAME As String,
                                  ByVal IMAGE_FOLDER As String,
                                  ByRef ImageRows As Integer,
                                  ByRef ImageRowsBig As Integer,
                                  ByRef i As Integer)
        Dim imageFileStyle As String = IMAGE_FOLDER & "\" & iMAGE_NAME
        If Not System.IO.File.Exists(imageFileStyle) Then
            iMAGE_NAME = ""
        End If

        If iMAGE_NAME <> "" _
                AndAlso My.Computer.FileSystem.FileExists(imageFileStyle) Then

            Dim widthStyle As Double
            Dim heightStyle As Double

            Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
            Try
                widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution / 3
                heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution / 3
            Finally
                imageStyle.Dispose()
            End Try

            Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

            Dim col_adj As Decimal = 0
            If heightStyle > widthStyle Then
                col_adj = 0.3
            Else
                col_adj = 0.05
            End If

            Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(0) + col_adj
            Dim topStyle As Double = windowInfoStyle.RowToPoints(i - 1) + 0.1

            ImageRows = windowInfoStyle.PointsToRow(heightStyle)
            worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
        End If
    End Sub

    Private Sub Excel_StyleHeader(ByRef worksheet As IWorksheet, ByRef COL As Integer, ByRef i As Integer, ByVal COL0 As Integer)
        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange

        worksheet.Cells(i, COL - 1).Value = "" & Chr(13) & Chr(10) & "Color"
        worksheet.Cells(i, COL - 1).Font.Size = 12
        worksheet.Cells(i, COL).Value = "" & Chr(13) & Chr(10) & "Description"
        worksheet.Cells(i, COL).Font.Size = 12
        worksheet.Cells(i, COL + 1).Value = "" & Chr(13) & Chr(10) & "Net Position"
        worksheet.Cells(i, COL + 1).Font.Size = 12


        ''    COL += 1
        'With worksheet.Cells(i, COL)
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .Value = "Net Position"
        'End With

        'COL += 1
        'With worksheet.Cells(i, COL)
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .Value = "Ship"
        'End With

        'COL += 1
        'With worksheet.Cells(i, COL)
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .Value = "Cancel"
        'End With

        COL += 1



        range = worksheet.Cells(i, COL0 - 1, i, COL)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.Gold

        If chkStyleStats.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "On Hand"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "In Pick"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "OTS"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "In Transit"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "WIP"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "Open"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "Net Pos"
            End With

            range = worksheet.Cells(i, COL - 6, i, COL)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.Aquamarine

        End If




    End Sub

    Private Sub Excel_Header(worksheet As IWorksheet)
        Dim H0 As Integer = 8 + 6

        worksheet.Cells(0, H0).Value = "Prep"
        worksheet.Cells(1, H0).Value = "By"
        worksheet.Cells(2, H0).Value = "XNo"

        worksheet.Cells(0, H0, 2, H0).Interior.Color = SpreadsheetGear.Colors.LightGray


        worksheet.Cells(0, H0 + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Left
        worksheet.Cells(0, H0 + 1).Value = Now
        worksheet.Cells(0, H0 + 1).NumberFormat = "MM/dd/yy"

        worksheet.Cells(1, H0 + 1).Value = ASCMAIN1.USER_ID
        worksheet.Cells(2, H0 + 1).Value = "'" & Mid(XNO, 5)

        With worksheet.Cells(0, H0, 2, H0 + 1)
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Font.Color = SpreadsheetGear.Colors.Black
            .Font.Size = 10
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        Dim H1 As Integer = 11
        Dim HEAD1 As String = ""
        Dim HEAD2 As String = ""


        'If chkOpen.Checked Then
        '    HEAD2 = "Open"
        'End If
        'If chkPick.Checked Then
        '    If HEAD2 = "" Then
        '        HEAD2 = "Pick"
        '    Else
        '        HEAD2 = HEAD2 & "," & "Pick"
        '    End If
        'End If
        'If chkReservations.Checked Then
        '    If HEAD2 = "" Then
        '        HEAD2 = "Res"
        '    Else
        '        HEAD2 = HEAD2 & "," & "Res"
        '    End If
        'End If

        worksheet.Cells(0, 2).Value = "Net Position Report"
        worksheet.Cells(0, 2).Font.Bold = True
        '   worksheet.Cells(1, 2).Value = "Customer: " & txtCUST_CODE.Text & "   Styles: " & HEAD1
        worksheet.Cells(1, 2).Font.Bold = True



        worksheet.Cells(0, H1).Value = "Note"
        worksheet.Cells(1, H1).Value = "For"

        worksheet.Cells(0, H1, 2, H1).Interior.Color = SpreadsheetGear.Colors.LightGray

        worksheet.Cells(0, H1 + 1).NumberFormat = "MM/dd/yy"
        worksheet.Cells(0, H1 + 1).Value = "Notes"
        '     worksheet.Cells(1, H1 + 1).Value = txtCUST_CODE.Text & String.Empty

        With worksheet.Cells(0, H1, 2, H1 + 2)
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Font.Color = SpreadsheetGear.Colors.Black
            .Font.Size = 10
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        With worksheet.Cells(3, 3)
            .Font.Color = SpreadsheetGear.Colors.Purple
            .Font.Size = 20
            .Font.Bold = True
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With
    End Sub

    Private Sub Excel_PageSetup(ByRef worksheet As IWorksheet)
        With worksheet.PageSetup
            .TopMargin = 0.25
            .LeftMargin = 0.25
            .RightMargin = 0.25
            .BottomMargin = 0.25
            .FitToPagesWide = 1
            .FitToPagesTall = Nothing
            .PrintTitleRows = "A1:S5"
            .CenterFooter = "&P"
        End With
    End Sub

    Private Sub Excel_DefaultColumns(ByRef worksheet As IWorksheet, ByRef COL As Int64)
        worksheet.Cells("A1:Z1").EntireColumn.Font.Size = 16

        Dim CWC() As String = Split("A,B, C,D,E,F,G,H,I,J,K,L, M", ",")
        Dim CWS() As String = Split("1,1,40,6,6,6,6,6,6,6,6,6,20", ",")
        CWS(2) = 45
        For CWCi As Integer = 0 To CWC.Length - 1
            worksheet.Cells(Trim(CWC(CWCi)) & "1").EntireColumn.ColumnWidth = Val(CWS(CWCi))
        Next

        worksheet.Cells(0, 0).EntireColumn.Hidden = True
        worksheet.Cells(0, 1).EntireColumn.Hidden = True

        Dim _COL As Int64 = 1
        ''PO Column
        'COL += 1
        'With worksheet.Cells(_COL, COL)
        '    .ColumnWidth = 20
        '    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        'End With

        ''Ship Date Column
        'COL += 1
        '_COL += 1
        'With worksheet.Cells(_COL, COL)
        '    .ColumnWidth = 15
        '    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .EntireColumn.NumberFormat = "MM/dd/yy"
        'End With

        ''Cancel Date Column
        'COL += 1
        '_COL += 1
        'With worksheet.Cells(_COL, COL)
        '    .ColumnWidth = 15
        '    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '    .EntireColumn.NumberFormat = "MM/dd/yy"
        'End With

        'Ship 1 Column
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 25
            .EntireColumn.NumberFormat = "#,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With


        If chkStyleStats.Checked Then
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
        End If


    End Sub
    Private Function GetAltColorCode(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal COLOR_DESC_ORIG As String) As String
        Dim RetVal As String = COLOR_DESC_ORIG
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        Dim SIZE_SCALE As String = rowICTSTYL1.Item("SIZE_SCALE") & String.Empty
        Dim MAX_LENGTH As Integer = 60
        Dim I As Integer = InStr(SIZE_SCALE, COLOR_CODE)
        If I <> 0 Then
            Dim S As String = Trim(Mid(SIZE_SCALE, I + 3))
            Dim J As Integer = InStr(Mid(S & "  ", 1, MAX_LENGTH), "  ")
            Dim K As Integer = InStr(Mid(S & vbCrLf, 1, MAX_LENGTH), vbCrLf)
            If J = 0 And K = 0 Then
                J = InStr(Mid(S & " ", 1, MAX_LENGTH), " ")
            End If
            If J = 0 Or J > K Then J = K
            Dim SC As String = ""
            If J <> 0 Then
                SC = Mid(S, 1, J)
                SIZE_SCALE = Mid(SIZE_SCALE, 1, I - 1) & Mid(S, J)
                For C As Integer = 1 To SC.Length - 1
                    If C = 1 Or (C > 1 AndAlso Mid(SC, C + 1, 1) <> " " AndAlso (Mid(SC, C - 1, 1) = " " Or Mid(SC, C - 1, 1) = "/")) Then
                        Mid(SC, C, 1) = Mid(SC, C, 1).ToUpper
                    End If
                Next
                If Trim(SC) <> "" Then
                    If SC.Length > 35 Then
                        RetVal = SC.Substring(0, 34)
                    Else
                        RetVal = SC
                    End If

                End If
            End If
        End If
        If RetVal = COLOR_DESC_ORIG Then
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT NVL(STYLE_COLOR_DESC,'') STYLE_COLOR_DESC")
            SQLS.AppendLine("FROM ICTSTYC1")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = SQLS.ToString()
            Dim COLOR_DESC_MF As String = ASCDATA1.GetDataValue
            If COLOR_DESC_MF.Length > 35 Then
                COLOR_DESC_MF = COLOR_DESC_MF.Substring(0, 35)
            End If
            If COLOR_DESC_MF.Length > 0 Then
                RetVal = COLOR_DESC_MF
            End If
        End If
        Return RetVal
    End Function

    Sub RESEQ()

        Dim SEQ As Integer = 0
        Dim OLDSTYLE As String = ""


        dst.Tables.Item("SOTCADSZ").Rows.Clear()
        Dim SORTSOTCUSTS As String = "SUB_BODY_CODE,FABRIC_CODE,STYLE_CODE,COLOR_CODE"
        ''If chkSortStyle.Checked Then
        ''    SORTSOTCUSTQ = "STYLE_CODE,COLOR_CODE"
        ''End If

        Dim sqlWB As String = ""
        ''If chk1Sheet.Checked Then
        ''    sqlWB = ",SALES_DIVISION_CODE," & SORTSOTCUSTS
        ''Else
        ''    sqlWB = ","
        ''End If

        SORTSOTCUSTS = Mid(sqlWB, 2)


        For Each row As DataRow In dst.Tables("ICTSTYL1").Select("", "")
            If OLDSTYLE = "" Or OLDSTYLE <> row.Item("STYLE_CODE") Then
                SEQ += 10
                ''row.Item("SEQ") = SEQ
                ''row.Item("STYLE_CODE_PLM") = row.Item("STYLE_CODE")
                ''row.Item("SELECTED") = "1"
                OLDSTYLE = row.Item("STYLE_CODE")

                Dim rowSOTCADSZ As DataRow = dst.Tables("SOTCADSZ").NewRow
                rowSOTCADSZ.Item("SEQ") = SEQ
                rowSOTCADSZ.Item("STYLE_CODE") = row.Item("STYLE_CODE") & ""
                rowSOTCADSZ.Item("IMAGE_NAME") = row.Item("IMAGE_NAME") & ""
                rowSOTCADSZ.Item("SELECTED") = "1"
                dst.Tables("SOTCADSZ").Rows.Add(rowSOTCADSZ)


            End If
        Next
    End Sub
    Sub Print_Full_CAD_Print(eItemKey As String, Optional STYLE_CODE As String = "")
        Dim ListPDFSheets As New List(Of String)
        Dim MISSING_IMAGES As New List(Of String)


        RESEQ()

        Dim EXCUDE_FUTURE As String = ""

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        If Not FOLDER_NAME.EndsWith("\") Then FOLDER_NAME &= "\"
        FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")

        For Each row As DataRow In dst.Tables("SOTCADSZ").Select("SELECTED='1'")

            Dim STYLE_CODE_PLM As String = row.Item("STYLE_CODE")
            'If STYLE_CODE_PLM = "500498AVR" And ASCMAIN1.Running_in_VS Then Stop
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "dgj")) Then
                'Stop
                FOLDER_NAME = "S:\VAN\images\"
                'FOLDER_NAME = "\\192.168.180.32\g\VAN\images\"
            End If

            If Not My.Computer.FileSystem.FileExists(FOLDER_NAME & row.Item("IMAGE_NAME")) Then
                row.Item("SELECTED") = "0"
                MISSING_IMAGES.Add(STYLE_CODE_PLM)
            End If

        Next

        Dim RPT As String = ""
        ''If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        ''    RPT = "ICRQUOT2"
        ''End If

        Dim ColVisible(4) As Boolean
        RPT = "SORFCADX"

        If eItemKey = "email" Then
            ''Dim tempFileName As String = rowICTQUOT1.Item("QUOTE_NO")

            ''Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)
            ''' Dim FILENAME As String = REPORT_FILENAMES(REPORT_NO)
            ''Print_Report_End(, True)
            ''email_Quote(tempFileName)
        Else

            For Each row As DataRow In dst.Tables("SOTCADSZ").Select("SELECTED='1'")
                row.Item("SELECTED") = "2"
            Next

            Dim REPORT_INDEX As Integer = 0
            Dim PDF_FN As String = ""
            Dim PDF_LINKS As String = ""
            Dim SUB_BODY_DESC As String = ""
            Dim SALES_DIVISION_NAME As String = ""
            Dim FABRIC_DESC As String = ""
            Dim DESCHASH As String = ""

            Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/"
            'Dim LINEPFX_NEW As String = "https://docs.vandalequotes.com/"
            Dim LINEPFX_NEW As String = "https://vandaledocs.azurewebsites.net/Documents/"

            Dim SESSION_NO As String = ASCMAIN1.Next_Control_No("SOTCUSTS.SESSION_NO")
            Dim FILE_NO As Integer = 0

            Do While dst.Tables("SOTCADSZ").Select("SELECTED='2'").Length <> 0

                Print_Report_Begin()

                ''    Dim STYLE_count As Integer = 0
                ''    Dim SRT As String = "SEQ"
                ''    Select Case opt1Sheet.Value
                ''        Case "S"
                ''            SRT = "SUB_BODY_CODE, STYLE_CODE_PLM"
                ''        Case "FS"
                ''            SRT = "FABRIC_CODE, SUB_BODY_CODE, STYLE_CODE_PLM"
                ''        Case "G"
                ''            SRT = "STYLE_GROUP_CODE, STYLE_CODE_PLM"
                ''        Case "D"
                ''            SRT = "SALES_DIVISION_CODE"
                ''    End Select
                ''    For Each row As DataRow In dst.Tables("ICTQUOT2").Select(sqlw, SRT)
                ''        STYLE_count += 1
                ''        row.Item("SELECTED") = "1"
                ''        SetRowImage(row)
                ''    Next
                For Each row As DataRow In dst.Tables("SOTCADSZ").Select()
                    row.Item("IMAGE") = Null
                Next
                Dim STYLE_count As Integer = 0
                For Each row As DataRow In dst.Tables("SOTCADSZ").Select("SELECTED='2'", "SEQ")
                    STYLE_count += 1
                    row.Item("SELECTED") = "1"
                    SetRowImage(row)
                    If STYLE_count >= 50 Then Exit For
                Next
                Application.DoEvents()

                CR_params.Add("IMAGES_FOLDER", FOLDER_NAME)

                CR_params.Add("TXTSTYLE_CODE", "")
                '    CR_params.Add("RECAP", "")


                Dim tempFileName As String = ""
                Do
                    REPORT_INDEX += 1
                    tempFileName = "SORCUSTS" & "-" & Format(REPORT_INDEX, "000")
                Loop While My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")


                ''          Generate_Report(RPT, "Sales Order Summary", , SUBT)

                Dim REPORT_NO As String = Generate_Report(RPT, "Sales Order Summary", "", "", "PDF", tempFileName, False)



                Dim tempNotMade As Boolean = Not System.IO.File.Exists(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")

                If Not tempNotMade Then
                    'Show_Document(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
                    ListPDFSheets.Add(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
                    Print_Report_End(, True)
                End If

                For Each row As DataRow In dst.Tables("SOTCADSZ").Select("SELECTED='1'")
                    row.Item("SELECTED") = "3"
                    row.Item("IMAGE") = DBNull.Value
                Next
            Loop

        End If

        For Each row As DataRow In dst.Tables("SOTCADSZ").Select("")
            row.Item("IMAGE") = Nothing
        Next

        For Each PDF As String In ListPDFSheets
            Show_Document(PDF)
        Next

        If MISSING_IMAGES.Count > 0 Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Missing Images"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("The Following Styles Did Not Have")
            iMSG.AppendLine("Set-up In The Style Masterfile:")
            For Each MI As String In MISSING_IMAGES
                iMSG.AppendLine("-> " & MI)
            Next
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        End If
    End Sub


    Private Sub SetRowImage(row As DataRow)
        Dim STYLE_CODE As String
        If row.Table.TableName = "SOTCUSTQ" Then
            STYLE_CODE = row.Item("STYLE_CODE_PLM") & ""
        Else
            STYLE_CODE = row.Item("STYLE_CODE") & ""
        End If

        Dim IMAGE_NAME As String = row.Item("IMAGE_NAME") & ""

        If IMAGE_NAME = "" Then IMAGE_NAME = STYLE_CODE

        'Dim imgba() As Byte = Nothing
        Dim imgb As System.Drawing.Bitmap = Nothing
        If IMAGE_NAME <> "" Then
            Dim ex_err As Exception = Nothing
            Dim IMAGE_FILE_USED As String = ""
            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")
                ''If chkLowRes.Checked Then
                ''Dim FILE_NAME_LOW_RES As String = String.Format("{0}{1}{2}", FOLDER_NAME, "_lowres\", IMAGE_NAME)
                ''If System.IO.File.Exists(FILE_NAME_LOW_RES) Then
                ''    FOLDER_NAME = FOLDER_NAME & "_lowres"
                ''    IMAGE_FILE_USED = FILE_NAME_LOW_RES
                ''Else
                ''    IMAGE_FILE_USED = String.Format("{0}{1}{2}", FOLDER_NAME, "\", IMAGE_NAME)
                ''End If
                ''End If
            End If
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "dgj")) Then
                'Stop
                FOLDER_NAME = "S:\VAN\images\"
                'FOLDER_NAME = "\\192.168.180.32\g\VAN\images\"
            End If


            Dim img As System.Drawing.Bitmap = Nothing

            Dim image_file_found As Boolean = True

            If IMAGE_NAME = "\.jpg" Then
                image_file_found = False
                Exit Sub
            End If

            If Not FOLDER_NAME.EndsWith("\") Then FOLDER_NAME &= "\"
            Dim IMAGE_FILENAME As String = FOLDER_NAME & IMAGE_NAME
            Try
                If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then

                ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".PNG") Then
                    IMAGE_FILE_USED &= ".PNG"
                ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".JPG") Then
                    IMAGE_FILE_USED &= ".JPG"
                Else
                    image_file_found = False
                    img = Nothing
                End If
            Catch ex As Exception
                image_file_found = False
                img = Nothing
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    ex_err = ex
                End If
            End Try

            Dim fs As System.IO.FileStream = New System.IO.FileStream(IMAGE_FILENAME, System.IO.FileMode.Open)
            Dim newBMP As System.Drawing.Bitmap = New System.Drawing.Bitmap(System.Drawing.Image.FromStream(fs))
            Dim scaleFactor As Double = 1  ' 1 (trkScaleImage.Value / 100)
            Dim newBMP2 As System.Drawing.Bitmap = New System.Drawing.Bitmap(newBMP, newBMP.Width * scaleFactor, newBMP.Height * scaleFactor)
            Application.DoEvents()
            Try
                'newBMP.MakeTransparent(System.Drawing.Color.White)
                Dim converter As New System.Drawing.ImageConverter
                'row.Item("IMAGE") = converter.ConvertTo(newBMP, GetType(Byte()))
                row.Item("IMAGE") = converter.ConvertTo(newBMP2, GetType(Byte()))
                newBMP.Dispose()
                newBMP2.Dispose()
            Catch ex As Exception
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    ex_err = ex
                End If
            End Try
            fs.Close()
            Application.DoEvents()
            If Not IsNothing(ex_err) Then
                If Not IMG_Error_Reported Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Error Getting Image"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("The Following Error Occured While Attempting To ")
                    iMSG.AppendLine("Get An Image:")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Style: " & STYLE_CODE)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Image: " & IMAGE_NAME)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Error: " & ex_err.Message)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Please Relay This Information To Wayne At ABS.")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
                    IMG_Error_Reported = True
                End If
            End If
            'Dim converter As New ImageConverter
            'row.Item("IMAGE") = converter.ConvertTo(imgb, GetType(Byte()))
            'row.Item("IMAGE") = imgb
            'UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE & "-" & COLOR_CODE
        Else
            'row.Item("IMAGE") = DBNull.Value
            'UltraExplorerBar1.Groups("Style Image").Text = "Style Image"
        End If

    End Sub

End Class