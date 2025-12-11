Imports System.Net.Mail
Public Class ARFCALL1
    Dim SQLSB As New Text.StringBuilder
    Dim CallInProgress As Boolean
    Dim AllSelected As Boolean
    Dim IsFollowUpCall As Boolean
    Dim ARTSTMTR As String
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    'Notes On Call Status:
    '-- 0 = Finished.
    '-- 1 = Open.   To Be Followed Up
    '-- 2 = Finished.   Was Followed Up on.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ARFCALLI" Then
            InquiryMode = True
        End If

        'Check_Form_Options()

        With dst

            Get_PARM("ARTPARM1")

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT")
            SQLSB.AppendLine("O1.CUST_CODE,")
            SQLSB.AppendLine("C1.CUST_NAME,")
            SQLSB.AppendLine("C1.SREP_CODE,")
            SQLSB.AppendLine("C1.CUST_STATE,")
            SQLSB.AppendLine("C1.CUST_CREDIT_HOLD,")
            SQLSB.AppendLine("N1.NOTIFY_ACTION,")
            SQLSB.AppendLine("N1.NOTIFY_DATE,")
            SQLSB.AppendLine("NVL(C1.POST_CODE,'REG') AS POST_CODE,")
            SQLSB.AppendLine("MIN(O1.INV_DATE) AS INV_OLDEST,")
            SQLSB.AppendLine("MAX(O1.INV_DATE) AS INV_NEWEST,")
            SQLSB.AppendLine("COUNT(*) AS INVOICE_CNT,")
            SQLSB.AppendLine("SUM(")
            SQLSB.AppendLine("  CASE WHEN SYSDATE - INV_DUE_DATE <= 30 THEN")
            SQLSB.AppendLine("    NVL(O1.INV_BALANCE,0)")
            SQLSB.AppendLine("  ELSE")
            SQLSB.AppendLine("    0")
            SQLSB.AppendLine("  END ) AS DAYS00,")
            SQLSB.AppendLine("SUM(")
            SQLSB.AppendLine("  CASE WHEN (SYSDATE - INV_DUE_DATE > 30) AND (SYSDATE - INV_DUE_DATE <= 60) THEN")
            SQLSB.AppendLine("    NVL(O1.INV_BALANCE,0)")
            SQLSB.AppendLine("  ELSE")
            SQLSB.AppendLine("    0")
            SQLSB.AppendLine("  END ) AS DAYS30,")
            SQLSB.AppendLine("SUM(")
            SQLSB.AppendLine("  CASE WHEN (SYSDATE - INV_DUE_DATE > 60) AND (SYSDATE - INV_DUE_DATE <= 90) THEN")
            SQLSB.AppendLine("    NVL(O1.INV_BALANCE,0)")
            SQLSB.AppendLine("  ELSE")
            SQLSB.AppendLine("    0")
            SQLSB.AppendLine("  END ) AS DAYS60,")
            SQLSB.AppendLine("SUM(")
            SQLSB.AppendLine("  CASE WHEN (SYSDATE - INV_DUE_DATE > 90) AND (SYSDATE - INV_DUE_DATE <= 120) THEN")
            SQLSB.AppendLine("    NVL(O1.INV_BALANCE,0)")
            SQLSB.AppendLine("  ELSE")
            SQLSB.AppendLine("    0")
            SQLSB.AppendLine("  END ) AS DAYS90,")
            SQLSB.AppendLine("SUM(")
            SQLSB.AppendLine("  CASE WHEN (SYSDATE - INV_DUE_DATE > 120) THEN")
            SQLSB.AppendLine("    NVL(O1.INV_BALANCE,0)")
            SQLSB.AppendLine("  ELSE")
            SQLSB.AppendLine("    0")
            SQLSB.AppendLine("  END ) AS DAYS120,")
            SQLSB.AppendLine("SUM(NVL(O1.INV_BALANCE,0)) AS TOTAL_BAL,")
            SQLSB.AppendLine("0 AS ORDS_RELD,")
            SQLSB.AppendLine("0 AS ORDS_PEND,")
            SQLSB.AppendLine("0 AS CUST_CREDIT_LIMIT,")
            SQLSB.AppendLine("'' AS CREDIT_PCT,")
            SQLSB.AppendLine("SYSDATE AS LAST_FU,")
            SQLSB.AppendLine("SYSDATE AS NEXT_FU,")
            SQLSB.AppendLine("'MAP THIS!' AS EMP_NAME_FU")
            SQLSB.AppendLine("FROM ARTOPEN1 O1, ARTCUST1 C1, ARTCUSTN N1")
            SQLSB.AppendLine("WHERE O1.CUST_CODE = C1.CUST_CODE")
            SQLSB.AppendLine("AND C1.CUST_CODE = N1.CUST_CODE (+)")
            SQLSB.AppendLine("AND O1.INV_BALANCE <> 0")
            SQLSB.AppendLine("GROUP BY O1.CUST_CODE, C1.CUST_NAME, C1.SREP_CODE, C1.CUST_CREDIT_HOLD, C1.CUST_STATE, N1.NOTIFY_ACTION, N1.NOTIFY_DATE, NVL(C1.POST_CODE,'REG')")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "ARTOPENX", "**", 0, False)
            .Tables("ARTOPENX").Columns.Add("UPDATED")

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT")
            SQLSB.AppendLine("H1.INV_NO,")
            SQLSB.AppendLine("H1.INV_DATE,")
            SQLSB.AppendLine("O1.INV_DUE_DATE,")
            SQLSB.AppendLine("H1.ORDR_CUST_PO,")
            SQLSB.AppendLine("H1.ORDR_NO,")
            SQLSB.AppendLine("H1.WHSE_CODE,")
            SQLSB.AppendLine("H1.INV_COMMENT,")
            SQLSB.AppendLine("NVL(H1.INV_TOTAL_AMT_CURR,0) AS INV_TOTAL_AMT,")
            SQLSB.AppendLine("NVL(O1.INV_BALANCE,0) AS INV_BALANCE")
            SQLSB.AppendLine("FROM SOTINVH1 H1, (SELECT * FROM ARTOPEN1 UNION SELECT * FROM ARTOPENX) O1")
            SQLSB.AppendLine("WHERE H1.INV_NO = O1.INV_NUM")
            SQLSB.AppendLine("AND H1.CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "PMTINVHX", "**", 0, False, "V", 1)

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT")
            SQLSB.AppendLine("P3.INV_NUM AS INV_NO,")
            SQLSB.AppendLine("P3.PYMT_BATCH_NO,")
            SQLSB.AppendLine("P1.PYMT_BATCH_DATE,")
            SQLSB.AppendLine("P1.BANK_CODE,")
            SQLSB.AppendLine("P3.REASON_CODE,")
            SQLSB.AppendLine("SUM(P3.INV_PMT) AS INV_PMT")
            SQLSB.AppendLine("FROM ARTPYMT3 P3, ARTPYMT1 P1, (SELECT * FROM ARTOPEN1 UNION SELECT * FROM ARTOPENX) O1")
            SQLSB.AppendLine("WHERE P3.PYMT_BATCH_NO = P1.PYMT_BATCH_NO")
            SQLSB.AppendLine("AND P3.INV_NUM = O1.INV_NUM")
            SQLSB.AppendLine("AND O1.CUST_CODE = :PARM1")
            SQLSB.AppendLine("AND P3.INV_TYPE <> 'O'")
            SQLSB.AppendLine("AND P3.INV_NUM IN (")
            SQLSB.AppendLine("  SELECT")
            SQLSB.AppendLine("  H1.INV_NO")
            SQLSB.AppendLine("  FROM SOTINVH1 H1, (SELECT * FROM ARTOPEN1 UNION SELECT * FROM ARTOPENX) O1")
            SQLSB.AppendLine("  WHERE H1.INV_NO = O1.INV_NUM")
            SQLSB.AppendLine("  AND H1.CUST_CODE = :PARM1")
            SQLSB.AppendLine(")")
            SQLSB.AppendLine("GROUP BY")
            SQLSB.AppendLine("P3.INV_NUM,")
            SQLSB.AppendLine("P3.PYMT_BATCH_NO,")
            SQLSB.AppendLine("P1.PYMT_BATCH_DATE,")
            SQLSB.AppendLine("P1.BANK_CODE,")
            SQLSB.AppendLine("P3.REASON_CODE")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "ARTPYMTX", "**", 0, False, "V", 2)

            .Relations.Add("PMTINVHX_ARTPYMTX",
                New DataColumn() { .Tables("PMTINVHX").Columns("INV_NO")},
                New DataColumn() { .Tables("ARTPYMTX").Columns("INV_NO")})

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT TATCONV1.*,")
            SQLSB.AppendLine("' ' AS SELCALL")
            SQLSB.AppendLine("FROM TATCONV1")
            SQLSB.AppendLine("WHERE TATCONV1.TABLE_NAME = 'ARTCUST1'")
            SQLSB.AppendLine("AND TATCONV1.TABLE_KEY = :PARM1")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "TATCONV1", "**", 0, True, "V")

            .Tables("TATCONV1").Columns.Add("CONV_ATTACHMENTS", GetType(System.Int64))


            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT")
            SQLSB.AppendLine("USER_ID AS EMPLOYEE_CODE,")
            SQLSB.AppendLine("USER_NAME AS EMPLOYEE_NAME")
            SQLSB.AppendLine("FROM ASTUSER1")
            SQLSB.AppendLine("WHERE USER_STATUS = 'A'")
            SQLSB.AppendLine("ORDER BY USER_ID")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "PMTEMPL1", "**", 0, False)

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT")
            SQLSB.AppendLine(" *")
            SQLSB.AppendLine(" FROM ASTATTA2")
            Create_TDA(.Tables.Add, "ASTATTA2", "*", 0, True)

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT O1.INIT_DATE ,  O1.INV_DATE, C1.CUST_NAME, O1.INV_BALANCE  FROM")
            SQLSB.AppendLine(" ARTOPEN1 O1, ARTCUST1 C1")
            SQLSB.AppendLine(" WHERE O1.CUST_CODE = C1.CUST_CODE")
            SQLSB.AppendLine(" AND O1.INV_BALANCE > 0")
            SQLSB.AppendLine(" AND O1.INV_NUM NOT IN")
            SQLSB.AppendLine(" (")
            SQLSB.AppendLine(" SELECT T1.TABLE_KEY ")
            SQLSB.AppendLine(" FROM TATCONV1 T1")
            SQLSB.AppendLine(" WHERE T1.TABLE_NAME = 'PMTINVH1'")
            SQLSB.AppendLine(" AND T1.CONV_STATUS = '1'")
            SQLSB.AppendLine(" )")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "MISSING", "**", 0, False)

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "ARTCUSTR", "**", 0, False, "V")
            .Tables("ARTCUSTR").Columns.Add("SIGNATURE_IMAGE", GetType(System.Byte()))
            .Tables("ARTCUSTR").Columns.Add("SIGNATURE_NAME")
            .Tables("ARTCUSTR").Columns.Add("SIGNATURE_TITLE")
            .Tables("ARTCUSTR").Columns.Add("SIGNATURE_PREP")

            'SQLSB.Length = 0
            'SQLSB.AppendLine("SELECT * FROM PMTCONT1 WHERE CONTACT_ID = @PARM1")
            'ASCMAIN1.sql = SQLSB.ToString
            'Create_TDA(.Tables.Add, "PMTCONTR", "**", 0, False, "V")

            'SQLSB.Length = 0
            'SQLSB.AppendLine("SELECT * FROM PMTFIRM1 WHERE FIRM_ID = :PARM1")
            'ASCMAIN1.sql = SQLSB.ToString
            'Create_TDA(.Tables.Add, "PMTFIRMR", "**", 0, False, "V")

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT CUST_CODE, MIN(INV_DATE) OLDEST_INV, SUM(NVL(INV_BALANCE_CURR,0)) TOT_DUE")
            SQLSB.AppendLine(" FROM ARTOPEN1")
            SQLSB.AppendLine(" WHERE CUST_CODE = :PARM1")
            SQLSB.AppendLine(" AND NVL(INV_BALANCE_CURR,0) > 0")
            SQLSB.AppendLine(" GROUP BY CUST_CODE")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "ARTOPENR", "**", 0, False, "V", 1)

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT * FROM ARTOPEN1 WHERE INV_BALANCE > 0 AND CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "ARTOPENF", "**", 0, False, "V")

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT *")
            SQLSB.AppendLine("FROM ARTCUSTD")
            SQLSB.AppendLine("WHERE NVL(CONTACT_NOTE,'NULL') <> 'DELETED'")
            SQLSB.AppendLine("UNION")
            SQLSB.AppendLine("SELECT")
            SQLSB.AppendLine("CUST_CODE,")
            SQLSB.AppendLine("888 AS CONTACT_NO,")
            SQLSB.AppendLine("CUST_CONTACT AS CONTACT_NAME,")
            SQLSB.AppendLine("'MASTER CONTACT' AS CONTACT_TITLE,")
            SQLSB.AppendLine("CUST_EMAIL AS CONTACT_EMAIL,")
            SQLSB.AppendLine("SUBSTR(CUST_PHONE,15) AS CONTACT_PHONE,")
            SQLSB.AppendLine("CUST_EXT AS CONTACT_EXT,")
            SQLSB.AppendLine("CUST_FAX AS CONTACT_FAX,")
            SQLSB.AppendLine("'X' AS CONTACT_TYPE,")
            SQLSB.AppendLine("'1' AS CONTACT_PRIMARY,")
            SQLSB.AppendLine("NULL AS CONTACT_NOTE,")
            SQLSB.AppendLine("NULL AS INIT_OPER,")
            SQLSB.AppendLine("NULL AS LAST_DATE,")
            SQLSB.AppendLine("NULL AS LAST_OPER,")
            SQLSB.AppendLine("NULL AS INIT_DATE,")
            SQLSB.AppendLine("NULL AS CONTACT_CELL")
            SQLSB.AppendLine("FROM ARTCUST1")
            ASCMAIN1.sql = SQLSB.ToString
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "V", 2)

            'SQLSB.Length = 0
            'SQLSB.AppendLine("SELECT *")
            'SQLSB.AppendLine(" FROM pmtjobm1")
            'SQLSB.AppendLine(" WHERE SITE_CODE = :PARM1")
            'SQLSB.AppendLine(" AND (JOB_STATUS = 'O'")
            'SQLSB.AppendLine(" OR (DATE_COMPLETED >= dateadd(month,-6, getdate())))")
            'ASCMAIN1.sql = SQLSB.ToString
            'Create_TDA(.Tables.Add, "PMTJOBMB", "**", 0, False, "V")

            'Create_TDA(.Tables.Add, "ASTATTV1", "*")
            'Fill_Records("ASTATTV1")

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


                ASCMAIN1.sql = "Select ARTCUST1.* from ARTCUST1 WHERE ARTCUST1.CUST_CODE = :PARM1"
                Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)




            End If






        End With

        grdARTOPENX.DataSource = dst.Tables("ARTOPENX")
        grdPMTINVHX.DataSource = dst.Tables("PMTINVHX")
        grdTATCONV1.DataSource = dst.Tables("TATCONV1")
        grdARTCUSTD.DataSource = dst.Tables("ARTCUSTD")

        Create_Summary(grdARTOPENX, "INVOICE_CNT")
        Create_Summary(grdARTOPENX, "DAYS00")
        Create_Summary(grdARTOPENX, "DAYS30")
        Create_Summary(grdARTOPENX, "DAYS60")
        Create_Summary(grdARTOPENX, "DAYS90")
        Create_Summary(grdARTOPENX, "DAYS120")
        Create_Summary(grdARTOPENX, "TOTAL_BAL")

        Create_Summary(grdPMTINVHX, "INV_TOTAL_AMT")
        Create_Summary(grdPMTINVHX, "INV_BALANCE")

        With grdARTOPENX.DisplayLayout.Bands(0)
            For Each COL_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "NOTIFY_DATE", "NOTIFY_ACTION", "UPDATED"}
                .Columns(COL_NAME).Header.Fixed = True
            Next
        End With

        With grdARTOPENX.DisplayLayout.Bands(0).Columns("INVOICE_CNT")
            '.MaskInput = "####"
            .Format = "#,##0"
        End With

        With grdARTOPENX.DisplayLayout.Bands(0).Columns("CUST_CREDIT_HOLD")
            .CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        End With

        Sort_grdColumns(grdTATCONV1, "CONV_DATE".ToLower(), False)

        Add_Attachment_Column(grdTATCONV1, 2, "Y", "TATCONV1", "CONV_NO")
        Add_Attachment_Column(grdPMTINVHX, 2, "Y", "PMTINVH1", "INV_NO")

        ASCMAIN1.Add_Value_List(grdTATCONV1, "CONV_STATUS", , New String() {":", ":Unknown", "0:Closed", "1:Open", "2:Followed Up"})
        ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")
        TABLE_NAME = "ARTOPENX"

        EntryMode = "E"
        Call Load_Record()
        Call Mode_Settings(True)
    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Done"
                If CallInProgress Then
                    EMsg &= "Please Finish The Call In Progress Before Closing This Screen."
                End If
            Case "Refresh"
                Dim iResult As MsgBoxResult
                iResult = MsgBox("Refreshing Grid Will Lose Any un-Saved (O's) Changes.", vbOKCancel, "Refresh Grid")
                If iResult <> vbOK Then
                    EMsg &= "Refresh Canceled"
                End If
            Case "Cancel"
                Dim iResult As MsgBoxResult
                iResult = MsgBox("Cancelling Will Lose Any un-Saved (O's) Changes.", vbOKCancel, "Cancel?")
                If iResult <> vbOK Then
                    EMsg &= "OK"
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
            Case "Done", "Cancel"
                Call Mode_Settings(False)
                Close()
            Case "Refresh"
                Call RefreshOpenx()
                'Case "Save"
                'Call SaveUnfinished()
                'Call RefreshOpenx()
            Case "Print Report"
                Call Print_Report()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If
        grdTATCONV1.DisplayLayout.Bands(0).Columns("SELCALL").CellActivation = UltraWinGrid.Activation.AllowEdit

        grdARTOPENX.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.True
        grdARTOPENX.DisplayLayout.Bands(0).Columns("NOTIFY_DATE").Format = "MM/dd/yy"
        With grdARTOPENX.DisplayLayout.Bands(0)
            For Each col As UltraWinGrid.UltraGridColumn In grdARTOPENX.DisplayLayout.Bands(0).Columns
                col.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("NOTIFY_ACTION").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("NOTIFY_DATE").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        grdARTCUSTD.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.False
        grdARTCUSTD.DisplayLayout.Bands(0).Override.AllowAddNew = DefaultableBoolean.False
        With grdARTCUSTD.DisplayLayout.Bands(0)
            For Each col As UltraWinGrid.UltraGridColumn In grdARTCUSTD.DisplayLayout.Bands(0).Columns
                col.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With


        Fill_Records("MISSING", "", True)
        If dst.Tables.Item("MISSING").Rows.Count > 0 Then
            If ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "cwalsh" Then
                btnBadCalls.Visible = True
            Else
                btnBadCalls.Visible = False
            End If
        Else
            btnBadCalls.Visible = False
        End If
    End Sub

    Sub Clear_Record()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Cursor = Cursors.WaitCursor

        dst.EnforceConstraints = False

        Fill_Records("ARTOPENX")
        ComputeCreditSnapshot()
        Fill_Records("PMTEMPL1")
        Fill_Records("ARTCUSTD")

        dst.EnforceConstraints = True
        SetFollowUpData()
        Setup_ARTCUSTD()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Cursor = Cursors.Default
    End Sub

    Sub Print_Report()

        Print_Report_Begin()

        Generate_Report("ARRCALLC", "Collections Report")

        Print_Report_End()
    End Sub

    Sub RefreshOpenx()

        'Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Refreshing Data")
        Cursor = Cursors.WaitCursor

        dst.EnforceConstraints = False
        dst.Tables("ARTOPENX").Rows.Clear()
        Fill_Records("ARTOPENX")
        ComputeCreditSnapshot()
        'Fill_Records("PMTEMPL1")
        dst.EnforceConstraints = True
        SetFollowUpData()

        'Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Cursor = Cursors.Default
    End Sub

    Sub Update_Record()
        BeginTrans()
        'INIT_LAST("PMTVIST1", True, "", True)
        Update_Record_TDA("TATCONV1")
        CommitTrans()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdARTOPENX, "SSB", "Show Filter", "Show GroupBox", "Show Customer")
        Call Load_Popup_Menu(grdPMTINVHX, "SSB", "Show Filter", "Show GroupBox", "Show Job", "Show Invoice", "e-mail Invoice", "email Cust Statement")
        Call Load_Popup_Menu(grdTATCONV1, "SSBBBBS", "Show Filter", "Show GroupBox", "Add to Log", "Show Log", "Edit Log", "Start Follow-up", "Select All")

        '    Load_Popup_Menu(grdTATCONV1, "BBBB", "Add to Log", "Show Log", "Edit Log", "Follow-Up")

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
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Show Job"
                Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
                Context_Launch("Load", Column_Values("JOB_NO", JOB_NO), e.Tool.Key, "PMFJOBMI")

            Case "Show Customer"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", Column_Values("CUST_CODE", CUST_CODE), e.Tool.Key, "ARFCINQ1")
            Case "Start Follow-up"
                If Not CallInProgress Then
                    If grdTATCONV1.Selected.Rows.Count <> 1 Then
                        MsgBox("You Need To Select A Call To Follow-up on.", vbOKOnly, "Follow-up Selection")
                    Else
                        StartNewCall(True)
                    End If
                Else
                    MsgBox("Call Already in Progress", MsgBoxStyle.Exclamation, "In Progress")
                End If
            Case "Select All", "Un-Select All"
                Dim SelChar As String = "0"
                If AllSelected Then
                    SelChar = "0"
                    AllSelected = False
                Else
                    SelChar = "1"
                    AllSelected = True
                End If
                For Each Grid_Row As UltraWinGrid.UltraGridRow In grdTATCONV1.Rows
                    Grid_Row.Cells("SELCALL").Value = SelChar
                Next
                grdTATCONV1.UpdateData()
            Case "Show Invoice"
                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Text
                Dim FILENAME As String = ASCMAIN1.Folders("Archive") & "Invoices\" & INV_NO & ".pdf"
                Show_Document(FILENAME)
            Case "e-mail Invoice"
                If grdARTOPENX.ActiveRow.IsDataRow Then
                    Dim FILENAME As String = ""
                    Dim ATTACHMENT As String = ""
                    Dim SUBJECT As String = ""
                    Dim INV_NO As String = ""
                    Dim CUST_CODE As String = grdARTOPENX.ActiveRow.Cells.Item("CUST_CODE").Text
                    Dim CUST_NAME As String = Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(grdARTOPENX.ActiveRow.Cells.Item("CUST_NAME").Text.ToLower())
                    Dim CONTACT_EMAIL As String = ""
                    Dim CONTACT_NAME As String = ""
                    Dim Emgs As String = ""
                    If grdARTCUSTD.Selected.Rows.Count = 1 Then
                        CONTACT_EMAIL = grdARTCUSTD.Selected.Rows(0).Cells.Item("CONTACT_EMAIL").Text
                        CONTACT_NAME = Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(grdARTCUSTD.Selected.Rows(0).Cells.Item("CONTACT_NAME").Text.ToLower())
                        If CONTACT_EMAIL.Length = 0 Then
                            Emgs += "Invalid E-Mail For Selected Contact." & vbCrLf
                        Else
                            If Not CONTACT_EMAIL.Contains("@") Then
                                Emgs += "Invalid E-Mail For Selected Contact." & vbCrLf
                            End If
                        End If
                        If CONTACT_NAME.Length = 0 Then
                            Emgs += "Invalid Name For Selected Contact." & vbCrLf
                        End If
                    Else
                        Emgs += "You Must Select A Contact To E-mail" & vbCrLf
                    End If
                    If Emgs.Length <> 0 Then
                        MsgBox(Emgs.ToString(), vbOKOnly, "Error With Selection")
                        Exit Sub
                    End If
                    If grdPMTINVHX.Selected.Rows.Count <= 1 Then
                        INV_NO = grd.ActiveRow.Cells("INV_NO").Value & ""
                    Else
                        For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdPMTINVHX.Selected.Rows
                            INV_NO &= "," & grdRow.Cells("INV_NO").Value & ""
                        Next
                        INV_NO = INV_NO.Substring(1).Trim
                    End If
                    FILENAME = SOCMAIN1.Create_Invoice(Me, INV_NO)
                    SUBJECT = "Invoice " & INV_NO
                    SOCMAIN1.email_Invoice(Me,
                        CUST_CODE,
                        CUST_NAME,
                        CONTACT_EMAIL,
                        CONTACT_NAME,
                        FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT, INV_NO)

                End If

            Case "email Cust Statement"
                If grd.ActiveRow Is Nothing Then
                    Exit Sub
                End If

                Dim CUST_CODE As String = grdARTOPENX.ActiveRow.Cells("CUST_CODE").Value & ""
                If CUST_CODE = "" Then
                    Exit Sub
                End If


                Call AR_STATEMENT()
                Dim FILENAME As String = Print_Hard_Copy_Statement(True, True)

                dst.Tables("ARTSTMTZ").Rows.Clear()


                '  Show_Document(ASCMAIN1.Folders("Temp") & FILENAME & ".PDF")

                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    ' Context_Launch("Select", CUST_CODE, e.Tool.Key, "ARFCINQ1")
                End If





                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                '    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                Dim CUST_NAME As String = rowARTCUST1.Item("CUST_NAME") & ""
                Dim CUST_EMAIL As String = rowARTCUST1.Item("CUST_EMAIL") & ""
                Dim CUST_CONTACT As String = rowARTCUST1.Item("CUST_CONTACT") & ""
                If CUST_EMAIL <> "" Then
                    EMAIL_ADDRESSs.Add(CUST_EMAIL, IIf(CUST_CONTACT = "", CUST_EMAIL, CUST_CONTACT))
                End If

                Dim ATTACHMENTs As New Dictionary(Of String, String)
                ATTACHMENTs.Add(FILENAME & ".pdf", ASCMAIN1.Folders("Temp") & FILENAME & ".PDF")

                Dim SUBJECT As String = "Statement " & " - Customer No " & CUST_CODE
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

                    ''Dim F2 As ASFCONV1 = DirectCast(tabMain.Tabs("Log").TabPage.Controls(0), ASFCONV1)
                    ''If F2.tblTATCONV1 IsNot Nothing Then
                    ''    F2.tblTATCONV1.Rows.Add(rowTATCONV1.ItemArray)
                    ''End If

                    dst.Tables("TATCONV1").Rows.Clear()


                End If

                Exit Sub

            Case "Add to Log"

                'If Not E_.TABLE_KEY_locked Then
                '    If Not ASCMAIN1.Logical_Lock(_E.TABLE_NAME, _E.TABLE_KEY, , , , 1) Then Exit Sub
                'End If
                Dim CUST_CODE As String = grdARTOPENX.ActiveRow.Cells("CUST_CODE").Value
                Dim F As New ASFCONV2(Me, "ARTCUST1", CUST_CODE, "")
                F.EntryMode = "N"
                F.ShowDialog()
                If F.result = "U" Then
                    dst.Tables("TATCONV1").Rows.Add(F.rowTATCONV1.ItemArray)
                    Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)

                    Update_Record_TDA("TATCONV1")
                End If
                F.Dispose()
                ASCMAIN1.MultiTask_Release(, , 1)


            Case "Show Log"
                Dim CONV_NO As String
                CONV_NO = grd.ActiveRow.Cells("CONV_NO").Text
                Dim CUST_CODE As String = grdARTOPENX.ActiveRow.Cells("CUST_CODE").Value


                Dim F As New ASFCONV2(Me, "ARTCUST1", CUST_CODE, "")
                F.EntryMode = "V"
                F.rowTATCONV1 = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                F.ShowDialog()
                F.Dispose()

            Case "Edit Log"
                Dim CONV_NO As String
                CONV_NO = grd.ActiveRow.Cells("CONV_NO").Text
                Dim CUST_CODE As String = grdARTOPENX.ActiveRow.Cells("CUST_CODE").Value


                'If Not _E.TABLE_KEY_locked Then
                '    If Not ASCMAIN1.Logical_Lock(_E.TABLE_NAME, _E.TABLE_KEY, , , , 1) Then Exit Sub
                'End If
                If Not ASCMAIN1.Logical_Lock("TATCONV1", CONV_NO, , , , 1) Then Exit Sub

                Dim F As New ASFCONV2(Me, "ARTCUST1", CUST_CODE, "")
                F.EntryMode = "E"
                F.followup_is_mandatory = Not (EntryMode = "N" Or EntryMode = "E")
                F.rowTATCONV1 = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                F.ShowDialog()
                If F.result = "U" Then
                    Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)

                    'If tabTATCONV1.SelectedTab.Key = "Tree View" Then
                    '    Setup_tvwTATCONV1()
                    'End If
                    Update_Record_TDA("TATCONV1")
                End If
                F.Dispose()
                ASCMAIN1.MultiTask_Release(, , 1)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

#End Region

#Region "grdARTOPENX"

    Private Sub grdARTOPENX_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTOPENX.AfterCellUpdate
        If grdARTOPENX.ActiveRow.IsDataRow Then
            If e.Cell.Column.Key = "NOTIFY_ACTION" Then
                e.Cell.Row.Cells("NOTIFY_DATE").Value = DBNull.Value
                e.Cell.Row.Cells("UPDATED").Value = "O"
                SetModeStatus()
            End If
        End If
    End Sub

    Private Sub grdARTOPENX_AfterRowActivate(ByVal sender As Object, ByVal e As EventArgs) Handles grdARTOPENX.AfterRowActivate
        If grdARTOPENX.ActiveRow.IsDataRow Then
            FetchINVX()
            FetchCONVH()
            Setup_ARTCUSTD()
        End If
    End Sub

    Private Sub grdARTOPENX_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTOPENX.BeforeRowUpdate
        With grdARTOPENX
            If e.Row.Cells("NOTIFY_ACTION").Text <> "" Then
                If IsDate(e.Row.Cells("INV_OLDEST").Text) Then
                    If Not QualifyNotification(e.Row.Cells("NOTIFY_ACTION").Text, e.Row.Cells("INV_OLDEST").Text) Then
                        e.Cancel = True
                        e.Row.Cells("NOTIFY_ACTION").Value = DBNull.Value
                        e.Row.Cells("NOTIFY_DATE").Value = DBNull.Value
                        e.Row.Cells("UPDATED").Value = ""
                        If Not ScreenMode Then
                            Mode_Settings(True)
                        End If
                    End If
                End If
            End If
            'If Not e.Cancel Then
            '    .ActiveRow.Cells("NOTIFY_DATE").Value = System.DBNull.Value
            'End If
        End With
    End Sub

    Private Sub grdARTOPENX_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTOPENX.ClickCellButton

        Select Case e.Cell.Column.Key
            Case "NOTIFY_ACTION"
                Dim sql_where As String = ""
                grdARTOPENX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdClickCellButton(grdARTOPENX, sql_where)
            Case "UPDATED"
                SendNotification()
            Case ""
        End Select
    End Sub
#End Region

#Region "Form Controls"

    Private Sub btnCancelCall_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles btnCancelCall.Click
        If CallInProgress Then
            Dim iResult As MsgBoxResult = MsgBox("Are You Sure You Want To Cancel This Call?", MsgBoxStyle.YesNo, "Cancel")
            If iResult = MsgBoxResult.Yes Then
                SetCallStatus(False)
            End If
        Else
            MsgBox("There Is No Call In Progress")
        End If
    End Sub

    Private Sub btnFinishCall_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles btnFinishCall.Click
        '-- CONV_NO_PREV 'Who did I Follow up on
        '-- CONV_FOLLOWUP_CONV_NO 'Who Followed up on me
        If CallInProgress Then
            Dim msg As String = ""

            If txtPMTEMPL1_EMPLOYEE_CODE.Text = "" Then
                msg = msg & vbCrLf & "You Must Specify The Employee."
            End If
            If Not IsDate(txtCALLDATE.Text) Then
                msg = msg & vbCrLf & "Invalid Call Date."
            End If
            If Not IsDate(txtFOLLOWUPDT.Text) Then
                msg = msg & vbCrLf & "Invalid Follow-Up Date."
            End If
            If txtNotes.Text = "" Then
                msg = msg & vbCrLf & "You Must Supply A Note."
            End If
            If msg <> "" Then
                MsgBox(msg, MsgBoxStyle.OkOnly, "Information Still Required")
                Exit Sub
            End If


            Dim CUST_CODE As String = grdARTOPENX.ActiveRow.Cells("CUST_CODE").Value
            'Dim CUST_NAME As String = grdARTOPENX.ActiveRow.Cells("CUST_NAME").Value


            Dim THISCONV_NO As String = ASCMAIN1.Next_Control_No("TATCONV1.CONV_NO")
            Dim FU_CONV_NO As String = ""
            If IsFollowUpCall Then
                FU_CONV_NO = grdTATCONV1.Selected.Rows(0).Cells.Item("CONV_NO").Text.ToString
                Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Select("CONV_NO = '" & FU_CONV_NO & "'").FirstOrDefault()
                rowTATCONV1.Item("CONV_FOLLOWUP_CONV_NO") = THISCONV_NO
                rowTATCONV1.Item("CONV_STATUS") = "2"
            End If

            Dim rowTATCONV_NEW As DataRow = dst.Tables("TATCONV1").NewRow
            rowTATCONV_NEW.Item("CONV_NO") = THISCONV_NO
            rowTATCONV_NEW.Item("CONV_DATE") = txtCALLDATE.Text
            rowTATCONV_NEW.Item("CONV_SUBJECT") = "Collections Call"
            rowTATCONV_NEW.Item("CONV_NOTES") = txtNotes.Text
            rowTATCONV_NEW.Item("CONV_FOLLOWUP_BY") = txtPMTEMPL1_EMPLOYEE_CODE.Text
            rowTATCONV_NEW.Item("CONV_FOLLOWUP_DATE") = txtFOLLOWUPDT.Text
            rowTATCONV_NEW.Item("CONV_FOLLOWUP_CONV_NO") = Null
            rowTATCONV_NEW.Item("CONV_NO_PREV") = FU_CONV_NO
            rowTATCONV_NEW.Item("CONV_STATUS") = "1"
            rowTATCONV_NEW.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowTATCONV_NEW.Item("INIT_DATE") = DateTime.Now
            rowTATCONV_NEW.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowTATCONV_NEW.Item("LAST_DATE") = DateTime.Now
            rowTATCONV_NEW.Item("TABLE_NAME") = "ARTCUST1"
            rowTATCONV_NEW.Item("TABLE_KEY") = CUST_CODE
            dst.Tables("TATCONV1").Rows.Add(rowTATCONV_NEW)

            'Next
            Update_Record()
            dst.Tables("TATCONV1").Rows.Clear()
            FetchCONVH()
            SetCallStatus(False)
        Else
            MsgBox("There Is No Call In Progress")
        End If
    End Sub

    Private Sub btnStartCall_Click(sender As Object, e As EventArgs) Handles btnStartCall.Click
        StartNewCall(False)
    End Sub

#End Region

#Region "Custom Methods"

    Private Sub FetchCONVH()
        With grdARTOPENX
            If .ActiveRow Is Nothing OrElse .ActiveRow.IsGroupByRow Then

            Else
                Dim Cust_Code As String = .ActiveRow.Cells("CUST_CODE").Text
                Dim Cust_Name As String = .ActiveRow.Cells("CUST_NAME").Text
                EnforceConstraints(False)
                dst.Tables("TATCONV1").Rows.Clear()
                Fill_Records("TATCONV1", Cust_Code)

                If ASCMAIN1.CLIENT = "RGI" Then
                    Fill_Records("ARTSTMTR", Cust_Code)
                    Fill_Records("ARTCUSTA", Cust_Code)
                    Fill_Records("ARTCUST1", Cust_Code)
                End If

                EnforceConstraints(True)
                grdTATCONV1.Text = "Conversations for " & Cust_Name
            End If
            FilterCONVH()
        End With
    End Sub

    Private Sub FetchINVX()
        With grdARTOPENX
            If .ActiveRow Is Nothing OrElse .ActiveRow.IsGroupByRow Then
                'grdARTOPENX.Visible = False
            Else
                Dim CUST_CODE As String = .ActiveRow.Cells("CUST_CODE").Text
                EnforceConstraints(False)
                Fill_Records("PMTINVHX", CUST_CODE, True)
                Fill_Records("ARTPYMTX", CUST_CODE, True)
                Fill_Records("TATCONV1", CUST_CODE, True)
                If ASCMAIN1.CLIENT = "RGI" Then
                    Fill_Records("ARTSTMTR", CUST_CODE, True)
                    Fill_Records("ARTCUSTA", CUST_CODE, True)
                    Fill_Records("ARTCUST1", CUST_CODE, True)
                End If

                EnforceConstraints(True)
                dst.EnforceConstraints = True
            End If
            FilterINVX()

        End With
    End Sub

    Private Sub FilterCONVH()
        Dim dvw As DataView
        Dim SFilter As String = ""
        SFilter = ""

        dvw = DirectCast(grdTATCONV1.DataSource, DataTable).DefaultView
        dvw.RowFilter = SFilter
        grdTATCONV1.Refresh()
    End Sub

    Private Sub FilterINVX()
        Dim dvw As DataView
        Dim SFilter As String = ""

        'If chkALLINV.Checked Then
        '    SFilter = ""
        'Else
        SFilter = "INV_BALANCE <> 0"
        'End If
        dvw = DirectCast(grdPMTINVHX.DataSource, DataTable).DefaultView
        dvw.RowFilter = SFilter
        grdPMTINVHX.Refresh()
    End Sub

    Private Sub FindAndReplace(ByVal doc As Microsoft.Office.Interop.Word.Document,
                               ByVal FindText As String, ByVal ReplaceText As String)
        Dim WordRange As Microsoft.Office.Interop.Word.Range

        Try
            For Each WordRange In doc.StoryRanges

                With WordRange.Find
                    .Text = FindText ' "<<FULLNAME>>"
                    .Replacement.Text = ReplaceText
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                End With


            Next WordRange
        Catch ex As Exception
            ' Do Nothing
        Finally
        End Try

    End Sub

    Function GeneratePDF(ByVal REPORT_NAME As String, ByVal Report_TITLE As String, ByVal CUST_CODE As String) As String
        Dim RPTNO As String = ASCMAIN1.Next_Control_No("CLETTERS")
        Call Print_Report_Begin()
        'CR_params.Add("Y1", Format(YYYY + 0, "0000"))
        Dim FILENAME_temp As String = ASCMAIN1.Folders("Temp") & RPTNO & ".pdf"
        Dim FILENAME As String = ASCMAIN1.Folders("Archive") & "Cletters\" & RPTNO & ".pdf"
        'Generate_Report("PMRINVP1", "Invoices to be emailed", , , "PDF", RPTNO, False)
        Generate_Report(REPORT_NAME, Report_TITLE, "", "", "PDF", RPTNO, False)
        My.Computer.FileSystem.CopyFile(FILENAME_temp, FILENAME, True)
        Call Print_Report_End(True, True)
        Return FILENAME
    End Function

    Private Function QualifyNotification(ByVal NOTIFY_ACTION As String, ByVal INV_OLDEST As Date) As Boolean
        Dim RetVal As Boolean = True
        If NOTIFY_ACTION <> "" Then
            If IsDate(INV_OLDEST) Then
                If DateDiff(DateInterval.Day, INV_OLDEST, Now()) <= 30 Then
                    Dim S As New Text.StringBuilder() With {.Length = 0}
                    S.AppendLine("SELECT NVL(NOTIFY_OVERDAYS,0) AS NOTIFY_OVERDAYS")
                    S.AppendLine("FROM ARTNOTF1")
                    S.AppendLine("WHERE NOTIFY_ACTION =  '" & NOTIFY_ACTION & "'")
                    ASCMAIN1.sql = S.ToString()
                    Dim NOTIFY_OVERDAYS As String = ASCDATA1.GetDataValue
                    If NOTIFY_OVERDAYS = "1" Then
                        MsgBox("You May Not Set A Notification On Accounts Less Than 31 Days Overdue", MsgBoxStyle.Critical, "Young Invoices")
                        RetVal = False
                    End If
                End If
            End If
        End If
        Return RetVal
    End Function

    Private Sub SetCallStatus(ByVal IsRunning As Boolean, Optional IsFollow As Boolean = False)
        If IsRunning Then
            Dim USER_ID As String = ASCMAIN1.USER_ID
            If USER_ID = "wayne" Then
                USER_ID = "ana"
            End If

            Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", USER_ID)
            If rowASTUSER1 IsNot Nothing Then
                txtPMTEMPL1_EMPLOYEE_NAME.Text = rowASTUSER1.Item("USER_NAME").ToString()
                txtPMTEMPL1_EMPLOYEE_CODE.Text = rowASTUSER1.Item("USER_ID").ToString()
            End If
            lblMessage.Text = "Call In Progress"
            lblCurrentCall.Text = "Currently Calling"
            txtFOLLOWUPDT.Value = DateTime.Now.AddDays(14).Date
        Else
            txtPMTEMPL1_EMPLOYEE_NAME.Text = ""
            txtPMTEMPL1_EMPLOYEE_CODE.Text = ""
            txtNotes.Text = ""
            lblMessage.Text = ""
            lblCurrentCall.Text = ""
        End If
        UltraTabControl2.Tabs.Item(3).Visible = IsRunning
        UltraTabControl2.Tabs.Item(3).Active = True
        lblMessage.Visible = IsRunning
        lblCurrentCall.Visible = IsRunning
        btnStartCall.Visible = Not IsRunning
        btnFinishCall.Enabled = IsRunning
        btnCancelCall.Enabled = IsRunning
        btnFinishCall.Visible = IsRunning
        btnCancelCall.Visible = IsRunning
        CallInProgress = IsRunning
        IsFollowUpCall = IsFollow
        Mode_Settings(IsRunning)
    End Sub

    Private Sub SetFollowUpData()
        Dim FirstFound As Boolean = False
        Dim LAST_FU As Date
        Dim NEXT_FU As Date
        Dim EMP_NAME_FU As String = ""

        For Each rowARTOPENX As DataRow In dst.Tables("ARTOPENX").Select()
            EMP_NAME_FU = ""
            LAST_FU = "1/1/1800"
            NEXT_FU = "1/1/1800"

            SQLSB.Length = 0
            SQLSB.AppendLine("SELECT")
            SQLSB.AppendLine("TATCONV1.CONV_NO,")
            SQLSB.AppendLine("TATCONV1.CONV_STATUS,")
            SQLSB.AppendLine("TATCONV1.CONV_DATE,")
            SQLSB.AppendLine("TATCONV1.CONV_SUBJECT,")
            SQLSB.AppendLine("TATCONV1.CONV_NOTES,")
            SQLSB.AppendLine("ASTUSER1.USER_NAME  AS FU_EMPLOYEE_NAME,")
            SQLSB.AppendLine("TATCONV1.CONV_FOLLOWUP_DATE")
            SQLSB.AppendLine("FROM TATCONV1, ASTUSER1")
            SQLSB.AppendLine("WHERE TATCONV1.CONV_FOLLOWUP_BY = ASTUSER1.USER_ID")
            SQLSB.AppendLine("AND TATCONV1.TABLE_NAME = 'ARTCUST1'")
            SQLSB.AppendLine("AND TATCONV1.TABLE_KEY = '" & rowARTOPENX.Item("CUST_CODE").ToString & "'")
            'SQLSB.AppendLine("AND TATCONV1.CONV_STATUS = '1'")
            SQLSB.AppendLine("AND TATCONV1.CONV_SUBJECT = 'Collections Call'")
            'SQLSB.AppendLine("AND NVL(TATCONV1.CONV_FOLLOWUP_BY,'') <> ''")
            SQLSB.AppendLine("ORDER BY TATCONV1.CONV_FOLLOWUP_DATE")
            ASCMAIN1.sql = SQLSB.ToString
            For Each rowTATCONV1 As DataRow In ASCDATA1.GetDataTable.Rows
                If FirstFound = False Then
                    FirstFound = True
                    NEXT_FU = rowTATCONV1.Item("CONV_FOLLOWUP_DATE").ToString
                    EMP_NAME_FU = Mid(rowTATCONV1.Item("FU_EMPLOYEE_NAME").ToString, 1, 9)
                End If
                LAST_FU = rowTATCONV1.Item("CONV_DATE").ToString

            Next
            FirstFound = False

            If LAST_FU = "1/1/1800" Then
                rowARTOPENX.Item("LAST_FU") = DBNull.Value
            Else
                rowARTOPENX.Item("LAST_FU") = LAST_FU
            End If
            If NEXT_FU = "1/1/1800" Then
                rowARTOPENX.Item("NEXT_FU") = DBNull.Value
            Else
                rowARTOPENX.Item("NEXT_FU") = NEXT_FU
            End If
            rowARTOPENX.Item("EMP_NAME_FU") = EMP_NAME_FU
        Next
    End Sub

    Private Sub SetModeStatus()
        Dim UnSavedFound As Boolean = False
        For Each rowARTOPENX As DataRow In dst.Tables("ARTOPENX").Select()
            If rowARTOPENX.Item("UPDATED").ToString = "O" Then
                UnSavedFound = True
                Exit For
            End If
        Next

        If UnSavedFound Then
            Mode_Settings(True)
        Else
            Mode_Settings(False)
        End If
    End Sub

    Sub Setup_ARTCUSTD()
        Dim Filter As String = "CUST_CODE = 'XXXXX'"
        Dim dvw As DataView = DirectCast(grdARTCUSTD.DataSource, DataTable).DefaultView
        If grdARTOPENX.ActiveRow Is Nothing OrElse Not grdARTOPENX.ActiveRow.IsDataRow Then
            grdARTCUSTD.Text = "Select A Customer to See Contacts"
        Else
            Dim CUST_CODE As String = grdARTOPENX.ActiveRow.Cells("CUST_CODE").Value
            Dim CUST_NAME As String = grdARTOPENX.ActiveRow.Cells("CUST_NAME").Value
            Filter = "CUST_CODE = '" & CUST_CODE & "'"
            grdARTCUSTD.Text = "Contacts On File For " & CUST_NAME
        End If
        dvw.RowFilter = Filter
    End Sub

    Private Sub StartNewCall(ByVal FU As Boolean)
        If CallInProgress Then
            MsgBox("There Is Already A Call In Progress")
        Else
            SetCallStatus(True, FU)
        End If
    End Sub

    Private Sub SendNotification()
        Dim Emgs As String = ""
        Dim CUST_CODE As String = ""
        Dim CUST_NAME As String = ""
        Dim CONTACT_EMAIL As String = ""
        Dim CONTACT_NAME As String = ""
        Dim NOTIFY_ACTION As String = ""
        Dim rowARTNOTF1 As DataRow
        Dim WORDDOC As String = ""

        'These all need to be parameterized before Wayne Retires.
        Const FROM_ADDRESS As String = "claims@regency-rib.com"
        Const FROM_NAME As String = "Regency Claims"
        Const SERVER_IP As String = "192.168.110.221"
        Const SERVER_PORT As Integer = 25
        Const SERVER_ACCOUNT As String = "claims@regency-rib.com"
        Const SERVER_PASSWORD As String = "ineedavacation"

        If grdARTOPENX.ActiveRow.IsDataRow Then
            CUST_CODE = grdARTOPENX.ActiveRow.Cells.Item("CUST_CODE").Text
            CUST_NAME = grdARTOPENX.ActiveRow.Cells.Item("CUST_NAME").Text
            CUST_NAME = Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CUST_NAME.ToLower())
            NOTIFY_ACTION = grdARTOPENX.ActiveRow.Cells.Item("NOTIFY_ACTION").Text
            If NOTIFY_ACTION.Length = 0 Then
                Emgs += "No Notification Type Selected" & vbCrLf
            Else
                rowARTNOTF1 = LookUp("ARTNOTF1", NOTIFY_ACTION)
                If IsNothing(rowARTNOTF1) Then
                    Emgs += "Invalid Notification Type Defined" & vbCrLf
                Else
                    WORDDOC = rowARTNOTF1.Item("NOTIFY_REPORT").ToString()
                End If
            End If
        Else
            Emgs += "Invalid Active Row for Selected Customer" & vbCrLf
        End If
        If grdARTCUSTD.Selected.Rows.Count = 1 Then
            CONTACT_EMAIL = grdARTCUSTD.Selected.Rows(0).Cells.Item("CONTACT_EMAIL").Text
            CONTACT_NAME = grdARTCUSTD.Selected.Rows(0).Cells.Item("CONTACT_NAME").Text
            If CONTACT_EMAIL.Length = 0 Then
                Emgs += "Invalid E-Mail For Selected Contact." & vbCrLf
            Else
                If Not CONTACT_EMAIL.Contains("@") Then
                    Emgs += "Invalid E-Mail For Selected Contact." & vbCrLf
                End If
            End If
            If CONTACT_NAME.Length = 0 Then
                Emgs += "Invalid Name For Selected Contact." & vbCrLf
            End If
        Else
            Emgs += "You Must Select A Contact To E-mail" & vbCrLf
        End If

        If Emgs.Length <> 0 Then
            MsgBox(Emgs.ToString(), vbOKOnly, "Error With Selection")
            Exit Sub
        End If

        Dim W As New Microsoft.Office.Interop.Word.Application
        Dim WD As Microsoft.Office.Interop.Word.Document = W.Documents.Open(WORDDOC)
        FindAndReplace(WD, "<<FullName>>", Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CONTACT_NAME.ToLower()))
        WD.SaveAs(ASCMAIN1.Folders("Temp") & "CollectionsHTML.doc", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatHTML)
        WD.SaveAs(ASCMAIN1.Folders("Temp") & "CollectionsText.doc", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDOSTextLineBreaks)
        WD.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges,
                      Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat)
        W.Quit()
        W = Nothing

        Dim HTMLString As String = New Net.WebClient().DownloadString(ASCMAIN1.Folders("Temp") & "CollectionsHTML.doc")
        Dim TextString As String = New Net.WebClient().DownloadString(ASCMAIN1.Folders("Temp") & "CollectionsText.doc")

        Dim mail As New MailMessage()
        mail.From = New MailAddress(FROM_ADDRESS, FROM_NAME)
        mail.To.Add(New MailAddress(CONTACT_EMAIL, CONTACT_NAME))
        mail.Bcc.Add(New MailAddress("whr@waynerichmond.net", "Wayne Richmond"))
        mail.Subject = String.Format("Collections Notice for {0} (Account: {1})", CUST_NAME, CUST_CODE)
        mail.IsBodyHtml = True
        mail.Body = HTMLString
        'mail.CC.Add(New MailAddress(ccSREPEmail, ccSPREPName))

        Dim smtp As New SmtpClient(SERVER_IP, SERVER_PORT)
        If smtp IsNot Nothing Then
            smtp.Credentials = New Net.NetworkCredential(SERVER_ACCOUNT, SERVER_PASSWORD)
        Else
            Dim eMsg As String = "SMTP Client could not be created."
            MsgBox(eMsg, MsgBoxStyle.OkOnly, "Error")
        End If
        If ASCMAIN1.Running_in_VS Then
            Stop
        Else
            Try
                smtp.Send(mail)
            Catch ex As Exception
                MsgBox(ex.InnerException, MsgBoxStyle.OkOnly, "Error Sending E-mail")
            End Try
        End If

        Dim SQLS As New Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine(String.Format("Select Count(*) from ARTCUSTN where CUST_CODE = '{0}'", CUST_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim RecCnt As Int16 = Val(ASCDATA1.GetDataValue)
        If RecCnt = 0 Then
            SQLS.Length = 0
            SQLS.AppendLine(String.Format("INSERT INTO ARTCUSTN VALUES('{0}','{1}',SYSDATE)", CUST_CODE, NOTIFY_ACTION))
        Else
            SQLS.Length = 0
            SQLS.AppendLine(String.Format("UPDATE ARTCUSTN SET NOTIFY_ACTION = '{0}', NOTIFY_DATE = SYSDATE WHERE CUST_CODE = '{1}'", NOTIFY_ACTION, CUST_CODE))
        End If
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        grdARTOPENX.ActiveRow.Cells.Item("NOTIFY_DATE").Value = Now()
        grdARTOPENX.ActiveRow.Cells.Item("NOTIFY_ACTION").Value = NOTIFY_ACTION

        'TODO: Add record to TATCONV1.
        Dim NextFU As Date = Now.AddDays(15)
        Dim THISCONV_NO As String = ASCMAIN1.Next_Control_No("TATCONV1.CONV_NO")
        Dim rowTATCONV_NEW As DataRow = dst.Tables("TATCONV1").NewRow
        rowTATCONV_NEW.Item("CONV_NO") = THISCONV_NO
        rowTATCONV_NEW.Item("CONV_DATE") = Now
        rowTATCONV_NEW.Item("CONV_SUBJECT") = "Collections Letter Sent"
        rowTATCONV_NEW.Item("CONV_NOTES") = TextString
        rowTATCONV_NEW.Item("CONV_FOLLOWUP_BY") = ASCMAIN1.USER_ID
        rowTATCONV_NEW.Item("CONV_FOLLOWUP_DATE") = NextFU
        rowTATCONV_NEW.Item("CONV_FOLLOWUP_CONV_NO") = Null
        rowTATCONV_NEW.Item("CONV_NO_PREV") = Null
        rowTATCONV_NEW.Item("CONV_STATUS") = "1"
        rowTATCONV_NEW.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowTATCONV_NEW.Item("INIT_DATE") = DateTime.Now
        rowTATCONV_NEW.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowTATCONV_NEW.Item("LAST_DATE") = DateTime.Now
        rowTATCONV_NEW.Item("TABLE_NAME") = "ARTCUST1"
        rowTATCONV_NEW.Item("TABLE_KEY") = CUST_CODE
        dst.Tables("TATCONV1").Rows.Add(rowTATCONV_NEW)
        Update_Record()
        dst.Tables("TATCONV1").Rows.Clear()
        FetchCONVH()
    End Sub
    Private Sub ComputeCreditSnapshot()
        If dst.Tables("ARTOPENX") Is Nothing OrElse dst.Tables("ARTOPENX").Rows.Count = 0 Then Exit Sub

        Dim distinctCusts As DataTable = dst.Tables("ARTOPENX").DefaultView.ToTable(True, "CUST_CODE")
        Dim sbIn As New Text.StringBuilder()
        Dim ids As New List(Of String)
        For Each r As DataRow In distinctCusts.Rows
            Dim c As String = CStr(r("CUST_CODE") & "")
            If c <> "" Then
                ids.Add(c)
                c = c.Replace("'", "''")
                If sbIn.Length > 0 Then sbIn.Append(",")
                sbIn.Append("'").Append(c).Append("'")
            End If
        Next
        If ids.Count = 0 Then Exit Sub
        Dim inList As String = sbIn.ToString()

        Dim sb As New Text.StringBuilder()
        sb.Length = 0
        sb.AppendLine("SELECT C1.CUST_CODE, NVL(C1.CUST_CREDIT_LIMIT,0) CUST_CREDIT_LIMIT")
        sb.AppendLine("FROM ARTCUST1 C1")
        sb.AppendLine("WHERE " & BuildOracleInChunks("C1.CUST_CODE", ids))
        ASCMAIN1.sql = sb.ToString()
        Dim dtCL As DataTable = ASCDATA1.GetDataTable()

        Dim CL As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        For Each r As DataRow In dtCL.Rows
            CL(CStr(r("CUST_CODE"))) = Val(r("CUST_CREDIT_LIMIT") & "")
        Next

        sb.Length = 0
        sb.AppendLine("SELECT S1.CUST_CODE,")
        sb.AppendLine("  SUM(NVL(S2.ORDR_QTY_PICK,0)*NVL(S2.ORDR_UNIT_PRICE,0)) ORDS_RELD,")
        sb.AppendLine("  SUM(NVL(S2.ORDR_QTY_OPEN,0)*NVL(S2.ORDR_UNIT_PRICE,0)) ORDS_PEND")
        sb.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2")
        sb.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
        sb.AppendLine("  AND S1.ORDR_STATUS IN ('O','P')")
        sb.AppendLine("  AND " & BuildOracleInChunks("S1.CUST_CODE", ids))
        sb.AppendLine("GROUP BY S1.CUST_CODE")
        ASCMAIN1.sql = sb.ToString()
        Dim dtOO As DataTable = ASCDATA1.GetDataTable()

        Dim REL As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        Dim PEND As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        For Each r As DataRow In dtOO.Rows
            Dim k As String = CStr(r("CUST_CODE"))
            REL(k) = Val(r("ORDS_RELD") & "")
            PEND(k) = Val(r("ORDS_PEND") & "")
        Next
        sb.Length = 0
        sb.AppendLine("SELECT O1.CUST_CODE,")
        sb.AppendLine("       SUM(NVL(O1.INV_BALANCE,0)) AS TOTAL_DUE ")
        sb.AppendLine("FROM ARTOPEN1 O1")
        sb.AppendLine("WHERE " & BuildOracleInChunks("O1.CUST_CODE", ids))
        sb.AppendLine("GROUP BY O1.CUST_CODE")
        ASCMAIN1.sql = sb.ToString()
        Dim dtDue As DataTable = ASCDATA1.GetDataTable()
        Dim DUE As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        For Each r As DataRow In dtDue.Rows
            Dim k As String = CStr(r("CUST_CODE"))
            DUE(k) = Val(r("TOTAL_DUE") & "")
        Next

        For Each r As DataRow In dst.Tables("ARTOPENX").Rows
            Dim cust As String = CStr(r("CUST_CODE"))
            Dim limit As Decimal = If(CL.ContainsKey(cust), CL(cust), 0D)
            Dim ordR As Decimal = If(REL.ContainsKey(cust), REL(cust), 0D)
            Dim ordP As Decimal = If(PEND.ContainsKey(cust), PEND(cust), 0D)

            Dim totalDue As Decimal = If(DUE.ContainsKey(cust), DUE(cust), 0D)

            r("CUST_CREDIT_LIMIT") = limit
            r("ORDS_RELD") = ordR
            r("ORDS_PEND") = ordP

            Dim avail As Decimal = limit - totalDue - ordR
            If avail < 0D Then avail = 0D
            r("CREDIT_PCT") = If(limit > 0D, Math.Round(100D * avail / limit, 0, MidpointRounding.AwayFromZero), 0D)

        Next

        With grdARTOPENX.DisplayLayout.Bands(0)
            For Each col As String In New String() {"ORDS_RELD", "ORDS_PEND", "CREDIT_PCT", "CUST_CREDIT_LIMIT"}
                If grdARTOPENX.DisplayLayout.Bands(0).Columns.Exists(col) Then
                    grdARTOPENX.DisplayLayout.Bands(0).Columns(col).Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    grdARTOPENX.DisplayLayout.Bands(0).Columns(col).Header.Appearance.BackColor = System.Drawing.Color.White
                    grdARTOPENX.DisplayLayout.Bands(0).Columns(col).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With
    End Sub
    Private Function BuildOracleInChunks(columnName As String,
                                     values As IEnumerable(Of String),
                                     Optional chunkSize As Integer = 999) As String
        'using batches of 999 to get around 1000 limit in Oracle
        Dim list As New List(Of String)
        For Each v As String In values
            If Not String.IsNullOrEmpty(v) Then
                list.Add("'" & v.Replace("'", "''") & "'")
            End If
        Next
        If list.Count = 0 Then Return "1=0"

        Dim sb As New Text.StringBuilder()
        Dim i As Integer = 0
        Do While i < list.Count
            If sb.Length > 0 Then sb.Append(" OR ")
            Dim take As Integer = Math.Min(chunkSize, list.Count - i)
            sb.Append(columnName).Append(" IN (").Append(String.Join(",", list.GetRange(i, take))).Append(")")
            i += take
        Loop
        Return "(" & sb.ToString() & ")"
    End Function

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

#End Region
End Class