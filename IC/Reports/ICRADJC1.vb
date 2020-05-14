Imports System.Math

Public Class ICRADJC1
    Dim CURR_EXCH_RATE_response As String
    Dim WithEvents http1 As New nsoftware.IPWorks.Http

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("ICTPARM1")
        http1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwarehttpkey")
    End Sub

    Protected Overrides Sub Build_Workfile()

        'Dim conversion As Boolean = False

        'If ASCMAIN1.DBS_COMPANY = "COS" And ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.Running_in_VS Then
        '    conversion = True
        '    Stop
        'End If

        RWU = "R"
        Dim sqlw As String = ""
        Prepare_dst(True, sqlw)

        Dim GP_ELIM As String = ""
        ASCMAIN1.sql = "Select DIVISION_CODE from SOTSDIV1 where ACCT_CODE_GP_ELIM is Not Null"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            GP_ELIM = GP_ELIM & ",'" & row.Item("DIVISION_CODE") & "'"
        Next

        'Dim WHSE_CODE As String = ""
        'Dim WHSE_CODE_rate As String = ""
        'Dim PROD_CODE As String = ""
        'Dim PACK_CODE As String = ""
        'Dim WHSE_STORAGE_CLASS_CODE As String = ""
        'Dim LOT_NO As String = ""
        'Dim WHSE_NO_STG_ACCRUAL As String = ""

        'Dim DATE_WHSE_ANNIV As Date

        'Dim WHSE_STORAGE_per_unit As Decimal
        'Dim PACK_FACTOR As Decimal
        'Dim WHSE_MIN_STORAGE As Decimal
        'Dim ADDED_STORAGE As Decimal


        'Dim MIN_STORAGE As String = ""

        ' this routine is currently not applying minimum storage charges correctly for slack lots
        ' problems - what if products are different, what if DATE_WHSE_ANNIV dates differ
        ' weird things will occur if the product is sold to zero and then re-activated by a return

        Prepare_GL_Interface("ICSR")

        'Check_if_Empty("ICTLOTDX")
    End Sub

    Public Overrides Sub Print_Report()
        ' Generate_Report(RPT, , SUBT)
        Print_GL()
    End Sub

    Overrides Sub Update_Record()

        ' Add a Day's worth of Interest to Adjusted Cost and Adjusted Cost of ICTLOTD1

        Dim IC_PARM_INT_PCT As Decimal = Val(ROWs("ICTPARM1").Item("IC_PARM_INT_PCT") & "")
        Dim IC_PARM_DATE_LAST_INT As Date
        If ROWs("ICTPARM1").Item("IC_PARM_DATE_LAST_INT") & "" = "" Then
            IC_PARM_DATE_LAST_INT = DATETIME_STAMP.Date
        Else
            IC_PARM_DATE_LAST_INT = ROWs("ICTPARM1").Item("IC_PARM_DATE_LAST_INT")
        End If
        Dim days_int As Integer = DATETIME_STAMP.Date.Subtract(IC_PARM_DATE_LAST_INT).Days

        ASCMAIN1.sql = "Insert into ICTLOTDI Select WHSE_CODE, LOT_NO, LOT_SEQ_NO" _
        & ", '" & XNO & "' REGISTER_XNO" _
        & ", NVL(ADJUSTED_COST,0) " _
        & " * " & CStr(days_int) & " * " & CStr(IC_PARM_INT_PCT) & " / 100 / 365" _
        & ", SYSDATE INIT_DATE, '" & ASCMAIN1.USER_ID & "' INIT_OPER" _
        & " from ICTLOTD1" _
        & " where QTY_ON_HAND <> 0 AND ADJUSTED_COST <> 0"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update ICTLOTD1 set ADJUSTED_COST = NVL(ADJUSTED_COST,0) + NVL(ADJUSTED_COST,0) " _
        & " * " & CStr(days_int) & " * " & CStr(IC_PARM_INT_PCT) & " / 100 / 365" _
        & " where QTY_ON_HAND <> 0"
        ASCDATA1.ExecuteSQL()

        ' Update Warehouse Storage Accruals

        Update_Record_TDA("ICTLOTDX")

        ASCMAIN1.sql = "" _
        & " Begin" _
        & "  Declare Cursor C1 is " _
        & "   Select * from ICTLOTDX " _
        & "     where REGISTER_XNO = '" & XNO & "'" _
        & "     order by DATE_WHSE_ANNIV_NEW;" _
        & "  Begin" _
        & "   For R1 in C1 Loop" _
        & "    Update ICTLOTD1 set" _
        & "      ADJUSTED_COST = NVL(ADJUSTED_COST,0) + R1.ADDED_STORAGE" _
        & "     ,STANDARD_COST = NVL(STANDARD_COST,0) + R1.ADDED_STORAGE" _
        & "     ,DATE_WHSE_ANNIV = R1.DATE_WHSE_ANNIV_NEW" _
        & "    Where WHSE_CODE = R1.WHSE_CODE and LOT_NO = R1.LOT_NO " _
        & "      and LOT_SEQ_NO = R1.LOT_SEQ_NO;" _
        & "   End Loop;" _
        & "  End;" _
        & " End;"
        ASCDATA1.ExecuteSQL()

        ' Update Parameter Record to reflect Interest Calculated til Date

        ASCMAIN1.sql = "Update ICTPARM1 Set IC_PARM_DATE_LAST_INT = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "D", DATETIME_STAMP.Date)

        ' Update Daily On Hand

        '  TAC.ICCMAIN1.Calc_Daily_On_Hand(ASCMAIN1.CYP)

        ' Update Foreign Currency History Table



        Fill_Records("ICTCURR1")
        For Each rowICTCURR1 As DataRow In dst.Tables("ICTCURR1").Rows
            Dim CURR_CODE As String = rowICTCURR1.Item("CURR_CODE")
            Dim rowICTCURR2 As DataRow = dst.Tables("ICTCURR2").NewRow
            rowICTCURR2.Item("CURR_CODE") = CURR_CODE
            rowICTCURR2.Item("CURR_DATE") = DATETIME_STAMP.Date

            CURR_EXCH_RATE_response = ""
            ASCMAIN1.sql = "http://finance.yahoo.com/d/quotes.csv?s=" & CURR_CODE & "USD=X&f=l1&e=.txt"
            Try
                'System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
                http1.FollowRedirects = nsoftware.IPWorks.HttpFollowRedirects.frAlways
                http1.TransferredDataLimit = 0
                http1.Get(ASCMAIN1.sql)
                'System.Windows.Forms.Cursor.Current = Cursors.Default
            Catch ex1 As Exception
                MessageBox.Show("Error: " & ex1.Message, "Error", MessageBoxButtons.OK)
            End Try
            Dim CURR_EXCH_RATE As Decimal = 0
            If CURR_EXCH_RATE_response <> "" Then
                CURR_EXCH_RATE = Val(CURR_EXCH_RATE_response)
            End If
            rowICTCURR2.Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
            dst.Tables("ICTCURR2").Rows.Add(rowICTCURR2)
        Next
        Update_Record_TDA("ICTCURR2", "CURR_DATE = '" & Format(DATETIME_STAMP.Date, "dd-MMM-yyyy") & "'")

        '   Prepare_Journal()
        GL_Update()

        Rebuild_ICTLOTD2()

    End Sub


    Private Sub Http1_OnTransfer(ByVal sender As System.Object, ByVal e As nsoftware.IPWorks.HttpTransferEventArgs) Handles Http1.OnTransfer
        If e.Direction = 1 Then CURR_EXCH_RATE_response = e.Text
    End Sub

    Sub Rebuild_ICTLOTD2()

        ' Rebuild ICTLOTD2

        ASCMAIN1.sql = "Delete from ICTLOTD2"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into ICTLOTD2 SELECT ICTLOTD1.*, SOTTERR1.DIVISION_CODE," _
        & " ICTWHSE1.TERR_CODE, ICTPROD1.CLASS_CODE, ICTCLAS1.CATEGORY_CODE" _
        & " From ICTLOTD1, ICTWHSE1, SOTTERR1, ICTPROD1, ICTCLAS1" _
        & " Where ((QTY_ON_HAND <> 0 Or QTY_COMMITTED <> 0) or DATE_LAST_TRAN > SYSDATE - 100)" _
        & " AND ICTWHSE1.WHSE_CODE (+) = ICTLOTD1.WHSE_CODE" _
        & " AND SOTTERR1.TERR_CODE (+) = ICTWHSE1.TERR_CODE" _
        & " AND ICTPROD1.PROD_CODE (+) = ICTLOTD1.PROD_CODE" _
        & " AND ICTCLAS1.CLASS_CODE (+) = ICTPROD1.CLASS_CODE"
        ASCDATA1.ExecuteSQL()

        ' the transfer lots (found using the PL/SQL below) are required for the In Transit
        ' combination with afloat in Production Requirements

        ASCMAIN1.sql = "BEGIN" _
        & " DECLARE CURSOR C1 IS" _
        & " SELECT ICTTRAN2.WHSE_CODE, ICTTRAN2.LOT_NO, ICTTRAN2.LOT_SEQ_NO" _
        & " FROM ICTTRAN2,ICTTRAN1" _
        & "  WHERE ICTTRAN1.TRANSFER_STATUS = 'O'" _
        & "  AND ICTTRAN1.TRANSFER_NO = ICTTRAN2.TRANSFER_NO" _
        & " MINUS" _
        & " SELECT WHSE_CODE, LOT_NO, LOT_SEQ_NO FROM ICTLOTD2;" _
        & " BEGIN" _
        & " FOR R1 IN C1 LOOP" _
        & " INSERT INTO ICTLOTD2 SELECT ICTLOTD1.*, SOTTERR1.DIVISION_CODE," _
        & " ICTWHSE1.TERR_CODE, ICTPROD1.CLASS_CODE, ICTCLAS1.CATEGORY_CODE" _
        & " From ICTLOTD1, ICTWHSE1, SOTTERR1, ICTPROD1, ICTCLAS1" _
        & " where ICTWHSE1.WHSE_CODE (+) = ICTLOTD1.WHSE_CODE" _
        & "   and SOTTERR1.TERR_CODE (+) = ICTWHSE1.TERR_CODE" _
        & "   and ICTPROD1.PROD_CODE (+) = ICTLOTD1.PROD_CODE" _
        & "   and ICTCLAS1.CLASS_CODE (+) = ICTPROD1.CLASS_CODE" _
        & "   and ICTLOTD1.WHSE_CODE = R1.WHSE_CODE AND ICTLOTD1.LOT_NO = R1.LOT_NO" _
        & "   and ICTLOTD1.LOT_SEQ_NO = R1.LOT_SEQ_NO;" _
        & " END LOOP;" _
        & " END;" _
        & " END;"
        ASCDATA1.ExecuteSQL()
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        'Dim sqlw As String = CStr(parms(0))
        'If sqlw = "" Then sqlw = "ROWNUM < 1"


        ASCMAIN1.sql = "Select * from ICTLOTD1 " _
        & " where DATE_WHSE_ANNIV <= :PARM1" _
        & " and QTY_ON_HAND <> 0" _
        & " order by WHSE_CODE, PROD_CODE, LOT_NO, LOT_SEQ_NO"
        Create_TDA(dst.Tables.Add, "ICTLOTD1", "**", 0, False, "D", 3)

        ASCMAIN1.sql = "Select * from ICTCURR1 where CURR_CODE <> 'USD'"
        Create_TDA(dst.Tables.Add, "ICTCURR1", "**", 0, False)
        Create_TDA(dst.Tables.Add, "ICTCURR2", "*")

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        Create_TDA(dst.Tables.Add, "ICTLOTDX", "*")

        ASCMAIN1.sql = "Select * from ICTPACK1"
        Create_TDA(dst.Tables.Add, "ICTPACK1", "**", 0, False)

        ASCMAIN1.sql = "Select * from ICTPROD1"
        Create_TDA(dst.Tables.Add, "ICTPROD1", "**", 0, False)

        ASCMAIN1.sql = "Select * from ICTCOSTE"
        Create_TDA(dst.Tables.Add, "ICTCOSTE", "**", 0, False)


        ASCMAIN1.sql = "Select SUM (DECODE(ICTLOTD1.PACK_CODE,'" & TAC.TACMAIN1.CATCH_PACK _
        & "',ICTLOTD1.CATCH_WEIGHT,ICTPACK1.PACK_FACTOR) * ICTLOTD1.QTY_ON_HAND) UNITS_SLACK" _
        & " from ICTLOTD1,ICTPACK1 " _
        & " where ICTLOTD1.WHSE_CODE = :PARM1" _
        & "   and ICTLOTD1.LOT_NO = :PARM2" _
        & "   and ICTLOTD1.LOT_SEQ_NO <> :PARM3" _
        & "   and ICTLOTD1.QTY_ON_HAND <> 0" _
        & "   and ICTLOTD1.PACK_CODE = ICTPACK1.PACK_CODE"
        Create_TDA(dst.Tables.Add, "ICTLOTDU", "**", 0, False, "VVN")


        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)
        EnforceConstraints(False)
        Fill_Records("ICTLOTD1", New String() {Format(DATETIME_STAMP.Date, "dd-MMM-yyyy")})
        Fill_Records("ICTPACK1")
        Fill_Records("ICTPROD1")
        Fill_Records("ICTCOSTE")
        EnforceConstraints(True)
    End Sub

    Sub Prepare_GL_Interface(ByVal JOURNAL_TYPE As String)

        Get_PARM("GLTPARM1")

        Dim JOURNAL_NO As String = ""
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            JOURNAL_NO = ASCMAIN1.Next_Control_No("GLTJRNL1")
        Else
            JOURNAL_NO = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        End If
        Dim JOURNAL_LNO As Integer = 0

        Dim rowPOTCATG1 As DataRow = LookUp("POTCATG1", "STG")

        ' Storage Accrual Side

        For Each row As DataRow In ASCDATA1.SelectDistinct("ICTLOTDX", "WHSE_CODE", "CON_REG_IND").Rows
            Dim WHSE_CODE As String = row.Item("WHSE_CODE")
            Dim CON_REG_IND As String = row.Item("CON_REG_IND") & ""
            Dim sqlw As String = "WHSE_CODE = '" & WHSE_CODE & "' and CON_REG_IND = '" & CON_REG_IND & "'"
            Dim DETL_POSTING_AMT As Decimal = -1 * Val(dst.Tables("ICTLOTDX").Compute("SUM(ADDED_STORAGE_TOTAL)", sqlw) & "")
            If DETL_POSTING_AMT <> 0 Then
                Dim DETL_CTL_NO As String = CON_REG_IND & ":" & WHSE_CODE
                Dim ACCT_CODE As String = rowPOTCATG1.Item("ACCT_CODE_" & CON_REG_IND)
                Dim SEG2_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                Write_GLTINTF1(JOURNAL_NO, JOURNAL_LNO, ACCT_CODE, _
                    SEG2_CODE, DETL_CTL_NO, DETL_POSTING_AMT, JOURNAL_TYPE)
            End If
        Next

        ' Inventory Side

        For Each row As DataRow In ASCDATA1.SelectDistinct("ICTLOTDX", "CON_REG_IND").Rows
            Dim CON_REG_IND As String = row.Item("CON_REG_IND")
            Dim sqlw As String = "CON_REG_IND = '" & CON_REG_IND & "'"
            Dim DETL_POSTING_AMT As Decimal = Val(dst.Tables("ICTLOTDX").Compute("SUM(ADDED_STORAGE_TOTAL)", sqlw) & "")
            If DETL_POSTING_AMT <> 0 Then
                Dim DETL_CTL_NO As String = CON_REG_IND
                Dim rowICTCREG1 As DataRow = LookUp("ICTCREG1", CON_REG_IND)
                Dim ACCT_CODE As String = rowICTCREG1.Item("ACCT_INVTY")
                Dim SEG2_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                Write_GLTINTF1(JOURNAL_NO, JOURNAL_LNO, ACCT_CODE, _
                                SEG2_CODE, DETL_CTL_NO, DETL_POSTING_AMT, JOURNAL_TYPE)
            End If
        Next
    End Sub

    Sub Write_GLTINTF1( _
    ByVal JOURNAL_NO As String, _
    ByRef JOURNAL_LNO As Integer, _
    ByVal ACCT_CODE As String, _
    ByVal SEG2_CODE As String, _
    ByVal DETL_CTL_NO As String, _
    ByVal DETL_POSTING_AMT As Decimal, _
    ByVal JOURNAL_TYPE As String)

        Dim rowGLTINTF1 As DataRow = ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1("OPS_YYYYPP") = ASCMAIN1.CYP
        rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
        JOURNAL_LNO += 1
        rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
        rowGLTINTF1("ACCT_CODE") = ACCT_CODE
        rowGLTINTF1("SEG2_CODE") = SEG2_CODE
        rowGLTINTF1("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowGLTINTF1("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowGLTINTF1("DETL_CTL_DATE") = DATETIME_STAMP.Date
        rowGLTINTF1("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
        rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
        rowGLTINTF1("DETL_CTL_NO") = DETL_CTL_NO
        rowGLTINTF1("DETL_CTL_LNO") = DBNull.Value
        rowGLTINTF1("DETL_CVX_NO") = DBNull.Value
        rowGLTINTF1("DETL_CVX_REF_DATE") = DBNull.Value
        rowGLTINTF1("DETL_CVX_REF_NO") = DBNull.Value
        rowGLTINTF1("DETL_CVX_REF_LNO") = DBNull.Value
        rowGLTINTF1("DETL_DESC") = DBNull.Value
        rowGLTINTF1("DETL_CTL_TYPE") = DBNull.Value
        rowGLTINTF1("DETL_CVX_TYPE") = DBNull.Value
        rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
        rowGLTINTF1("DIST_CODE") = DBNull.Value
        ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
    End Sub
End Class