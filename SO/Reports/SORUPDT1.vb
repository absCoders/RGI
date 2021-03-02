Public Class SORUPDT1
    Dim SOTINVH1 As String
    Dim SOTINVH2 As String
    Dim SOTORDRS As String
    Dim CURR_CODE As String = ""
    Dim CURR_EXCH_RATE As Decimal = 0
    Dim GST_TAX As Decimal = 0
    Dim NYP As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")
        Get_PARM("WHTPARM1")
        Get_PARM("GLTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""
        sqlw = "  and SOTINVH1.ORDR_YYYYPP_UPDATED is Null"

        Dim EMsg As String = TAC.SOCMAIN1.Prepare_Sales_Invoices(Me, sqlw, SOTINVH1, SOTINVH2)
        If EMsg <> "" Then
            RWU &= "0"
            xErrMsg = EMsg
        Else
            ASCMAIN1.sql = "Select ORDR_NO, SUM (ORDR_QTY_OPEN) OPEN, SUM (ORDR_QTY_PICK) PICK" & vbCrLf _
                & " from SOTORDR2 where ROWNUM < 1" & vbCrLf _
                & " group by ORDR_NO"
            SOTORDRS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRS & " Add Primary Key (ORDR_NO)")
        End If

        ASCMAIN1.sql = "Select * from TATTERM1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "TATTERM1", 1))


        ASCMAIN1.sql = "Select * from ICTWHSE1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTWHSE1", 1))


        Check_if_Empty("SOTINVH1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If ASCDATA1.GetDataValue("Select PRD_CLOSE_IND from ASTPCTL1") & "" = "1" Then
            '    EMsg = EMsg & vbCr & "Period-End Already Initialized"

            'End If

        End If

    End Sub

    Overrides Sub Update_Record()

        ASCMAIN1.Progress("Adjusting Overshipped Pick Tickets", "")

        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_QTY_CONF > PICK_QTY")
            With rowSOTPICK2
                Dim SHIP_OVER As Int64 = Val(.Item("PICK_QTY_CONF") & "") - Val(.Item("PICK_QTY") & "")
                Dim rowSOTORDR2 = LookUp("SOTORDR2", New String() {.Item("ORDR_NO"), .Item("ORDR_LNO")})
                Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                Dim SHIP_OVER_ADJ As Int64 = 0
                If SHIP_OVER <= ORDR_QTY_OPEN Then
                    SHIP_OVER_ADJ = SHIP_OVER
                Else
                    SHIP_OVER_ADJ = ORDR_QTY_OPEN
                End If
                If SHIP_OVER_ADJ <> 0 Then
                    .Item("PICK_QTY_BACK") = -1 * SHIP_OVER_ADJ
                End If
            End With
            Update_Record_TDA("SOTPICK2")
            '   Dim R() As DataRow = dst.Tables("").Select("", "", DataViewRowState.ModifiedCurrent)
        Next

        ASCMAIN1.sql = "Select Distinct SHIP_BOL_NO, CUST_CODE from " & SOTINVH1
        Dim SOTSHIPX As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")

        Dim WHSE_CODEs_locator As New List(Of String)
        ASCMAIN1.sql = "SELECT WHSE_CODE FROM ICTWHSE1 WHERE WHSE_LOCATOR = '1'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select()
            WHSE_CODEs_locator.Add(row.Item("WHSE_CODE"))
        Next

        ASCMAIN1.Progress("Updating Invoice", "")
        Dim r As Int32 = 0
        NYP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)

        Dim RMAX As Integer = dst.Tables("SOTINVH1").Select("").Length

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("", "INV_NO")
            r += 1

            With rowSOTINVH1

                Dim INV_NO As String = .Item("INV_NO") & ""
                Dim PICK_NO As String = .Item("PICK_NO") & ""
                Dim WHSE_CODE As String = .Item("WHSE_CODE") & ""

                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(RMAX)) ' & ":" & INV_NO)

                If WHSE_CODEs_locator.Contains(WHSE_CODE) Then
                    'ASCMAIN1.sql = "Begin WHPLOCB2('S','" & INV_NO & "', '" & ASCMAIN1.SESSION_NO & "'); END;"
                    'ASCDATA1.ExecuteSQL() ' this is taking almost 1 sec per invoice'
                End If

                Dim INV_TOTAL_AMOUNT As Decimal = Val(.Item("INV_TOTAL_AMOUNT") & "")
                Dim INV_SALES As Decimal = Val(.Item("INV_SALES") & "")
                Dim INV_BALANCE As Decimal = INV_TOTAL_AMOUNT
                Dim INV_TYPE As String = .Item("INV_TYPE") & ""
                Dim ORDR_BILL_TO_CUST As String = .Item("ORDR_BILL_TO_CUST") & ""
                Dim CUST_CODE As String = .Item("CUST_CODE") & ""
                If ORDR_BILL_TO_CUST = "" Then
                    ORDR_BILL_TO_CUST = CUST_CODE
                End If
                Dim INV_DATE As Date = .Item("INV_DATE")


                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

                CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""

                If CURR_CODE = "" Then
                    CURR_CODE = "USD"
                    CURR_EXCH_RATE = 1
                    GST_TAX = 0
                Else
                    CURR_CODE = rowARTCUST1.Item("CURR_CODE")

                    'Stop ' THIS HAS BEEN MOVED TO GL TABLES NOW
                    If rowARTCUST1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                        CURR_EXCH_RATE = 1
                        GST_TAX = 0
                    Else
                        'Stop ' NO MORE ICTCURR1
                        Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", rowARTCUST1.Item("CURR_CODE"))
                        CURR_EXCH_RATE = rowTATCURR1.Item("CURR_EXCH_CUR")
                        GST_TAX = 0.07
                    End If
                End If

                If r Mod 100 = 0 Then
                    ASCMAIN1.Progress("-", INV_NO & " - " & CStr(r))
                End If

                Dim TERM_CODE As String = .Item("TERM_CODE") & ""
                Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, TERM_CODE, Nothing, INV_DATE)

                If INV_TYPE = "I" Then
                    'sql = "Begin SOPORDR2_P('" & PICK_NO & "', '" & INV_NO & "'); END;"
                    'ASCDATA1.ExecuteSQL(sql)
                    ASCDATA1.ExecuteSP("SOPORDR2_P", "VV", New Object() {PICK_NO, INV_NO}, New String() {"PICK_NO_X", "INV_NO_X"})
                End If

                ' Update Open A/R Items File

                If INV_BALANCE <> 0 And .Item("INV_NO_CONS") & "" = "" Then
                    Dim YP As String = Get_ARTOPEN1_OPS_YYYYPP(INV_DATE)
                    Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
                    rowARTOPEN1.Item("CUST_CODE") = ORDR_BILL_TO_CUST
                    rowARTOPEN1.Item("INV_TYPE") = INV_TYPE
                    'rdw 3/8/2019
                    'rowARTOPEN1.Item("INV_NO") = INV_NO
                    rowARTOPEN1.Item("INV_NUM") = INV_NO
                    rowARTOPEN1.Item("INV_DATE") = INV_DATE
                    If CURR_CODE = "" Or CURR_EXCH_RATE = 0 Then
                        Stop
                    ElseIf CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Or CURR_EXCH_RATE <> 1 Then
                        Stop
                        'NEED TO ADDRESS CALCULATIONS HERE
                    End If

                    For Each COLUMN_NAME As String In New String() _
                        {"CUST_STORE_NO", "POST_CODE", "TERM_CODE", _
                         "INV_SALES", "INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT", _
                         "REASON_CODE", "SALES_DIVISION_CODE", "SREP_CODE", _
                         "INV_SALES_CURR", "INV_FREIGHT_CURR", "INV_MISC_CHG_CURR", "INV_TOTAL_AMOUNT_CURR", _
                         "GST_TAX", "GST_TAX_CURR"}
                        rowARTOPEN1.Item(COLUMN_NAME) = .Item(COLUMN_NAME)
                    Next
                    rowARTOPEN1.Item("INV_DUE_DATE") = INV_DUE_DATE
                    rowARTOPEN1.Item("INV_CUST_PO") = .Item("ORDR_CUST_PO") & ""
                    'rdw 3/8/19
                    'rowARTOPEN1.Item("INV_ORDR_NO") = .Item("ORDR_NO") & ""
                    rowARTOPEN1.Item("ORDR_NO") = .Item("ORDR_NO") & ""
                    rowARTOPEN1.Item("INV_DISC") = 0
                    rowARTOPEN1.Item("INV_STAX") = 0
                    rowARTOPEN1.Item("INV_BALANCE") = INV_BALANCE
                    rowARTOPEN1.Item("CUST_CODE_SO") = CUST_CODE
                    rowARTOPEN1.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowARTOPEN1.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    rowARTOPEN1.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    rowARTOPEN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowARTOPEN1.Item("INIT_DATE") = DATETIME_STAMP
                    rowARTOPEN1.Item("CURR_CODE") = CURR_CODE
                    rowARTOPEN1.Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                    rowARTOPEN1.Item("INV_DISC_CURR") = 0
                    rowARTOPEN1.Item("INV_STAX_CURR") = 0
                    rowARTOPEN1.Item("OPS_YYYYPP") = Get_ARTOPEN1_OPS_YYYYPP(INV_DATE)
                    rowARTOPEN1.Item("INV_SALES_CURR") = IIf(CURR_CODE = "USD", Val(.Item("INV_SALES") & ""), Val(.Item("INV_SALES_CURR") & ""))
                    rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = IIf(CURR_CODE = "USD", Val(.Item("INV_TOTAL_AMOUNT") & ""), Val(.Item("INV_TOTAL_AMT_CURR") & ""))

                    rowARTOPEN1.Item("INV_BALANCE_CURR") = INV_BALANCE
                    dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
                End If
            End With
        Next
        Update_Record_TDA("ARTOPEN1")

        ' Update Locator

        ASCMAIN1.sql = "" _
            & "Insert into WHTLOCB2 Select" & vbCrLf _
            & "SOTINVH1.WHSE_CODE," & vbCrLf _
            & "ICTWHSE1.WHSE_LOC_SHP LOCATION_CODE," & vbCrLf _
            & "ICTWHSE1.WHSE_DEF_BAR_CODE BAR_CODE," & vbCrLf _
            & "SOTINVH2.STYLE_CODE,  SOTINVH2.COLOR_CODE," & vbCrLf _
            & "-1 * SOTINVH2.ORDR_QTY_SHIP, 'S' WHSE_TRAN_TYPE, SOTINVH2.INV_NO, SOTINVH2.INV_LNO," & vbCrLf _
            & "SOTINVH1.INIT_DATE, SOTINVH1.INIT_OPER," & vbCrLf _
            & "NULL LOCATION_CODE_OTHER," & vbCrLf _
            & "'" & ASCMAIN1.SESSION_NO & "' SESSION_NO," & vbCrLf _
            & "ICTWHSE1.WHSE_DEF_LOAD_NO LOAD_NO, NULL LOAD_NO_OTHER, NULL BAR_CODE_OTHER" & vbCrLf _
            & "    from SOTINVH1,SOTINVH2,ICTWHSE1,ICTSTYC1," & SOTINVH1 & " SOTINVHX" & vbCrLf _
            & "     where SOTINVH1.INV_NO = SOTINVHX.INV_NO" & vbCrLf _
            & "       and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & "       and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "       and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "       and ICTWHSE1.WHSE_CODE = SOTINVH1.WHSE_CODE" _
            & "       and ICTWHSE1.WHSE_LOCATOR = '1'" & vbCrLf _
            & "       and ICTSTYC1.STYLE_CODE = SOTINVH2.STYLE_CODE" & vbCrLf _
            & "       and ICTSTYC1.COLOR_CODE = SOTINVH2.COLOR_CODE"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare" & vbCrLf _
            & "  Cursor C1 is Select" & vbCrLf _
            & "SOTINVH1.WHSE_CODE," & vbCrLf _
            & "ICTWHSE1.WHSE_LOC_SHP LOCATION_CODE," & vbCrLf _
            & "ICTWHSE1.WHSE_DEF_BAR_CODE BAR_CODE," & vbCrLf _
            & "SOTINVH2.STYLE_CODE,  SOTINVH2.COLOR_CODE," & vbCrLf _
            & "Sum (-1 * SOTINVH2.ORDR_QTY_SHIP) WHSE_TRAN_QTY," & vbCrLf _
            & "Max (SOTINVH1.INIT_DATE) INIT_DATE, Max (SOTINVH1.INIT_OPER) INIT_OPER" & vbCrLf _
            & "    from SOTINVH1,SOTINVH2,ICTWHSE1,ICTSTYC1," & SOTINVH1 & " SOTINVHX" & vbCrLf _
            & "     where SOTINVH1.INV_NO = SOTINVHX.INV_NO" & vbCrLf _
            & "       and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & "       and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "       and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "       and ICTWHSE1.WHSE_CODE = SOTINVH1.WHSE_CODE" _
            & "       and ICTWHSE1.WHSE_LOCATOR = '1'" & vbCrLf _
            & "       and ICTSTYC1.STYLE_CODE = SOTINVH2.STYLE_CODE" & vbCrLf _
            & "       and ICTSTYC1.COLOR_CODE = SOTINVH2.COLOR_CODE" & vbCrLf _
            & " group by SOTINVH1.WHSE_CODE, ICTWHSE1.WHSE_LOC_SHP, ICTWHSE1.WHSE_DEF_BAR_CODE," & vbCrLf _
            & "          SOTINVH2.STYLE_CODE,  SOTINVH2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update WHTLOCB1" & vbCrLf _
            & "    Set LOCATION_QTY = NVL(LOCATION_QTY,0) + NVL(R1.WHSE_TRAN_QTY,0)," & vbCrLf _
            & "     LAST_DATE = R1.INIT_DATE, LAST_OPER = R1.INIT_OPER" & vbCrLf _
            & "    where WHSE_CODE = R1.WHSE_CODE and LOCATION_CODE = R1.LOCATION_CODE" & vbCrLf _
            & "      and BAR_CODE = R1.BAR_CODE and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "   If SQL%NOTFOUND Then" & vbCrLf _
            & "    Insert into WHTLOCB1" & vbCrLf _
            & "     (WHSE_CODE,LOCATION_CODE,BAR_CODE,STYLE_CODE,COLOR_CODE,LOCATION_QTY," & vbCrLf _
            & "      INIT_DATE,INIT_OPER,LAST_DATE,LAST_OPER)" & vbCrLf _
            & "     Values" & vbCrLf _
            & "     (R1.WHSE_CODE,R1.LOCATION_CODE,R1.BAR_CODE,R1.STYLE_CODE,R1.COLOR_CODE,NVL(R1.WHSE_TRAN_QTY,0)," & vbCrLf _
            & "      R1.INIT_DATE,R1.INIT_OPER,R1.INIT_DATE,R1.INIT_OPER);" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop;  " & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()



        ' Update SOTSHIP1

        ASCMAIN1.Progress("Shipments", "")
        ASCMAIN1.sql = "Update SOTSHIP1 Set REGISTER_XNO = '" & Mid(XNO, 5) & "'" & vbCrLf _
            & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTSHIP1 Set CUST_FACTOR_TRANS_IND = '0' where SHIP_BOL_NO IN (" & vbCrLf _
            & " Select SHIP_BOL_NO from " & SOTSHIPX & ", ARTCUST1" & vbCrLf _
            & " Where " & SOTSHIPX & ".CUST_CODE = ARTCUST1.CUST_CODE" & vbCrLf _
            & " AND ARTCUST1.CUST_FACTOR_TRANS_IND = '1')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_810_BATCH_NO = 'N'" & vbCrLf _
            & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_810_BATCH_NO = NULL" & vbCrLf _
            & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & vbCrLf _
            & " where CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '810' AND EDI_STATUS = 'P') or CUST_CODE = 'WALMARTCOM')"
        'Replaced with the line above for Rick 08/15/07 - WR
        '  where CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '810') or CUST_CODE = 'NORDS1')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_856_BATCH_NO = 'N'" & vbCrLf _
            & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_856_BATCH_NO = NULL" & vbCrLf _
            & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & vbCrLf _
            & " where CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '856' AND EDI_STATUS = 'P')  or CUST_CODE = 'WALMARTCOM')"
        'Replaced with the line above for Rick 08/15/07 - WR
        '  where CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '856')  or CUST_CODE = 'NORDS1')"
        ASCDATA1.ExecuteSQL()


        ' Update Order Status

        ASCMAIN1.Progress("Order Status", "")
        ASCMAIN1.sql = "Insert into " & SOTORDRS & vbCrLf _
            & " Select ORDR_NO, SUM (ORDR_QTY_OPEN) OPEN, SUM (ORDR_QTY_PICK) PICK" & vbCrLf _
            & " from SOTORDR2 where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVH1 & ")" & vbCrLf _
            & " group by ORDR_NO"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT ORDR_NO FROM " & SOTORDRS & " WHERE OPEN <> 0;" & vbCrLf _
            & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & " UPDATE SOTORDR1 SET ORDR_STATUS = 'O' WHERE ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            & " UPDATE SOTORDR2 SET ORDR_STATUS = 'O' WHERE ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT ORDR_NO FROM " & SOTORDRS & " WHERE OPEN = 0 AND PICK <> 0;" & vbCrLf _
            & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & " UPDATE SOTORDR1 SET ORDR_STATUS = 'P' WHERE ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            & " UPDATE SOTORDR2 SET ORDR_STATUS = 'P' WHERE ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT ORDR_NO FROM " & SOTORDRS & " WHERE OPEN = 0 AND PICK = 0;" & vbCrLf _
            & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & " UPDATE SOTORDR1 SET ORDR_STATUS = 'F', ORDR_YYYYPP_CLOSED = '" & ASCMAIN1.CYP & "', ORDR_DATE_CLOSED = SYSDATE WHERE ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            & " UPDATE SOTORDR2 SET ORDR_STATUS = 'F' WHERE ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        ' Update Style Status

        ASCMAIN1.Progress("Style Status Update", "")

        Dim sql_Driver As String = "" _
            & "Select SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.WHSE_CODE" & vbCrLf _
            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", SUM (SOTPICK2.PICK_QTY) as PICK_QTY" & vbCrLf _
            & ", SUM (SOTPICK2.PICK_QTY_CONF) as PICK_QTY_CONF " & vbCrLf _
            & ", SUM (SOTPICK2.PICK_QTY_BACK) as PICK_QTY_BACK" & vbCrLf _
            & " from SOTPICK1,SOTPICK2," & SOTINVH1 & " SOTINVH1,SOTORDR2" & vbCrLf _
            & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTPICK1.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.WHSE_CODE" & vbCrLf _
            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"

        ASCMAIN1.sql = "" _
           & "Begin" & vbCrLf _
           & " Declare Cursor C1 is " & sql_Driver & ";" & vbCrLf _
           & " Begin" & vbCrLf _
           & "  For R1 in C1 Loop" & vbCrLf _
           & "   Begin" & vbCrLf _
           & "    Update ICTSTAT1 Set " & vbCrLf _
           & "     WHSE_QTY_SHP = NVL(WHSE_QTY_SHP,0) + NVL(R1.PICK_QTY_CONF,0)" & vbCrLf _
           & "     where STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
           & "       and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
           & "       and OPS_YYYYPP = R1.OPS_YYYYPP" & vbCrLf _
           & "       and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
           & "    If SQL%NOTFOUND Then" & vbCrLf _
           & "     Insert into ICTSTAT1" & vbCrLf _
           & "       (WHSE_CODE,STYLE_CODE,COLOR_CODE,OPS_YYYYPP,WHSE_QTY_SHP)" & vbCrLf _
           & "       Values (R1.WHSE_CODE,R1.STYLE_CODE,R1.COLOR_CODE,R1.OPS_YYYYPP,NVL(R1.PICK_QTY_CONF,0));" & vbCrLf _
           & "    End If;" & vbCrLf _
           & "   End;" & vbCrLf _
           & "  End Loop;" & vbCrLf _
           & " End;" & vbCrLf _
           & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
           & "Begin" & vbCrLf _
           & " Declare Cursor C1 is " & sql_Driver & ";" & vbCrLf _
           & " Begin" & vbCrLf _
           & "  For R1 in C1 Loop" & vbCrLf _
           & "   Begin" & vbCrLf _
           & "    Update ICTSTAT2 Set " & vbCrLf _
           & "     WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) + NVL(R1.PICK_QTY_BACK,0)," & vbCrLf _
           & "     WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) - NVL(R1.PICK_QTY,0)," & vbCrLf _
           & "     WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) - NVL(R1.PICK_QTY_CONF,0)" & vbCrLf _
           & "     where STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
           & "       and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
           & "       and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
           & "    If SQL%NOTFOUND Then" & vbCrLf _
           & "     Insert into ICTSTAT2" & vbCrLf _
           & "       (WHSE_CODE,STYLE_CODE,COLOR_CODE,WHSE_QTY_OPEN,WHSE_QTY_PICK,WHSE_QTY_ON_HAND)" & vbCrLf _
           & "       Values (R1.WHSE_CODE,R1.STYLE_CODE,R1.COLOR_CODE,NVL(R1.PICK_QTY_BACK,0),-1*NVL(R1.PICK_QTY,0),-1*NVL(R1.PICK_QTY_CONF,0));" & vbCrLf _
           & "    End If;" & vbCrLf _
           & "   End;" & vbCrLf _
           & "  End Loop;" & vbCrLf _
           & " End;" & vbCrLf _
           & "End;"
        ASCDATA1.ExecuteSQL()



        Dim sql_Driver2 As String = "Select SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.WHSE_CODE" & vbCrLf _
            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & ", DECODE(NVL(WHSE_PHYS_STATUS,'0'),'1','" _
            & ROWs("WHTPARM1").Item("WH_PARM_LOC_SPH") & "','" _
            & ROWs("WHTPARM1").Item("WH_PARM_LOC_SHP") & "') LOCATION_CODE" & vbCrLf _
            & ", SUM (SOTPICK2.PICK_QTY_CONF) as PICK_QTY_CONF " & vbCrLf _
            & " from ICTWHSE1,SOTPICK1,SOTPICK2," & SOTINVH1 & " SOTINVH1,SOTORDR2" & vbCrLf _
            & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTPICK1.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and ICTWHSE1.WHSE_CODE = SOTINVH1.WHSE_CODE" & vbCrLf _
            & "   and ICTWHSE1.WHSE_LOCATOR = '1'" & vbCrLf _
            & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.WHSE_CODE" & vbCrLf _
            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & ", DECODE(NVL(WHSE_PHYS_STATUS,'0'),'1','" _
            & ROWs("WHTPARM1").Item("WH_PARM_LOC_SPH") & "','" _
            & ROWs("WHTPARM1").Item("WH_PARM_LOC_SHP") & "')"

        'updates to the WHTLOCBx tables were remarked in V1, so I remarked them in V2 as well
        'ASCMAIN1.sql = "" _
        '    & "Begin" & vbCrLf _
        '    & " Declare Cursor C1 is " & sql_Driver2 & ";" & vbCrLf _
        '    & "  BAR_CODE_SHP VARCHAR2(10);" & vbCrLf _
        '    & "  WHSE_TRAN_LNO NUMBER(6,0);" & vbCrLf _
        '    & " Begin" & vbCrLf _
        '    & "  BAR_CODE_SHP := '0000000000';" & vbCrLf _
        '    & "  For R1 in C1 Loop" & vbCrLf _
        '    & "   Begin" & vbCrLf _
        '    & "    WHSE_TRAN_LNO := NVL(WHSE_TRAN_LNO,0) + 1;" & vbCrLf _
        '    & "    Update WHTLOCB1 Set LOCATION_QTY = NVL(LOCATION_QTY,0) - NVL(R1.PICK_QTY_CONF,0)" & vbCrLf _
        '    & "     where WHSE_CODE = R1.WHSE_CODE" & vbCrLf _
        '    & "       and LOCATION_CODE = R1.LOCATION_CODE" & vbCrLf _
        '    & "       and BAR_CODE = BAR_CODE_SHP" & vbCrLf _
        '    & "       and STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
        '    & "       and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
        '    & "    If SQL%NOTFOUND Then" & vbCrLf _
        '    & "     Insert into WHTLOCB1" & vbCrLf _
        '    & "       (WHSE_CODE,LOCATION_CODE,BAR_CODE,STYLE_CODE,COLOR_CODE)" & vbCrLf _
        '    & "       Values (R1.WHSE_CODE,R1.LOCATION_CODE,BAR_CODE_SHP,R1.STYLE_CODE,R1.COLOR_CODE);" & vbCrLf _
        '    & "    End If;" & vbCrLf _
        '    & "   Insert into WHTLOCB2 (WHSE_CODE,LOCATION_CODE,BAR_CODE," & vbCrLf _
        '    & "                         STYLE_CODE,COLOR_CODE,WHSE_TRAN_QTY," & vbCrLf _
        '    & "                         WHSE_TRAN_TYPE,WHSE_TRAN_NO,WHSE_TRAN_LNO," & vbCrLf _
        '    & "                         INIT_DATE,INIT_OPER)" & vbCrLf _
        '    & "    values (R1.WHSE_CODE,R1.LOCATION_CODE,BAR_CODE_SHP," & vbCrLf _
        '    & "            R1.STYLE_CODE,R1.COLOR_CODE,-1*R1.PICK_QTY_CONF," & vbCrLf _
        '    & "            'S',R1.SHIP_BOL_NO,WHSE_TRAN_LNO," & vbCrLf _
        '    & "            SYSDATE,'" & ASCMAIN1.USER_ID & "');" & vbCrLf _
        '    & "   End;" & vbCrLf _
        '    & "  End Loop;" & vbCrLf _
        '    & " End;" & vbCrLf _
        '    & "End;"
        'ASCDATA1.ExecuteSQL()


        ' Record Consolidated Invoice

        ASCMAIN1.Progress("Consolidated Invoice", "")
        ASCMAIN1.sql = "Select INV_NO_CONS, ORDR_BILL_TO_CUST as CUST_CODE" & vbCrLf _
            & ", MAX (INV_DATE) AS INV_DATE, MAX (REASON_CODE) as REASON_CODE" & vbCrLf _
            & ", MAX (SALES_DIVISION_CODE) as SALES_DIVISION_CODE, MAX (ORDR_CUST_PO) AS ORDR_CUST_PO" & vbCrLf _
            & ", SUM (INV_SALES) AS INV_SALES, SUM (INV_FREIGHT) AS INV_FREIGHT, SUM (INV_MISC_CHG) AS INV_MISC_CHG" & vbCrLf _
            & ", SUM (INV_TOTAL_AMOUNT) AS INV_TOTAL_AMOUNT, MAX(EDI_APPOINTMENT) AS EDI_APPOINTMENT" & vbCrLf _
            & ", SUM (INV_SALES_CURR) AS INV_SALES_CURR, SUM (INV_FREIGHT_CURR) AS INV_FREIGHT_CURR" & vbCrLf _
            & ", SUM (INV_MISC_CHG_CURR) AS INV_MISC_CHG_CURR, SUM (INV_TOTAL_AMOUNT_CURR) AS INV_TOTAL_AMOUNT_CURR" & vbCrLf _
            & " from " & SOTINVH1 & vbCrLf _
            & " where INV_NO_CONS is Not Null" & vbCrLf _
            & " group by INV_NO_CONS, ORDR_BILL_TO_CUST"
        For Each rowARM As DataRow In ASCDATA1.GetDataTable.Select("")
            Consolidated_Invoice(rowARM)
        Next

        ' Record Sales Summary

        ASCMAIN1.Progress("Customer Sales Summary", "")


        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
        Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        If ASCMAIN1.Running_in_VS Then Stop ' SEPARATE CREDITS FROM INVOICES
        If ASCMAIN1.Running_in_VS Then Stop ' DOES THIS ROUTINE WORK FOR REVERSED INVOICES PROPERLY?
        ASCMAIN1.sql = "Select CUST_CODE, Count (*) as INVOICES" & vbCrLf _
            & ", Sum (INV_SALES) as INV_SALES" & vbCrLf _
            & ", Max (INV_DATE) as LAST_DATE" & vbCrLf _
            & ", Min (INV_DATE) as FIRST_DATE" & vbCrLf _
            & " from " & SOTINVH1 & vbCrLf _
            & " where INV_DATE <= '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'" & vbCrLf _
            & " group by CUST_CODE"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim LAST_DATE As Date = row.Item("LAST_DATE")
            ASCMAIN1.sql = "Select * from " & SOTINVH1 & " where CUST_CODE = :PARM1 and INV_DATE = :PARM2"
            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VD", New Object() {CUST_CODE, LAST_DATE})

            Dim rowARTCUST6 As DataRow = dst.Tables("ARTCUST6").Rows.Find(CUST_CODE)
            If rowARTCUST6 Is Nothing Then
                rowARTCUST6 = Fill_Record("ARTCUST6", CUST_CODE, False, False)
                If rowARTCUST6 Is Nothing Then
                    rowARTCUST6 = dst.Tables("ARTCUST6").NewRow
                    rowARTCUST6.Item("CUST_CODE") = CUST_CODE
                    dst.Tables("ARTCUST6").Rows.Add(rowARTCUST6)
                End If
            End If

            With rowARTCUST6
                .Item("CUST_LAST_INV_NUM") = rowSOTINVH1.Item("INV_NO")
                .Item("CUST_LAST_INV_DATE") = rowSOTINVH1.Item("INV_DATE")
                .Item("CUST_LAST_INV_AMT") = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & "")

                If .Item("CUST_FIRST_PURCH") & "" = "" Then
                    .Item("CUST_FIRST_PURCH") = row.Item("FIRST_DATE")
                End If
                .Item("CUST_SALES_MTD") = Val(.Item("CUST_SALES_MTD") & "") + Val(row.Item("INV_SALES") & "")
                .Item("CUST_SALES_YTD") = Val(.Item("CUST_SALES_YTD") & "") + Val(row.Item("INV_SALES") & "")
                .Item("CUST_NUM_INV_MTD") = Val(.Item("CUST_NUM_INV_MTD") & "") + Val(row.Item("INVOICES") & "")
                .Item("CUST_NUM_INV_YTD") = Val(.Item("CUST_NUM_INV_YTD") & "") + Val(row.Item("INVOICES") & "")
            End With
        Next

        ' A/R Summary

        ASCMAIN1.Progress("Customer A/R Summary", "")

        ASCMAIN1.sql = "Select ORDR_BILL_TO_CUST CUST_CODE, MAX(INV_DATE) INV_DATE" _
            & " from " & SOTINVH1 & " group by ORDR_BILL_TO_CUST"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim INV_DATE As Date = row.Item("INV_DATE")

            Dim rowARTCUST6 As DataRow = dst.Tables("ARTCUST6").Rows.Find(CUST_CODE)
            If rowARTCUST6 Is Nothing Then
                rowARTCUST6 = Fill_Record("ARTCUST6", CUST_CODE, False, False)
                If rowARTCUST6 Is Nothing Then
                    rowARTCUST6 = dst.Tables("ARTCUST6").NewRow
                    rowARTCUST6.Item("CUST_CODE") = CUST_CODE
                    dst.Tables("ARTCUST6").Rows.Add(rowARTCUST6)
                End If
            End If

            ASCMAIN1.sql = "Select Sum (INV_BALANCE) from ARTOPEN1 where CUST_CODE = :PARM1"
            Dim INV_BALANCE As Decimal = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {CUST_CODE}))
            With rowARTCUST6
                If INV_BALANCE > Val(.Item("CUST_HIGH_BAL_AMT") & "") Then
                    .Item("CUST_HIGH_BAL_AMT") = INV_BALANCE
                    .Item("CUST_HIGH_BAL_DATE") = INV_DATE
                End If
            End With
        Next

        ' Stamp Ops Period

        ASCMAIN1.Progress("Mark Records with Period", "")

        ASCMAIN1.sql = "Update SOTINVH1 set ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & ", ORDR_DATE_UPDATED = SYSDATE, REGISTER_XNO = '" & Mid(XNO, 5) & "'" & vbCrLf _
            & " where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1 & ")" & vbCrLf _
            & "   and INV_DATE <= '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update SOTINVH2 set ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " where INV_TYPE = 'I'" & vbCrLf _
            & "   and INV_NO in (SELECT INV_NO FROM SOTINVH1 where INV_TYPE = 'I'" & vbCrLf _
            & "   and INV_NO in (Select INV_NO from " & SOTINVH1 & ") " _
            & "   and INV_DATE <= '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "')"
        ASCDATA1.ExecuteSQL()


        'THESE SQL STATEMENT WILL UPDATE INVOICES FOR PERIODS NOT IN THE CURRENT CYP

        ASCMAIN1.sql = "Update SOTINVH1 set ORDR_YYYYPP_UPDATED = '" & NYP & "'" & vbCrLf _
            & ", ORDR_DATE_UPDATED = SYSDATE, REGISTER_XNO = '" & Mid(XNO, 5) & "'" & vbCrLf _
            & " where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1 & ") AND " & vbCrLf _
            & "INV_DATE > '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTINVH2 set ORDR_YYYYPP_UPDATED = '" & NYP & "'" & vbCrLf _
            & " where INV_TYPE = 'I'" & vbCrLf _
            & "   and INV_NO in (SELECT INV_NO FROM SOTINVH1 where INV_TYPE = 'I'" & vbCrLf _
            & "   and INV_NO in (Select INV_NO from " & SOTINVH1 & ") " & vbCrLf _
            & "   and INV_DATE > '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "')"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare XFNO VARCHAR2(6);" & vbCrLf _
            & "  Cursor C1 is" & vbCrLf _
            & "   Select * from SOTINVH1" & vbCrLf _
            & "    where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1 & ")" & vbCrLf _
            & "      and ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "      and ORDR_TYPE_CODE = 'XFR';" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   XFNO := SOPSHIP1_XFR(R1.INV_NO);" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()
        ' A VERSION OF THIS UPDATE IS ALSO FOUND IN PERIOD END FOR NYP INVOICES


        'temp table used for validation
        'ASCMAIN1.sql = "Create table SOTINVH1_SJU AS SELECT REGISTER_XNO, INV_NO, ORDR_NO, INV_NO_CONS, ORDR_BILL_TO_CUST, CUST_CODE " & vbCrLf _
        '    & " FROM SOTINVH1 WHERE REGISTER_XNO IN ('" & Mid(XNO, 5) & "','" & XNO & "')"
        'ASCDATA1.ExecuteSQL()

        ' Order Group Summary

        ASCMAIN1.Progress("Updating Order Group Summary", "")
        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is Select Distinct ORDR_GROUP_NO from " & SOTINVH1 & ";" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   SOPORDR0_G(R1.ORDR_GROUP_NO);" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ' Create a Cash Receipts Batch Matching off All ReverSals with Originals

        Dim PYMT_BATCH_NO As String = ""
        Dim PYMT_BATCH_ILNO_ctr As Integer = 0

        Dim SHIP_BOL_NO As String = ""
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select _
                ("INV_NO_REV is Not Null", "SHIP_BOL_NO,INV_NO")
            With rowSOTINVH1
                Dim CUST_CODE As String = .Item("ORDR_BILL_TO_CUST")
                Dim INV_DATE As Date = .Item("INV_DATE")
                If .Item("SHIP_BOL_NO") <> SHIP_BOL_NO Then
                    SHIP_BOL_NO = .Item("SHIP_BOL_NO")
                    RYP = .Item("ORDR_YYYYPP_UPDATED")
                    PYMT_BATCH_NO = Update_ARTPYMT1(CUST_CODE, INV_DATE)
                    PYMT_BATCH_ILNO_ctr = 0
                End If

                Update_ARTPYMT3(rowSOTINVH1, rowSOTINVH1.Item("INV_NO"), PYMT_BATCH_ILNO_ctr, PYMT_BATCH_NO, 1)
                Update_ARTPYMT3(rowSOTINVH1, rowSOTINVH1.Item("INV_NO_REV"), PYMT_BATCH_ILNO_ctr, PYMT_BATCH_NO, -1)
            End With
        Next

        For Each TABLE_NAME As String In New String() {"ARTPYMT1", "ARTPYMT2", "ARTPYMT3", "ARTOPEN1", "ARTCUST6"}
            Update_Record_TDA(TABLE_NAME)
        Next

        'Update New Control File To Force SJ&U between Confirm and Deconfirm - WR - 20051024
        ASCMAIN1.sql = "Update SOTCTLU1" & vbCrLf _
            & " Set CTL_UPDATE_REQ = 'U'" & vbCrLf _
            & " where UPPER(CTL_KEY) = 'Z'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Consolidated_Invoice(rowARM As DataRow)

        Dim CURR_CODE As String = "USD"
        Dim CURR_EXCH_RATE As Decimal = 1

        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", rowARM.Item("CUST_CODE"))
        Dim TERM_CODE As String = rowARTCUST1.Item("TERM_CODE") & ""

        'INV_DUE_DATE = Format$(DateValue(INV_DATE) + TERM_DAYS_DUE, "MM/DD/YYYY")
        Dim INV_DATE As Date = rowARM.Item("INV_DATE")
        Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, TERM_CODE, Nothing, INV_DATE)

        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        With rowARTOPEN1
            .Item("CUST_CODE") = rowARM.Item("CUST_CODE")
            .Item("INV_TYPE") = "I"
            .Item("INV_NUM") = rowARM.Item("INV_NO_CONS")
            .Item("INV_DATE") = rowARM.Item("INV_DATE")
            .Item("CUST_STORE_NO") = ""
            .Item("POST_CODE") = rowARTCUST1.Item("POST_CODE") & ""
            .Item("TERM_CODE") = TERM_CODE
            .Item("INV_DUE_DATE") = INV_DUE_DATE
            .Item("SREP_CODE") = ""
            If rowARM.Item("EDI_APPOINTMENT") & "" <> "" Then
                .Item("INV_CUST_PO") = rowARM.Item("EDI_APPOINTMENT")
            Else
                .Item("INV_CUST_PO") = rowARM.Item("ORDR_CUST_PO")
            End If
            '.Item("INV_ORDR_NO") = ""
            .Item("ORDR_NO") = ""
            .Item("INV_SALES") = rowARM.Item("INV_SALES")
            .Item("INV_DISC") = 0
            .Item("INV_FREIGHT") = rowARM.Item("INV_FREIGHT")
            .Item("INV_STAX") = 0
            .Item("INV_MISC_CHG") = rowARM.Item("INV_MISC_CHG")
            .Item("INV_TOTAL_AMOUNT") = rowARM.Item("INV_TOTAL_AMOUNT")
            .Item("INV_BALANCE") = rowARM.Item("INV_TOTAL_AMOUNT")
            .Item("CUST_CODE_SO") = rowARM.Item("CUST_CODE")
            .Item("REASON_CODE") = rowARM.Item("REASON_CODE")
            .Item("SALES_DIVISION_CODE") = rowARM.Item("SALES_DIVISION_CODE")
            .Item("SEG2_CODE") = "000"
            .Item("SEG3_CODE") = "000"
            .Item("SEG4_CODE") = "000"
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            'CURRENCY
            If CURR_CODE = "" Or CURR_EXCH_RATE = 0 Then
                Stop
            End If
            .Item("CURR_CODE") = CURR_CODE
            .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
            .Item("INV_SALES_CURR") = Val(rowARM.Item("INV_SALES") & "") / CURR_EXCH_RATE
            .Item("INV_DISC_CURR") = 0
            .Item("INV_FREIGHT_CURR") = Val(rowARM.Item("INV_FREIGHT") & "") / CURR_EXCH_RATE
            .Item("INV_STAX_CURR") = 0
            .Item("INV_MISC_CHG_CURR") = Val(rowARM.Item("INV_MISC_CHG") & "") / CURR_EXCH_RATE
            .Item("GST_TAX") = (Val(rowARM.Item("INV_SALES") & "") * GST_TAX)
            .Item("GST_TAX_CURR") = (Val(rowARM.Item("INV_SALES") & "") * GST_TAX) / CURR_EXCH_RATE
            .Item("OPS_YYYYPP") = Get_ARTOPEN1_OPS_YYYYPP(rowARM.Item("INV_DATE"))
            .Item("INV_TOTAL_AMOUNT_CURR") = Val(rowARM.Item("INV_TOTAL_AMOUNT") & "") / CURR_EXCH_RATE
            .Item("INV_BALANCE_CURR") = Val(rowARM.Item("INV_TOTAL_AMOUNT") & "") / CURR_EXCH_RATE
        End With
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
    End Sub

    Sub Update_ARTPYMT3(rowSOTINVH1 As DataRow, INV_NO As String, ByRef PYMT_BATCH_ILNO_ctr As Integer, PYMT_BATCH_NO As String, S As Integer)

        Dim INV_DATE As String = rowSOTINVH1.Item("INV_DATE")
        Dim INV_BALANCE_expected As Decimal = S * Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & "")
        Dim INV_BALANCE_expected_CURR As Decimal = S * Val(rowSOTINVH1.Item("INV_TOTAL_AMT_CURR") & "")

        Dim rowSOTINVH1_CONS As DataRow = LookUp("SOTINVH1", New String() {"I", INV_NO})

        If rowSOTINVH1_CONS.Item("INV_NO_CONS") & "" <> "" Then
            INV_NO = rowSOTINVH1_CONS.Item("INV_NO_CONS") & ""
        End If
        Dim CUST_CODE As String = rowSOTINVH1_CONS.Item("ORDR_BILL_TO_CUST")
        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").Rows.Find(New String() {CUST_CODE, "I", INV_NO})
        If rowARTOPEN1 Is Nothing Then
            Fill_Records("ARTOPEN1", New String() {CUST_CODE, "I", INV_NO}, False)
            rowARTOPEN1 = dst.Tables("ARTOPEN1").Rows.Find(New String() {CUST_CODE, "I", INV_NO})
            If rowARTOPEN1 Is Nothing Then
                ASCMAIN1.sql = "Select * from ARTOPENX where CUST_CODE = :PARM1 and INV_TYPE = :PARM2 and INV_NUM = :PARM3"
                Dim rowARTOPENX As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New Object() {CUST_CODE, "I", INV_NO})

                'Dim rowARTOPENX As DataRow = LookUp("ARTOPENX", New String() {CUST_CODE, "I", INV_NO})
                ASCMAIN1.sql = "Delete from ARTOPENX where CUST_CODE = :PARM1 and INV_TYPE = :PARM2 and INV_NUM = :PARM3"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {CUST_CODE, "I", INV_NO})
                rowARTOPEN1 = dst.Tables("ARTOPEN1").NewRow
                rowARTOPEN1.ItemArray = rowARTOPENX.ItemArray
                rowARTOPEN1.Item("OPS_YYYYPP_F") = ""
            End If
        End If

        If rowARTOPEN1 Is Nothing Then
            MsgBox("Can Not Find Invoice " & INV_NO & " Required for Update!!", vbCritical, "Update Error")
            Stop ' PROBLEM WITH FINDING INVOICE NUMBER TO NET TO ZERO
        ElseIf rowSOTINVH1.Item("INV_NO_CONS") & "" = "" And INV_BALANCE_expected <> Val(rowARTOPEN1.Item("INV_BALANCE") & "") Then
            Stop    ' PROBLEM WITH FINDING INVOICE BALANCE EXPECTED
            ' WR - 2/16/04 - You may be finding this problem when you they are de-confirming a shipment whose
            ' invoices have had cash applied against them already.  
            ' If this is true AND they are deconfirming in order to
            ' change something other than anounts and re-confirm then it is OK to step through here.
        End If

        If Val(rowARTOPEN1.Item("INV_BALANCE") & "") <> 0 Then

            Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow
            With rowARTPYMT3
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                .Item("PYMT_BATCH_LNO") = 1
                PYMT_BATCH_ILNO_ctr += 1
                .Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO_ctr
                .Item("INV_TYPE") = rowARTOPEN1.Item("INV_TYPE")
                .Item("INV_NUM") = rowARTOPEN1.Item("INV_NUM")
                .Item("REASON_CODE") = rowARTOPEN1.Item("REASON_CODE")
                .Item("INV_DATE") = rowARTOPEN1.Item("INV_DATE")
                .Item("INV_DUE_DATE") = rowARTOPEN1.Item("INV_DUE_DATE")
                .Item("CUST_CODE_SO") = rowARTOPEN1.Item("CUST_CODE_SO")
                .Item("CUST_STORE_NO") = rowARTOPEN1.Item("CUST_STORE_NO")
                .Item("INV_CUST_PO") = rowARTOPEN1.Item("INV_CUST_PO")
                .Item("INV_PMT") = INV_BALANCE_expected '.Item("INV_BALANCE")
                .Item("INV_DISC_TAKEN") = 0
                .Item("INV_WRITE_OFF") = 0
                .Item("INV_BALANCE_NEW") = Val(rowARTOPEN1.Item("INV_BALANCE") & "") - INV_BALANCE_expected
                .Item("POST_CODE") = rowARTOPEN1.Item("POST_CODE")
                .Item("SEG2_CODE") = rowARTOPEN1.Item("SEG2_CODE")
                .Item("SEG3_CODE") = rowARTOPEN1.Item("SEG3_CODE")
                .Item("SEG4_CODE") = rowARTOPEN1.Item("SEG4_CODE")

                .Item("INV_BALANCE") = Val(rowARTOPEN1.Item("INV_BALANCE") & "")
                .Item("INV_BALANCE_CURR") = Val(rowARTOPEN1.Item("INV_BALANCE_CURR") & "")
                .Item("INV_PMT_CURR") = INV_BALANCE_expected 'dynARTOPEN1.Fields("INV_BALANCE")
                .Item("INV_DISC_TAKEN_CURR") = 0
                .Item("INV_WRITE_OFF_CURR") = 0
                .Item("INV_BALANCE_NEW_CURR") = Val(rowARTOPEN1.Item("INV_BALANCE") & "") - INV_BALANCE_expected
                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE


            End With
            dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)

            With rowARTOPEN1
                .Item("INV_LAST_PMT") = INV_DATE
                .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + INV_BALANCE_expected '.Item("INV_BALANCE")
                .Item("INV_PMT_CURR") = (Val(.Item("INV_PMT_CURR") & "") + INV_BALANCE_expected_CURR)
                .Item("INV_BALANCE") = Val(.Item("INV_BALANCE") & "") - INV_BALANCE_expected
                .Item("INV_BALANCE_CURR") = (Val(.Item("INV_BALANCE_CURR") & "") - INV_BALANCE_expected_CURR)
                .Item("INV_LAST_PMT_REF") = "AUTO"
                .Item("INV_LAST_PMT_REF_DT") = INV_DATE
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
            End With
        End If
    End Sub

    Function Update_ARTPYMT1(CUST_CODE As String, INV_DATE As Date) As String

        Dim PYMT_BATCH_NO As String = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")

        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        Dim CUST_NAME As String = rowARTCUST1.Item("CUST_NAME") & ""

        Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
        With rowARTPYMT1
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
            .Item("PYMT_BATCH_DATE") = INV_DATE
            .Item("BANK_CODE") = DBNull.Value
            .Item("STATUS") = "1"
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("PYMT_APPL_ONLY") = "1"
            .Item("OPS_YYYYPP") = RYP
            .Item("CURR_CODE") = CURR_CODE
            .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
            .Item("PYMT_SOURCE") = "K"
        End With
        dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

        Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
        With rowARTPYMT2
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
            .Item("PYMT_BATCH_LNO") = 1
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_NAME") = CUST_NAME
            .Item("CUST_PYMT_REF_NO") = "AUTO"
            .Item("CUST_PYMT_REF_DATE") = INV_DATE
            .Item("CUST_PYMT_AMT") = 0
            .Item("PYMT_STATUS") = "2"
            .Item("CUST_PYMT_AMT_CURR") = 0
            .Item("CURR_CODE") = CURR_CODE
            .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
        End With
        dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)


        Return PYMT_BATCH_NO
    End Function

    Function Get_ARTOPEN1_OPS_YYYYPP(INV_DATE As String) As String
        Dim NYM As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, 1)
        Dim CYP_or_NYP As String = ASCMAIN1.CYP
        If Format(CDate(INV_DATE), "yyyyMM") >= Format(CDate(Microsoft.VisualBasic.Strings.Right(NYM, 2) & "/01/" & Microsoft.VisualBasic.Strings.Left(NYM, 4)), "yyyyMM") Then
            CYP_or_NYP = NYP
        End If
        Return CYP_or_NYP
    End Function
End Class