Public Class TARPEND1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.CYP, 0, 0, 0)

        ASCMAIN1.sql = "SELECT ASTOPST1.MENU_ITEM_OBJECT, ASTOPST1.INIT_DATE" _
            & ", ASTOPST1.USER_ID, ASTMENU1.MENU_ITEM_DESC" _
            & " FROM ASTOPST1,ASTMENU1 " _
            & " WHERE ASTOPST1.YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " AND ASTOPST1.PRD_CLOSE_IND = '1' AND ASTOPST1.UPDATED = '1'" _
            & " AND ASTMENU1.MENU_ID = ASTOPST1.MENU_ID" _
            & " AND ASTMENU1.MENU_ITEM_OBJECT <> 'TARPEND1'" _
            & " AND ASTMENU1.MENU_ITEM_TYPE = ASTOPST1.MENU_ITEM_TYPE" _
            & " AND ASTMENU1.MENU_ITEM_OBJECT = ASTOPST1.MENU_ITEM_OBJECT"
        Dim sql As String = ASCMAIN1.sql

        Dim tblTATPEND1 As DataTable = ASCDATA1.GetDataTable("", "TATPEND1")

        grdTATPEND1.DataSource = tblTATPEND1
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "U"

    End Sub

    Public Overrides Sub Print_Report()
        'Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            If ASCDATA1.GetDataValue("Select PRD_CLOSE_IND from ASTPCTL1") & "" <> "1" Then
                EMsg = EMsg & vbCr & "Period-End has not been Initialized"
            End If
            Dim z As String = Absx1.cmbFor("RYP").Text
            z = Mid(z, 1, 4) & Mid(z, 6, 2)
            Dim zctl As String = ASCDATA1.GetDataValue("Select CURR_YEAR || CURR_PERIOD from ASTPCTL1") & ""
            If zctl <> z Then
                EMsg = EMsg & vbCr & "Incorrect Period to Finalize"
            End If

            If EMsg = "" Then
                'Check_for_Records("ICTIADJ1", "Inventory Adjustment Journal", "NVL(JOURNAL_IND,'0') = '0'")
                'Check_for_Records("ICTIXFR1", "Warehouse Transfer Journal", "NVL(JOURNAL_IND,'0') = '0'")
                'Check_for_Records("ICTIREC1", "PO Receipts Journal", "NVL(REGISTER_IND,'0') = '0'")


                If ASCMAIN1.CLIENT = "VAN" Then
                    Check_for_Records("SOTINVH1", "Sales Journal (Month End) Must be Updated", "NVL(REGISTER_IND,'0') = '0' AND NVL(ORDR_YYYYPP_UPDATED,'0') = '" & z & "'")

                    ASCMAIN1.sql = "Select count (*) from POTACCR1 where NVL(OPS_YYYYPP,'0') = '" & z & "'"
                    Dim sql As String = ASCMAIN1.sql
                    Dim r As Long = Val(ASCDATA1.GetDataValue() & "")
                    If r = 0 Then
                        EMsg = "Accrued PO Month-End Report Must Be Updated for Current Period " & z & EMsg
                    End If

                End If

                If EMsg <> "" Then
                    EMsg = "Cannot Proceed because a Clean Cut-off has not been established as follows:" & vbCr & EMsg
                End If
            End If
        End If

    End Sub

    Sub Check_for_Records( _
    ByVal TABLE_NAME As String, _
    ByVal TABLE_DESC As String, _
    Optional ByVal where_clause As String = "", _
    Optional ByVal custom_sql As String = "")

        If custom_sql <> "" Then
            ASCMAIN1.sql = custom_sql
        Else
            ASCMAIN1.sql = "Select count (*) from " & TABLE_NAME
            If where_clause <> "" Then
                ASCMAIN1.sql &= " where " & where_clause
            End If
        End If

        Dim sql As String = ASCMAIN1.sql
        Dim r As Long = Val(ASCDATA1.GetDataValue() & "")
        If r <> 0 Then
            EMsg &= vbCr & TABLE_DESC & " (" & TABLE_NAME & ") " & CStr(r) & " Records"
        End If

    End Sub

    Overrides Sub Update_Record()

        Get_PARM("ICTPARM1")

        Dim NYP As String = ASCMAIN1.CYP
        If Mid$(NYP, 5, 2) = "12" Then
            Dim YYYY As Integer = Val(Mid$(NYP, 1, 4))
            Mid$(NYP, 5, 2) = "01"
            Mid$(NYP, 1, 4) = Format$(YYYY + 1, "0000")
        Else
            Dim PP As Integer = Val(Mid$(NYP, 5, 2))
            Mid$(NYP, 5, 2) = Format$(PP + 1, "00")
        End If

        ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        Dim CYPdt As Date = CDate(ASCDATA1.GetDataValue) '.AddDays(1)


        Dim LYP As String = ASCMAIN1.Period_Calc(NYP, -1)
        ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 " _
        & "where OPS_YYYYPP = '" & LYP & "'"
        Dim DATE_FIRST As Date = CDate(ASCDATA1.GetDataValue).AddDays(1)


        ' A/R

        Get_PARM("ARTPARM1")
        TAC.ARCMAIN1.Get_Aging_Data(ROWs("ARTPARM1"), CYPdt, True, True)

        Dim sql_type As String = ""
        For Each T As String In New String() {"I", "R", "C", "D", "B", "O"}
            sql_type &= ", Sum (CASE WHEN INV_TYPE = '" & T & "' THEN INV_BALANCE ELSE 0 END) TYP_" & T & "_OPEN" & vbCrLf
            sql_type &= ", Sum (CASE WHEN INV_TYPE = '" & T & "' THEN 1 ELSE 0 END) TYP_" & T & "_CNT" & vbCrLf
        Next

        Dim sql_ARTOPEN1 = "Select ARTCUST1.CUST_CODE CUST_BILL_TO_CUST" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_BALANCE) INV_BALANCE " & vbCrLf _
            & TAC.ARCMAIN1.AGED_TOTALS _
            & sql_type _
            & " from ARTOPEN1,ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = ARTOPEN1.CUST_CODE " & vbCrLf _
            & " and ARTOPEN1.INV_BALANCE <> 0" & vbCrLf _
            & " and NVL(ARTOPEN1.OPS_YYYYPP,'" & ASCMAIN1.CYP & "') <= '" & ASCMAIN1.CYP & "'" _
            & " group by ARTCUST1.CUST_CODE"

        ASCMAIN1.sql = "Insert into ARTSTMT1 (OPS_YYYYPP,CUST_CODE" _
            & ",AGE_0,AGE_1,AGE_2,AGE_3,AGE_4" _
            & ",DUE_0,DUE_1,DUE_2,DUE_3,DUE_4" _
            & ",TYP_I_OPEN,TYP_R_OPEN,TYP_C_OPEN,TYP_D_OPEN,TYP_B_OPEN,TYP_O_OPEN" _
            & ",TYP_I_CNT,TYP_R_CNT,TYP_C_CNT,TYP_D_CNT,TYP_B_CNT,TYP_O_CNT) " _
            & " Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, CUST_CODE" _
            & ",AGE_0,AGE_1,AGE_2,AGE_3,AGE_4" _
            & ",DUE_0,DUE_1,DUE_2,DUE_3,DUE_4" _
            & ",TYP_I_OPEN,TYP_R_OPEN,TYP_C_OPEN,TYP_D_OPEN,TYP_B_OPEN,TYP_O_OPEN" _
            & ",TYP_I_CNT,TYP_R_CNT,TYP_C_CNT,TYP_D_CNT,TYP_B_CNT,TYP_O_CNT" _
            & " from (Select CUST_CODE, X.*" _
            & " from ARTCUST1,(" & sql_ARTOPEN1 & ") X where X.CUST_BILL_TO_CUST (+) = ARTCUST1.CUST_CODE)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update ARTSTMT1 Set TOTAL_DUE = NVL(DUE_1,0)+NVL(DUE_2,0)+NVL(DUE_3,0)+NVL(DUE_4,0) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update ARTSTMT1 Set TOTAL_OPEN_AMT = NVL(AGE_0,0)+NVL(AGE_1,0)+NVL(AGE_2,0)+NVL(AGE_3,0)+NVL(AGE_4,0) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " & vbCrLf _
            & "SELECT CUST_CODE, TOTAL_DUE FROM ARTSTMT1" & vbCrLf _
            & "WHERE OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "';" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "UPDATE ARTSTMT1 SET BALFWD = R1.TOTAL_DUE WHERE OPS_YYYYPP = '" & ASCMAIN1.CYP & "' AND CUST_CODE= R1.CUST_CODE;" & vbCrLf _
            & "IF SQL%NOTFOUND THEN INSERT INTO ARTSTMT1 (CUST_CODE, OPS_YYYYPP, BALFWD) " & vbCrLf _
            & " VALUES (R1.CUST_CODE,'" & ASCMAIN1.CYP & "',R1.TOTAL_DUE); END IF;" _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        Dim sql_act As String = "Select ARTCUST1.CUST_CODE CUST_BILL_TO_CUST, ARTCUST1.CUST_CREDIT_GROUP_CUST, " & vbCrLf _
            & " ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO, " & vbCrLf _
            & " ARTPYMT2.CUST_PYMT_AMT, ARTPYMT1.PYMT_BATCH_DATE," & vbCrLf _
            & " SUM (ARTPYMT3.INV_PMT) AR_DOLLARS," & vbCrLf _
            & " SUM (ARTPYMT3.INV_PMT * (ARTPYMT1.PYMT_BATCH_DATE - ARTPYMT3.INV_DATE)) DAY_DOLLARS," & vbCrLf _
            & " COUNT (DECODE (ARTPYMT3.INV_TYPE, 'I', 1, 0)) NO_INVOICES," & vbCrLf _
            & " SUM (DECODE (ARTPYMT3.INV_TYPE, 'O', ARTPYMT3.INV_PMT, 0)) DOLLARS_PAID_OA" & vbCrLf _
            & " From ARTPYMT1, ARTPYMT2, ARTPYMT3, ARTCUST1" & vbCrLf _
            & " WHERE ARTPYMT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " AND ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE" & vbCrLf _
            & " AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
            & " AND ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & " AND ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
            & " GROUP BY ARTCUST1.CUST_CODE, ARTCUST1.CUST_CREDIT_GROUP_CUST," & vbCrLf _
            & " ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO," & vbCrLf _
            & " ARTPYMT2.CUST_PYMT_AMT , ARTPYMT1.PYMT_BATCH_DATE"

        'ASCMAIN1.sql = "" _
        '    & "Begin " _
        '    & " Declare Cursor C1 is " _
        '    & "  Select CUST_BILL_TO_CUST CUST_CODE" _
        '    & ",  SUM (CUST_PYMT_AMT) CUST_PYMT_AMT" _
        '    & ",  SUM (DOLLARS_PAID_OA) DOLLARS_PAID_OA" _
        '    & ",  SUM (DAY_DOLLARS) DAY_DOLLARS" _
        '    & ",  SUM (NO_INVOICES) NO_INVOICES" _
        '    & "   from (" & sql_act & ") X group by CUST_BILL_TO_CUST;" _
        '    & " Begin" _
        '    & "  For R1 in C1 Loop" _
        '    & "   Update ARTSTMT1 Set " _
        '    & "    DAY_DOLLARS = R1.DAY_DOLLARS," _
        '    & "    DOLLARS_PAID = R1.CUST_PYMT_AMT + R1.DOLLARS_PAID_OA," _
        '    & "    NO_INVOICES = R1.NO_INVOICES" _
        '    & "    Where CUST_CODE = R1.CUST_CODE;" _
        '    & "  End Loop;" _
        '    & " End;" _
        '    & "End;"
        'ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is " _
            & "  Select * from ARTCUST6;" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update ARTSTMT1 Set " _
            & "    CUST_HIGH_BAL_DATE = R1.CUST_HIGH_BAL_DATE," _
            & "    CUST_HIGH_BAL_AMT = R1.CUST_HIGH_BAL_AMT" _
            & "    Where CUST_CODE = R1.CUST_CODE;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("A/R Closed Items Purge", "")
        ASCDATA1.ExecuteSQL("Update ARTOPEN1 set OPS_YYYYPP_F = '" & ASCMAIN1.CYP & "' where INV_BALANCE = 0 ")
        ASCDATA1.ExecuteSQL("Insert into ARTOPENX Select * from ARTOPEN1 where OPS_YYYYPP_F is Not Null")
        ASCDATA1.ExecuteSQL("Delete from ARTOPEN1 where OPS_YYYYPP_F is Not Null")


        If ASCMAIN1.CLIENT = "VAN" Then
            '  REM INSERT RECORDS TO POTSHIPH HISTORICAL IN-TRANSIT
            ASCMAIN1.sql = "Insert into POTSHIPH Select '" & RYP & "' OPS_YYYYPP,POTSHIP3.*,POTSHIP2.ACCRUAL_STATUS,POTSHIP2.VOUCHER_NO,POTSHIP1.PO_DATE_SHIPPED FROM POTSHIP3,POTSHIP2,POTSHIP1" _
            & " WHERE  POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" _
            & " And POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" _
            & " And POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" _
            & " And (POTSHIP2.PO_SHIP_STATUS  = 'O' OR POTSHIP2.OPS_YYYYPP > '" & RYP & "')" _
            & " And POTSHIP1.PO_DATE_SHIPPED < '" & Format(DATE_FIRST, "dd-MMM-yyyy") & "'" '01-OCT-2022'" 1ST DATE OF NECT MONTH
            ASCDATA1.ExecuteSQL()
        End If

        '' Sales Rep Snapshots
        'ASCMAIN1.Progress("Sales Rep Snapshots", "")
        'ASCMAIN1.sql = "Insert into ARTCUST4 Select '" & RYP & "' OPS_YYYYPP, ARTCUST1.CUST_CODE" _
        '    & ", ARTCUST1.SREP_CODE, ARTCUST1.SREP2_CODE, ARTCUST1.SREP_CODE_OVER" _
        '    & ", DECODE(SOTSREP2.SREP_CODE,NULL,SOTSREP1.SREP_COMM_PCT,SOTSREP2.SREP_COMM_PCT) SREP_COMM_PCT" _
        '    & ", DECODE(SOTSREP2.SREP_CODE,NULL,SOTSREP1.SREP_COMM_PCT_SPEC,SOTSREP2.SREP_COMM_PCT_SPEC) SREP_COMM_PCT_SPEC" _
        '    & ", DECODE(SOTSREP2_2.SREP_CODE,NULL,SOTSREP1_2.SREP_COMM_PCT,SOTSREP2_2.SREP_COMM_PCT) SREP2_COMM_PCT" _
        '    & ", DECODE(SOTSREP2_2.SREP_CODE,NULL,SOTSREP1_2.SREP_COMM_PCT_SPEC,SOTSREP2_2.SREP_COMM_PCT_SPEC) SREP2_COMM_PCT_SPEC" _
        '    & ", ARTCUST1.SREP_COMM_PCT_OVER" _
        '    & " from ARTCUST1,SOTSREP1 SOTSREP1, SOTSREP1 SOTSREP1_2, SOTSREP2, SOTSREP2 SOTSREP2_2" _
        '    & " where SOTSREP1.SREP_CODE (+) = ARTCUST1.SREP_CODE" _
        '    & "   and SOTSREP1_2.SREP_CODE (+) = ARTCUST1.SREP2_CODE" _
        '    & "   and SOTSREP2.SREP_CODE (+) = ARTCUST1.SREP_CODE" _
        '    & "   and SOTSREP2.CUST_CODE (+) = ARTCUST1.CUST_CODE" _
        '    & "   and SOTSREP2_2.SREP_CODE (+) = ARTCUST1.SREP2_CODE" _
        '    & "   and SOTSREP2_2.CUST_CODE (+) = ARTCUST1.CUST_CODE"
        'ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Insert into ARTCUST7 Select '" & RYP & "' OPS_YYYYPP" _
        '    & ", ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" _
        '    & ", ARTCUST2.SREP_CODE, ARTCUST2.SELL_CODE" _
        '    & " from ARTCUST2"
        'ASCDATA1.ExecuteSQL()



        ASCMAIN1.Progress("Reset Customer MTD Activity Summary", "")
        ASCMAIN1.sql = "Update ARTCUST6 Set" & vbCrLf _
            & " CUST_SALES_MTD = 0" & vbCrLf _
            & ", CUST_COGS_MTD = 0" & vbCrLf _
            & ", CUST_CASH_MTD = 0" & vbCrLf _
            & ", CUST_FIN_CHG_MTD = 0" & vbCrLf _
            & ", CUST_NUM_INV_MTD = 0" & vbCrLf _
            & ", CUST_NUM_FIN_MTD = 0" & vbCrLf _
            & ", CUST_CRED_MTD = 0" & vbCrLf _
            & ", CUST_HIGH_BAL_DATE = NULL" & vbCrLf _
            & ", CUST_HIGH_BAL_AMT = 0"
        ASCDATA1.ExecuteSQL()

        If Mid$(ASCMAIN1.CYP, 5, 2) = "12" Then
            ASCMAIN1.sql = "Update ARTCUST6 Set" & vbCrLf _
                & "  CUST_SALES_LYR = CUST_SALES_YTD" & vbCrLf _
                & ", CUST_COGS_LYR = CUST_COGS_YTD" & vbCrLf _
                & ", CUST_CASH_LYR = CUST_CASH_YTD" & vbCrLf _
                & ", CUST_FIN_CHG_LYR = CUST_FIN_CHG_YTD" & vbCrLf _
                & ", CUST_NUM_INV_LYR = CUST_NUM_INV_YTD" & vbCrLf _
                & ", CUST_NUM_FIN_LYR = CUST_NUM_FIN_YTD" & vbCrLf _
                & ", CUST_CRED_LYR = CUST_CRED_YTD"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update ARTCUST6 Set" & vbCrLf _
                & "  CUST_SALES_YTD = 0" & vbCrLf _
                & ", CUST_COGS_YTD = 0" & vbCrLf _
                & ", CUST_CASH_YTD = 0" & vbCrLf _
                & ", CUST_FIN_CHG_YTD = 0" & vbCrLf _
                & ", CUST_NUM_INV_YTD = 0" & vbCrLf _
                & ", CUST_NUM_FIN_YTD = 0" & vbCrLf _
                & ", CUST_CRED_YTD = 0"
            ASCDATA1.ExecuteSQL()
        End If


        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select CUST_CODE, SUM (INV_BALANCE) INV_BALANCE" & vbCrLf _
            & " from ARTOPEN1 group by CUST_CODE;" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ARTCUST6 Set CUST_HIGH_BAL_DATE = TRUNC(SYSDATE)," & vbCrLf _
            & " CUST_HIGH_BAL_AMT = R1.INV_BALANCE" & vbCrLf _
            & " where CUST_CODE = R1.CUST_CODE;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "  INSERT INTO ARTCUST6 (CUST_CODE, CUST_HIGH_BAL_DATE, CUST_HIGH_BAL_AMT)" & vbCrLf _
            & "  VALUES (R1.CUST_CODE, TRUNC(SYSDATE), R1.INV_BALANCE);" & vbCrLf _
            & " End If;" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()


        If ASCMAIN1.CLIENT = "VAN" Then
            'when I run this sql in database to fix data cyp = 200206
            ASCMAIN1.sql = " BEGIN DECLARE CURSOR C1 IS " & vbCrLf _
            & "    select CUST_CODE, SUM(INV_SALES) INV_SALES, COUNT(*) NUM_INV, MAX(ORDR_DATE_UPDATED) ORDR_DATE_UPDATED " & vbCrLf _
            & "    from (SELECT * FROM SOTINVH1 WHERE ORDR_YYYYPP_UPDATED > '" & ASCMAIN1.CYP & "') WHERE INV_TYPE = 'I' GROUP BY CUST_CODE;" & vbCrLf _
            & "    " & vbCrLf _
            & "    SOTINVH1_TYP SOTINVH1%ROWTYPE;" & vbCrLf _
            & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "    SELECT * INTO SOTINVH1_TYP FROM SOTINVH1 WHERE INV_TYPE = 'I' " & vbCrLf _
            & "    AND CUST_CODE = R1.CUST_CODE AND ORDR_YYYYPP_UPDATED > '" & ASCMAIN1.CYP & "' " & vbCrLf _
            & "    AND ORDR_DATE_UPDATED = R1.ORDR_DATE_UPDATED AND ROWNUM <= 1;" & vbCrLf _
            & "    Update ARTCUST6" & vbCrLf _
            & "    SET CUST_LAST_INV_NUM = SOTINVH1_TYP.INV_NO" & vbCrLf _
            & "    ,CUST_LAST_INV_DATE = SOTINVH1_TYP.INV_DATE" & vbCrLf _
            & "    ,CUST_LAST_INV_AMT = SOTINVH1_TYP.INV_TOTAL_AMOUNT" & vbCrLf _
            & "    ,CUST_SALES_MTD = NVL(CUST_SALES_MTD,0) + R1.INV_SALES" & vbCrLf _
            & "    ,CUST_SALES_YTD = NVL(CUST_SALES_YTD,0) + R1.INV_SALES" & vbCrLf _
            & "    ,CUST_NUM_INV_MTD = NVL(CUST_NUM_INV_MTD,0) + R1.NUM_INV" & vbCrLf _
            & "    ,CUST_NUM_INV_YTD = NVL(CUST_NUM_INV_YTD,0) + R1.NUM_INV" & vbCrLf _
            & "    WHERE CUST_CODE = R1.CUST_CODE;" & vbCrLf _
            & "    IF SQL%NOTFOUND THEN" & vbCrLf _
            & "       INSERT INTO ARTCUST6 (CUST_CODE, CUST_LAST_INV_NUM, CUST_LAST_INV_DATE, CUST_LAST_INV_AMT" & vbCrLf _
            & "       , CUST_SALES_MTD, CUST_SALES_YTD, CUST_NUM_INV_MTD, CUST_NUM_INV_YTD) VALUES (R1.CUST_CODE, SOTINVH1_TYP.INV_NO,SOTINVH1_TYP.INV_DATE, " & vbCrLf _
            & "       SOTINVH1_TYP.INV_TOTAL_AMOUNT, R1.INV_SALES, R1.INV_SALES, R1.NUM_INV, R1.NUM_INV);" & vbCrLf _
            & "    END IF; " & vbCrLf _
            & " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL()


            ' Sales Order Processing

            ASCMAIN1.sql = "Insert into SOTFPCT1 (OPS_YYYYPP, CUST_FACTOR_PERCENT, CUST_SURCHARGE_PERCENT) " & vbCrLf _
                & " Select '" & ASCMAIN1.CYP & "', SO_PARM_FACTOR_PCT, SO_PARM_SURCHARGE_PCT from SOTPARM1"
            ASCDATA1.ExecuteSQL()

        End If


        ' A/P
        ASCMAIN1.Progress("Reset Vendor MTD Activity Summary", "")

        ASCMAIN1.sql = "Update APTVEND5 Set" & vbCrLf _
            & "  VEND_PURCHASES_MTD = 0" & vbCrLf _
            & ", VEND_PAYMENTS_MTD = 0" & vbCrLf _
            & ", VEND_DISC_TAKEN_MTD = 0" & vbCrLf _
            & ", VEND_NUM_INV_MTD = 0" & vbCrLf _
            & ", VEND_NUM_CHKS_MTD = 0" & vbCrLf
        ASCDATA1.ExecuteSQL()

        If Mid$(ASCMAIN1.CYP, 5, 2) = "12" Then
            ASCMAIN1.sql = "Update APTVEND5 Set" & vbCrLf _
                & "  VEND_PURCHASES_LYR = VEND_PURCHASES_YTD" & vbCrLf _
                & ", VEND_PAYMENTS_LYR = VEND_PAYMENTS_YTD" & vbCrLf _
                & ", VEND_DISC_TAKEN_LYR = VEND_DISC_TAKEN_YTD" & vbCrLf _
                & ", VEND_NUM_INV_LYR = VEND_NUM_INV_YTD" & vbCrLf _
                & ", VEND_NUM_CHKS_LYR = VEND_NUM_CHKS_YTD"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update APTVEND5 Set" & vbCrLf _
                & "  VEND_PURCHASES_YTD = 0" & vbCrLf _
                & ", VEND_PAYMENTS_YTD = 0" & vbCrLf _
                & ", VEND_DISC_TAKEN_YTD = 0" & vbCrLf _
                & ", VEND_NUM_INV_YTD = 0" & vbCrLf _
                & ", VEND_NUM_CHKS_YTD = 0"
            ASCDATA1.ExecuteSQL()
        End If

        ' Purge Files
        Purge_Files()

        ' Inventory Statistics recorded with NYP
        ASCMAIN1.Progress("Inventory Statistics for Next Period", "")
        ASCMAIN1.sql = "Update ICTSTAT1 " & vbCrLf _
            & " Set OPS_YYYYPP = '000000' where OPS_YYYYPP = '" & NYP & "'"
        ASCDATA1.ExecuteSQL()

        ' Inventory Statistics
        ASCMAIN1.Progress("Inventory Statistics", "")
        ASCMAIN1.sql = "Insert into ICTSTAT1 " & vbCrLf _
            & " (OPS_YYYYPP, STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_BEG) " & vbCrLf _
            & " Select '" & NYP & "'" & vbCrLf _
            & " OPS_YYYYPP, STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_ON_HAND" & vbCrLf _
            & " from ICTSTAT2 where WHSE_QTY_ON_HAND <> 0"
        ASCDATA1.ExecuteSQL()

        'If ASCMAIN1.CLIENT = "VAN" Then
        '    ASCMAIN1.sql = "Delete from ICTSTAT2" & vbCrLf _
        '        & " Where NVL(WHSE_QTY_ON_HAND,0) = 0" & vbCrLf _
        '        & " and NVL(WHSE_QTY_ON_ORDER,0) = 0" & vbCrLf _
        '        & " and NVL(WHSE_QTY_TRAN,0) = 0" & vbCrLf _
        '        & " and NVL(WHSE_QTY_OPEN,0) = 0" & vbCrLf _
        '        & " and NVL(WHSE_QTY_PICK,0) = 0" & vbCrLf _
        '        & " and NVL(WHSE_QTY_ALLO,0) = 0"
        '    ASCDATA1.ExecuteSQL()
        'End If

        ' Inventory Status History

        ASCMAIN1.Progress("Inventory Status History", "")

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Insert into ICTSTAT5 Select '" & ASCMAIN1.CYP & "',ICTSTAT2.* from ICTSTAT2"
            ASCDATA1.ExecuteSQL()
        Else
            ASCMAIN1.sql = "Insert into ICTSTAT5 Select '" & ASCMAIN1.CYP & "',ICTSTAT2.* from ICTSTAT2 where " & vbCrLf _
                & " NVL(WHSE_QTY_ON_HAND,0) <> 0 or NVL(WHSE_QTY_ON_ORDER,0) <> 0 or NVL(WHSE_QTY_TRAN,0) <> 0 or " & vbCrLf _
                & " NVL(WHSE_QTY_OPEN,0) <> 0 or NVL(WHSE_QTY_PICK,0) <> 0 or " & vbCrLf _
                & " NVL(WHSE_QTY_ALLO,0) <> 0 or NVL(WHSE_QTY_COMM,0) <> 0 or NVL(WHSE_QTY_PROD,0) <> 0"
            ASCDATA1.ExecuteSQL()
        End If


        ' need to do this after writing an evidence record to ICTSTAT5 - VAN Stat/Act report 803528xmj1x
        ' 12/09/17 - don't know why just VAN

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Delete from ICTSTAT2" & vbCrLf _
                & " Where NVL(WHSE_QTY_ON_HAND,0) = 0" & vbCrLf _
                & " and NVL(WHSE_QTY_ON_ORDER,0) = 0" & vbCrLf _
                & " and NVL(WHSE_QTY_TRAN,0) = 0" & vbCrLf _
                & " and NVL(WHSE_QTY_OPEN,0) = 0" & vbCrLf _
                & " and NVL(WHSE_QTY_PICK,0) = 0" & vbCrLf _
                & " and NVL(WHSE_QTY_ALLO,0) = 0"
            ASCDATA1.ExecuteSQL()
        End If

        ' Correct Beg/End using Inventory Statistics recorded with NYP
        ASCMAIN1.Progress("Correct Beg/End Invty using Statistics for Next Period", "")
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select * from ICTSTAT1 where OPS_YYYYPP = '000000' for Update;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT1" & vbCrLf _
            & "    Set WHSE_QTY_BEG = NVL(WHSE_QTY_BEG,0) + NVL(R1.WHSE_QTY_SHP,0) - NVL(R1.WHSE_QTY_RTN,0) - NVL(R1.WHSE_QTY_REC,0) - NVL(R1.WHSE_QTY_ADJ,0)" & vbCrLf _
            & "      , WHSE_QTY_SHP = R1.WHSE_QTY_SHP, WHSE_QTY_RTN = R1.WHSE_QTY_RTN, WHSE_QTY_REC = R1.WHSE_QTY_REC, WHSE_QTY_ADJ = R1.WHSE_QTY_ADJ" & vbCrLf _
            & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & NYP & "';" & vbCrLf _
            & "   If SQL%NOTFOUND Then " & vbCrLf _
            & "    Insert into ICTSTAT1 (OPS_YYYYPP,STYLE_CODE,COLOR_CODE,WHSE_CODE,WHSE_QTY_BEG,WHSE_QTY_SHP,WHSE_QTY_RTN,WHSE_QTY_REC,WHSE_QTY_ADJ)" & vbCrLf _
            & "     Values ('" & NYP & "',R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,NVL(R1.WHSE_QTY_SHP,0) - NVL(R1.WHSE_QTY_RTN,0) - NVL(R1.WHSE_QTY_REC,0) - NVL(R1.WHSE_QTY_ADJ,0),R1.WHSE_QTY_SHP,R1.WHSE_QTY_RTN,R1.WHSE_QTY_REC,R1.WHSE_QTY_ADJ);" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "   Update ICTSTAT5" & vbCrLf _
            & "    Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.WHSE_QTY_SHP,0) - NVL(R1.WHSE_QTY_RTN,0) - NVL(R1.WHSE_QTY_REC,0) - NVL(R1.WHSE_QTY_ADJ,0)" & vbCrLf _
            & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "   If SQL%NOTFOUND Then " & vbCrLf _
            & "    Insert into ICTSTAT5 (OPS_YYYYPP,STYLE_CODE,COLOR_CODE,WHSE_CODE,WHSE_QTY_ON_HAND)" & vbCrLf _
            & "     Values ('" & ASCMAIN1.CYP & "',R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,NVL(R1.WHSE_QTY_SHP,0) - NVL(R1.WHSE_QTY_RTN,0) - NVL(R1.WHSE_QTY_REC,0) - NVL(R1.WHSE_QTY_ADJ,0));" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "   Delete from ICTSTAT1 where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Dim SQLW As String = ""
        SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_PHYS_STATUS = 'C')"
        ASCMAIN1.sql = "" _
          & "Update ICTWHSE1 X Set WHSE_PHYS_STATUS = NULL" & ASCMAIN1.SQL_Add_WHERE(SQLW)
        ASCDATA1.ExecuteSQL()

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare XFNO VARCHAR2(6);" & vbCrLf _
                & "  Cursor C1 is" & vbCrLf _
                & "   Select * from SOTINVH1" & vbCrLf _
                & "    where ORDR_YYYYPP_UPDATED = '" & NYP & "'" & vbCrLf _
                & "      and ORDR_TYPE_CODE = 'XFR';" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   XFNO := SOPSHIP1_XFR(R1.INV_NO);" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
            ' A VERSION OF THIS UPDATE IS ALSO FOUND IN SORUPDT1



            'CARRY OVER DUTY RATES FOR THE NEW YEAR; ADDED BY RDW 2-13-02
            Dim P As String = Mid(ASCMAIN1.CYP, 5, 2)
            If P = "11" Then
                ASCMAIN1.sql = "Select * from ICTDUTY3 where OPS_YYYY = '" & Mid(ASCMAIN1.CYP, 1, 4) & "'"
                Create_TDA(dst.Tables.Add, "ICTDUTY3", "**", 0)
                Dim NY As String = Format(1 + Val(Mid(ASCMAIN1.CYP, 1, 4)), "0000")
                Fill_Records("ICTDUTY3")
                For Each row As DataRow In dst.Tables("ICTDUTY3").Select
                    row.Item("OPS_YYYY") = NY
                    row.AcceptChanges()
                    row.SetAdded()
                Next
                Update_Record_TDA("ICTDUTY3")
            End If
        End If

        ' WHY DO WE NEED TO INITIALIZE THESE?
        'ASCDATA1.ExecuteSQL("Delete from ICTPHYC1")
        'ASCDATA1.ExecuteSQL("Delete from ICTPHYC2")

        ' Aged PO History
        If ASCMAIN1.CLIENT = "VAN" Then
            Dim ICTSTKL2 As String = TAC.POCMAIN1.Aged_PO(Me, ASCMAIN1.CYP)
            ASCMAIN1.sql = "Insert into ICTSTKL2 Select * from " & ICTSTKL2
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Insert into ICTSTKL0" & vbCrLf _
                & " Select '" & NYP & "' OPS_YYYYPP" & vbCrLf _
                & ", OPS_YYYYPP_FC, SALES_DIVISION_CODE, CATGY_CODE, STK, CUST_CODE, QTY_PROJ, SLS_PROJ, CST_PROJ, AVG_PRICE, AVG_COST" & vbCrLf _
                & " from ICTSTKL0" & vbCrLf _
                & " where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
                & "   and OPS_YYYYPP_FC >= '" & NYP & "'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into ICTSTKL0" _
                & "Select '" & NYP & "', YP, SALES_DIVISION_CODE, CATGY_CODE, DECODE(CUST_CODE,NULL,'S','N'), CUST_CODE" & vbCrLf _
                & ", QTY, SLS, CGS, SLS / QTY, CGS / QTY" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select X.YP, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.CUST_CODE, ICTBODY2.CATGY_CODE" & vbCrLf _
                & ", SUM (QTY) QTY, SUM (SLS) SLS, SUM (CGS) CGS" & vbCrLf _
                & " from ICTSTYL1,ICTBODY2,(" & vbCrLf _
                & "Select TRIM(TO_CHAR(TO_NUMBER(SUBSTR(ORDR_YYYYPP_UPDATED,1,4))+1,'0000')) || SUBSTR(ORDR_YYYYPP_UPDATED,5,2) YP, STYLE_CODE" & vbCrLf _
                & ", SUM (ORDR_QTY_SHIP) QTY" & vbCrLf _
                & ", SUM (ORDR_QTY_SHIP * ORDR_UNIT_PRICE) SLS" & vbCrLf _
                & ", SUM (ORDR_QTY_SHIP * ORDR_UNIT_COST) CGS" & vbCrLf _
                & " from SOTINVH2 where SOTINVH2.INV_TYPE = 'I' AND SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
                & " group by ORDR_YYYYPP_UPDATED, STYLE_CODE) X" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE AND ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE" & vbCrLf _
                & " group by X.YP, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.CUST_CODE, ICTBODY2.CATGY_CODE" & vbCrLf _
                & ") Y" & vbCrLf _
                & " where Y.QTY <> 0"

        End If


        ' Sales Commissions
        '     TAC.SOCMAIN1.SetMonthlyCommissions()

        If ASCMAIN1.CLIENT = "VAN" Then
            ' DO NOTHING
        Else
            CreateCommissionData(ASCMAIN1.CYP)
        End If

        ' GL Control Account Subsidiary Snapshots
        Load_GLREC()

        If ASCMAIN1.CLIENT = "VAN" Then

            Create_TDA(dst.Tables.Add, "ICTSEAS1", "*")

            Dim YP3MOS As String = ASCMAIN1.Period_Calc(NYP, 12 + 3) ' Calculate the period which is 1 year and 3 months from new YP
            Dim SEASON_TYPE As String = ""
            Dim SEASON_TYPE_DESC As String = ""

            If Mid(YP3MOS, 5, 2) = "02" Then SEASON_TYPE = "S" : SEASON_TYPE_DESC = "SPRING"
            If Mid(YP3MOS, 5, 2) = "05" Then SEASON_TYPE = "M" : SEASON_TYPE_DESC = "SUMMER"
            If Mid(YP3MOS, 5, 2) = "08" Then SEASON_TYPE = "F" : SEASON_TYPE_DESC = "FALL"
            If Mid(YP3MOS, 5, 2) = "11" Then SEASON_TYPE = "H" : SEASON_TYPE_DESC = "HOLIDAY"

            If SEASON_TYPE <> "" Then
                Dim SEASON_YEAR As String = Mid(YP3MOS, 1, 4)
                Dim SEASON_CODE As String = Mid(SEASON_TYPE_DESC, 1, 3) & Mid(SEASON_YEAR, 3, 2)
                Dim rowICTSEAS1 As DataRow = Fill_Record("ICTSEAS1", SEASON_CODE) '  LookUp("ICTSEAS1", SEASON_CODE)

                If rowICTSEAS1 Is Nothing Then
                    rowICTSEAS1 = dst.Tables("ICTSEAS1").NewRow
                    rowICTSEAS1.Item("SEASON_CODE") = SEASON_CODE
                    dst.Tables("ICTSEAS1").Rows.Add(rowICTSEAS1)
                End If

                With rowICTSEAS1
                    .Item("SEASON_TYPE") = SEASON_TYPE
                    .Item("SEASON_YEAR") = SEASON_YEAR
                    .Item("SEASON_DESC") = SEASON_TYPE_DESC & " " & SEASON_YEAR
                    .Item("SEASON_ACTIVE") = "1"
                    .Item("SEASON_YP_BEG") = YP3MOS
                    .Item("SEASON_YP_END") = ASCMAIN1.Period_Calc(YP3MOS, 2)
                End With
                Update_Record_TDA("ICTSEAS1")

                ASCMAIN1.sql = "Update ICTSEAS1 Set SEASON_ACTIVE = '0' where SEASON_YP_BEG < '" & ASCMAIN1.Period_Calc(NYP, -15) & "'"
                ASCDATA1.ExecuteSQL()

            End If
        End If



        If ASCMAIN1.CLIENT = "RGI" Then

            Create_TDA(dst.Tables.Add, "ICTSEAS1", "*")

            Dim YP6MOS As String = ASCMAIN1.Period_Calc(NYP, 12 + 6) ' Calculate the period which is 1 year and 6 months from new YP
            Dim SEASON_TYPE As String = ""
            Dim SEASON_TYPE_DESC As String = ""

            If Mid(YP6MOS, 5, 2) = "02" Then SEASON_TYPE = "S" : SEASON_TYPE_DESC = "Spring"
            If Mid(YP6MOS, 5, 2) = "08" Then SEASON_TYPE = "F" : SEASON_TYPE_DESC = "Fall"

            If SEASON_TYPE <> "" Then
                Dim SEASON_YEAR As String = Mid(YP6MOS, 1, 4)
                Dim SEASON_CODE As String = SEASON_YEAR & SEASON_TYPE
                Dim rowICTSEAS1 As DataRow = Fill_Record("ICTSEAS1", SEASON_CODE)

                If rowICTSEAS1 Is Nothing Then
                    rowICTSEAS1 = dst.Tables("ICTSEAS1").NewRow
                    rowICTSEAS1.Item("SEASON_CODE") = SEASON_CODE
                    dst.Tables("ICTSEAS1").Rows.Add(rowICTSEAS1)
                End If

                With rowICTSEAS1
                    .Item("SEASON_TYPE") = SEASON_TYPE
                    .Item("SEASON_YEAR") = SEASON_YEAR
                    .Item("SEASON_DESC") = SEASON_TYPE_DESC & " " & SEASON_YEAR
                    .Item("SEASON_ACTIVE") = "1"
                    .Item("SEASON_YP_BEG") = YP6MOS
                    .Item("SEASON_YP_END") = ASCMAIN1.Period_Calc(YP6MOS, 5)
                End With
                Update_Record_TDA("ICTSEAS1")

                ASCMAIN1.sql = "Update ICTSEAS1 Set SEASON_ACTIVE = '0' where SEASON_YP_BEG < '" & ASCMAIN1.Period_Calc(NYP, -18) & "'"
                ASCDATA1.ExecuteSQL()

            End If
        End If

        ' Close Period
        ASCMAIN1.Progress("Updating Period Control Record", "")
        ASCMAIN1.sql = "Update ASTPCTL1 set CURR_YEAR = '" & Mid$(NYP, 1, 4) & "'," & vbCrLf _
            & " CURR_PERIOD = '" & Mid$(NYP, 5, 2) & "'," & vbCrLf _
            & " PRD_CLOSE_IND = Null"
        ASCDATA1.ExecuteSQL()

        ' Purge SQL Execution Log
        ASCMAIN1.Progress("Cleaning up SQL Execution Log", "")
        ASCMAIN1.sql = "Delete from ASTSQLX1"
        ASCDATA1.ExecuteSQL()

        ' Purge Control Number Generation Log
        ASCMAIN1.Progress("Cleaning up Control Number Generation Log", "")
        ASCMAIN1.sql = "Delete from TATCTLN2"
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Load_GLREC()

        Get_PARM("GLTPARM1")
        Get_PARM("SOTPARM1")

        ASCMAIN1.sql = "Delete from GLTCREC1"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from GLTCREC2"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from GLTCREC4"
        ASCDATA1.ExecuteSQL()

        'Inventory On Hand
        ASCMAIN1.Progress("Subsidiary Snapshots", "Inventory On Hand")
        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'IC', 'Inventory On Hand', 'D' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'IC', STYLE_CLASS_CODE, STYLE_CLASS_DESC from ICTCLAS1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'IC' CREC_TYPE_CODE, ICTSTYL1.STYLE_CLASS_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", ICTCLAS1.ACCT_CODE_ONH ACCT_CODE" & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            & ", NULL DETL_CTL_TYPE, NULL DETL_CTL_NO, 'G' DETL_CVX_TYPE, ICTSTYL1.STYLE_GROUP_CODE DETL_CVX_NO" & vbCrLf _
            & ", SUM(NVL(ICTSTYL1.STYLE_COST,0) * NVL(ICTSTAT5.WHSE_QTY_ON_HAND,0)) CREC_AMT" & vbCrLf _
            & " from ICTSTAT5,ICTSTYL1,ICTCLAS1" & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE = ICTSTAT5.STYLE_CODE" & vbCrLf _
            & "   and ICTSTAT5.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and ICTCLAS1.STYLE_CLASS_CODE (+) = ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
            & " group by ICTSTYL1.STYLE_CLASS_CODE, ICTCLAS1.ACCT_CODE_ONH" & vbCrLf _
            & ", ICTSTYL1.STYLE_GROUP_CODE" & vbCrLf
        ASCMAIN1.sql = "Insert into GLTCREC3 " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()


        ' Accrued Purchases
        ASCMAIN1.Progress("Subsidiary Snapshots", "Accrued Purchases")
        ASCMAIN1.sql = "INSERT INTO GLTCREC1 SELECT 'ICP', 'Accrued Purchases', 'C' FROM DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'ICP', STYLE_CLASS_CODE, STYLE_CLASS_DESC from ICTCLAS1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "SELECT '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'ICP' CREC_TYPE_CODE," & vbCrLf _
            & "       ICTSTYL1.STYLE_CLASS_CODE CREC_CLASS_CODE, '" & ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_INVTY_PUR") & "' ACCT_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE," & vbCrLf _
            & "       'R' DETL_CTL_TYPE, ICTIREC1.RECEIPT_NO DETL_CTL_NO, 'V' DETL_CVX_TYPE, ICTIREC1.VEND_CODE DETL_CVX_NO," & vbCrLf _
            & "       SUM(ICTIREC2.PO_COST * ICTIREC2.QTY_REC) CREC_AMT" & vbCrLf _
            & " from ICTIREC1,ICTIREC2,ICTSTYL1" & vbCrLf _
            & "  where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = ICTIREC2.STYLE_CODE" & vbCrLf _
            & "   and ICTIREC1.ACCRUAL_STATUS = '0'" & vbCrLf _
            & " group by ICTSTYL1.STYLE_CLASS_CODE, ICTIREC1.RECEIPT_NO, ICTIREC1.VEND_CODE"
        ASCMAIN1.sql = "INSERT INTO GLTCREC3 " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()


        'Accounts Receivable
        ASCMAIN1.Progress("Subsidiary Snapshots", "Accounts Receivable")
        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'AR', 'Accounts Receivable', 'D' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'AR', POST_CODE, POST_DESC from ARTPOST1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'AR' CREC_TYPE_CODE, ARTOPEN1.POST_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", ARTPOST1.ACCT_CODE ACCT_CODE" & vbCrLf _
            & ", ARTOPEN1.SEG2_CODE, ARTOPEN1.SEG3_CODE, ARTOPEN1.SEG4_CODE" & vbCrLf _
            & ", ARTOPEN1.INV_TYPE DETL_CTL_TYPE, ARTOPEN1.INV_NUM DETL_CTL_NO" & vbCrLf _
            & ", 'C' DETL_CVX_TYPE, ARTOPEN1.CUST_CODE DETL_CVX_NO" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE CREC_AMT" & vbCrLf _
            & " from ARTOPEN1,ARTPOST1" & vbCrLf _
            & " where ARTPOST1.POST_CODE (+) = ARTOPEN1.POST_CODE" & vbCrLf _
            & "   and NVL(ARTOPEN1.OPS_YYYYPP,'" & ASCMAIN1.CYP & "') <= '" & ASCMAIN1.CYP & "'"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into GLTCREC4 Select 'C', CUST_CODE, CUST_NAME from ARTCUST1"
        ASCDATA1.ExecuteSQL()

        'Accounts Payable
        ASCMAIN1.Progress("Subsidiary Snapshots", "Accounts Payable")
        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'AP', 'Accounts Payable', 'C' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'AP', POST_CODE, POST_DESC from APTPOST1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'AP' CREC_TYPE_CODE, APTINVH1.POST_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", APTPOST1.ACCT_CODE ACCT_CODE" & vbCrLf _
            & ", APTINVH1.SEG2_CODE, APTINVH1.SEG3_CODE, APTINVH1.SEG4_CODE" & vbCrLf _
            & ", APTINVH1.INV_TYPE DETL_CTL_TYPE, APTINVH1.VOUCHER_NO DETL_CTL_NO" & vbCrLf _
            & ", 'V' DETL_CVX_TYPE, APTINVH1.VEND_CODE DETL_CVX_NO" & vbCrLf _
            & ", APTINVH1.INV_BALANCE CREC_AMT" & vbCrLf _
            & " from APTINVH1,APTPOST1" & vbCrLf _
            & " where APTPOST1.POST_CODE (+) = APTINVH1.POST_CODE" & vbCrLf _
            & "   and (APTINVH1.INV_STATUS = 'O' or APTINVH1.INV_STATUS = 'H')"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into GLTCREC4 Select 'V', VEND_CODE, VEND_NAME from APTVEND1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("")
    End Sub

    Sub Purge_Files()
        Get_PARM("ASTPARM1")

        Dim AS_PARM_ARCHIVE_FOLDER As String = ROWs("ASTPARM1").Item("AS_PARM_ARCHIVE_FOLDER") & "\"
        AS_PARM_ARCHIVE_FOLDER = ASCMAIN1.Folders("Archive")
        Dim AS_PARM_REPORTS_ARCHIVE_DAYS As Integer = Val(ROWs("ASTPARM1").Item("AS_PARM_REPORTS_ARCHIVE_DAYS") & "")

        Dim PURGE_DATE As String = Format$(Now.AddDays(-1 * AS_PARM_REPORTS_ARCHIVE_DAYS), "dd-MMM-yyyy")

        Dim sql As String = "Select * from ASTSPRF1 where REPORT_DATE <= '" & PURGE_DATE & "'"
        For Each rowASTSPRF1 As DataRow In ASCDATA1.GetDataTable(sql, "ASTSPRF1").Rows
            On Error Resume Next
            Dim FILENAME As String = AS_PARM_ARCHIVE_FOLDER _
                                   & "Reports\" _
                                   & rowASTSPRF1.Item("REPORT_NO") _
                                   & "." & ROWs("ASTPARM1").Item("AS_PARM_REPORTS_SFX")
            My.Computer.FileSystem.DeleteFile(FILENAME)
            On Error GoTo 0
            rowASTSPRF1.Delete()
        Next
    End Sub

    Private Sub CreateCommissionData(ByVal period As String)

        ASCMAIN1.sql = "INSERT INTO SOTCOMH1 SELECT '" & period & "', SREP_CODE, NVL(SREP_COMM_RATE, 0) FROM SOTSREP1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "INSERT INTO SOTCOMH4 SELECT '" & period & "', SREP_CODE, STYLE_GROUP_CODE, NVL(SREP_COMM_RATE, 0) FROM SOTSREP4"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "INSERT INTO SOTCOMH5 SELECT '" & period & "', SREP_CODE, CUST_CODE, NVL(SREP_COMM_RATE, 0), SREP_COMM_USE_STD FROM SOTSREP5"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "INSERT INTO SOTCOMH6 SELECT '" & period & "', SREP_CODE, CUST_CODE, STYLE_GROUP_CODE, NVL(SREP_COMM_RATE, 0) FROM SOTSREP6"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = " INSERT INTO SOTINVHS"
        ASCMAIN1.sql &= " ("
        ASCMAIN1.sql &= " SELECT DISTINCT INV_TYPE, INV_NO, SALES_DIVISION_CODE, SREP_CODE, SREP_COMM_RATE FROM"
        ASCMAIN1.sql &= " ("
        ASCMAIN1.sql &= " SELECT INV_TYPE, INV_NO, SALES_DIVISION_CODE"
        ASCMAIN1.sql &= " ,DECODE(SREP_CODE_SD, NULL, SREP_CODE, SREP_CODE_SD) SREP_CODE"
        ASCMAIN1.sql &= " ,DECODE(SREP_CODE_SD, NULL, NVL(SREP_COMM_RATE,0), NVL(SREP_COMM_RATE_SD,0)) SREP_COMM_RATE"
        ASCMAIN1.sql &= " FROM"
        ASCMAIN1.sql &= " ("
        ASCMAIN1.sql &= " select xx.inv_type, xx.inv_no, xx.cust_code, xx.srep_code, xx.SALES_DIVISION_CODE"
        ASCMAIN1.sql &= " , artsrep1.srep_code srep_code_SD "
        ASCMAIN1.sql &= " , sotsrep1.SREP_COMM_RATE "
        ASCMAIN1.sql &= " , sotsrep1SD.SREP_COMM_RATE SREP_COMM_RATE_SD"
        ASCMAIN1.sql &= " from"
        ASCMAIN1.sql &= " (select sotinvh1.inv_type, sotinvh1.inv_no, sotinvh1.cust_code, sotinvh1.srep_code, ictstyl1.SALES_DIVISION_CODE"
        ASCMAIN1.sql &= " from sotinvh1, sotinvh2, ictstyl1"
        ASCMAIN1.sql &= " where sotinvh1.inv_type = sotinvh2.inv_type"
        ASCMAIN1.sql &= " and sotinvh1.inv_no = sotinvh2.inv_no"
        ASCMAIN1.sql &= " and sotinvh2.STYLE_CODE = ictstyl1.STYLE_CODE (+)"
        ASCMAIN1.sql &= " and sotinvh1.ORDR_YYYYPP_UPDATED = '" & period & "'"
        ASCMAIN1.sql &= " ) XX, artsrep1, sotsrep1 , sotsrep1 sotsrep1SD"
        ASCMAIN1.sql &= " where artsrep1.CUST_CODE (+) = xx.cust_code "
        ASCMAIN1.sql &= " and artsrep1.SALES_DIVISION_CODE (+) = xx.SALES_DIVISION_CODE"
        ASCMAIN1.sql &= " AND XX.SREP_CODE = sotsrep1.SREP_CODE (+)"
        ASCMAIN1.sql &= " AND artsrep1.SREP_CODE = sotsrep1SD.SREP_CODE (+)"
        ASCMAIN1.sql &= " 	)"
        ASCMAIN1.sql &= " 	) WHERE INV_TYPE IS NOT NULL AND INV_NO IS NOT NULL AND SALES_DIVISION_CODE IS NOT NULL AND SREP_CODE IS NOT NULL"
        ASCMAIN1.sql &= " 	)"
        ASCDATA1.ExecuteSQL()

    End Sub

End Class