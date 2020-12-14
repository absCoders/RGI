Public Class ARRATBR1
    Dim ARTOPEN1 As String
    Dim ARTOPEN1_SUMMARY As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Range_Events(grpINV_DATE_RANGE)

        Get_PARM("GLTPARM1")
        Get_PARM("ARTPARM1")

        Absx1.dteFor("AGING_DATE").Value = (Now + ASCMAIN1.NowTSD).Date

        Set_DAYS()

        ASCMAIN1.Add_Value_List(Absx1.cbeFor("CBESORT"), "", Nothing, New String() {":", "INV_NUM:Invoice No", "INV_BALANCE:Balance (Desc)", "INV_CUST_PO:Customer PO", "REASON_CODE:Reason"})
        Absx1.cbeFor("CBESORT").SelectedIndex = 0

        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, -1, -1)

        chkBOL.Visible = (ASCMAIN1.CLIENT = "VAN")
        If ASCMAIN1.CLIENT = "INT" Then grpFX.Visible = False
        If ASCMAIN1.CLIENT = "VAN" Then
            Absx1.chkFor("INV_TYPE_R").Visible = True
            Absx1.optFor("OPTID").Value = "D"
        End If

        If ASCMAIN1.CLIENT = "RGI" Then grpSREP.Visible = True

    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("CUST_CODE")
                If Absx1.optFor("OPTDS").Value = "D" And Val(rowASTDSQLA("SEQUENCE") & "") = 0 Then
                    Dim rowASTDSQLA2 As DataRow = tblASTDSQLA.Rows.Find("CUST_CODE_SO")
                    If Absx1.optFor("OPTDS").Value = "D" And (rowASTDSQLA2 Is Nothing OrElse Val(rowASTDSQLA2("SEQUENCE") & "") = 0) Then
                        EMsg &= vbCr & "You Must include Customer (Bill-To or Sold-To) in the Sort when reporting AR Item Details"
                    End If
                End If

                If Absx1.dteFor("AGING_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "You Must Select an Aging Date"
                End If

                If Get_INV_TYPEs(False) = "" Then
                    EMsg &= vbCr & "You Must Select at Least 1 A/R Item Type"
                End If

                If Val(Absx1.numFor("DAYS1").Value & "") >= Val(Absx1.numFor("DAYS2").Value & "") + 1 _
                Or Val(Absx1.numFor("DAYS2").Value & "") >= Val(Absx1.numFor("DAYS3").Value & "") + 1 _
                Or Val(Absx1.numFor("DAYS3").Value & "") >= Val(Absx1.numFor("DAYS4").Value & "") + 1 Then
                    EMsg &= vbCr & "Please Verify Values in Aging Days Buckets"
                End If

                If Absx1.chkFor("NON_USD").Checked Then
                    If Absx1.txtFor("CURR_CODE").Text = "" Then
                        EMsg &= vbCr & "A currency code must be selected when the Non USD Items Only checkbox is checked"
                    End If
                End If

                If EMsg = "" Then

                    For i As Integer = 1 To 4
                        If Absx1.optFor("OPTID").Value = "I" Then
                            TAC.ARCMAIN1.AGE_DAYS(i) = Val(Absx1.numFor("DAYS" & CStr(i)).Value & "")
                        Else
                            TAC.ARCMAIN1.AGE_DAYS(i) = Val(ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_" & CStr(i)) & "")
                        End If
                        If Absx1.optFor("OPTID").Value = "D" Then
                            TAC.ARCMAIN1.DUE_DAYS(i) = Val(Absx1.numFor("DAYS" & CStr(i)).Value & "")
                        Else
                            TAC.ARCMAIN1.DUE_DAYS(i) = Val(ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_" & CStr(i)) & "")
                        End If
                    Next
                End If
        End Select
    End Sub

    Public Overrides Sub Retrieve_Settings_Post()
        ' MyBase.Retrieve_Settings_Post()

        Absx1.dteFor("AGING_DATE").Value = (Now + ASCMAIN1.NowTSD).Date

    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Prepare Working Variables

        ASCMAIN1.Progress("Building Open AR Work File")

        Dim ARTOPENI As String = ""
        If optARITEMS.Value = "C" Then
            Page0.Add("Using Currently Open AR Items")
            ARTOPENI = "ARTOPEN1"
        Else
            '  Dim RYP As String = Absx1.cmbFor("RYP").Value

            Page0.Add("Using AR Items open at End of " & Absx1.cmbFor("RYP").Text)
            ARTOPENI = ASCMAIN1.Temp_Table("Select * from ARTOPEN1 where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & ARTOPENI & " Add Primary Key (CUST_CODE,INV_TYPE,INV_NUM)")

            'ASCMAIN1.sql = "Select * from ARTOPEN1 where (CUST_CODE,INV_TYPE,INV_NUM) in " _
            '    & " (Select DETL_CVX_NO,DETL_CTL_TYPE,DETL_CTL_NO from GLTCREC3" _
            '    & " where CREC_TYPE_CODE = 'AR' and OPS_YYYYPP = '" & RYP & "')"
            'ASCDATA1.ExecuteSQL("Insert into " & ARTOPENI & " " & ASCMAIN1.sql)
            'ASCDATA1.ExecuteSQL("Insert into " & ARTOPENI & " " & Replace(ASCMAIN1.sql, "ARTOPEN1", "ARTOPENX"))

            Dim SQL_ARTOPEN1 As String = "Select * from ARTOPEN1 where (CUST_CODE,INV_TYPE,INV_NUM) in " _
                & " (Select DETL_CVX_NO,DETL_CTL_TYPE,DETL_CTL_NO from GLTCREC3" _
                & " where CREC_TYPE_CODE = 'AR' and OPS_YYYYPP = '" & RYP & "')"
            ASCDATA1.ExecuteSQL("Insert into " & ARTOPENI & " " & SQL_ARTOPEN1)

            ASCMAIN1.sql = "Delete from " & ARTOPENI & " where (CUST_CODE,INV_TYPE,INV_NUM) in " _
                & " (Select CUST_CODE,INV_TYPE,INV_NUM from ARTOPENX where (CUST_CODE,INV_TYPE,INV_NUM) in " _
                & " (Select DETL_CVX_NO,DETL_CTL_TYPE,DETL_CTL_NO from GLTCREC3" _
                & " where CREC_TYPE_CODE = 'AR' and OPS_YYYYPP = '" & RYP & "'))"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCDATA1.ExecuteSQL("Insert into " & ARTOPENI & " " & Replace(SQL_ARTOPEN1, "ARTOPEN1", "ARTOPENX"))

            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & "  Select DETL_CVX_NO CUST_CODE, DETL_CTL_TYPE INV_TYPE, DETL_CTL_NO INV_NUM, CREC_AMT from GLTCREC3" & vbCrLf _
                & "   where CREC_TYPE_CODE = 'AR' and OPS_YYYYPP = '" & RYP & "';" & vbCrLf _
                & " Begin " & vbCrLf _
                & "   For R1 in C1 Loop" & vbCrLf _
                & "    Update " & ARTOPENI & " Set INV_BALANCE = R1.CREC_AMT, INV_BALANCE_CURR = R1.CREC_AMT" & vbCrLf _
                & "     where CUST_CODE = R1.CUST_CODE and INV_TYPE = R1.INV_TYPE and INV_NUM = R1.INV_NUM;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        Dim sql As String

        sql = "Select ARTOPEN1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.CUST_PHONE, ARTCUST1.CUST_EXT" _
        & ", ARTCUST1.CUST_CONTACT, ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.TERM_CODE, TATTERM1.TERM_DESC, ARTCUST1.TRADE_CLASS_CODE, ARTCUST1.CUST_CRED_LIMIT_REV" _
        & ", ARTCUST6.CUST_LAST_PMT_DATE, ARTCUST6.CUST_LAST_PMT_AMT, ARTCUST6.CUST_LAST_INV_DATE, ARTCUST6.CUST_LAST_INV_AMT" _
        & ", SUM (ARTOPEN1.INV_BALANCE) INV_BALANCE " _
        & ", SUM (DECODE(ARTOPEN1.INV_TYPE,'I',ARTOPEN1.INV_BALANCE,0)) INV_BALANCE_I " _
        & ", SUM (DECODE(ARTOPEN1.INV_TYPE,'B',ARTOPEN1.INV_BALANCE,0)) INV_BALANCE_B " _
        & ", SUM (DECODE(ARTOPEN1.INV_TYPE,'D',ARTOPEN1.INV_BALANCE,0)) INV_BALANCE_D " _
        & ", SUM (DECODE(ARTOPEN1.INV_TYPE,'R',ARTOPEN1.INV_BALANCE,0)) INV_BALANCE_R " _
        & ", SUM (DECODE(ARTOPEN1.INV_TYPE,'C',ARTOPEN1.INV_BALANCE,0)) INV_BALANCE_C " _
        & ", SUM (DECODE(ARTOPEN1.INV_TYPE,'O',ARTOPEN1.INV_BALANCE,0)) INV_BALANCE_O " _
        & " from " & ARTOPENI & " ARTOPEN1, ARTCUST1, ARTCUST6, TATTERM1 " _
        & " where ARTOPEN1.CUST_CODE = ARTCUST1.CUST_CODE" _
        & "   and TATTERM1.TERM_CODE (+) = ARTCUST1.TERM_CODE" _
        & " and ARTOPEN1.CUST_CODE = ARTCUST6.CUST_CODE (+)" _
        & " and ARTOPEN1.INV_BALANCE <> 0" _
        & IIf(ASCMAIN1.CLIENT = "VAN-not no more", "", " and ARTOPEN1.OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'") _
        & " group by ARTOPEN1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.CUST_PHONE, ARTCUST1.CUST_EXT" _
        & ", ARTCUST1.CUST_CONTACT, ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.TERM_CODE, TATTERM1.TERM_DESC, ARTCUST1.TRADE_CLASS_CODE, ARTCUST1.CUST_CRED_LIMIT_REV" _
        & ", ARTCUST6.CUST_LAST_PMT_DATE, ARTCUST6.CUST_LAST_PMT_AMT, ARTCUST6.CUST_LAST_INV_DATE, ARTCUST6.CUST_LAST_INV_AMT"

        Dim ARTCUSTX As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUSTX & " Add Primary Key (CUST_CODE)")


        'TAC.ARCMAIN1.Get_Aging_Data(ROWs("ARTPARM1"), Now.Date)
        TAC.ARCMAIN1.Get_Aging_Data( _
        ROWs("ARTPARM1"), _
        Absx1.dteFor("AGING_DATE").Value, False)

        Page0.Add("Age using Days from Base Date " & Format(Absx1.dteFor("AGING_DATE").Value, "MM/dd/yyyy"))

        sql = "Select ARTOPEN1.* " & TAC.ARCMAIN1.DAYS_AND_BUCKETS

        Dim sql_age_options As String = ""
        If Absx1.chkFor("AGE_CHARGEBACKS").Checked Then
            Page0.Add("Chargebacks ARE reflected in the Aging Columns")
        Else
            Page0.Add("Chargebacks are NOT reflected in the Aging Columns")
            sql_age_options &= " AND INV_TYPE <> 'B'"
        End If
        If Absx1.chkFor("AGE_CREDITS").Checked Then
            Page0.Add("Credits ARE reflected in the Aging Columns")
        Else
            Page0.Add("Credits are NOT reflected in the Aging Columns")
            sql_age_options &= " AND INV_TYPE <> 'C' AND INV_TYPE <> 'O' AND INV_TYPE <> 'R'"
        End If

        sql = sql _
        & ", DECODE (ARTOPEN1.INV_TYPE,'B',ARTOPEN1.INV_BALANCE,0) CHARGEBACKS " & vbCrLf _
        & ", CASE WHEN ARTOPEN1.INV_TYPE = 'C' OR ARTOPEN1.INV_TYPE = 'O' OR ARTOPEN1.INV_TYPE = 'R' THEN ARTOPEN1.INV_BALANCE ELSE 0 END CREDITS" & vbCrLf _
        & " from " & ARTOPENI & " ARTOPEN1, " & ARTCUSTX & " ARTCUSTX" & vbCrLf

        sql = sql & " where ARTOPEN1.INV_BALANCE <> 0" & vbCrLf
        sql = sql & " and ARTCUSTX.CUST_CODE = ARTOPEN1.CUST_CODE" & vbCrLf

        Select Case Absx1.optFor("OPTBALANCE").Value
            Case "M"
                sql = sql & " and ARTCUSTX.INV_BALANCE > " & CStr(Absx1.numFor("BALANCE").Value)
                Page0.Add("Customers with Balance > " & CStr(Absx1.numFor("BALANCE").Value))
            Case "L"
                sql = sql & " and ARTCUSTX.INV_BALANCE < " & CStr(Absx1.numFor("BALANCE").Value)
                Page0.Add("Customers with Balance < " & CStr(Absx1.numFor("BALANCE").Value))
        End Select

        If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
            Dim z As String = Format(Absx1.dteFor("INV_DATE_F").Value, "dd-MMM-yyyy")
            sql = sql & " and ARTOPEN1.INV_DATE >= '" & z & "'"
            Page0.Add("Invoices dated >= " & z)
        End If
        If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
            Dim z As String = Format(Absx1.dteFor("INV_DATE_L").Value, "dd-MMM-yyyy")
            sql = sql & " and ARTOPEN1.INV_DATE <= '" & z & "'"
            Page0.Add("Invoices dated <= " & z)
        End If

        Dim not_all_types_selected As Boolean
        Dim INV_TYPEs As String = Get_INV_TYPEs(not_all_types_selected)
        If not_all_types_selected Then
            sql = sql & " and ARTOPEN1.INV_TYPE in (" & Mid(INV_TYPEs, 2) & ")"
        End If

        If ASCMAIN1.CLIENT = "VAN-not no more" Then
        Else
            sql = sql & " and ARTOPEN1.OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
        End If
 

        ' Prepare filters from Run-Time Options
        Dim sql_filter As String = ""



        If Absx1.txtFor("CURR_CODE").Text <> "" Then
            sql_filter = " AND CURR_CODE = '" & Absx1.txtFor("CURR_CODE").Text & "'"
        End If

        ' Extracts from Data Sources
        MyBase.Get_SQL("*")
        sql = Replace(sql, " from " & ARTOPENI & " ARTOPEN1", " from " & ARTOPENI & " ARTOPEN1" & sql_TABLE_NAMEs)
        sql = sql & sql_WHERE & sql_JOIN & sql_filter & vbCr

        ARTOPEN1 = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add Primary Key (CUST_CODE, INV_TYPE, INV_NUM)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add SORT_BY VARCHAR2(20)")


        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add MASTER_BILL_OF_LADING_NO VARCHAR2(20)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add MASTER_SHIP_BOL_NO VARCHAR2(20)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add BILL_OF_LADING_NO VARCHAR2(20)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add SHIP_BOL_NO VARCHAR2(10)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ARTOPEN1 & "_BOL on " & ARTOPEN1 & " (INV_TYPE,INV_NUM)")

        If chkBOL.Checked Then
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select SOTINVH1.INV_NO, SOTINVH1.SHIP_BOL_NO, SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.MASTER_SHIP_BOL_NO, SOTSHIP1.MASTER_BILL_OF_LADING_NO " & vbCrLf _
                & "   from SOTINVH1,SOTSHIP1" & vbCrLf _
                & "   where SOTINVH1.INV_NO in (Select INV_NO from " & ARTOPEN1 & " where INV_TYPE = 'I')" & vbCrLf _
                & "     and SOTSHIP1.SHIP_BOL_NO = SOTINVH1.SHIP_BOL_NO;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & ARTOPEN1 & vbCrLf _
                & "    Set SHIP_BOL_NO = R1.SHIP_BOL_NO" & vbCrLf _
                & "      , BILL_OF_LADING_NO = R1.BILL_OF_LADING_NO, MASTER_SHIP_BOL_NO = R1.MASTER_SHIP_BOL_NO, MASTER_BILL_OF_LADING_NO = R1.MASTER_BILL_OF_LADING_NO" & vbCrLf _
                & "    where INV_TYPE = 'I' and INV_NUM = R1.INV_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If


        Dim SORT_BY As String = Absx1.cbeFor("CBESORT").Value
        If SORT_BY = "INV_BALANCE" Then
            SORT_BY = "TRIM(TO_CHAR(99999999999999 - 100 * TRUNC(NVL(INV_BALANCE,0)),'00000000000000000000'))"
        End If
        ASCDATA1.ExecuteSQL("Update " & ARTOPEN1 & " Set SORT_BY = " & SORT_BY)

        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add INV_BALANCE_1 NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add INV_BALANCE_2 NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add INV_BALANCE_3 NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add INV_BALANCE_4 NUMBER (13,2)")

        If Absx1.chkFor("NON_USD").Checked Then
            ASCMAIN1.sql = "Update " & ARTOPEN1 & " Set INV_BALANCE = INV_BALANCE_CURR, INV_TOTAL_AMOUNT = INV_TOTAL_AMOUNT_CURR"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Update " & ARTOPEN1 & " Set INV_BALANCE_1 = INV_BALANCE where " & IIf(optID.Value = "I", "AGE_BUCKET", "DUE_BUCKET") & " = 1" & sql_age_options
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ARTOPEN1 & " Set INV_BALANCE_2 = INV_BALANCE where " & IIf(optID.Value = "I", "AGE_BUCKET", "DUE_BUCKET") & " = 2" & sql_age_options
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ARTOPEN1 & " Set INV_BALANCE_3 = INV_BALANCE where " & IIf(optID.Value = "I", "AGE_BUCKET", "DUE_BUCKET") & " = 3" & sql_age_options
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ARTOPEN1 & " Set INV_BALANCE_4 = INV_BALANCE where " & IIf(optID.Value = "I", "AGE_BUCKET", "DUE_BUCKET") & " = 4" & sql_age_options
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUSTX & " Add INV_BALANCE_1 NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUSTX & " Add INV_BALANCE_2 NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUSTX & " Add INV_BALANCE_3 NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUSTX & " Add INV_BALANCE_4 NUMBER (13,2)")

        ASCMAIN1.sql = "Begin Declare Cursor C1 is Select CUST_CODE" _
        & ", SUM (INV_BALANCE_1) INV_BALANCE_1, SUM (INV_BALANCE_2) INV_BALANCE_2" _
        & ", SUM (INV_BALANCE_3) INV_BALANCE_3, SUM (INV_BALANCE_4) INV_BALANCE_4" _
        & " from " & ARTOPEN1 & " GROUP BY CUST_CODE;" _
        & " Begin For R1 in C1 Loop" _
        & "  Update " & ARTCUSTX & " Set " _
        & "  INV_BALANCE_1 = R1.INV_BALANCE_1" _
        & ", INV_BALANCE_2 = R1.INV_BALANCE_2" _
        & ", INV_BALANCE_3 = R1.INV_BALANCE_3" _
        & ", INV_BALANCE_4 = R1.INV_BALANCE_4" _
        & " where CUST_CODE = R1.CUST_CODE;" _
        & " End Loop; End; End;"

        ASCDATA1.ExecuteSQL()

        If Absx1.optFor("OPTAP").Value = "P" Then 'Past Due Customers Only
            Dim sqld As String = ""
            For i As Integer = Val(Absx1.optFor("OPTPAST_DUE_SEL").Value & "") To 4
                sqld &= " and NVL(INV_BALANCE_" & CStr(i) & ",0) = 0"
            Next
            ASCMAIN1.sql = "Delete from " & ARTOPEN1 _
            & " where CUST_CODE in (Select CUST_CODE from " & ARTCUSTX & ASCMAIN1.SQL_Add_WHERE(sqld) & ")"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Delete from " & ARTCUSTX & ASCMAIN1.SQL_Add_WHERE(sqld)
            ASCDATA1.ExecuteSQL()
            Page0.Add("Past Due Customers Only (at least " & Absx1.optFor("OPTPAST_DUE_SEL").Text & ")")
        End If

        ASCMAIN1.sql = "Select * from " & ARTCUSTX
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTCUSTX", 1))

        ASCMAIN1.sql = "Select * from TATTERM1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "TATTERM1", 1))


        If Absx1.optFor("OPTDS").Value = "D" Then 'DETAIL
            ASCMAIN1.sql = "Select * from " & ARTOPEN1
        Else 'SUMMARY
            sql = "Select * from " & ARTOPEN1 & " where ROWNUM < 1"
            ARTOPEN1_SUMMARY = ASCMAIN1.Temp_Table(sql)

            sql = "Insert into " & ARTOPEN1_SUMMARY & vbCr _
            & " (CUST_CODE, SREP_CODE, POST_CODE, REASON_CODE, INV_TYPE, INV_NUM, CUST_CODE_SO, CURR_CODE, CURR_EXCH_RATE" & vbCrLf _
            & ", INV_TOTAL_AMOUNT, INV_TOTAL_AMOUNT_CURR, INV_BALANCE_CURR" & vbCrLf _
            & ", INV_BALANCE, INV_BALANCE_1 " & vbCr _
            & ", INV_BALANCE_2, INV_BALANCE_3, INV_BALANCE_4, CREDITS, CHARGEBACKS) " & vbCrLf _
            & "Select ARTOPEN1.CUST_CODE, ARTOPEN1.SREP_CODE, ARTOPEN1.POST_CODE, ARTOPEN1.REASON_CODE" & vbCrLf _
            & ", ARTOPEN1.INV_TYPE" & vbCrLf _
            & ", 'S' INV_NUM" & vbCrLf _
            & ", ARTOPEN1.CUST_CODE, 'USD' CURR_CODE" & vbCrLf _
            & ", 1 CURR_EXCH_RATE" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_TOTAL_AMOUNT) INV_TOTAL_AMOUNT" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_TOTAL_AMOUNT_CURR) INV_TOTAL_AMOUNT_CURR" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_BALANCE_CURR) INV_BALANCE_CURR" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_BALANCE) INV_BALANCE" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_BALANCE_1) INV_BALANCE_1" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_BALANCE_2) INV_BALANCE_2" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_BALANCE_3) INV_BALANCE_3" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_BALANCE_4) INV_BALANCE_4" & vbCrLf _
            & ", SUM(ARTOPEN1.CREDITS) CREDITS" & vbCrLf _
            & ", SUM(ARTOPEN1.CHARGEBACKS) CHARGEBACKS" & vbCrLf _
            & " from " & ARTOPEN1 & " ARTOPEN1"

            sql &= " group by ARTOPEN1.CUST_CODE, ARTOPEN1.SREP_CODE, ARTOPEN1.POST_CODE, ARTOPEN1.REASON_CODE, ARTOPEN1.INV_TYPE"

            ASCDATA1.ExecuteSQL(sql)
            ASCDATA1.ExecuteSQL("Update " & ARTOPEN1_SUMMARY & " SET INV_NUM = ROWNUM")
            ASCMAIN1.sql = "Select * from " & ARTOPEN1_SUMMARY

        End If

        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTATBR1", 3))

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ", ARTOPEN1.CUST_CODE" & vbCrLf _
            & ", ARTOPEN1.INV_TYPE" & vbCrLf _
            & ", ARTOPEN1.INV_NUM" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE_1" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE_2" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE_3" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE_4" & vbCrLf _
            & ", ARTOPEN1.CREDITS" & vbCrLf _
            & ", ARTOPEN1.CHARGEBACKS" & vbCrLf

        If Absx1.optFor("OPTDS").Value = "D" Then 'DETAIL
            sql = sql & " from " & ARTOPEN1 & " ARTOPEN1 " & sql_TABLE_NAMEs & vbCrLf
        Else
            sql = sql & " from " & ARTOPEN1_SUMMARY & " ARTOPEN1 " & sql_TABLE_NAMEs & vbCrLf
        End If

        sql = sql & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

    End Sub

    Public Overrides Sub Print_Report()
        If ASCMAIN1.CLIENT = "RGI" And chkSREP_SEPERATE.Checked Then
            Print_Report_RGI()
        Else
            Dim i As Integer
            Dim z As String

            Dim SUBT As String = ""

            If Absx1.optFor("OPTID").Value = "I" Then
                SUBT &= " Aged by Invoice Date"
            Else
                SUBT &= " Aged by Due Date"
            End If
            If Absx1.optFor("OPTAP").Value = "P" Then
                SUBT = SUBT & ", Past Due Customers"
            End If

            Dim AGING_DATE As Date = Absx1.dteFor("AGING_DATE").Value
            SUBT &= ", Age Date " & Format(AGING_DATE, "MM/dd/yyyy")

            If optARITEMS.Value = "P" Then
                SUBT &= ", Balance as of " & Absx1.cmbFor("RYP").Text
            Else
                SUBT &= ", as of " & Format(System.DateTime.Now, "MM/dd/yyyy")
            End If

            If Not Absx1.chkFor("CHKINV_DATE_F").Checked _
        Or Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
                SUBT = SUBT & ", "
                SUBT = SUBT & "Showing A/R Items Dated"
                If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
                    SUBT = SUBT & " from " & Format(Absx1.dteFor("INV_DATE_F").Value, "MM/dd/yyyy")
                End If
                If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
                    SUBT = SUBT & " thru " & Format(Absx1.dteFor("INV_DATE_L").Value, "MM/dd/yyyy")
                End If
            End If
            Select Case Absx1.optFor("OPTBALANCE").Value
                Case "M"
                    SUBT = SUBT & ", Balances > " & Format(Absx1.numFor("BALANCE").Value, "$0")
                Case "L"
                    SUBT = SUBT & ", Balances < " & Format(Absx1.numFor("BALANCE").Value, "$0")
            End Select
            If Absx1.chkFor("NON_USD").Checked Then
                '    SUBT = SUBT & ", Non USD Currency Code Items"
                'End If
                'If Absx1.txtFor("CURR_CODE").Text <> "" Then
                SUBT = SUBT & ", " & Absx1.txtFor("CURR_CODE").Value & " Items Only"
            End If

            For i = 1 To 4
                z = "DATE" & Format$(i, "0")
                CR_params.Add(z, Absx1.txtFor("LBL_DAYS" & CStr(i)).Text)
            Next i

            CR_params.Add("DATE0", Format(AGING_DATE, "MM/dd"))

            If Absx1.optFor("OPTID").Value = "I" Then
                z = "Days Old"
            Else
                z = "Past Due"
            End If
            CR_params.Add("AGEBY", z)
            CR_params.Add("DTL", Absx1.optFor("OPTDS").Value)
            CR_params.Add("NO_CENTS", "0")
            CR_params.Add("NOTES", "1")
            CR_params.Add("SHOW_BOL", IIf(chkBOL.Checked, "1", "0"))

            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("CUST_CODE")
            If Val(rowASTDSQLA("SEQUENCE") & "") = 0 Then
                CR_params.Add("CUST_SORT", "0")
            Else
                CR_params.Add("CUST_SORT", "1")
            End If

            Generate_Report(RPT, , SUBT)

            If ASCMAIN1.CLIENT = "NYA" And Not ASCMAIN1.Running_in_VS Then
                ' skip for NYA - 
            Else
                Prepare_Data_Extracts()
            End If
        End If
    End Sub

    Public Sub Print_Report_RGI()

        Dim Report_RGI As String = "ARRATBR3"
        Dim SREP_CODES_RPT As New List(Of String)
        Dim MultiRep As Boolean = False
        If chkSREP_SEPERATE.Checked Then
            MultiRep = True
            For Each rowARRATBR1 As DataRow In dst.Tables("ARTATBR1").Select("", "SREP_CODE")
                If Not SREP_CODES_RPT.Contains(rowARRATBR1.Item("SREP_CODE").ToString) Then
                    SREP_CODES_RPT.Add(rowARRATBR1.Item("SREP_CODE").ToString)
                End If
            Next
        Else
            SREP_CODES_RPT.Add("")
        End If

        CR_params.Add("SREP_CODE", "")
        For Each SREP_CODE As String In SREP_CODES_RPT
            CR_params("SREP_CODE") = SREP_CODE

            Dim i As Integer
            Dim z As String

            Dim SUBT As String = ""

            If Absx1.optFor("OPTID").Value = "I" Then
                SUBT &= " Aged by Invoice Date"
            Else
                SUBT &= " Aged by Due Date"
            End If
            If Absx1.optFor("OPTAP").Value = "P" Then
                SUBT = SUBT & ", Past Due Customers"
            End If

            Dim AGING_DATE As Date = Absx1.dteFor("AGING_DATE").Value
            SUBT &= ", Age Date " & Format(AGING_DATE, "MM/dd/yyyy")

            If optARITEMS.Value = "P" Then
                SUBT &= ", Balance as of " & Absx1.cmbFor("RYP").Text
            Else
                SUBT &= ", as of " & Format(System.DateTime.Now, "MM/dd/yyyy")
            End If

            If Not Absx1.chkFor("CHKINV_DATE_F").Checked _
        Or Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
                SUBT = SUBT & ", "
                SUBT = SUBT & "Showing A/R Items Dated"
                If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
                    SUBT = SUBT & " from " & Format(Absx1.dteFor("INV_DATE_F").Value, "MM/dd/yyyy")
                End If
                If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
                    SUBT = SUBT & " thru " & Format(Absx1.dteFor("INV_DATE_L").Value, "MM/dd/yyyy")
                End If
            End If
            Select Case Absx1.optFor("OPTBALANCE").Value
                Case "M"
                    SUBT = SUBT & ", Balances > " & Format(Absx1.numFor("BALANCE").Value, "$0")
                Case "L"
                    SUBT = SUBT & ", Balances < " & Format(Absx1.numFor("BALANCE").Value, "$0")
            End Select
            If Absx1.chkFor("NON_USD").Checked Then
                SUBT = SUBT & ", " & Absx1.txtFor("CURR_CODE").Value & " Items Only"
            End If

            For i = 1 To 4
                z = "DATE" & Format$(i, "0")
                CR_params.Add(z, Absx1.txtFor("LBL_DAYS" & CStr(i)).Text)
            Next i

            CR_params.Add("DATE0", Format(AGING_DATE, "MM/dd"))

            If Absx1.optFor("OPTID").Value = "I" Then
                z = "Days Old"
            Else
                z = "Past Due"
            End If
            CR_params.Add("AGEBY", z)
            CR_params.Add("DTL", Absx1.optFor("OPTDS").Value)
            CR_params.Add("NO_CENTS", "0")
            CR_params.Add("NOTES", "1")
            CR_params.Add("SHOW_BOL", IIf(chkBOL.Checked, "1", "0"))

            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("CUST_CODE")
            If Val(rowASTDSQLA("SEQUENCE") & "") = 0 Then
                CR_params.Add("CUST_SORT", "0")
            Else
                CR_params.Add("CUST_SORT", "1")
            End If

            If SREP_CODE <> "" Then
                SUBT = SUBT & " Generated For Sales Rep " & SREP_CODE
            End If

            If chkSREP_EMAIL.Checked Then
                Dim tempFileName As String = ASCMAIN1.Next_Control_No("ARRATBR1.TEMP_REPORT_NO")
                Dim RPT_NO As String = Generate_Report(Report_RGI, "Open AR Items - " & SREP_CODE, SUBT,, "PDF", tempFileName)
                emailReport(tempFileName, SREP_CODE)
            Else
                Generate_Report(Report_RGI, "Open AR Items - " & SREP_CODE, SUBT)
            End If
        Next
    End Sub

    Private Sub emailReport(ByVal tempFileName As String, ByVal SREP_CODE As String)
        Dim rowSOTSREP1 As DataRow = LookUp("SOTSREP1", SREP_CODE)
        Dim SREP_problem As Boolean = False
        If IsNothing(rowSOTSREP1) Then
            SREP_problem = True
        Else
            If rowSOTSREP1.Item("SREP_EMAIL").ToString & "" = "" Then
                SREP_problem = True
            Else
                If rowSOTSREP1.Item("SREP_NAME").ToString & "" = "" Then
                    SREP_problem = True
                End If
            End If
        End If

        If SREP_problem Then
            MsgBox("Problem With E-mail for Sales Rep " & SREP_CODE, vbExclamation, "No E-mail Sent")
        Else
            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(tempFileName & ".pdf", ASCMAIN1.Folders("Temp") & tempFileName & ".pdf")

            Dim SUBJECT As String = "Open AR Items Report"
            Dim PFX As String = ""

            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

            'EMAIL_ADDRESSs.Add("whr@waynerichmond.net", "Wayne Richmond")
            EMAIL_ADDRESSs.Add(rowSOTSREP1.Item("SREP_EMAIL").ToString & "", rowSOTSREP1.Item("SREP_NAME").ToString & "")

            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                        SUBJECT, "ARRATBR1", True, True)
            If SEND_NO <> "" Then
                'TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "QUOEML", "Quote Sheet emailed", SEND_NO)
            End If
        End If
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        With dst.Tables("ASTSRPT1").Columns
            .Add("TERM_CODE")
            .Add("REASON_CODE")
            .Add("SREP_CODE")
            .Add("SALES_DIVISION_CODE")
            .Add("POST_CODE")
            .Add("CUST_CODE_SO")
            .Add("INV_CUST_PO")
            .Add("BILL_OF_LADING_NO")
            .Add("INV_DATE", GetType(System.DateTime))
            .Add("INV_DUE_DATE", GetType(System.DateTime))
            .Add("INV_TOTAL_AMOUNT", GetType(System.Decimal))
            .Add("AGE", GetType(System.Int64))
            .Add("INV_NOTES")
        End With

        For Each row As DataRow In dst.Tables("ASTSRPT1").Select("")
            For I As Integer = 1 To COLUMN_NAMEs.Count
                Dim CODE_VALUE As String = row.Item("G" & CStr(I))
                row.Item("G" & CStr(I)) = Split(CODE_VALUE, ":")(1)
            Next
            Dim INV_TYPE As String = row.Item("INV_TYPE")
            Dim INV_NUM As String = row.Item("INV_NUM")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim rowD As DataRow = dst.Tables("ARTATBR1").Rows.Find(New String() {CUST_CODE, INV_TYPE, INV_NUM})
            row.Item("TERM_CODE") = rowD.Item("TERM_CODE")
            row.Item("REASON_CODE") = rowD.Item("REASON_CODE")
            row.Item("SREP_CODE") = rowD.Item("SREP_CODE")
            row.Item("SALES_DIVISION_CODE") = rowD.Item("SALES_DIVISION_CODE")
            row.Item("POST_CODE") = rowD.Item("POST_CODE")
            row.Item("CUST_CODE_SO") = rowD.Item("CUST_CODE_SO")
            row.Item("INV_CUST_PO") = rowD.Item("INV_CUST_PO")
            row.Item("BILL_OF_LADING_NO") = rowD.Item("BILL_OF_LADING_NO")
            row.Item("INV_DATE") = rowD.Item("INV_DATE")
            row.Item("INV_DUE_DATE") = rowD.Item("INV_DUE_DATE")
            row.Item("INV_TOTAL_AMOUNT") = rowD.Item("INV_TOTAL_AMOUNT")
            row.Item("AGE") = rowD.Item("AGE")
            row.Item("INV_NOTES") = rowD.Item("INV_NOTES")
        Next

        grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")

        grdASTEXPT1.Text = "Open AR"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Dim Cs As New List(Of String)
        Dim G As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Cs.Add(COLUMN_NAME)
            G += 1
            Set_DX_Column(grdASTEXPT1, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , System.Drawing.Color.Silver)
        Next
        If Not Cs.Contains("CUST_CODE") Then Set_DX_Column(grdASTEXPT1, "CUST_CODE", "Bill-To", 100, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "INV_TYPE", "Type", 40, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "INV_NUM", "Doc No", 100, , , System.Drawing.Color.Gold)
        If Not Cs.Contains("TERM_CODE") Then Set_DX_Column(grdASTEXPT1, "TERM_CODE", "Terms", 60, , , System.Drawing.Color.Gold)
        If Not Cs.Contains("REASON_CODE") Then Set_DX_Column(grdASTEXPT1, "REASON_CODE", "Reason", 60, , , System.Drawing.Color.Gold)
        If Not Cs.Contains("SREP_CODE") Then Set_DX_Column(grdASTEXPT1, "SREP_CODE", "SRep", 60, , , System.Drawing.Color.Gold)
        If Not Cs.Contains("SALES_DIVISION_CODE") Then Set_DX_Column(grdASTEXPT1, "SALES_DIVISION_CODE", "SDiv", 60, , , System.Drawing.Color.Gold)
        If Not Cs.Contains("POST_CODE") Then Set_DX_Column(grdASTEXPT1, "POST_CODE", "Post", 60, , , System.Drawing.Color.Gold)
        If Not Cs.Contains("CUST_CODE_SO") Then Set_DX_Column(grdASTEXPT1, "CUST_CODE_SO", "Sold-To", 60, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "INV_CUST_PO", "Customer PO", 80, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "BILL_OF_LADING_NO", "ASN BOL No", 80, , , System.Drawing.Color.Gold)

        Set_DX_Column(grdASTEXPT1, "INV_DATE", "Date", 80, "MM/dd/yy", , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "INV_DUE_DATE", "Due", 80, "MM/dd/yy", , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "INV_TOTAL_AMOUNT", "Orig Amt", 100, "#,##0.00", "Sum", System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT1, "INV_BALANCE", "Balance", 100, "#,##0.00", "Sum", System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT1, "AGE", "Days", 50, "#,##0", , System.Drawing.Color.Orange)

        For i As Integer = 1 To 4
            Set_DX_Column(grdASTEXPT1, "INV_BALANCE_" & CStr(i), Absx1.txtFor("LBL_DAYS" & CStr(i)).Text, 100, "#,##0.00", "Sum", System.Drawing.Color.LightBlue)
        Next i

        Set_DX_Column(grdASTEXPT1, "CHARGEBACKS", "Chargebacks", 100, "#,##0.00", "Sum", System.Drawing.Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "CREDITS", "Credits", 100, "#,##0.00", "Sum", System.Drawing.Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "INV_NOTES", "Notes", 350, , , System.Drawing.Color.LightGreen)


        grdASTEXPT1.DisplayLayout.Bands(0).Columns("CUST_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "CUST_CODE")

    End Sub

    Private Sub optBALANCE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optBALANCE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Absx1.numFor("BALANCE").Enabled = (optBALANCE.Value = "M" Or optBALANCE.Value = "L")
        Absx1.numFor("BALANCE").Visible = (optBALANCE.Value = "M" Or optBALANCE.Value = "L")

        If Absx1.numFor("BALANCE").Visible Then
            Absx1.numFor("BALANCE").Top = optBALANCE.Top + IIf(optBALANCE.Value = "L", 15, 0)
        End If
    End Sub

    Function Get_INV_TYPEs(ByRef not_all_types_selected As Boolean) As String
        not_all_types_selected = False
        Dim INV_TYPEs As String = ""
        For Each INV_TYPE As String In New String() {"I", "B", "D", "R", "C", "O"}
            If Absx1.chkFor("INV_TYPE_" & INV_TYPE).Checked Then
                INV_TYPEs &= ",'" & INV_TYPE & "'"
            Else
                not_all_types_selected = True
            End If
        Next

        Return INV_TYPEs
    End Function

    Private Sub optID_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optID.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_DAYS()
    End Sub

    Sub Set_DAYS()
        For Each z As String In New String() {"1", "2", "3", "4"}
            If Absx1.optFor("OPTID").Value = "I" Then
                Absx1.numFor("DAYS" & z).Value = Val(ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_" & z) & "")
            Else
                Absx1.numFor("DAYS" & z).Value = Val(ROWs("ARTPARM1").Item("AR_PARM_DUE_CATG_" & z) & "")
            End If
        Next
    End Sub

    Sub Set_DAYS_Captions()
        If SELECTION_NO = 0 Then Exit Sub
        Dim DAYS(4) As Integer
        For i As Integer = 1 To 4
            DAYS(i) = Val(Absx1.numFor("DAYS" & CStr(i)).Value & "")
        Next
        For i As Integer = 1 To 4
            If i = 1 Then
                If optID.Value = "D" Then
                    Absx1.txtFor("LBL_DAYS" & CStr(i)).Text = "Current"
                Else
                    Absx1.txtFor("LBL_DAYS" & CStr(i)).Text = CStr(0) & " - " & CStr(DAYS(i + 1))
                End If
            ElseIf i = 4 Then
                Absx1.txtFor("LBL_DAYS" & CStr(i)).Text = "Over " & CStr(DAYS(i)) & IIf(optID.Value = "D", " PD", "")
            Else
                Absx1.txtFor("LBL_DAYS" & CStr(i)).Text = CStr(DAYS(i) + 1) & " - " & CStr(DAYS(i + 1))
            End If
        Next

        For Each vli As ValueListItem In Absx1.optFor("OPTPAST_DUE_SEL").ValueList.ValueListItems
            Dim i As Integer = vli.DataValue
            vli.DisplayText = Absx1.txtFor("LBL_DAYS" & CStr(i)).Text
        Next
    End Sub

    Private Sub numDAYS4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDAYS4.ValueChanged
        Set_DAYS_Captions()
    End Sub

    Private Sub numDAYS3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDAYS3.ValueChanged
        Set_DAYS_Captions()
    End Sub

    Private Sub numDAYS2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDAYS2.ValueChanged
        Set_DAYS_Captions()
    End Sub

    Private Sub optAP_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optAP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Absx1.optFor("OPTPAST_DUE_SEL").Visible = (optAP.Value = "P")
    End Sub

    Private Sub optARITEMS_ValueChanged(sender As Object, e As EventArgs) Handles optARITEMS.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Absx1.cmbFor("RYP").Visible = (optARITEMS.Value = "P")
        If optARITEMS.Value = "P" Then
            Set_Aging_Date()
        End If
    End Sub

    Sub Set_Aging_Date()
        Dim YP As String = Absx1.cmbFor("RYP").Value
        YP = Mid(YP, 1, 4) & Mid(YP, 6, 2)
        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
        Absx1.dteFor("AGING_DATE").Value = rowGLTPARM2.Item("PRD_END_DATE")
    End Sub

    Private Sub cmbRYP_ValueChanged(sender As Object, e As EventArgs) Handles cmbRYP.ValueChanged

    End Sub

    Private Sub UltraCheckEditor8_CheckedChanged(sender As Object, e As EventArgs) Handles UltraCheckEditor8.CheckedChanged

    End Sub

    Private Sub UltraTextEditor29_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor29.ValueChanged
        If Absx1.chkFor("NON_USD").Checked Then
            If Absx1.txtFor("CURR_CODE").Text = "USD" Then
                MsgBox("USD currency Cannot be selected when Non USD items ony checkbox is checked ", MsgBoxStyle.OkOnly, "USD Invalid Currency Selected")
                Absx1.txtFor("CURR_CODE").Text = ""

            End If
        End If
    End Sub

    Private Sub chkSREP_SEPERATE_CheckedChanged(sender As Object, e As EventArgs) Handles chkSREP_SEPERATE.CheckedChanged
        If chkSREP_SEPERATE.Checked Then
            chkSREP_EMAIL.Visible = True
            chkSREP_EMAIL.Checked = False
        Else
            chkSREP_EMAIL.Visible = False
            chkSREP_EMAIL.Checked = False
        End If
    End Sub
End Class