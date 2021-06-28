Public Class SOCMAIN1

    Public Shared AGING_DATES() As Date


    Public Shared Function UPC( _
        ByRef frmASFBASE0 As ASFBASE0, _
        ByVal UPC_SEQUENCE_NO As String, _
        ByRef SO_PARM_UPC_VENDOR_ID As String, _
        Optional ByVal prefix_with_VENDOR_ID As Boolean = True) As String

        ' Note: Check Digit Calculation applies to the 19-digits prior to the check digit
        '       These 11 digits are made up from the 6 digit Vendor ID prepended to the 5 digit UPC Serial Number
        '       19 digits = '0000' + 6 digit SO_PARM_UPC_VENDOR_ID + 9 digit Carton Serial Number

        Dim Check_Digit_Seed As String

        If prefix_with_VENDOR_ID Then
            If SO_PARM_UPC_VENDOR_ID = "" Then
                SO_PARM_UPC_VENDOR_ID = frmASFBASE0.ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & ""
            End If

            If Len(UPC_SEQUENCE_NO) <> 5 Then
                If Len(UPC_SEQUENCE_NO) <> 9 Then
                    Stop
                End If
            End If

            Check_Digit_Seed = Mid(SO_PARM_UPC_VENDOR_ID, 1) & UPC_SEQUENCE_NO
        Else
            Check_Digit_Seed = UPC_SEQUENCE_NO
        End If

        Dim odd_digits As Integer
        Dim even_digits As Integer

        For i As Integer = 1 To Len(Check_Digit_Seed) Step 2
            odd_digits = odd_digits + Val(Mid(Check_Digit_Seed, 1, 1))
            Check_Digit_Seed = Mid(Check_Digit_Seed, 2)
            If Check_Digit_Seed <> "" Then
                even_digits = even_digits + Val(Mid(Check_Digit_Seed, 1, 1))
                Check_Digit_Seed = Mid(Check_Digit_Seed, 2)
            End If
        Next i

        Dim check_digit As Integer
        check_digit = (odd_digits * 3 + even_digits) Mod 10
        If check_digit <> 0 Then
            check_digit = 10 - check_digit
        End If

        If prefix_with_VENDOR_ID Then
            UPC = SO_PARM_UPC_VENDOR_ID & UPC_SEQUENCE_NO & Format(check_digit, "0")
        Else
            UPC = UPC_SEQUENCE_NO & Format(check_digit, "0")
        End If

    End Function

    Public Shared Function Credit_Check( _
    ByVal rowSOTORDR1 As DataRow, _
    ByVal rowARTCREDC As DataRow, _
    ByVal ORDR_AMT_OPEN As Decimal, _
    ByVal rowSOTPARM1 As DataRow, _
    Optional ByVal ARTCREDC As String = "", _
    Optional ByVal sqlARTCREDC As String = "") As String

        ' REVIEW THESE CHANGES WITH ED
        ' USE AS A DEMO OF WHEN A CLASS SHOULD BE INSTANTIATED
        ' GET CLS OUT OF THIS COLLECTION IN TATTERM1

        'TERM_C TERM_DESC                          
        '------ -----------------------------------
        'C1     COD CERTIFI CHK                    
        'C2     Credit On Acct                     
        'C3     Credit Card                        
        'C4     COD COMPANY CHK                    
        'P9     WEB CREDIT CARD                    

        Dim TERM_CODE_C As New List(Of String)
        TERM_CODE_C.Add("C1")
        TERM_CODE_C.Add("C2")
        TERM_CODE_C.Add("C3")
        TERM_CODE_C.Add("C4")
        TERM_CODE_C.Add("P9")

        Dim credit_card_auth As Boolean = False
        Dim C1_C2_C3_P9 As Boolean = False
        Dim past_due As Boolean = False

        Dim CCPA_NO As String = rowSOTORDR1.Item("CCPA_NO") & String.Empty

        If rowSOTORDR1.Item("ORDR_HOLD_CREDIT_REL_BY") & String.Empty <> "" Then
            Return String.Empty ' Order was Released by Credit Dept
        End If
        If rowSOTORDR1.Item("ORDR_TYPE_CODE") & String.Empty = "B2C" Then
            Return String.Empty ' B2C Orders are pre-paid by CC
        End If

        Dim TERM_CODE As String = rowSOTORDR1.Item("TERM_CODE") & String.Empty

        If Val(rowSOTORDR1.Item("ORDR_COD_ADDON_AMT") & "") <> 0 _
        Or rowSOTORDR1.Item("ORDR_HOLD_CREDIT_SPECIAL") & "" = "1" _
        Then
            ' order must visit the credit release screen
        Else

            If TERM_CODE_C.Contains(TERM_CODE) And TERM_CODE <> "C4" Then ' COMPANY COD CHECK MUST STILL BE WITHIN CREDIT LIMITS AND CREDIT HOLD CRITERIA
                'Return String.Empty ' Terms are a CC Type or a COD Type - Order is Pre-paid prior to Shipment
                C1_C2_C3_P9 = True
            End If

            If rowSOTORDR1.Item("CCPA_NO") & "" <> "" Then ' If Order is to be paid by CC and the CC was Authorized, then let it go
                'Return String.Empty
                credit_card_auth = True
            End If
        End If

        If ARTCREDC <> "" Then

            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")

            ' SIMILAR CODE EXISTS IN SOFPICK1
            Dim SOTPICKC As String = "(" _
            & "Select SOTORDR1.CUST_CODE" _
            & ", Sum (NVL(SOTPICK2.PICK_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_PT" _
            & " from SOTPICK2,SOTPICK1,SOTORDR1,SOTORDR2" _
            & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO " _
            & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
            & "   and (SOTPICK1.PICK_STATUS = 'N' or SOTPICK1.PICK_STATUS = 'P')" _
            & "   and SOTORDR1.CUST_CODE = '" & CUST_CODE & "'" _
            & "   and SOTORDR1.ORDR_STATUS = 'P'" _
            & " group by SOTORDR1.CUST_CODE" _
            & ")"


            If AGING_DATES Is Nothing Then
                Initialize_AGING_DATES()
            End If

            ' SIMILAR CODE EXISTS IN SOFPICK1
            Dim ARTOPENA As String = "(" _
            & " Select CUST_CODE" _
            & ", SUM (CASE WHEN INV_DATE > '" & Format(AGING_DATES(1), "dd-MMM-yyyy") & "'                                 THEN INV_BALANCE ELSE 0 END) AGE_1" _
            & ", SUM (CASE WHEN INV_DATE > '" & Format(AGING_DATES(2), "dd-MMM-yyyy") & "' AND INV_DATE <= '" & Format(AGING_DATES(1), "dd-MMM-yyyy") & "' THEN INV_BALANCE ELSE 0 END) AGE_2" _
            & ", SUM (CASE WHEN INV_DATE > '" & Format(AGING_DATES(3), "dd-MMM-yyyy") & "' AND INV_DATE <= '" & Format(AGING_DATES(2), "dd-MMM-yyyy") & "' THEN INV_BALANCE ELSE 0 END) AGE_3" _
            & ", SUM (CASE WHEN INV_DATE                                                                <= '" & Format(AGING_DATES(3), "dd-MMM-yyyy") & "' THEN INV_BALANCE ELSE 0 END) AGE_4" _
            & ", SUM (CASE WHEN INV_DATE                                                                <= '" & Format(AGING_DATES(3), "dd-MMM-yyyy") & "' AND INV_TYPE IN ('I','D','B','O','C','R') THEN INV_BALANCE ELSE 0 END) PAST_DUE_DR_AMT" _
            & ", SUM (INV_BALANCE) TOTAL_DUE" _
            & " from ARTOPEN1 where INV_BALANCE <> 0" _
            & " and CUST_CODE = '" & CUST_CODE & "'" _
            & " group by CUST_CODE" _
            & ")"

            'Prepare_ARTCREDC(ARTCREDC, Replace(sqlARTCREDC, ":PARM1", "'" & CUST_CODE & "'"), SOTPICKC, "", ARTOPENA, "")
            'rowARTCREDC = ASCDATA1.GetDataRow("Select * from " & ARTCREDC)

            ASCMAIN1.sql = "Select Sum (ORDR_SALES) from SOTORDR1 " _
            & " where ORDR_STATUS = 'O' and CUST_CODE = '" & CUST_CODE & "'"
            ORDR_AMT_OPEN = Val(ASCDATA1.GetDataValue)

        End If

        Dim SO_PARM_CRH_MOS_INACTIVE As Int32 = Val(rowSOTPARM1.Item("SO_PARM_CRH_MOS_INACTIVE") & "")
        Dim SO_PARM_CR_LIMIT_GRACE As Int32 = Val(rowSOTPARM1.Item("SO_PARM_CR_LIMIT_GRACE") & "")

        Dim ORDR_HOLD_CREDIT_REASON As String = ""

        ' CRH - Customer Master indicates a Credit Hold
        If rowARTCREDC.Item("CUST_CREDIT_HOLD") & "" = "1" Then
            ORDR_HOLD_CREDIT_REASON &= ",CRH"
        End If

        ' SLH - Customer Master indicates a Sales Hold
        Dim CUST_SALES_HOLD As String = rowARTCREDC.Item("CUST_SALES_HOLD") & ""
        If rowARTCREDC.Item("CUST_SALES_HOLD") & "" = "1" Then
            ORDR_HOLD_CREDIT_REASON &= ",SLH"
        End If

        ' REV - Customer Master indicates that the Credit Limit Review date is expired
        If rowARTCREDC.Item("CUST_CRED_LIMIT_REV") & "" = "" Then
            'ORDR_HOLD_CREDIT_REASON &= ",REV" ' DON'T HOLD THE ORDER UP IF THERE IS NO REVIEW DATE
        Else
            Dim CUST_CRED_LIMIT_REV As Date = rowARTCREDC.Item("CUST_CRED_LIMIT_REV")
            If Format(CUST_CRED_LIMIT_REV, "yyyyMMdd") _
             < Format(Now + ASCMAIN1.NowTSD, "yyyyMMdd") Then
                ORDR_HOLD_CREDIT_REASON &= ",REV"
            End If
        End If

        If rowARTCREDC.Item("NO_CREDIT_CHECK") & "" <> "1" Then

            ' INA - Last Sale was more than 3 months ago
            If rowARTCREDC.Item("LAST_SALE") & "" = "" Then
                ORDR_HOLD_CREDIT_REASON &= ",INA"
            Else
                Dim LAST_SALE As Date = rowARTCREDC.Item("LAST_SALE")
                If Format(LAST_SALE.AddMonths(SO_PARM_CRH_MOS_INACTIVE), "yyyyMMdd") _
                 < Format(Now + ASCMAIN1.NowTSD, "yyyyMMdd") Then
                    ORDR_HOLD_CREDIT_REASON &= ",INA"
                End If
            End If

            ' LIM - This order puts customer over the Credit Limit, or Customer was already over the Credit Limit
            Dim CUST_CREDIT_LIMIT As Decimal = Val(rowARTCREDC.Item("CUST_CREDIT_LIMIT") & "")
            Dim TOTAL_DUE As Decimal = Val(rowARTCREDC.Item("TOTAL_DUE") & "")
            Dim ORDR_AMT_PT As Decimal = Val(rowARTCREDC.Item("ORDR_AMT_PT") & "")
            Dim ORDR_AMT_PT_REL As Decimal = Val(rowARTCREDC.Item("ORDR_AMT_PT_REL") & "")
            If TOTAL_DUE + ORDR_AMT_PT + ORDR_AMT_OPEN + ORDR_AMT_PT_REL _
             > CUST_CREDIT_LIMIT * (1 + SO_PARM_CR_LIMIT_GRACE / 100) Then
                ORDR_HOLD_CREDIT_REASON &= ",LIM"
            End If

            ' P/D - This Customer has Past Due DR Balances (over 60 days)
            ' P/D - This Customer has Past Due DR Balances (over 90 days) - as per ciro
            Dim PAST_DUE_DR_AMT As Decimal = Val(rowARTCREDC.Item("PAST_DUE_DR_AMT") & "")
            If PAST_DUE_DR_AMT > 0 Then
                ORDR_HOLD_CREDIT_REASON &= ",P/D"
                past_due = True
            End If
        End If

        ' ADD - Customer Master has a COD Add-On Amt Defined, loaded into Order at Order Entry time
        If Val(rowSOTORDR1.Item("ORDR_COD_ADDON_AMT") & "") <> 0 Then
            ORDR_HOLD_CREDIT_REASON &= ",ADD"
        End If

        ' SPC - CSM indicates that order needs special processing
        If rowSOTORDR1.Item("ORDR_HOLD_CREDIT_SPECIAL") & String.Empty = "1" Then
            ORDR_HOLD_CREDIT_REASON &= ",SPC"
        End If

        If TERM_CODE = "C3" And Not credit_card_auth Then
            ' IF TERMS ARE CC AND WE DO NOT HAVE AN AUTH THEN HOLD THE ORDER
            'If ORDR_HOLD_CREDIT_REASON = "" Then
            ORDR_HOLD_CREDIT_REASON &= ",CCA"
            'End If
        Else
            If C1_C2_C3_P9 Or credit_card_auth Then
                If Not past_due And CUST_SALES_HOLD <> "1" Then
                    Return String.Empty
                End If
            End If
        End If

        Return ORDR_HOLD_CREDIT_REASON
    End Function

    Public Shared Sub Initialize_AGING_DATES()
        ReDim Preserve AGING_DATES(4)
        For i As Integer = 0 To 4
            'ascmain1.sql = "Select TO_CHAR(PRD_END_DATE,'dd-MMM-yyyy') from GLTPARM2 where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i)) & "'"
            ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i) & "'"
            AGING_DATES(i) = ASCDATA1.GetDataValue
        Next
    End Sub

    Public Shared Sub Update_ARTCUST6( _
    ByVal frmASFBASE0 As ASFBASE0, _
    ByVal CUST_CODE As String, _
    ByVal rowSOTORDR1 As DataRow, _
    ByVal S As Integer)

        Dim rowARTCUST6 As DataRow = frmASFBASE0.Fill_Record("ARTCUST6", CUST_CODE)
        If rowARTCUST6 Is Nothing Then
            rowARTCUST6 = frmASFBASE0.dst.Tables("ARTCUST6").NewRow
            rowARTCUST6.Item("CUST_CODE") = CUST_CODE
            frmASFBASE0.dst.Tables("ARTCUST6").Rows.Add(rowARTCUST6)
        End If

        With rowARTCUST6
            .Item("CUST_LAST_INV_NUM") = rowSOTORDR1.Item("ORDR_INV_NO")
            .Item("CUST_LAST_INV_DATE") = rowSOTORDR1.Item("ORDR_INV_DATE")
            .Item("CUST_LAST_INV_AMT") = Val(rowSOTORDR1.Item("ORDR_TOTAL_AMT") & "")
            If .Item("CUST_FIRST_PURCH") & "" = "" Then
                .Item("CUST_FIRST_PURCH") = rowSOTORDR1.Item("ORDR_INV_DATE")
            End If

            If rowSOTORDR1.Item("CUST_CODE") = CUST_CODE Then
                Dim INV_SALES As Decimal = Val(rowSOTORDR1.Item("ORDR_AMT") & "")
                If rowSOTORDR1.Item("ORDR_INV_TYPE") = "I" Then
                    .Item("CUST_SALES_MTD") = Val(.Item("CUST_SALES_MTD") & "") + INV_SALES
                    .Item("CUST_SALES_YTD") = Val(.Item("CUST_SALES_YTD") & "") + INV_SALES
                    .Item("CUST_NUM_INV_MTD") = Val(.Item("CUST_NUM_INV_MTD") & "") + S
                    .Item("CUST_NUM_INV_YTD") = Val(.Item("CUST_NUM_INV_YTD") & "") + S
                ElseIf rowSOTORDR1.Item("ORDR_INV_TYPE") = "C" Then
                    .Item("CUST_CRED_MTD") = Val(.Item("CUST_CRED_MTD") & "") + INV_SALES
                    .Item("CUST_CRED_YTD") = Val(.Item("CUST_CRED_YTD") & "") + INV_SALES
                End If
            End If

            If rowSOTORDR1.Item("CUST_BILL_TO_CUST") = CUST_CODE Then
                ASCMAIN1.sql = "Select Sum (INV_BALANCE) from ARTOPEN1 where CUST_CODE = '" & CUST_CODE & "'"
                Dim CUST_BAL As Decimal = Val(ASCDATA1.GetDataValue)
                If CUST_BAL > Val(.Item("CUST_HIGH_BAL_AMT") & "") Then
                    .Item("CUST_HIGH_BAL_DATE") = rowSOTORDR1.Item("ORDR_INV_DATE")
                    .Item("CUST_HIGH_BAL_AMT") = CUST_BAL
                End If
            End If
        End With
        frmASFBASE0.Update_Record_TDA("ARTCUST6")
    End Sub

    Public Shared Function Get_EDI_Custs(EDI_DOC_NO As String) As List(Of String)

        ASCMAIN1.sql = "Select DISTINCT CUST_CODE from EDTTRPM1 where CUST_CODE is Not Null "
        If EDI_DOC_NO <> "" Then
            ASCMAIN1.sql &= " and EDI_DOC_NO = '" & EDI_DOC_NO & "'"
        End If

        Dim c As New List(Of String)
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            c.Add(row.Item(0))
        Next

        Return c
    End Function

    Public Shared Function Prepare_Sales_Invoices( _
        F As ASFBASE1, _
        sqlw As String, _
        ByRef SOTINVH1 As String, _
        ByRef SOTINVH2 As String) As String

        ASCMAIN1.Progress("Building Work File")

        Dim rowGLTPARM2 As DataRow = F.LookUp("GLTPARM2", ASCMAIN1.CYP)
        Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")
        Dim NYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)

        ' wjz removing the and that filters to I only because this makes it impossible to bring in credits to SOTINVHR down below
        '& "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _

        ASCMAIN1.sql = "Select SOTINVH2.*" & vbCrLf _
            & ", ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTINVH1.INV_DATE" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE as ORDR_AMT_SHIP" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST as ORDR_CGS_SHIP" & vbCrLf _
            & " from SOTINVH2, SOTINVH1, ICTSTYL1" & vbCrLf _
            & " where SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & sqlw

        SOTINVH2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add Primary Key (INV_TYPE,INV_NO,INV_LNO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add ORDR_QTY_CANC NUMBER(6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add ORDR_AMT_CANC NUMBER(13,2)")
        ASCMAIN1.AnalyzeTable(SOTINVH2)


        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.CUST_DC_NO" & vbCrLf _
            & ", SOTORDR1.EDI_APPOINTMENT" & vbCrLf _
            & " from SOTINVH1, SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and (INV_TYPE, INV_NO) in (Select Distinct INV_TYPE, INV_NO from " & SOTINVH2 & ")"
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE,INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_1 on " & SOTINVH1 & " (INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_2 on " & SOTINVH1 & " (PICK_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_3 on " & SOTINVH1 & " (ORDR_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_4 on " & SOTINVH1 & " (SHIP_BOL_NO)")
        ASCMAIN1.AnalyzeTable(SOTINVH1)

        If F.MENU_ITEM_OBJECT = "SORUPDT1" Then
            ASCMAIN1.sql = "Update " & SOTINVH1 & " Set ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update " & SOTINVH1 & " Set ORDR_YYYYPP_UPDATED = '" & NYP & "'" _
                & " where INV_DATE > '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
            ASCDATA1.ExecuteSQL()
        End If


        '& ", 0 TOTAL_UNITS " & vbCrLf _
        '& ", 0 TOTAL_UNITS_CANC " & vbCrLf _
        '& ", 0 TOTAL_UNITS_BACK " & vbCrLf _

        ASCMAIN1.sql = "Select SOTINVH1.* from " & SOTINVH1 & " SOTINVH1"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVH1", 2))


        If F.MENU_ITEM_OBJECT = "SORMTDS1" Then
        Else
            ASCMAIN1.sql = "Select SOTINVH2.* from " & SOTINVH2 & " SOTINVH2"
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVH2", 3))
        End If


        ' Credits

        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", NULL ORDR_GROUP_NO, NULL ORDR_ADDR_TYPE_ST, NULL CUST_DC_NO" & vbCrLf _
            & ", NULL EDI_APPOINTMENT" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = 'C'"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHR", 2))

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC" _
            & ", ICTSTYL1.STYLE_COST, ICTSTYL1.SALES_DIVISION_CODE from ICTSTYL1" _
            & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & SOTINVH2 & ")"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTSTYL1", 1))

        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", Decode(E.CUST_CODE,NULL,'N','Y') EDI" & vbCrLf _
            & ", Decode(M.CUST_CODE,NULL,'N','Y') MULTI_STORE" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & ", (Select Distinct CUST_CODE from EDTTRPM1 where EDI_STATUS = 'P' and EDI_DOC_NO = '810') E" & vbCrLf _
            & ", (Select CUST_CODE from ARTCUST2 where CUST_ADDR_TYPE = 'MK' group by CUST_CODE having COUNT (*) > 1) M" & vbCrLf _
            & " where E.CUST_CODE (+) = ARTCUST1.CUST_CODE" & vbCrLf _
            & "   and M.CUST_CODE (+) = ARTCUST1.CUST_CODE"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTCUST1", 1))

        ASCMAIN1.sql = "" _
            & "Select SOTFPCT1.OPS_YYYYPP, SOTFPCT1.CUST_FACTOR_PERCENT, SOTFPCT1.CUST_SURCHARGE_PERCENT" & vbCrLf _
            & " from SOTFPCT1" & vbCrLf _
            & " union " & vbCrLf _
            & "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, SO_PARM_FACTOR_PCT CUST_FACTOR_PERCENT, SO_PARM_SURCHARGE_PCT CUST_SURCHARGE_PERCENT" & vbCrLf _
            & " from SOTPARM1 where SO_PARM_KEY = 'Z'" & vbCrLf _
            & " union " & vbCrLf _
            & "Select '" & NYP & "' OPS_YYYYPP, SO_PARM_FACTOR_PCT CUST_FACTOR_PERCENT, SO_PARM_SURCHARGE_PCT CUST_SURCHARGE_PERCENT" & vbCrLf _
            & " from SOTPARM1 where SO_PARM_KEY = 'Z'"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTFPCT1", 1))


        If F.MENU_ITEM_OBJECT = "SORUPDT1" Then
            ASCMAIN1.Progress("-", "Pick Tickets")
            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1" & vbCrLf _
                & " where SOTPICK1.PICK_NO in (Select PICK_NO from " & SOTINVH1 & ")"
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTPICK1", 1))

            ASCMAIN1.Progress("-", "Shipments")
            ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
                & ", DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'MK','G:' || SOTSHIP1.ORDR_GROUP_NO, 'S:' || SOTSHIP1.SHIP_BOL_NO) SHIP_BOL_NO_X" & vbCrLf _
                & ", 'N' MULTI_STORE from SOTSHIP1" & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO in (Select DISTINCT SHIP_BOL_NO from " & SOTINVH1 & ")"
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSHIP1", 1))

            ASCMAIN1.Progress("-", "Verify Integrity")
            Dim sql1 As String = "(Select SHIP_BOL_NO, Count (*), Min (INV_NO), Max (INV_NO), Sum (INV_SALES) from " & SOTINVH1 & vbCrLf _
                & " group by SHIP_BOL_NO)" & vbCrLf
            Dim sql2 As String = "(Select SOTSHIP1.SHIP_BOL_NO, Count (*), Min (SOTINVH1.INV_NO), Max (SOTINVH1.INV_NO)" & vbCrLf _
                & ", Sum (SOTINVH1.INV_SALES) from SOTINVH1,SOTSHIP1" & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTINVH1.SHIP_BOL_NO (+) and SOTINVH1.ORDR_YYYYPP_UPDATED IS NULL" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from SOTSHIP1 where REGISTER_XNO is Null)" & vbCrLf _
                & " group by SOTSHIP1.SHIP_BOL_NO)"
            'ASCMAIN1.sql = "" _
            '    & sql1 & " minus " & sql2 & vbCrLf _
            '    & " union " & vbCrLf _
            '    & sql2 & " minus " & sql1
            ASCMAIN1.sql = sql1 & " minus " & sql2
            If ASCDATA1.GetDataTable.Rows.Count <> 0 Then
                Return "Shipments Header Record not in synch w/Invoices; Call ABS"
            End If

            'ASCMAIN1.sql = "Select * from SOTPICK1 where PICK_STATUS <> 'F'"
            'If ASCDATA1.GetDataTable.Rows.Count <> 0 Then
            '    Return "Unconfirmed Pick Tickets linked to Invoices; Call ABS"
            'End If

            If F.dst.Tables("SOTPICK1").Select("ISNULL(PICK_STATUS,'?') <> 'F'").Length <> 0 Then
                Return "Unconfirmed Pick Tickets linked to Invoices; Call ABS"
            End If

            ASCMAIN1.Progress("-", "Pick Ticket Details")
            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE from SOTPICK2,SOTORDR2" & vbCrLf _
                & " where SOTPICK2.PICK_NO in (Select PICK_NO from " & SOTINVH1 & ")" & vbCrLf _
                & " and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & " and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO"
            'F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTPICK2", 2))
            F.Create_TDA(F.dst.Tables.Add, "SOTPICK2", "**", 0, True, "", 2, "PICK_QTY_BACK")
            F.Fill_Records("SOTPICK2")

            F.Create_Relation("SOTINVH1", "SOTPICK2", "PICK_NO")
            F.dst.Tables("SOTINVH1").Columns.Add("TOTAL_UNITS", GetType(System.Int64), "SUM(CHILD(SOTINVH1_SOTPICK2).PICK_QTY_CONF)")
            F.dst.Tables("SOTINVH1").Columns.Add("TOTAL_UNITS_CANC", GetType(System.Int64), "SUM(CHILD(SOTINVH1_SOTPICK2).PICK_QTY_CANC)")
            F.dst.Tables("SOTINVH1").Columns.Add("TOTAL_UNITS_BACK", GetType(System.Int64), "SUM(CHILD(SOTINVH1_SOTPICK2).PICK_QTY_BACK)")

            F.Create_TDA(F.dst.Tables.Add, "ARTOPEN1", "*")
            F.Create_TDA(F.dst.Tables.Add, "ARTPYMT1", "*")
            F.Create_TDA(F.dst.Tables.Add, "ARTPYMT2", "*")
            F.Create_TDA(F.dst.Tables.Add, "ARTPYMT3", "*")
            F.Create_TDA(F.dst.Tables.Add, "ARTCUST6", "*")

        End If




        If F.MENU_ITEM_OBJECT = "SORUPDT1" Then
            ' NOT REQUIRED FOR INVOICE UPDATE
        Else


            ASCMAIN1.Progress("-", "Report Summaries")
            ASCMAIN1.sql = "Select SOTINVH2.SALES_DIVISION_CODE, SOTINVH1.CUST_CODE, SOTINVH1.INV_DATE" & vbCrLf _
                & ", SOTINVH1.SREP_CODE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','G:' || SOTORDR1.ORDR_GROUP_NO, 'S:' || SOTINVH1.SHIP_BOL_NO) AS SHIP_BOL_NO_X" & vbCrLf _
                & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','MK','DC') AS SHIP_ADDR_TYPE" & vbCrLf _
                & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','000000',SOTORDR1.CUST_DC_NO) AS SHIP_ADDR_CODE" & vbCrLf _
                & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
                & ", Sum (SOTINVH2.ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
                & ", Sum (SOTINVH2.ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
                & ", Sum (SOTINVH2.ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
                & "  from SOTINVH1," & SOTINVH2 & " SOTINVH2, SOTORDR1" & vbCrLf _
                & "  where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                & "    and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "    and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
                & " group by SOTINVH2.SALES_DIVISION_CODE, SOTINVH1.CUST_CODE,  SOTINVH1.INV_DATE" & vbCrLf _
                & ", SOTINVH1.SREP_CODE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','G:' || SOTORDR1.ORDR_GROUP_NO, 'S:' || SOTINVH1.SHIP_BOL_NO)" & vbCrLf _
                & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','MK','DC')" & vbCrLf _
                & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','000000',SOTORDR1.CUST_DC_NO)" & vbCrLf _
                & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf
            Dim SOTINVHD As String = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " Add ORDR_QTY_CANC NUMBER (6,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " Add ORDR_AMT_CANC NUMBER (13,2)")

            ASCMAIN1.sql = "Update " & SOTINVHD & " Set SHIP_BOL_NO_X = 'G:' || ORDR_GROUP_NO"
            ASCDATA1.ExecuteSQL()

            For Each GROUP As String In New String() {"G"} ' , "S"} NO NEED FOR S SINCE WE MAKE EVERYTHING G ABOVE
                ASCMAIN1.sql = "" _
                    & "Begin" _
                    & " Declare Cursor C1 is Select SOTINVHD.SHIP_BOL_NO_X, SOTINVHD.ORDR_GROUP_NO, SOTINVHD.SALES_DIVISION_CODE" _
                    & " , SOTINVHD.CUST_CODE, SOTINVHD.INV_DATE, SOTINVHD.SREP_CODE, SOTINVHD.SHIP_ADDR_TYPE, SOTINVHD.SHIP_ADDR_CODE" _
                    & " , SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                    & "  , SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0)) ORDR_QTY_CANC" _
                    & "  , SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_CANC" _
                    & "  from SOTORDR1,SOTORDR2," & SOTINVHD & " SOTINVHD" _
                    & IIf(GROUP = "S", ",SOTPICK1", "") _
                    & "  where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
                    & " and SOTORDR2.STYLE_CODE = SOTINVHD.STYLE_CODE and SOTORDR2.COLOR_CODE = SOTINVHD.COLOR_CODE" _
                    & "    and SOTORDR2.ORDR_QTY_CANC <> 0" _
                    & IIf(GROUP = "G", _
                            "    and SOTORDR1.ORDR_GROUP_NO = SOTINVHD.ORDR_GROUP_NO", _
                            "    and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO and SOTPICK1.SHIP_BOL_NO = SUBSTR(SOTINVHD.SHIP_BOL_NO_X,3)") _
                    & "    and SHIP_BOL_NO_X like '" & GROUP & "%'" _
                    & "  group by SOTINVHD.SHIP_BOL_NO_X, SOTINVHD.ORDR_GROUP_NO, SOTINVHD.SALES_DIVISION_CODE" _
                    & " , SOTINVHD.CUST_CODE, SOTINVHD.INV_DATE, SOTINVHD.SREP_CODE, SOTINVHD.SHIP_ADDR_TYPE, SOTINVHD.SHIP_ADDR_CODE" _
                    & " , SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" _
                    & " Begin" _
                    & "  For R1 in C1 Loop" _
                    & "   Begin" _
                    & "    Update " & SOTINVHD & " Set " _
                    & "      ORDR_QTY_CANC = R1.ORDR_QTY_CANC" _
                    & "     ,ORDR_AMT_CANC = R1.ORDR_AMT_CANC" _
                    & "     where SALES_DIVISION_CODE = R1.SALES_DIVISION_CODE" _
                    & IIf(GROUP = "G", _
                            " and ORDR_GROUP_NO = R1.ORDR_GROUP_NO;", _
                            " and SHIP_BOL_NO_X = R1.SHIP_BOL_NO_X;") _
                    & "   End;" _
                    & "  End Loop;" _
                    & " End;" _
                    & "End;"

                MsgBox("Please call ABS", MsgBoxStyle.OkOnly, "Need to check SQL for Performance")

                ASCDATA1.ExecuteSQL()

                MsgBox("Please call ABS", MsgBoxStyle.OkOnly, "Need to check SQL for Performance")
            Next

            ASCMAIN1.sql = "Select * from " & SOTINVHD
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHD", 0))

            ASCMAIN1.sql = "Select SHIP_BOL_NO_X, SALES_DIVISION_CODE, CUST_CODE, INV_DATE" & vbCrLf _
                & ", Sum (ORDR_QTY_CANC) as QTY_CANC" & vbCrLf _
                & ", Sum (ORDR_AMT_CANC) as AMT_CANC" & vbCrLf _
                & " from " & SOTINVHD & vbCrLf _
                & " group by SHIP_BOL_NO_X, SALES_DIVISION_CODE, CUST_CODE, INV_DATE"
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHN", 0))

            ASCMAIN1.sql = "Select SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
                & ", Sum (ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
                & ", Sum (ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
                & " from " & SOTINVH2 & vbCrLf _
                & " group by SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE"
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHC", 0))

            ASCMAIN1.sql = "Select INV_DATE, SALES_DIVISION_CODE, CUST_CODE" & vbCrLf _
                & ", STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
                & ", Sum (ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
                & ", Sum (ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
                & " from " & SOTINVH2 & vbCrLf _
                & " group by INV_DATE, SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE"
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHY", 0))

        End If



        ASCMAIN1.sql = "Select * from SOTORDR0 where ORDR_GROUP_NO in" & vbCrLf _
            & " (Select DISTINCT ORDR_GROUP_NO from " & SOTINVH1 & ")"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTORDR0", 0))

        ASCMAIN1.Progress("-", "Consolidated Invoices")

        For Each TABLE_NAME As String In New String() {"SOTINVHZ", "SOTINVHT"}
            ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE" & vbCrLf _
                & ", SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO" & vbCrLf _
                & ", Sum (SOTINVH2.ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
                & ", SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_MISC_CHG, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, SOTINVH2.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.INV_NO_CONS, SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO" & vbCrLf _
                & ", SOTINVH1.SALES_DIVISION_CODE as H_SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.GST_TAX" & vbCrLf _
                & " from SOTINVH1," & SOTINVH2 & " SOTINVH2" & vbCrLf _
                & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "   and SOTINVH1.INV_NO_CONS is " & IIf(TABLE_NAME = "SOTINVHT", "NOT", "") & " Null" & vbCrLf _
                & " group by " & vbCrLf _
                & " SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE" & vbCrLf _
                & ", SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO" & vbCrLf _
                & ", SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_MISC_CHG, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, SOTINVH2.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.INV_NO_CONS, SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO" & vbCrLf _
                & ", SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.GST_TAX" & vbCrLf
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", TABLE_NAME, 0))
        Next

        For Each rowSOTINVHT As DataRow In F.dst.Tables("SOTINVHT").Select("")
            Dim rowSOTINVHZ As DataRow = F.dst.Tables("SOTINVHZ").NewRow
            rowSOTINVHZ.ItemArray = rowSOTINVHT.ItemArray
            F.dst.Tables("SOTINVHZ").Rows.Add(rowSOTINVHZ)
        Next



        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", 0 AS TOTAL_UNITS" & vbCrLf _
            & ", 0 AS TOTAL_UNITS_CANC" & vbCrLf _
            & ", 0 AS TOTAL_UNITS_BACK" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1, SOTORDR1" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and SOTINVH1.INV_NO_CONS is Null"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHX", 2))

        ASCMAIN1.sql = "Select SOTINVH1.INV_NO_CONS, SOTINVH1.ORDR_BILL_TO_CUST as CUST_CODE" & vbCrLf _
            & ", Max(SOTINVH1.INV_DATE) AS INV_DATE" & vbCrLf _
            & ", Max(SOTINVH1.REASON_CODE) as REASON_CODE" & vbCrLf _
            & ", Max(SOTINVH1.SALES_DIVISION_CODE) as SALES_DIVISION_CODE" & vbCrLf _
            & ", Max(SOTINVH1.ORDR_CUST_PO) AS ORDR_CUST_PO" & vbCrLf _
            & ", Max(SOTINVH1.CUST_FACTOR_IND) AS CUST_FACTOR_IND" & vbCrLf _
            & ", Max(SOTINVH1.CUST_SURCHARGE_IND) AS CUST_SURCHARGE_IND" & vbCrLf _
            & ", Sum(SOTINVH1.INV_SALES) AS INV_SALES" & vbCrLf _
            & ", Sum(SOTINVH1.INV_FREIGHT) AS INV_FREIGHT" & vbCrLf _
            & ", Sum(SOTINVH1.INV_MISC_CHG) AS INV_MISC_CHG" & vbCrLf _
            & ", Sum(SOTINVH1.GST_TAX) AS GST_TAX" & vbCrLf _
            & ", Sum(SOTINVH1.INV_TOTAL_AMOUNT) AS INV_TOTAL_AMOUNT" & vbCrLf _
            & ", 0 AS TOTAL_UNITS" & vbCrLf _
            & ", 0 AS TOTAL_UNITS_CANC" & vbCrLf _
            & ", 0 AS TOTAL_UNITS_BACK" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1" & vbCrLf _
            & " where SOTINVH1.INV_NO_CONS is Not Null" & vbCrLf _
            & " group by SOTINVH1.INV_NO_CONS, SOTINVH1.ORDR_BILL_TO_CUST" & vbCrLf
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim rowSOTINVHX As DataRow = F.dst.Tables("SOTINVHX").NewRow
            With rowSOTINVHX
                For Each DCOL As DataColumn In row.Table.Columns
                    If DCOL.ColumnName = "INV_NO_CONS" Then
                        .Item("INV_TYPE") = "I"
                        .Item("INV_NO") = row.Item("INV_NO_CONS")
                    Else
                        .Item(DCOL.ColumnName) = row.Item(DCOL.ColumnName)
                    End If
                Next
                .Item("CURR_CODE") = "USD"
                .Item("CURR_EXCH_RATE") = 1
            End With
            F.dst.Tables("SOTINVHX").Rows.Add(rowSOTINVHX)
        Next



        If F.MENU_ITEM_OBJECT = "SORUPDT1" Then
            ' NOT REQUIRED FOR INVOICE UPDATE
        Else

            For Each TABLE_NAME As String In New String() {"SOTINVHG1", "SOTINVHG", "SOTINVHG2"}
                With F.dst.Tables.Add(TABLE_NAME)
                    .Columns.Add("SD")
                    .Columns.Add("CC")
                    If TABLE_NAME <> "SOWINVHG2" Then .Columns.Add("ID", GetType(System.DateTime))
                    .Columns.Add("QC", GetType(System.Int64))
                    .Columns.Add("AC", GetType(System.Decimal))
                    If TABLE_NAME = "SOTINVHG" Then .PrimaryKey = New DataColumn() {.Columns("SD"), .Columns("CC"), .Columns("ID")}
                    If TABLE_NAME = "SOTINVHG2" Then .PrimaryKey = New DataColumn() {.Columns("SD"), .Columns("CC")}
                End With
            Next

            For Each row As DataRow In F.dst.Tables("SOTINVHN").Select("")
                Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE")
                Dim CUST_CODE As String = row.Item("CUST_CODE")
                Dim INV_DATE As Date = row.Item("INV_DATE")
                F.dst.Tables("SOTINVHG1").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE, row.Item("QTY_CANC"), row.Item("AMT_CANC")})
                If F.dst.Tables("SOTINVHG").Rows.Find(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE}) Is Nothing Then
                    F.dst.Tables("SOTINVHG").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE})
                End If
                If F.dst.Tables("SOTINVHG2").Rows.Find(New Object() {SALES_DIVISION_CODE, CUST_CODE}) Is Nothing Then
                    F.dst.Tables("SOTINVHG2").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE})
                End If
            Next

            F.Create_Relation("SOTINVHG", "SOTINVHG1", "SD,CC,ID")
            F.dst.Tables("SOTINVHG").Columns("QC").Expression = "SUM(CHILD.QC)"
            F.dst.Tables("SOTINVHG").Columns("AC").Expression = "SUM(CHILD.AC)"

            F.Create_Relation("SOTINVHG2", "SOTINVHG", "SD,CC")
            F.dst.Tables("SOTINVHG2").Columns("QC").Expression = "SUM(CHILD.QC)"
            F.dst.Tables("SOTINVHG2").Columns("AC").Expression = "SUM(CHILD.AC)"


        End If

        ' Master Files

        ASCMAIN1.sql = "Select ARTREAS1.* from ARTREAS1"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTREAS1", 1))

        ASCMAIN1.sql = "Select SOTSDIV1.* from SOTSDIV1"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSDIV1", 1))


        Return ""
    End Function

    Public Shared Function Allocation_Initialization(frmASFBASE0 As ASFBASE0, _
                                       WHSE_CODE As String, _
                                       force_pick As Boolean, _
                                       allocation_only As Boolean, _
                                       fill_data As Boolean, _
                                       ORDR_GROUP_NO_sql As String, _
                                       SHIP_BY_DATE As Date, _
                                       Optional sql_where As String = "", _
                                       Optional manual_release As Boolean = False) As Dictionary(Of String, String)

        ' Creates Temporary Table used to drive Allocation
        ' May be called with fill_data option, 
        '  in which case this routine exits with Oracle Temp-Tables filled with data, 
        '  or else the sql's required to fill the data ioto the Temp-Tables are provided 

        ' SOTORDR1
        ' All Orders (regardless of any filters provided by the Order Release UI)
        ' - modified 05/01/13 by WJZ
        '   takes too long at RGI to consider all orders when doing a release
        '   so passing in an optional sql_where to limit the time required
        '    - allocation only should consider all orders, but a targeted order release does not need to do so
        '  are loaded into the demand work tables, so that we allocate based on all demand, 
        '  and then release based on filter criteria, including dates, customers, etc
        ' If Force-Picking, then only those orders that are being force-picked are loaded to save time
        ' If Allocating/Releasing by Warehouse, then only the demand (orders) shipping from that warehouse are loaded into Demand tables

        ' For SOTORDR1, SOTORDR2, SOTRSRV1, SOTRSRV2:
        '  The Data Adapter is set up for Update on Specific fields
        '  Is this used?  Or do we do all updates in Oracle

        ' Oracle Temp Tables (& ADO.Net TDAs & DataTables) are set up for:
        '  SOTORDR1 SOTORDR0 ARTCUST1 ICTSTDQ1 SOTORDR2 SOTRSRV1 SOTRSRV2
        ' ADO.Net Data Tables are set up for
        '  SOTSUPP0 SOTSUPPI SOTORDR7

        ' Truncate SOTORDR1 SOTORDR0 ARTCUST1 ICTSTDQ1 SOTORDR2 SOTRSRV1 SOTRSRV2
        ' Execute all sql's loaded into TABLE_NAMEs dictionary, in the order that they were placed
        ' Clear Rows for SOTSUPP0 SOTSUPPI SOTORDR7 and refill as necessary

        Dim TABLE_NAMEs As New Dictionary(Of String, String)

        ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, SOTORDR1.ORDR_DATE, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.CUST_STORE_NO" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.WHSE_CODE, SOTORDR1.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_SOURCE, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.ORDR_PRIORITY, SOTORDR1.ORDR_STATUS, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", SOTORDR1.ORDR_HOLD, SOTORDR1.ORDR_REL_HOLD_CODES, SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.CUST_DC_NO" & vbCrLf _
            & ", SOTORDR1.ORDR_REL_BATCH_NO, SOTORDR1.CUST_FACTOR_IND, SOTORDR1.ORDR_PRE_ALLOC" & vbCrLf _
            & ", SOTORDR1.CUST_BILL_TO_CUST CUST_CREDIT_GROUP_CUST" & vbCrLf _
            & ", SOTORDR1.INIT_DATE,SOTORDR1. ORDR_PICK_SEQ, SOTORDR1.SREP_CODE, SOTORDR1.FRT_TERMS, SOTORDR1.SHIP_VIA_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE_CLOSED, SOTORDR1.ORDR_YYYYPP_CLOSED, SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
            & ", SOTORDR1.ORDR_DEPT, SOTORDR1.TERM_CODE, SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.CC_TRANS_ID" & vbCrLf _
            & ", SOTORDR1.CCPA_NO, SOTORDR1.ORDR_SHIP_COMPLETE, '0' ALLOCATION_ONLY_SCOPE" _
            & " from SOTORDR1,ARTCUST1,SOTORDRG" & vbCrLf _
            & " where SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & "   and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & IIf(WHSE_CODE = "", _
                  "", _
                  "   and SOTORDR1.WHSE_CODE = '" & WHSE_CODE & "'")

        Dim sqlSOTORDR1_Allocation_Scope As String = ASCMAIN1.sql ' used below to include other orders, out of scope, calling for style/colors on orders that are in scope, to ensure a full allocation

        ASCMAIN1.sql &= sql_where ' used to limit the scope to manual release orders only

        If force_pick Or manual_release Then
            ASCMAIN1.sql &= " and SOTORDR1.ORDR_GROUP_NO in (" & ORDR_GROUP_NO_sql & ")"
        Else
            If ASCMAIN1.CLIENT = "VAN" Then
                '**************************** TEMPORARY WORKAROUND TO DEAL WITH WALMART GLUT
                'ASCMAIN1.sql &= " and CUST_CODE <> 'WALMART'"

                ' THIS SECTION DEALS WITH SETTING UP THE CUSTOMERS EXCLUDED FROM THE ALLOCATION PROCESS (UNLESS FORCE PICKED)
                ' maybe these should be presented in a grid in SOROREL1 and let the user uncheck?
                ' and then passed in using sql_where
                ' AND MAYBE WE SHOULD SEE THE EXCLUSION CUSTOMERS AND ORDERS ON THE RELEASE SCREEN SO THEY CAN BE EXPORTED
                If allocation_only Then
                    ' ALL CUSTOMERS, ALL ORDERS
                Else
                    ASCMAIN1.sql &= " and NVL(ARTCUST1.CUST_ALLO_EXCL,'0') <> '1'" & vbCrLf
                    ASCMAIN1.sql &= " and NVL(SOTORDRG.ORDR_ALLO_EXCL,'0') <> '1'" & vbCrLf
                End If


                'Dim sqlEXCL As String = ""
                'Dim sqlCUST_CODEsx = "Select CUST_CODE from ARTCUST1 where CUST_ALLO_EXCL = '1'"
                'For Each row As DataRow In ASCDATA1.GetDataTable(sqlCUST_CODEsx).Select("")
                '    sqlEXCL &= ",'" & row.Item("CUST_CODE") & "'"
                'Next
                'If sqlEXCL <> "" Then
                '    ASCMAIN1.sql &= " and SOTORDR1.CUST_CODE not in (" & Mid(sqlEXCL, 2) & ")"
                'End If
            End If

        End If




        If Not fill_data Then ASCMAIN1.sql &= " and ROWNUM < 1"
        Dim SOTORDR1 As String = ASCMAIN1.Temp_Table ' all Allocable Orders
        TABLE_NAMEs.Add("SOTORDR1", SOTORDR1)
        If Not fill_data Then
            TABLE_NAMEs.Add("sqlSOTORDR1", "Insert into " & SOTORDR1 & " " & Replace(ASCMAIN1.sql, " and ROWNUM < 1", " and ORDR_NO in (Select Distinct ORDR_NO from SOTORDR2 where ORDR_STATUS = 'O' and STYLE_CODE = 'STYLE_CODE')"))
        End If
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTORDR1 & "_1 on " & SOTORDR1 & " (ORDR_GROUP_NO)")

        If ASCMAIN1.CLIENT = "RGI" Then ' HOPEFULLY THIS WILL BE EVENTUALLY BE USED BY ALL
            ' Get all other orders that want the styles that are on the orders already placed into work table SOTORDR1
            If Not force_pick And fill_data And Not allocation_only And sql_where <> "" Then ' And manual_release Then APPARENTLY, manual_release is not the same as when rita picks release orders manually selected
                ' this condition is placed here only because a general release already includes all orders, and manual releases do not
                ASCMAIN1.sql = "Select Distinct SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
                    & " from SOTORDR2, " & SOTORDR1 & " SOTORDR1" & vbCrLf _
                    & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO and SOTORDR2.ORDR_STATUS = 'O'" & vbCrLf _
                    & IIf(WHSE_CODE = "", _
                      "", _
                      "   and SOTORDR1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf)

                Dim SOTORDRL As String = ASCMAIN1.Temp_Table ' all Style/Colors to consider for Allocation
                ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRL & " Add Primary Key (STYLE_CODE, COLOR_CODE, WHSE_CODE)")
                TABLE_NAMEs.Add("SOTORDRL", SOTORDRL)

                ASCMAIN1.sql = "Select Distinct SOTORDR2.ORDR_NO from SOTORDR2,SOTORDR1" & vbCrLf _
                    & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                    & "   and SOTORDR2.ORDR_STATUS = 'O'" & vbCrLf _
                    & "   and (SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE) in " & vbCrLf _
                    & " (Select Distinct STYLE_CODE, COLOR_CODE from " & SOTORDRL & ")" & vbCrLf _
                    & IIf(WHSE_CODE = "", _
                      "", _
                      "   and SOTORDR1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf) _
                    & " minus Select ORDR_NO from " & SOTORDR1
                ASCMAIN1.sql = "Insert into " & SOTORDR1 & " " & Replace(sqlSOTORDR1_Allocation_Scope, "'0' ALLOCATION_ONLY_SCOPE", "'1' ALLOCATION_ONLY_SCOPE") & " and SOTORDR1.ORDR_NO in (" & ASCMAIN1.sql & ")"
                ASCDATA1.ExecuteSQL()
            End If
        End If


        ' For those customers with a Fixed Priority, set that priority in the Orders Table 
        '  (in case it was tampered with) for all orders within Release Date Horizon

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is " _
            & "  Select CUST_CODE, CUST_PRIORITY_CODE from ARTCUST1" _
            & "   where CUST_PRIORITY_FIXED = '1'" _
            & "     and CUST_CODE in (Select Distinct CUST_CODE from " & SOTORDR1 & ");" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update " & SOTORDR1 _
            & "    Set ORDR_PRIORITY = R1.CUST_PRIORITY_CODE" _
            & "    where CUST_CODE = R1.CUST_CODE" _
            & "      and ORDR_SHIP_DATE <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "';" _
            & "  End Loop; " _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()
        If fill_data Then
            ASCMAIN1.AnalyzeTable(SOTORDR1)
        Else
            TABLE_NAMEs.Add("sqlSOTORDR1a", ASCMAIN1.sql)
        End If

        If Not allocation_only Then
            ' in the SOTORDR1 creation above, we assume that the Bill-To Customer is the CGC, 
            '  so now we need to go back and update those BTs with a specific CGC
            ' for those Bill-To Customers which have a specific Credit Group Customer, 
            '  let's use that specific CGC
            ' Since this is a credit concern, no need to do this for allocation only

            ASCMAIN1.sql = "Update " & SOTORDR1 _
                & " Set CUST_BILL_TO_CUST = CUST_CODE, CUST_CREDIT_GROUP_CUST = CUST_CODE" _
                & " where CUST_BILL_TO_CUST is Null"
            If fill_data Then
                ASCDATA1.ExecuteSQL()
            Else
                TABLE_NAMEs.Add("sqlSOTORDR1b", ASCMAIN1.sql)
            End If

            ASCMAIN1.sql = "Select CUST_BILL_TO_CUST, CUST_CREDIT_GROUP_CUST from ARTCUST1" _
                & " where CUST_CODE in (Select Distinct CUST_BILL_TO_CUST from " & SOTORDR1 & ")" _
                & "   and CUST_CREDIT_GROUP_CUST is Not Null"
            ASCMAIN1.sql = "Begin Declare Cursor C1 is " & ASCMAIN1.sql & ";" _
                & " Begin For R1 in C1 Loop" _
                & "  Update " & SOTORDR1 & " Set CUST_CREDIT_GROUP_CUST = R1.CUST_CREDIT_GROUP_CUST " _
                & "   where CUST_BILL_TO_CUST = R1.CUST_BILL_TO_CUST; " _
                & " End Loop; End; End; "
            If fill_data Then
                ASCDATA1.ExecuteSQL()
            Else
                TABLE_NAMEs.Add("sqlSOTORDR1c", ASCMAIN1.sql)
            End If
        End If

        ASCMAIN1.sql = "Select * from " & SOTORDR1 ' & " where ROWNUM < 1"
        frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("SOTORDR1"), SOTORDR1, "**", 0, , , , "ORDR_PICK_SEQ,ORDR_STATUS,ORDR_REL_HOLD_CODES,ORDR_REL_BATCH_NO")

        ' ICTSTDQ1

        Dim ICTSTDQ1 As String = frmASFBASE0.Create_Temporary_Table("ICTSTDQ1", "WHSE_CODE,STYLE_CODE,COLOR_CODE,STATUS_DATE")
        TABLE_NAMEs.Add("ICTSTDQ1", ICTSTDQ1)

        ' ICTSTDQ2

        Dim ICTSTDQ2 As String = frmASFBASE0.Create_Temporary_Table("ICTSTDQ2", "WHSE_CODE,STYLE_CODE,COLOR_CODE")
        TABLE_NAMEs.Add("ICTSTDQ2", ICTSTDQ2)

        ' ICTSTDQ3

        Dim ICTSTDQ3 As String = frmASFBASE0.Create_Temporary_Table("ICTSTDQ3", "ORDR_GROUP_NO,STYLE_CODE,COLOR_CODE")
        TABLE_NAMEs.Add("ICTSTDQ3", ICTSTDQ3)

        ' SOTORDR0

        ASCMAIN1.sql = "Select SOTORDR0.* " _
            & " from SOTORDR0 " _
            & " where SOTORDR0.ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ")"
        Dim SOTORDR0 As String = ASCMAIN1.Temp_Table
        TABLE_NAMEs.Add("SOTORDR0", SOTORDR0)
        If Not fill_data Then
            TABLE_NAMEs.Add("sqlSOTORDR0", "Insert into " & SOTORDR0 & " " & ASCMAIN1.sql)
        End If
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO)")
        If fill_data Then ASCMAIN1.AnalyzeTable(SOTORDR0)
        ASCMAIN1.sql = "Select * from " & SOTORDR0
        frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("SOTORDR0"), SOTORDR0, "**", 0, False)

        ' SOTORDR2
        ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" _
            & ", ICTSTYL1.SUB_BODY_CODE, round(decode(NVL(ICTSTYL1.CARTON_PACK_QTY,0),0,0,(NVL(ICTSTYL1.CASE_CUBE,0) / NVL(ICTSTYL1.CARTON_PACK_QTY,0))),5)  STANDARD_CUBE_PER_UNIT" _
            & " from SOTORDR2," & SOTORDR1 & " SOTORDR1, ICTSTYL1 where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & " and ICTSTYL1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE"

        'ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" _
        '    & ", ICTSTYL1.SUB_BODY_CODE, ICTBODY2.STANDARD_CUBE_PER_UNIT" _
        '    & " from SOTORDR2," & SOTORDR1 & " SOTORDR1, ICTSTYL1, ICTBODY2 where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
        '    & " and ICTSTYL1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE and ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE"
        Dim SOTORDR2 As String = ASCMAIN1.Temp_Table
        TABLE_NAMEs.Add("SOTORDR2", SOTORDR2)
        If Not fill_data Then TABLE_NAMEs.Add("sqlSOTORDR2", _
            "Insert into " & SOTORDR2 & " " _
            & Replace(ASCMAIN1.sql, " from", ", 0 ORDR_LAST_UNIT, 0 ORDR_QTY_ALLO_CUR, 0 ORDR_QTY_ALLO_FUT, 0 ORDR_QTY_ALLO_CXL, Null ORDR_BACKORDER, Null ORDR_RELEASE_SHIP, Null WIP_IND from"))
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add Primary Key (ORDR_NO,ORDR_LNO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add ORDR_LAST_UNIT NUMBER(8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add ORDR_QTY_ALLO_CUR NUMBER(8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add ORDR_QTY_ALLO_FUT NUMBER(8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add ORDR_QTY_ALLO_CXL NUMBER(8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add ORDR_BACKORDER VARCHAR2(1)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add ORDR_RELEASE_SHIP DATE")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add WIP_IND VARCHAR2(1)")
        'If ASCMAIN1.CLIENT = "VAN" Then
        '    ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add SUB_BODY_CODE VARCHAR2(6)")
        '    ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add STANDARD_CUBE_PER_UNIT NUMBER(8,5)")
        'End If
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTORDR2 & "_1 on " & SOTORDR2 & " (STYLE_CODE,COLOR_CODE,WHSE_CODE)")
        ASCMAIN1.sql = "Update " & SOTORDR2 & " Set ORDR_RELEASE = NULL where ORDR_RELEASE = ''"
        If fill_data Then
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.AnalyzeTable(SOTORDR2)
        Else
            TABLE_NAMEs.Add("sqlSOTORDR2a", ASCMAIN1.sql)
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Select * from " & SOTORDR2 & " where ROWNUM < 1"
            ' EVENTUALLY, EVERYONE GETS THIS
            ' NO NEED TO FILL SOTORDR2 BECAUSE IT GETS FILLED BY WHSE/STYLE/COLOR IN THE ALLOCATION PROCESS
        Else
            ASCMAIN1.sql = "Select * from " & SOTORDR2
        End If

        frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("SOTORDR2"), SOTORDR2, "**", 0, , , , _
                   "ORDR_STATUS,ORDR_QTY_OPEN,ORDR_QTY_ALLO,ORDR_QTY_PICK,ORDR_QTY_CANC," _
                   & "ORDR_RELEASE,ORDR_RELEASE_AVAIL,ORDR_QTY_ALLO_CUR,ORDR_QTY_ALLO_FUT,ORDR_QTY_ALLO_CXL," _
                   & "ORDR_RELEASE_SHIP,ORDR_LAST_UNIT,ORDR_BACKORDER,WIP_IND")

        'frmASFBASE0.Create_Relation("SOTORDR1", "SOTORDR2", "ORDR_NO")
        'frmASFBASE0.dst.Tables("SOTORDR2").Columns.Add("WHSE_CODE", GetType(System.String), "PARENT(SOTORDR1_SOTORDR2).WHSE_CODE")

        ' SOTRSRV1

        ASCMAIN1.sql = "Select SOTRSRV1.* from SOTRSRV1 where RSRV_STATUS = 'O'" _
            & IIf(WHSE_CODE = "", _
                  "", _
                  "   and WHSE_CODE = '" & WHSE_CODE & "'")
        If Not fill_data Then ASCMAIN1.sql &= " and ROWNUM < 1"
        Dim SOTRSRV1 As String = ASCMAIN1.Temp_Table
        TABLE_NAMEs.Add("SOTRSRV1", SOTRSRV1)
        If Not fill_data Then
            TABLE_NAMEs.Add("sqlSOTRSRV1", "Insert into " & SOTRSRV1 & " " & Replace(Replace(ASCMAIN1.sql, " from", ", Null ORDR_REL_HOLD_CODES from"), " and ROWNUM < 1", " and RSRV_NO in (Select Distinct RSRV_NO from SOTRSRV2 where STYLE_CODE = 'STYLE_CODE')"))
        End If
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRSRV1 & " Add Primary Key (RSRV_NO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRSRV1 & " Add ORDR_REL_HOLD_CODES VARCHAR2(20)")

        ASCMAIN1.sql = "Select SOTRSRV1.* from " & SOTRSRV1 & " SOTRSRV1"
        frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("SOTRSRV1"), SOTRSRV1, "**", 0, , , , "ORDR_REL_HOLD_CODES")

        ' SOTRSRV2

        ASCMAIN1.sql = "Select SOTRSRV2.*" _
            & " from SOTRSRV2," & SOTRSRV1 & " SOTRSRV1 where SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO"
        If Not fill_data Then ASCMAIN1.sql &= " and ROWNUM < 1"
        Dim SOTRSRV2 As String = ASCMAIN1.Temp_Table
        TABLE_NAMEs.Add("SOTRSRV2", SOTRSRV2)
        If Not fill_data Then
            TABLE_NAMEs.Add("sqlSOTRSRV2", "Insert into " & SOTRSRV2 & " " & Replace(Replace(ASCMAIN1.sql, " from ", ", Null ORDR_RELEASE_AVAIL, 0 ORDR_QTY_ALLO_CUR, 0 ORDR_QTY_ALLO_FUT, 0 ORDR_QTY_ALLO_CXL from "), " and ROWNUM < 1", " and STYLE_CODE = 'STYLE_CODE'"))
        End If
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRSRV2 & " Add Primary Key (RSRV_NO,RSRV_LNO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTRSRV2 & "_1 on " & SOTRSRV2 & " (STYLE_CODE,COLOR_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRSRV2 & " Add ORDR_RELEASE_AVAIL DATE")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRSRV2 & " Add ORDR_QTY_ALLO_CUR NUMBER(8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRSRV2 & " Add ORDR_QTY_ALLO_FUT NUMBER(8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTRSRV2 & " Add ORDR_QTY_ALLO_CXL NUMBER(8,0)")
        ASCMAIN1.sql = "Select SOTRSRV2.* from " & SOTRSRV2 & " SOTRSRV2" & " where ROWNUM < 1"

        If Not frmASFBASE0.dst.Tables.Contains("SOTRSRV2") Then ' SSI already has an ADO.Net DataTable
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("SOTRSRV2"), SOTRSRV2, "**", 0, , , , _
           "ORDR_RELEASE_AVAIL,ORDR_QTY_ALLO_CUR,ORDR_QTY_ALLO_FUT,ORDR_QTY_ALLO_CXL")
        End If

        ' ARTCUST1

        Dim ARTCUST1_cols As String = "" _
            & "CUST_CODE,CUST_NAME,CUST_SALES_HOLD,CUST_CREDIT_LIMIT,CUST_CRED_LIMIT_EST,CUST_CRED_LIMIT_REV," _
            & "CUST_CREDIT_HOLD,CUST_ALLOW_BACKORDER,TERM_CODE,SREP_CODE,CUST_PO_REQD,CUST_BILL_TO_CUST," _
            & "TRADE_CLASS_CODE,CUST_PRIORITY_CODE,CUST_PD_GRACE_DAYS,CUST_CANCEL_GRACE_DAYS,CUST_ROUTING_INST," _
            & "CUST_SHIP_COMPLETE,CUST_SHIP_PCT_ORDER,CUST_SHIP_PCT_LINE,CUST_SHIP_BY_CASE,CUST_CART_REQD," _
            & "CUST_REL_EXPLICITLY,CUST_CREDIT_GROUP_CUST,CUST_CREDIT_RELEASE,CUST_INCL_INV_SHIP,CUST_CODE_ALLO"

        ASCMAIN1.sql = "Select " & ARTCUST1_cols & " from ARTCUST1 where CUST_CODE in (" & vbCrLf _
            & " (Select DISTINCT CUST_CODE from " & SOTORDR1 & ")" & vbCrLf _
            & "  union" & vbCrLf _
            & " (Select DISTINCT CUST_BILL_TO_CUST from " & SOTORDR1 & ")" & vbCrLf _
            & "  union" & vbCrLf _
            & " (Select DISTINCT CUST_CREDIT_GROUP_CUST from " & SOTORDR1 & ")" & vbCrLf _
            & "  union" & vbCrLf _
            & " (Select DISTINCT CUST_CODE from SOTRSRV1 where RSRV_STATUS = 'O')" & vbCrLf _
            & ")"

        Dim ARTCUST1 As String = ASCMAIN1.Temp_Table
        TABLE_NAMEs.Add("ARTCUST1", ARTCUST1)
        If Not fill_data Then
            TABLE_NAMEs.Add("sqlARTCUST1", "Insert into " & ARTCUST1 & " " & ASCMAIN1.sql)
        End If
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Add Primary Key (CUST_CODE)")

        ASCMAIN1.sql = "Update " & ARTCUST1 & " Set CUST_CREDIT_GROUP_CUST = CUST_BILL_TO_CUST where CUST_CREDIT_GROUP_CUST is NULL"
        If Not fill_data Then
            TABLE_NAMEs.Add("sqlARTCUST1a", ASCMAIN1.sql)
        Else
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Update " & ARTCUST1 & " Set CUST_CANCEL_GRACE_DAYS = NVL(CUST_CANCEL_GRACE_DAYS,0)"
        If Not fill_data Then
            TABLE_NAMEs.Add("sqlARTCUST1b", ASCMAIN1.sql)
        Else
            ASCDATA1.ExecuteSQL()
        End If

        If fill_data Then ASCMAIN1.AnalyzeTable(ARTCUST1)

        ASCMAIN1.sql = "Select * from " & ARTCUST1
        frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("ARTCUST1"), ARTCUST1, "**", 0, False)

        'SOTSUPP0

        ASCMAIN1.sql = "Select Distinct WHSE_CODE, STYLE_CODE, COLOR_CODE from ICTSTAT2" _
            & IIf(WHSE_CODE = "", "", " where WHSE_CODE = '" & WHSE_CODE & "'")
        frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "SOTSUPP0", "**", , False, , 3)
        frmASFBASE0.dst.Tables("SOTSUPP0").Columns.Add("HAS_DEMAND")

        'SOTSUPPI

        With frmASFBASE0.dst.Tables.Add("SOTSUPPI")
            .Columns.Add("INDEX", GetType(System.Int32))
            .Columns.Add("SUPPLY_DATE")
            .Columns.Add("SHIP_DATE")
            .Columns.Add("WIP_IND")
            .PrimaryKey = New DataColumn() {.Columns("INDEX")}
        End With

        If Not frmASFBASE0.dst.Tables.Contains("SOTORDR7") Then ' SSI already has an ADO.Net DataTable
            ASCMAIN1.sql = "Select SOTORDR7.* from SOTORDR7,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTORDR7.ORDR_GROUP_NO" _
                & IIf(WHSE_CODE = "", "", " and SOTORDR0.WHSE_CODE = '" & WHSE_CODE & "'")
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "SOTORDR7", "**", 0, True, "", 3) ' WHY DO i NEED TO SPECIFY 3?
        End If

        '  frmASFBASE0.dst.Tables("ICTSTDQ1").Columns.Add("QTY_ATS_CUM", GetType(System.Int64))

        Return TABLE_NAMEs
    End Function

    Public Shared Sub Allocation( _
        frmASFBASE0 As ASFBASE0, _
        force_pick As Boolean, _
        allocation_only As Boolean, _
        WHSE_CODE_to_allocate As String, _
        ORDR_GROUP_NOs As String, _
        edi850cust As List(Of String), _
       ByRef SOTSUPP1 As String, _
       ByRef SOTDEMD1 As String, _
       TABLE_NAMEs As Dictionary(Of String, String), _
        Optional read_only As Boolean = False, _
        Optional allocate_as_late_as_possible As Boolean = False, _
        Optional STYLE_CODE_to_Allocate As String = "", _
        Optional COLOR_CODE_to_Allocate As String = "", _
        Optional manual_release As Boolean = False, _
        Optional doNotAllocateToCreditHoldCustomers As Boolean = False, _
        Optional showProgress As Boolean = True)

        Dim SOTORDR0 As String = TABLE_NAMEs("SOTORDR0")
        Dim SOTORDR1 As String = TABLE_NAMEs("SOTORDR1")
        Dim SOTORDR2 As String = TABLE_NAMEs("SOTORDR2")
        Dim SOTRSRV1 As String = TABLE_NAMEs("SOTRSRV1")
        Dim SOTRSRV2 As String = TABLE_NAMEs("SOTRSRV2")
        Dim ARTCUST1 As String = TABLE_NAMEs("ARTCUST1")
        Dim ICTSTDQ1 As String = TABLE_NAMEs("ICTSTDQ1")
        Dim ICTSTDQ2 As String = TABLE_NAMEs("ICTSTDQ2")
        Dim ICTSTDQ3 As String = TABLE_NAMEs("ICTSTDQ3")

        If Not frmASFBASE0.ROWs.ContainsKey("SOTPARM1") Then frmASFBASE0.Get_PARM("SOTPARM1")
        Dim SO_PARM_SHIP_WINDOW_DAYS As Integer = Val(frmASFBASE0.ROWs("SOTPARM1").Item("SO_PARM_SHIP_WINDOW_DAYS") & "")
        Dim SO_PARM_ARRIVAL_BUFFER_DAYS As Integer = Val(frmASFBASE0.ROWs("SOTPARM1").Item("SO_PARM_ARRIVAL_BUFFER_DAYS") & "")
        Dim SO_PARM_RELEASE_AT_ONCE As String = frmASFBASE0.ROWs("SOTPARM1").Item("SO_PARM_RELEASE_AT_ONCE") & ""
        Dim SO_PARM_DAYS_ADJ As Integer = 0
        If SO_PARM_RELEASE_AT_ONCE = "1" Then
            SO_PARM_DAYS_ADJ = -1 * SO_PARM_ARRIVAL_BUFFER_DAYS
        End If

        Dim SOTORDRL As String = ""
        If TABLE_NAMEs.ContainsKey("SOTORDRL") Then
            SOTORDRL = TABLE_NAMEs("SOTORDRL")
        End If

        ' Perform Allocation

        If showProgress Then
            ASCMAIN1.Progress("Setting Up On Hand & PO (Supply)", "")
        End If

        If SOTSUPP1 = "" Then
            ASCMAIN1.sql = "Select * from SOTSUPP1 where ROWNUM < 1"
            SOTSUPP1 = ASCMAIN1.Temp_Table
            '  ASCDATA1.ExecuteSQL("Alter Table " & SOTSUPP1 & " Add Primary Key (WHSE_CODE, STYLE_CODE, COLOR_CODE, SUPPLY_DATE, SHIP_DATE)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSUPP1 & " Modify WHSE_CODE NOT NULL")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSUPP1 & " Modify STYLE_CODE NOT NULL")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSUPP1 & " Modify COLOR_CODE NOT NULL")

            ASCMAIN1.sql = "Select * from " & SOTSUPP1
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "SOTSUPP1", "**", , False, , 4)
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSUPP1)
        End If

        If Not frmASFBASE0.ROWs.ContainsKey("POTPARM1") Then frmASFBASE0.Get_PARM("POTPARM1")
        Dim PO_PARM_DEF_DAYS_ETA_TO_ARR As Int32 = Val(frmASFBASE0.ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR") & "")

        Dim sqlORDR_GROUP_NO_STYLE_CODEs As String = "(Select STYLE_CODE from SOTORDR2,SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO and SOTORDR1.ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & "))"
        Dim sqlORDR_GROUP_NO_WHSE_CODEs As String = "(Select WHSE_CODE from SOTORDR1 where SOTORDR1.ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & "))"

        If Not force_pick And Not manual_release Then
            ASCMAIN1.sql = "" _
                & "Select WHSE_CODE,STYLE_CODE,COLOR_CODE,RECORD_NO,SUPPLY_DATE,SHIP_DATE,WIP_IND" & vbCrLf _
                & ",SUM (SUPPLY_QTY) SUPPLY_QTY" & vbCrLf _
                & ",SUM (SUPPLY_QTY_ALLO) SUPPLY_QTY_ALLO" & vbCrLf _
                & ",SUM (SUPPLY_QTY_SHIP) SUPPLY_QTY_SHIP" & vbCrLf _
                & ",SUM (SUPPLY_QTY_ORDR) SUPPLY_QTY_ORDR" & vbCrLf _
                & ",ORDR_NO,ORDR_LNO,PO_ORDER_NO,PO_ORDER_LNO,SUPPLY_TYPE,PO_REFERENCE,PO_SPEC_ORDR_NO,PO_SHIP_VESSEL,PO_SHIP_REF_NO,PO_SHIP_ETA,DAYS" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select POTSHIP1.WHSE_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, NULL RECORD_NO" & vbCrLf _
                & ", TO_CHAR(POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0),'YYYYMMDD') SUPPLY_DATE" & vbCrLf _
                & ", TO_CHAR(POTSHIP1.PO_DATE_SHIPPED,'YYYYMMDD') SHIP_DATE" & vbCrLf _
                & ", 'X' WIP_IND" & vbCrLf _
                & ", POTSHIP3.PO_QTY_SHP SUPPLY_QTY" & vbCrLf _
                & ", 0 SUPPLY_QTY_ALLO" & vbCrLf _
                & ", POTSHIP3.PO_QTY_SHP SUPPLY_QTY_SHIP" & vbCrLf _
                & ", 0 SUPPLY_QTY_ORDR" & vbCrLf _
                & ", POTSHIP3.PO_SHIPMENT_NO ORDR_NO, POTSHIP3.PO_ORDER_LNO ORDR_LNO" & vbCrLf _
                & ", POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
                & ", 'S' AS SUPPLY_TYPE" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", DECODE (POTSHIP1.PO_SHIP_VESSEL,NULL,'InXit',POTSHIP1.PO_SHIP_VESSEL) PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_REF_NO" & vbCrLf _
                & ", POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                & ", NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) DAYS" & vbCrLf _
                & " from POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2, POTORDR1" & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and NVL(POTORDR1.FOB_CMT,'?') <> 'B'" & vbCrLf _
                & IIf(ORDR_GROUP_NOs <> "", _
                      "   and POTORDR2.STYLE_CODE in " & sqlORDR_GROUP_NO_STYLE_CODEs, _
                      IIf(STYLE_CODE_to_Allocate = "", "", "   and POTORDR2.STYLE_CODE = '" & STYLE_CODE_to_Allocate & "'" & vbCrLf)) _
                & IIf(ORDR_GROUP_NOs <> "", _
                      "   and POTSHIP1.WHSE_CODE in " & sqlORDR_GROUP_NO_WHSE_CODEs, _
                      IIf(WHSE_CODE_to_allocate = "", "", "   and POTSHIP1.WHSE_CODE = '" & WHSE_CODE_to_allocate & "'" & vbCrLf)) _
                & ") group by " & vbCrLf _
                & "WHSE_CODE,STYLE_CODE,COLOR_CODE,RECORD_NO,SUPPLY_DATE,SHIP_DATE,WIP_IND" & vbCrLf _
                & ",ORDR_NO,ORDR_LNO,PO_ORDER_NO,PO_ORDER_LNO,SUPPLY_TYPE,PO_REFERENCE,PO_SPEC_ORDR_NO,PO_SHIP_VESSEL,PO_SHIP_REF_NO,PO_SHIP_ETA,DAYS" & vbCrLf _
                & "  Union " & vbCrLf _
                & "Select WHSE_CODE,STYLE_CODE,COLOR_CODE,RECORD_NO,SUPPLY_DATE,SHIP_DATE,WIP_IND" & vbCrLf _
                & ",SUM (SUPPLY_QTY) SUPPLY_QTY" & vbCrLf _
                & ",SUM (SUPPLY_QTY_ALLO) SUPPLY_QTY_ALLO" & vbCrLf _
                & ",SUM (SUPPLY_QTY_SHIP) SUPPLY_QTY_SHIP" & vbCrLf _
                & ",SUM (SUPPLY_QTY_ORDR) SUPPLY_QTY_ORDR" & vbCrLf _
                & ",ORDR_NO,ORDR_LNO,PO_ORDER_NO,PO_ORDER_LNO,SUPPLY_TYPE,PO_REFERENCE,PO_SPEC_ORDR_NO,PO_SHIP_VESSEL,PO_SHIP_REF_NO,PO_SHIP_ETA,DAYS" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select POTORDR1.WHSE_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, NULL RECORD_NO" & vbCrLf _
                & ", TO_CHAR(POTORDR2.PO_DATE_ETA + " & CStr(PO_PARM_DEF_DAYS_ETA_TO_ARR) & ",'YYYYMMDD') SUPPLY_DATE" & vbCrLf _
                & ", TO_CHAR(POTORDR2.PO_DATE_SHIP_BY,'YYYYMMDD') SHIP_DATE" & vbCrLf _
                & ", 'X' WIP_IND" & vbCrLf _
                & ", POTORDR2.PO_QTY_OPN SUPPLY_QTY, 0 SUPPLY_QTY_ALLO" & vbCrLf _
                & ", 0 SUPPLY_QTY_SHIP" & vbCrLf _
                & ", POTORDR2.PO_QTY_ORD SUPPLY_QTY_ORDR" & vbCrLf _
                & ", Null ORDR_NO, Null ORDR_LNO" & vbCrLf _
                & ", POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
                & ", 'S' AS SUPPLY_TYPE" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", 'OpenPO' PO_SHIP_VESSEL, POTORDR1.PO_REFERENCE PO_SHIP_REF_NO" & vbCrLf _
                & ", POTORDR2.PO_DATE_ETA PO_SHIP_ETA" & vbCrLf _
                & ", " & CStr(PO_PARM_DEF_DAYS_ETA_TO_ARR) & " DAYS" & vbCrLf _
                & " from POTORDR1, POTORDR2" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR1.PO_STATUS = 'O'" & vbCrLf _
                & "   and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
                & "   and NVL(POTORDR1.FOB_CMT,'?') <> 'B'" & vbCrLf _
                & IIf(ORDR_GROUP_NOs <> "", _
                      "   and POTORDR2.STYLE_CODE in " & sqlORDR_GROUP_NO_STYLE_CODEs, _
                      IIf(STYLE_CODE_to_Allocate = "", "", "   and POTORDR2.STYLE_CODE = '" & STYLE_CODE_to_Allocate & "'" & vbCrLf)) _
                & IIf(ORDR_GROUP_NOs <> "", _
                      "   and POTORDR1.WHSE_CODE in " & sqlORDR_GROUP_NO_WHSE_CODEs, _
                      IIf(WHSE_CODE_to_allocate = "", "", "   and POTORDR1.WHSE_CODE = '" & WHSE_CODE_to_allocate & "'" & vbCrLf)) _
                & " and POTORDR2.PO_QTY_OPN > 0" & vbCrLf _
                & ") group by " & vbCrLf _
                & "WHSE_CODE,STYLE_CODE,COLOR_CODE,RECORD_NO,SUPPLY_DATE,SHIP_DATE,WIP_IND" & vbCrLf _
                & ",ORDR_NO,ORDR_LNO,PO_ORDER_NO,PO_ORDER_LNO,SUPPLY_TYPE,PO_REFERENCE,PO_SPEC_ORDR_NO,PO_SHIP_VESSEL,PO_SHIP_REF_NO,PO_SHIP_ETA,DAYS"

            ' PROBABLY SHOULD FILTER TO JUST THOSE STYLES IN SOTORDRL, BUT PROBABLY NOT WORTH THE EFFORT/CLAUSE

            ASCDATA1.ExecuteSQL("Insert into " & SOTSUPP1 & " " & ASCMAIN1.sql)

            'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            '    ASCDATA1.ExecuteSQL("Update " & SOTSUPP1 & " set WHSE_CODE = 'NJ' where WHSE_CODE = 'NJE'")
            'End If

            'ASCMAIN1.sql = "Update " & SOTSUPP1 & " Set SUPPLY_QTY_ALLO = 0"

            ' ADD WORK ORDERS
            ' NEED TO ADD WORK ORDERS WITH DATES LIKE POS - NOT JUST STAUS QTYS

            ASCMAIN1.sql = "Select WHSE_CODE, STYLE_CODE, COLOR_CODE, NULL RECORD_NO" _
                & ", '00000000' SUPPLY_DATE" _
                & ", '00000000' SHIP_DATE, 'X' WIP_IND" _
                & ", NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) - NVL(WHSE_QTY_COMM,0) SUPPLY_QTY, 0 SUPPLY_QTY_ALLO " _
                & ", 0 SUPPLY_QTY_SHIP, 0 SUPPLY_QTY_ORDR" _
                & ", 'OTSInv' ORDR_NO, Null ORDR_LNO" & vbCrLf _
                & ", Null PO_ORDER_NO, Null PO_ORDER_LNO" & vbCrLf _
                & ", 'H' AS SUPPLY_TYPE" & vbCrLf _
                & ", Null PO_REFERENCE, Null PO_SPEC_ORDR_NO" & vbCrLf _
                & ", 'OTSInv' PO_SHIP_VESSEL, Null PO_SHIP_REF_NO" & vbCrLf _
                & ", Null PO_DATE_ETA" & vbCrLf _
                & ", 0 DAYS" & vbCrLf _
                & " from ICTSTAT2" & vbCrLf _
                & " where (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) - NVL(WHSE_QTY_COMM,0)) <> 0" & vbCrLf _
                & IIf(ORDR_GROUP_NOs <> "", _
                      "   and ICTSTAT2.STYLE_CODE in " & sqlORDR_GROUP_NO_STYLE_CODEs, _
                      IIf(STYLE_CODE_to_Allocate = "", "", "   and ICTSTAT2.STYLE_CODE = '" & STYLE_CODE_to_Allocate & "'")) _
                & IIf(ORDR_GROUP_NOs <> "", _
                      "   and ICTSTAT2.WHSE_CODE in " & sqlORDR_GROUP_NO_WHSE_CODEs, _
                      IIf(WHSE_CODE_to_allocate = "", "", "   and ICTSTAT2.WHSE_CODE = '" & WHSE_CODE_to_allocate & "'"))

            'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            '    Dim sqlNJ As String = "" _
            '        & "from (SELECT DECODE(WHSE_CODE,'NJ','NJE',WHSE_CODE) WHSE_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            '        & ", SUM (NVL(WHSE_QTY_ON_HAND,0)) WHSE_QTY_ON_HAND" & vbCrLf _
            '        & ", SUM (NVL(WHSE_QTY_PICK,0)) WHSE_QTY_PICK" & vbCrLf _
            '        & "GROUP BY DECODE(WHSE_CODE,'NJ','NJE',WHSE_CODE), STYLE_CODE, COLOR_CODE) ICTSTAT2"
            '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "from ICTSTAT2", sqlNJ)
            'End If
            ASCDATA1.ExecuteSQL("Insert into " & SOTSUPP1 & " " & ASCMAIN1.sql)

            ASCDATA1.ExecuteSQL("Update " & SOTSUPP1 & " Set RECORD_NO = TRIM(TO_CHAR(ROWNUM, '0000000000'))")

            If STYLE_CODE_to_Allocate = "" Then ASCMAIN1.AnalyzeTable(SOTSUPP1)

            'Calculate WIP Indicator for Supply
            'P - Supply coming from PO
            'W - Supply is on Water (VAN only)
            'S - Supply coming from a PO Shipment
            'M - Supply has Multiple sources
            '0 - On Hand

            ASCMAIN1.sql = "Update " & SOTSUPP1 & " set WIP_IND = 'P' where SUPPLY_QTY_SHIP = 0 and SUPPLY_QTY_ORDR >= 0 and SUPPLY_TYPE <> 'H'" ' MAYBE SHOULD SAY SUPPLY_TYPE = 'S'?
            ASCDATA1.ExecuteSQL()
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                ASCMAIN1.sql = "Update " & SOTSUPP1 & " set WIP_IND = 'W' where SUPPLY_QTY_SHIP > 0 and  SUPPLY_QTY_ORDR <=0"
            Else
                ASCMAIN1.sql = "Update " & SOTSUPP1 & " set WIP_IND = 'S' where SUPPLY_QTY_SHIP > 0 and  SUPPLY_QTY_ORDR <=0"
            End If
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Update " & SOTSUPP1 & " Set WIP_IND = 'M' where STYLE_CODE, COLOR_CODE, SUPPLY_DATE in " _
            & "(Select STYLE_CODE, COLOR_CODE, SUPPLY_DATE from " & SOTSUPP1 _
            & " group by STYLE_CODE, COLOR_CODE, SUPPLY_DATE" _
            & " having Count (*) > 1)"



        '        If ASCMAIN1.CLIENT = "VAN" AndAlso ("ICFQUOTV,ICRISTA2,SORAVAL1,SORGOPN1,SORAVAL3,SORAVAL1".Contains(frmASFBASE0.Name)) Then
        If ASCMAIN1.CLIENT = "VAN" AndAlso frmASFBASE0.Name <> "SOROREL1" Then
            ' ********************
            '        & ", SUM (NVL(WHSE_QTY_COMM,0)) WHSE_QTY_COMM from ICTSTAT2" & vbCrLf _****** SPECIAL MERGE NJC -> NJE, reversed on 11/24
            ' ASCMAIN1.sql = "Update " & SOTSUPP1 & " Set WHSE_CODE = 'NJE' where WHSE_CODE = 'NJC'"

            'ASCMAIN1.sql = "Update " & SOTSUPP1 & " Set WHSE_CODE = 'NJC' where WHSE_CODE = 'NJE'"
            'ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS SELECT * FROM " & SOTSUPP1 & " WHERE WHSE_CODE = 'NJE'" & vbCrLf _
            & " FOR UPDATE;" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "IF R1.SUPPLY_DATE = '00000000' THEN" & vbCrLf _
            & "UPDATE " & SOTSUPP1 & " SET " & vbCrLf _
            & "  SUPPLY_QTY = NVL(SUPPLY_QTY,0) + NVL(R1.SUPPLY_QTY,0)" & vbCrLf _
            & ", SUPPLY_QTY_ALLO = NVL(SUPPLY_QTY_ALLO,0) + NVL(R1.SUPPLY_QTY_ALLO,0)" & vbCrLf _
            & ", SUPPLY_QTY_ORDR = NVL(SUPPLY_QTY_ORDR,0) + NVL(R1.SUPPLY_QTY_ORDR,0)" & vbCrLf _
            & ", SUPPLY_QTY_SHIP = NVL(SUPPLY_QTY_SHIP,0) + NVL(R1.SUPPLY_QTY_SHIP,0)" & vbCrLf _
            & " WHERE WHSE_CODE = 'NJC' AND STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE AND SUPPLY_DATE = R1.SUPPLY_DATE;" & vbCrLf _
            & "IF SQL%NOTFOUND THEN" & vbCrLf _
            & " UPDATE " & SOTSUPP1 & " SET WHSE_CODE = 'NJC' WHERE CURRENT OF C1;" & vbCrLf _
            & "ELSE" & vbCrLf _
            & " DELETE FROM " & SOTSUPP1 & " WHERE CURRENT OF C1;" & vbCrLf _
            & "END IF;" & vbCrLf _
            & "ELSE" & vbCrLf _
            & " UPDATE " & SOTSUPP1 & " SET WHSE_CODE = 'NJC' WHERE CURRENT OF C1;" & vbCrLf _
            & "END IF;" & vbCrLf _
            & "END LOOP; END; END;"
            ASCDATA1.ExecuteSQL()
        End If







        If showProgress Then
            ASCMAIN1.Progress("Now Setting up Demand Data", "")
        End If

        If Not read_only Then
            ASCMAIN1.sql = "Update " & SOTORDR2 & " Set ORDR_QTY_ALLO = 0 "
            If Not allocation_only Then
                ASCMAIN1.sql &= ", ORDR_RELEASE = NULL, ORDR_RELEASE_AVAIL = NULL"
            End If
            ASCDATA1.ExecuteSQL()
        End If

        frmASFBASE0.Fill_Records("ARTCUST1")
        'frmASFBASE0.Fill_Records("ARTCUST1", "", True, "Select * from " & ARTCUST1)
        'For Each rowARTCUST1 As DataRow In frmASFBASE0.dst.Tables("ARTCUST1").Select("CUST_CANCEL_GRACE_DAYS IS NULL")
        '    rowARTCUST1.Item("CUST_CANCEL_GRACE_DAYS") = 0
        'Next

        'frmASFBASE0.Fill_Records("SOTRSRV1")
        'frmASFBASE0.Fill_Records("SOTRSRV2")

        If SOTDEMD1 = "" Then
            ASCMAIN1.sql = "Select * from SOTDEMD1 where ROWNUM < 1"
            SOTDEMD1 = ASCMAIN1.Temp_Table
            ' ASCDATA1.ExecuteSQL("Alter Table " & SOTSUPP1 & " Add Primary Key (STYLE_CODE, COLOR_CODE, SUPPLY_DATE, SHIP_DATE)")
            ASCDATA1.ExecuteSQL("Create Index I_" & SOTDEMD1 & "_1 on " & SOTDEMD1 & " (WHSE_CODE,STYLE_CODE,COLOR_CODE,ORDR_GROUP_NO)")
            ASCDATA1.ExecuteSQL("Create Index I_" & SOTDEMD1 & "_2 on " & SOTDEMD1 & " (ORDR_GROUP_NO,STYLE_CODE,COLOR_CODE,DEMAND_TYPE)")
            '            ASCDATA1.ExecuteSQL("Create Index I_" & SOTDEMD1 & "_3 on " & SOTDEMD1 & " (STYLE_CODE,COLOR_CODE,ORDR_GROUP_NO)")

            ASCMAIN1.sql = "Select * from " & SOTDEMD1
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "SOTDEMD1", "**", , False, , 0)
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTDEMD1)
        End If

        Dim WHSE_CODEs_which_are_ports As String = ""
        ASCMAIN1.sql = "Select WHSE_CODE from ICTWHSE1 where WHSE_TYPE = 'P'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            WHSE_CODEs_which_are_ports &= ",'" & row.Item(0) & "'"
        Next

        Dim sql_no_BTB_to_ports As String = ""
        If WHSE_CODEs_which_are_ports <> "" Then
            sql_no_BTB_to_ports = "   and NOT (NVL(SOTORDR1.ORDR_TYPE_CODE,'?') ='BTB' and SOTORDR1.WHSE_CODE in (" & Mid(WHSE_CODEs_which_are_ports, 2) & "))" & vbCrLf
        End If

        ASCMAIN1.sql = "Select SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", NVL(SOTORDR1.ORDR_PRIORITY,NVL(ARTCUST1.CUST_PRIORITY_CODE,'9')) as ORDR_PRIORITY" & vbCrLf _
            & ", (SOTORDR1.ORDR_CANCEL_DATE + NVL(ARTCUST1.CUST_CANCEL_GRACE_DAYS,0)) as ORDR_DEMAND_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_SHIP_DATE" & vbCrLf _
            & ", TRUNC(SOTORDR1.INIT_DATE) as ORDR_PRIORITY_DATE" & vbCrLf _
            & ", 'O' as DEMAND_TYPE" & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", SOTORDR1.CUST_CODE" & vbCrLf _
            & ", SOTORDR2.ORDR_NO" & vbCrLf _
            & ", SOTORDR2.ORDR_LNO" & vbCrLf _
            & ", ARTCUST1.CUST_CANCEL_GRACE_DAYS" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_OPEN" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO" & vbCrLf _
            & ", TRUNC(SOTORDR1.INIT_DATE) as ORDR_PRIORITY_DATE_ORIG" & vbCrLf _
            & ", 0 ORDR_QTY_ALLO" & vbCrLf _
            & ", 0 ORDR_QTY_ALLO_CUR" & vbCrLf _
            & ", 0 ORDR_QTY_ALLO_FUT" & vbCrLf _
            & ", SOTORDR2.ORDR_RELEASE_AVAIL" & vbCrLf _
            & ", 0 ORDR_LAST_UNIT" & vbCrLf _
            & ", NULL ORDR_BACKORDER" & vbCrLf _
            & ", NULL ORDR_RELEASE_SHIP" & vbCrLf _
            & ", NULL WIP_IND" & vbCrLf _
            & ", NULL ORDR_RELEASE" & vbCrLf _
            & ", 0 ORDR_QTY_ALLO_CXL" & vbCrLf _
            & ", SOTORDR1.INIT_DATE" & vbCrLf _
            & IIf(SO_PARM_RELEASE_AT_ONCE = "1",
                  ", SOTORDR1.ORDR_SHIP_DATE + " & CStr(SO_PARM_SHIP_WINDOW_DAYS) & " ORDR_SHIP_DATE_PLUS" & vbCrLf _
                & ", ICTSTYL1.STYLE_CLASS_CODE STYLE_CLASS_CODE" & vbCrLf _
                & ", DECODE(SOTORDR1.WHSE_CODE,'MS',ICTCLAS1.STYLE_CLASS_RELEASE_ATONCE,NULL) ATONCE" & vbCrLf _
                & ", TO_CHAR(SOTORDR1.ORDR_SHIP_DATE + " & CStr(SO_PARM_SHIP_WINDOW_DAYS) & ",'YYYYMMDD') ATONCE_DATE" & vbCrLf,
                  ", SOTORDR1.ORDR_SHIP_DATE ORDR_SHIP_DATE_PLUS" & vbCrLf _
                & ", NULL STYLE_CLASS_CODE" & vbCrLf _
                & ", '0' ATONCE" & vbCrLf _
                & ", '00000000' ATONCE_DATE" & vbCrLf
                      ) _
            & " from " & SOTORDR2 & " SOTORDR2," & SOTORDR1 & " SOTORDR1," & ARTCUST1 & " ARTCUST1" & vbCrLf _
            & IIf(SO_PARM_RELEASE_AT_ONCE = "1", ", ICTSTYL1, ICTCLAS1" & vbCrLf, "") _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & IIf(SO_PARM_RELEASE_AT_ONCE = "1", " and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE and ICTCLAS1.STYLE_CLASS_CODE (+) = ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf, "") _
            & sql_no_BTB_to_ports _
            & "   and SOTORDR2.ORDR_QTY_OPEN <> 0" & vbCrLf _
            & IIf(ORDR_GROUP_NOs <> "",
                      "   and SOTORDR2.STYLE_CODE in " & sqlORDR_GROUP_NO_STYLE_CODEs,
                      IIf(STYLE_CODE_to_Allocate = "", "", "   and SOTORDR2.STYLE_CODE = '" & STYLE_CODE_to_Allocate & "'" & vbCrLf)) _
            & IIf(ORDR_GROUP_NOs <> "",
                      "   and SOTORDR1.WHSE_CODE in " & sqlORDR_GROUP_NO_WHSE_CODEs,
                      IIf(WHSE_CODE_to_allocate = "", "", "   and SOTORDR1.WHSE_CODE = '" & WHSE_CODE_to_allocate & "'")) _
            & IIf(SOTORDRL <> "",
                      "   and (SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE) in (Select Distinct STYLE_CODE, COLOR_CODE from " & SOTORDRL & ")",
                      "")

        ' THE NEXT FEW LINES WOULD PREVENT ORDERS PAST CANCEL FROM SOAKING UP INVENTORY
        ' WHY ARE WE NOT CONSIDERING THIS?

        '    & " and SOWORDR1.ORDR_CANCEL_DATE "
        '    & "   + ARWCUST1.CUST_CANCEL_GRACE_DAYS "
        '    & "   - 0 < #" & Format$(Now + NowTSD, "mm/dd/yy") & "#"

        ASCDATA1.ExecuteSQL("Insert into " & SOTDEMD1 & " " & ASCMAIN1.sql)


        ' THERE ARE 2 DIFFERENT TYPES OF MANUAL RELEASE
        ' WHEN USING THE MANUAL RELEASE SCREEN, THE PRE_ALLO FIELDS ARE SET
        ' WHEN USING THE SOFCORD1 RIGHT CLICK TO SELECT - THE PRE_ALLO IS NOT SET, AND THE ORDER IS RELEASED AS IF IT IS THE ONLY ORDER TO RELEASE

        If manual_release Then
            ASCMAIN1.sql = "Update " & SOTORDR2 _
                & "  Set ORDR_QTY_ALLO = ORDR_QTY_PRE_ALLO" _
                & ", ORDR_QTY_ALLO_CUR = ORDR_QTY_PRE_ALLO" _
                & ", ORDR_QTY_ALLO_FUT = 0" _
                & ", ORDR_QTY_ALLO_CXL = 0" _
                & ", ORDR_RELEASE = Null" _
                & ", ORDR_RELEASE_AVAIL = Null" _
                & ", ORDR_RELEASE_SHIP = Null" _
                & ", WIP_IND = NULL" _
                & ", ORDR_BACKORDER = 'Y'" _
                & ", ORDR_LAST_UNIT = NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is " & vbCrLf _
                & " Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.ORDR_QTY_PRE_ALLO " & vbCrLf _
                & "   from SOTORDR2,SOTORDR1 " & vbCrLf _
                & "  where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO " & vbCrLf _
                & "    and SOTORDR1.ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & ") and SOTORDR2.ORDR_QTY_PRE_ALLO <> 0;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTDEMD1 & " Set ORDR_QTY_ALLO = R1.ORDR_QTY_PRE_ALLO " & vbCrLf _
                & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        If STYLE_CODE_to_Allocate = "" Then ASCMAIN1.AnalyzeTable(SOTDEMD1)

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is" _
            & "  Select * from SOTORDR7 " _
            & "   where ORDR_GROUP_NO in (Select ORDR_GROUP_NO from " & SOTORDR0 & ")" _
            & "     and PICK_BATCH_NO is Null;" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   If R1.ORDR_PRIORITY is Not Null Then" _
            & "    Update " & SOTDEMD1 & " Set ORDR_PRIORITY = R1.ORDR_PRIORITY" _
            & "     where ORDR_GROUP_NO = R1.ORDR_GROUP_NO" _
            & "       and STYLE_CODE = R1.STYLE_CODE" _
            & "       and COLOR_CODE = R1.COLOR_CODE;" _
            & "   End If;" _
            & "   If R1.ORDR_PRIORITY_DATE is Not Null Then" _
            & "    Update " & SOTDEMD1 & " Set ORDR_PRIORITY_DATE = R1.ORDR_PRIORITY_DATE" _
            & "     where ORDR_GROUP_NO = R1.ORDR_GROUP_NO" _
            & "       and STYLE_CODE = R1.STYLE_CODE" _
            & "       and COLOR_CODE = R1.COLOR_CODE;" _
            & "   End If;" _
            & "   If R1.ORDR_DEMAND_DATE is Not Null Then" _
            & "    Update " & SOTDEMD1 & " Set ORDR_DEMAND_DATE = R1.ORDR_DEMAND_DATE" _
            & "     where ORDR_GROUP_NO = R1.ORDR_GROUP_NO" _
            & "       and STYLE_CODE = R1.STYLE_CODE" _
            & "       and COLOR_CODE = R1.COLOR_CODE;" _
            & "   End If;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select SOTRSRV1.WHSE_CODE, SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE" & vbCrLf _
            & ", NVL(SOTRSRV2.RSRV_PRIORITY,NVL(ARTCUST1.CUST_PRIORITY_CODE,'9')) ORDR_PRIORITY" & vbCrLf _
            & ", NVL(SOTRSRV2.RSRV_DEMAND_DATE, (SOTRSRV1.ORDR_CANCEL_DATE + NVL(ARTCUST1.CUST_CANCEL_GRACE_DAYS,0))) ORDR_DEMAND_DATE" & vbCrLf _
            & ", SOTRSRV1.ORDR_CANCEL_DATE" & vbCrLf _
            & ", SOTRSRV1.ORDR_SHIP_DATE" & vbCrLf _
            & ", NVL(SOTRSRV2.RSRV_PRIORITY_DATE,TRUNC(SOTRSRV1.INIT_DATE)) ORDR_PRIORITY_DATE" & vbCrLf _
            & ", 'R' as DEMAND_TYPE" & vbCrLf _
            & ", SOTRSRV2.RSRV_NO as ORDR_GROUP_NO" & vbCrLf _
            & ", SOTRSRV1.CUST_CODE" & vbCrLf _
            & ", SOTRSRV2.RSRV_NO as ORDR_NO" & vbCrLf _
            & ", SOTRSRV2.RSRV_LNO as ORDR_LNO" & vbCrLf _
            & ", ARTCUST1.CUST_CANCEL_GRACE_DAYS" & vbCrLf _
            & ", SOTRSRV2.RSRV_QTY_OPEN as ORDR_QTY_OPEN " & vbCrLf _
            & ", SOTRSRV1.ORDR_CUST_PO" & vbCrLf _
            & ", NVL(SOTRSRV2.RSRV_PRIORITY_DATE,TRUNC(SOTRSRV1.INIT_DATE)) as ORDR_PRIORITY_DATE_ORIG" & vbCrLf _
            & ", 0 ORDR_QTY_ALLO" & vbCrLf _
            & ", 0 ORDR_QTY_ALLO_CUR" & vbCrLf _
            & ", 0 ORDR_QTY_ALLO_FUT" & vbCrLf _
            & ", SYSDATE ORDR_RELEASE_AVAIL" & vbCrLf _
            & ", 0 ORDR_LAST_UNIT" & vbCrLf _
            & ", NULL ORDR_BACKORDER" & vbCrLf _
            & ", NULL ORDR_RELEASE_SHIP" & vbCrLf _
            & ", NULL WIP_IND" & vbCrLf _
            & ", NULL ORDR_RELEASE" & vbCrLf _
            & ", 0 ORDR_QTY_ALLO_CXL" & vbCrLf _
            & ", SOTRSRV1.INIT_DATE" & vbCrLf _
            & IIf(SO_PARM_RELEASE_AT_ONCE = "1", _
                ", SOTRSRV1.ORDR_SHIP_DATE + " & CStr(SO_PARM_SHIP_WINDOW_DAYS) & " ORDR_SHIP_DATE_PLUS" & vbCrLf _
              & ", ICTSTYL1.STYLE_CLASS_CODE STYLE_CLASS_CODE" & vbCrLf _
              & ", ICTCLAS1.STYLE_CLASS_RELEASE_ATONCE ATONCE" & vbCrLf _
              & ", TO_CHAR(SOTRSRV1.ORDR_SHIP_DATE + " & CStr(SO_PARM_SHIP_WINDOW_DAYS) & ",'YYYYMMDD') ATONCE_DATE" & vbCrLf, _
                ", SOTRSRV1.ORDR_SHIP_DATE ORDR_SHIP_DATE_PLUS" & vbCrLf _
              & ", NULL STYLE_CLASS_CODE" & vbCrLf _
              & ", '0' ATONCE" & vbCrLf _
              & ", '00000000' ATONCE_DATE" & vbCrLf _
                    ) _
            & " from SOTRSRV2,SOTRSRV1,ARTCUST1" & vbCrLf _
            & IIf(SO_PARM_RELEASE_AT_ONCE = "1", ", ICTSTYL1, ICTCLAS1" & vbCrLf, "") _
            & " where SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTRSRV1.CUST_CODE" & vbCrLf _
            & IIf(SO_PARM_RELEASE_AT_ONCE = "1", " and ICTSTYL1.STYLE_CODE = SOTRSRV2.STYLE_CODE and ICTCLAS1.STYLE_CLASS_CODE (+) = ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf, "") _
            & "   and SOTRSRV2.RSRV_QTY_OPEN <> 0" & vbCrLf _
            & IIf(ORDR_GROUP_NOs <> "", _
                      "   and SOTRSRV2.STYLE_CODE in " & sqlORDR_GROUP_NO_STYLE_CODEs, _
                      IIf(STYLE_CODE_to_Allocate = "", "", "   and SOTRSRV2.STYLE_CODE = '" & STYLE_CODE_to_Allocate & "'" & vbCrLf)) _
            & IIf(ORDR_GROUP_NOs <> "", _
                      "   and SOTRSRV1.WHSE_CODE in " & sqlORDR_GROUP_NO_WHSE_CODEs, _
                      IIf(WHSE_CODE_to_allocate = "", "", "   and SOTRSRV1.WHSE_CODE = '" & WHSE_CODE_to_allocate & "'"))

        ASCDATA1.ExecuteSQL("Insert into " & SOTDEMD1 & " " & ASCMAIN1.sql)

        'If SO_PARM_RELEASE_AT_ONCE = "1" Then
        '    ASCMAIN1.sql = "Update " & SOTDEMD1 & " Set ORDR_PRIORITY_DATE = TRUNC(ORDR_SHIP_DATE_PLUS)" & vbCrLf _
        '        & "    where ATONCE = '1'"
        '    ASCDATA1.ExecuteSQL()
        'End If

        If force_pick Or manual_release Then
            ASCMAIN1.sql = "Update " & SOTDEMD1 & " Set ORDR_PRIORITY = '0'" _
                & " where ORDR_GROUP_NO in (" & ORDR_GROUP_NOs & ")"
            ASCDATA1.ExecuteSQL()
        End If

        '        If ASCMAIN1.CLIENT = "VAN" AndAlso ("ICFQUOTV,ICRISTA2,SORAVAL1,SORGOPN1,SORAVAL3,SORAVAL1".Contains(frmASFBASE0.Name)) Then
        If ASCMAIN1.CLIENT = "VAN" AndAlso frmASFBASE0.Name <> "SOROREL1" Then
            ' ************************** SPECIAL MERGE NJC -> NJE, reversed on 11/24
            ASCMAIN1.sql = "Update " & SOTDEMD1 & " Set WHSE_CODE = 'NJC' where WHSE_CODE = 'NJE'"
            ASCDATA1.ExecuteSQL()
        End If


        ASCMAIN1.sql = "Select Distinct WHSE_CODE, STYLE_CODE, COLOR_CODE, HAS_DEMAND from (" & vbCrLf _
            & "Select Distinct WHSE_CODE, STYLE_CODE, COLOR_CODE, '0' HAS_DEMAND from " & SOTSUPP1 & vbCrLf _
            & " union " & vbCrLf _
            & "Select Distinct WHSE_CODE, STYLE_CODE, COLOR_CODE, '0' HAS_DEMAND from " & SOTDEMD1 & vbCrLf _
            & ")"
        '    frmASFBASE0.dst.Tables("SOTSUPP0").Rows.Clear()
        frmASFBASE0.Fill_Records("SOTSUPP0", , , ASCMAIN1.sql)

        Dim WSC As String = ""

        ASCMAIN1.sql = "Select * from " & SOTSUPP1
        frmASFBASE0.Fill_Records("SOTSUPP1", , , ASCMAIN1.sql)
        ' DON'T KNOW WHY WE EVEN NEED THIS DATATABLE SINCE THE DEMAND LOOP BELOW USES AN ADHOC SQL
        'ASCMAIN1.sql = "Select * from " & SOTDEMD1
        'frmASFBASE0.Fill_Records("SOTDEMD1", , , ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from SOTORDR7 where ORDR_GROUP_NO in (Select ORDR_GROUP_NO from " & SOTORDR0 & ")"
        frmASFBASE0.Fill_Records("SOTORDR7", , , ASCMAIN1.sql)




        ' Allocation Main Process

        If showProgress Then
            ASCMAIN1.Progress("Now Allocating", "")
        End If

        ASCMAIN1.sql = "Select SOTDEMD1.WHSE_CODE, SOTDEMD1.STYLE_CODE, SOTDEMD1.COLOR_CODE" & vbCrLf _
            & ", SOTDEMD1.STYLE_CLASS_CODE, SOTDEMD1.ATONCE" & vbCrLf _
            & ", SOTDEMD1.DEMAND_TYPE, SOTDEMD1.ORDR_GROUP_NO, SOTDEMD1.CUST_CODE" & vbCrLf _
            & ", Case when SOTDEMD1.DEMAND_TYPE = 'R' Then SOTDEMD1.ORDR_LNO else 0 End as ORDR_LNO" & vbCrLf _
            & ", Sum (SOTDEMD1.ORDR_QTY_OPEN) as ORDR_QTY_OPEN" & vbCrLf _
            & ", Min (SOTDEMD1.ORDR_PRIORITY) as ORDR_PRIORITY" & vbCrLf _
            & ", Min (SOTDEMD1.ORDR_PRIORITY_DATE) as ORDR_PRIORITY_DATE" & vbCrLf _
            & ", Min (SOTDEMD1.ORDR_DEMAND_DATE) as ORDR_DEMAND_DATE" & vbCrLf _
            & ", Min (SOTDEMD1.ORDR_SHIP_DATE) as ORDR_SHIP_DATE" & vbCrLf _
            & ", Min (SOTDEMD1.INIT_DATE) as INIT_DATE" & vbCrLf _
            & ", Min (SOTDEMD1.ORDR_SHIP_DATE_PLUS) as ORDR_SHIP_DATE_PLUS" & vbCrLf _
            & ", Min (SOTDEMD1.ATONCE_DATE) as ATONCE_DATE" & vbCrLf _
            & " from " & SOTDEMD1 & " SOTDEMD1" & vbCrLf _
            & " group by SOTDEMD1.WHSE_CODE, SOTDEMD1.STYLE_CODE, SOTDEMD1.COLOR_CODE" & vbCrLf _
            & ", SOTDEMD1.STYLE_CLASS_CODE, SOTDEMD1.ATONCE" & vbCrLf _
            & ", SOTDEMD1.DEMAND_TYPE, SOTDEMD1.ORDR_GROUP_NO, SOTDEMD1.CUST_CODE, " & vbCrLf _
            & " Case when SOTDEMD1.DEMAND_TYPE = 'R' Then SOTDEMD1.ORDR_LNO else 0 End "
        Dim sqlDemand As String = ASCMAIN1.sql

        Dim WHSE_CODE As String = ""
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        Dim SQ(,) As Int64 = Nothing       ' Table of Supply Qty's for each of the Dates
        Dim imax As Integer = 0
        Dim imax_save As Integer = 0
        Dim DEMAND_TYPE_O As Integer = 0
        Dim DEMAND_TYPE_R As Integer = 0

        Dim order_by As String = ",ORDR_SHIP_DATE,DEMAND_TYPE"
        If Not frmASFBASE0.ROWs.ContainsKey("SOTPARM1") Then frmASFBASE0.Get_PARM("SOTPARM1")
        Dim SO_PARM_ALLO_SEQ As String = frmASFBASE0.ROWs("SOTPARM1").Item("SO_PARM_ALLO_SEQ") & ""
        If SO_PARM_ALLO_SEQ = "" Then SO_PARM_ALLO_SEQ = "0"
        If SO_PARM_ALLO_SEQ = "0" Then
            order_by = ",ORDR_PRIORITY_DATE,ORDR_GROUP_NO,ORDR_DEMAND_DATE" & order_by
        ElseIf SO_PARM_ALLO_SEQ = "1" Then
            order_by = ",ORDR_DEMAND_DATE,ORDR_GROUP_NO,ORDR_PRIORITY_DATE" & order_by
        ElseIf SO_PARM_ALLO_SEQ = "2" Then
            order_by = ",INIT_DATE" & order_by
        Else
            MsgBox("Cannot Determine Allocation Sequence")
        End If

        Dim order_by_at_once As String = ""

        If ASCMAIN1.CLIENT = "RGI" Then
            If SO_PARM_RELEASE_AT_ONCE = "1" Then
                order_by_at_once = ",ATONCE,ATONCE_DATE"
            End If
        End If

        order_by_at_once = ""

        If SO_PARM_ALLO_SEQ = "2" Then
            order_by = "WHSE_CODE,STYLE_CODE,COLOR_CODE" & order_by_at_once & order_by
        Else
            order_by = "WHSE_CODE,STYLE_CODE,COLOR_CODE,ORDR_PRIORITY" & order_by_at_once & order_by
        End If

        Dim rowICTSTYL1 As DataRow = Nothing
        Dim STYLE_CODE_ICTSTYL1 As String = ""
        Dim INNER_PACK_QTY As Int32 = 0
        Dim ATONCE As String = ""

        If manual_release Then
            ' ALREADY RELEASED
        Else

            Dim TBL As DataTable = ASCDATA1.GetDataTable()

            If ASCMAIN1.CLIENT = "RGI" Then
                If SO_PARM_RELEASE_AT_ONCE = "1" Then

                    ASCMAIN1.sql = "Select WHSE_CODE, STYLE_CODE, COLOR_CODE, SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                        & " from (" & Replace(sqlDemand, " group by", " where ATONCE = '1'" & vbCrLf & " group by") & ") where WHSE_CODE = 'MS' group by WHSE_CODE, STYLE_CODE, COLOR_CODE"
                    ASCMAIN1.sql = "Select X.*" & vbCrLf _
                        & ", NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) QTY_AVA" & vbCrLf _
                        & ", NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) QTY_SUP" & vbCrLf _
                        & " from ICTSTAT2, (" & ASCMAIN1.sql & ") X" & vbCrLf _
                        & " where ICTSTAT2.WHSE_CODE (+) = X.WHSE_CODE" & vbCrLf _
                        & "   and ICTSTAT2.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTAT2.COLOR_CODE (+) = X.COLOR_CODE"
                    ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") X where X.QTY_AVA < ORDR_QTY_OPEN"
                    Dim TBLX As DataTable = ASCDATA1.GetDataTable()

                    For Each ROW As DataRow In TBLX.Select("") ' For Each row As DataRow In ASCDATA1.SelectDistinct(TBL, New String() {"WHSE_CODE", "STYLE_CODE", "COLOR_CODE"}).Select("WHSE_CODE = 'MS'")
                        Dim WHSE_CODE_X As String = ROW.Item("WHSE_CODE")
                        Dim STYLE_CODE_X As String = ROW.Item("STYLE_CODE")
                        Dim COLOR_CODE_X As String = ROW.Item("COLOR_CODE")
                        Dim QTY_AVA As Integer = Val(ROW.Item("QTY_AVA") & "")
                        Dim QTY_SUP As Integer = Val(ROW.Item("QTY_SUP") & "")
                        Dim ORDR_QTY_OPEN As Integer = Val(ROW.Item("ORDR_QTY_OPEN") & "")
                        Dim sqlwx As String = "WHSE_CODE = '" & WHSE_CODE_X & "' and STYLE_CODE = '" & STYLE_CODE_X & "' and COLOR_CODE = '" & COLOR_CODE_X & "'"
                        Dim seqx As String = Replace(order_by, order_by_at_once, "")
                        Dim QTYX As Integer = 0
                        For Each row2 As DataRow In TBL.Select(sqlwx, seqx)
                            Dim ORDR_QTY_OPEN2 As Integer = Val(row2.Item("ORDR_QTY_OPEN") & "")
                            QTYX += ORDR_QTY_OPEN2
                            If QTYX > QTY_SUP Then
                                row2.Item("ATONCE") = "2"
                            End If
                        Next
                    Next
                End If
            End If

            'If ASCMAIN1.Running_in_VS Then
            '    Using FRM As New ASFMSGBF
            '        FRM.Show_grd(TBL, frmASFBASE0, "Demand Table")
            '    End Using
            'End If

            '    Debug.Print("ORDER BY  = " & order_by)
            '   Debug.Print("CUST_CODE" & vbTab & "ORDR_QTY_OPEN" & vbTab & "ATONCE" & vbTab & "ATONCE_DATE" & vbTab & "INIT_DATE" & vbTab & "ORDR_SHIP_DATE" & vbTab & "DEMAND_TYPE" & vbTab & "ORDR_PRIORITY_DATE")
            For Each rowSOTDEMDX As DataRow In TBL.Select("", order_by) ' ASCDATA1.GetDataTable.Select("", order_by)
                Dim DEMAND_TYPE As String = rowSOTDEMDX.Item("DEMAND_TYPE")
                If DEMAND_TYPE = "O" Then
                    DEMAND_TYPE_O += 1
                Else
                    DEMAND_TYPE_R += 1
                End If
                ' WHSE_CODE,STYLE_CODE,COLOR_CODE,ATONCE,ATONCE_DATE,INIT_DATE,ORDR_SHIP_DATE,DEMAND_TYPE
                'Debug.Print(rowSOTDEMDX.Item("CUST_CODE") & vbTab & rowSOTDEMDX.Item("ORDR_QTY_OPEN") _
                '            & vbTab & rowSOTDEMDX.Item("ATONCE") & vbTab & rowSOTDEMDX.Item("ATONCE_DATE") _
                '            & vbTab & rowSOTDEMDX.Item("INIT_DATE") & vbTab & rowSOTDEMDX.Item("ORDR_SHIP_DATE") _
                '            & vbTab & rowSOTDEMDX.Item("DEMAND_TYPE") & vbTab & rowSOTDEMDX.Item("ORDR_PRIORITY_DATE"))

                If STYLE_CODE_ICTSTYL1 <> rowSOTDEMDX.Item("STYLE_CODE") & "" Then
                    STYLE_CODE_ICTSTYL1 = rowSOTDEMDX.Item("STYLE_CODE") & ""

                    If frmASFBASE0.dst.Tables.Contains("ICTSTYL1") Then
                        rowICTSTYL1 = frmASFBASE0.dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE_ICTSTYL1)
                    Else
                        rowICTSTYL1 = Nothing
                    End If
                    ' rowICTSTYL1 = frmASFBASE0.dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE_ICTSTYL1)
                    If rowICTSTYL1 Is Nothing Then
                        rowICTSTYL1 = frmASFBASE0.LookUp("ICTSTYL1", STYLE_CODE_ICTSTYL1)
                    End If
                    INNER_PACK_QTY = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")
                End If

                If WSC <> rowSOTDEMDX.Item("WHSE_CODE") & "-" & rowSOTDEMDX.Item("STYLE_CODE") & "-" & rowSOTDEMDX.Item("COLOR_CODE") Then
                    If WSC <> "" Then

                        Move_Allocations_to_Earliest_Supply_Date(frmASFBASE0, SOTORDR2, WHSE_CODE, STYLE_CODE, COLOR_CODE, SQ, read_only, (allocate_as_late_as_possible Or ATONCE = "1"), SO_PARM_DAYS_ADJ)
                    End If
                    WHSE_CODE = rowSOTDEMDX.Item("WHSE_CODE")
                    STYLE_CODE = rowSOTDEMDX.Item("STYLE_CODE")
                    COLOR_CODE = rowSOTDEMDX.Item("COLOR_CODE")
                    ATONCE = rowSOTDEMDX.Item("ATONCE") & ""
                    'If COLOR_CODE = "NAGD" Then Stop
                    ' If ASCMAIN1.Running_in_VS And (STYLE_CODE = "MTX59907" Or STYLE_CODE = "MTX31805" Or STYLE_CODE = "MTF22249") Then Stop

                    SQ = Setup_WSC_Supply_by_Date(frmASFBASE0, WHSE_CODE, STYLE_CODE, COLOR_CODE, imax, WSC)
                End If

                '   Dim dtbl As DataTable = rowSOTDEMDX.Table

                Dim ORDR_GROUP_NO As String = rowSOTDEMDX.Item("ORDR_GROUP_NO")
                Dim ORDR_LNO As Integer = rowSOTDEMDX.Item("ORDR_LNO")
                Dim ORDR_RELEASE As String = ""
                Dim ORDR_LAST_UNIT As Int64 = 0
                Dim ORDR_RELEASE_AVAIL As Date = Nothing
                Dim ORDR_RELEASE_AVAIL_ADJ As Date = Nothing
                Dim ORDR_RELEASE_SHIP As Date = Nothing
                Dim WIP_IND As String = ""
                Dim QTY_OPEN As Int64 = Val(rowSOTDEMDX.Item("ORDR_QTY_OPEN") & "")

                If (ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI") Then
                    If doNotAllocateToCreditHoldCustomers Then
                        Dim CUST_CODE As String = rowSOTDEMDX.Item("CUST_CODE")
                        Dim rowARTCUST1 As DataRow = frmASFBASE0.dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
                        If rowARTCUST1.Item("CUST_CREDIT_HOLD") & String.Empty = "1" Then
                            Continue For
                        End If
                    End If
                End If

                '  If ASCMAIN1.Running_in_VS AndAlso ORDR_GROUP_NO = "0000781958" Then Stop

                If DEMAND_TYPE = "O" Then
                    Dim rowSOTORDR7 As DataRow = frmASFBASE0.dst.Tables("SOTORDR7").Rows.Find _
                                                 (New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})
                    If rowSOTORDR7 IsNot Nothing Then
                        If rowSOTORDR7.Item("ORDR_RELEASE") & "" = "X" Then
                            ORDR_RELEASE = "X"
                            QTY_OPEN = 0
                        End If
                    End If
                End If

                Dim BALANCE As Int64 = QTY_OPEN

                Dim ORDR_ALLO_CUR As Int64 = 0
                Dim ORDR_ALLO_FUT As Int64 = 0
                Dim ORDR_ALLO_CXL As Int64 = 0
                Dim ORDR_BACKORDER As String = ""
                ' frmASFBASE0.dst.Tables("ICTSTDQ1").Rows.Clear()

                ' Force Pick means to allocate whether you got the goods or not

                If rowSOTDEMDX.Item("ORDR_PRIORITY") = "0" Then
                    SQ(1, 0) = SQ(1, 0) - BALANCE
                    BALANCE = 0
                Else
                    Dim DD As String
                    If rowSOTDEMDX.Item("ORDR_DEMAND_DATE") & "" = "" Then
                        'WHEN WOULD THIS EVER HAPPEN?
                        DD = Format(Now, "yyyyMMdd")
                    Else
                        DD = Format(rowSOTDEMDX.Item("ORDR_DEMAND_DATE"), "yyyyMMdd")
                    End If

                    If ASCMAIN1.CLIENT = "RGI" Then
                        If Not allocate_as_late_as_possible And ATONCE <> "1" Then
                            ' DD = Format(rowSOTDEMDX.Item("ORDR_SHIP_DATE"), "yyyyMMdd")
                            DD = Format(rowSOTDEMDX.Item("INIT_DATE"), "yyyyMMdd")
                        Else
                            DD = Format(rowSOTDEMDX.Item("ORDR_SHIP_DATE_PLUS"), "yyyyMMdd")
                        End If

                    End If
                    '  Dim tbl2 As DataTable = rowSOTDEMDX.Table


                    Dim rowICTSTDQ3 As DataRow
                    rowICTSTDQ3 = frmASFBASE0.dst.Tables("ICTSTDQ3").Rows.Find(New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})
                    If rowICTSTDQ3 Is Nothing Then
                        rowICTSTDQ3 = frmASFBASE0.dst.Tables("ICTSTDQ3").NewRow
                        rowICTSTDQ3.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                        rowICTSTDQ3.Item("STYLE_CODE") = STYLE_CODE
                        rowICTSTDQ3.Item("COLOR_CODE") = COLOR_CODE
                        frmASFBASE0.dst.Tables("ICTSTDQ3").Rows.Add(rowICTSTDQ3)
                    End If

                    If imax > 0 Then

                        ' Save the Supply Array, in case we need to de-allocate this demand record
                        imax_save = imax
                        For i As Integer = 1 To imax
                            SQ(2, i) = SQ(1, i)
                        Next i

                        ' sq(0,) = Cum Supply prior to this date
                        ' sq(1,) = Remaining Supply Qty arriving on this date
                        ' sq(2,) = Saved Supply Qty's, in case we have to de-allocate
                        ' sq(3,) = Available to Sell from this Shipment
                        ' sq(4,) = FIFO Demand
                        ' sq(5,) = ABSolute Supply Qty arriving on this date

                        ' Look for supply from the last to arrive to the first to arrive, 
                        '  then On hand,  to satisfy this order based on cancel (ie demand) date

                        Dim imax_new As Integer = imax
                        Dim imax_old As Integer = imax
                        Dim imax_i As New List(Of Integer)

                        If ASCMAIN1.CLIENT = "RGI" Then
                            ' find the latest i that the ship date + 30 can tolerate and start the loop with that i value
                            For i As Integer = imax To 1 Step -1
                                Dim rowSOTSUPPI As DataRow = frmASFBASE0.dst.Tables("SOTSUPPI").Rows.Find(i)
                                Dim SUPPLY_DATE As String = rowSOTSUPPI.Item("SUPPLY_DATE") & ""
                                If ATONCE = "1" And SUPPLY_DATE <> "00000000" Then
                                    Dim SUPPLY_DATE_PLUS As Date = CDate(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                                    SUPPLY_DATE_PLUS = SUPPLY_DATE_PLUS.AddDays(SO_PARM_ARRIVAL_BUFFER_DAYS)
                                    SUPPLY_DATE = Format(SUPPLY_DATE_PLUS, "yyyyMMdd")
                                End If

                                'restore this

                                'Dim SHIP_DATE As String = rowSOTSUPPI.Item("SHIP_DATE") & ""

                                Dim ORDR_SHIP_DATE As Date = rowSOTDEMDX.Item("ORDR_SHIP_DATE")
                                Dim SHIP_DATE_PLUS As Date = ORDR_SHIP_DATE.AddDays(SO_PARM_SHIP_WINDOW_DAYS)
                                Dim SHIP_DATE As String = Format(SHIP_DATE_PLUS, "yyyyMMdd")

                                If SUPPLY_DATE = "" Then SUPPLY_DATE = Format(Now, "yyyyMMdd")

                                imax_new = i

                                ' remark out this

                                'Dim ORDR_SHIP_DATE As Date = rowSOTDEMDX.Item("ORDR_SHIP_DATE")
                                'Dim SHIP_DATE_PLUS As Date = ORDR_SHIP_DATE.AddDays(SO_PARM_SHIP_WINDOW_DAYS)
                                'Dim SHIP_DATE As String = Format(SHIP_DATE_PLUS, "yyyyMMdd")

                                'If SHIP_DATE > SUPPLY_DATE Then
                                '    imax = i
                                '    Exit For
                                'End If

                                ' restore this

                                'If SHIP_DATE <> "00000000" Then
                                '    Dim SHIP_DATE_PLUS As Date = CDate(Mid(SHIP_DATE, 5, 2) & "/" & Mid(SHIP_DATE, 7, 2) & "/" & Mid(SHIP_DATE, 1, 4))
                                '    SHIP_DATE_PLUS = SHIP_DATE_PLUS.AddDays(SO_PARM_SHIP_WINDOW_DAYS)
                                '    SHIP_DATE = Format(SHIP_DATE_PLUS, "yyyyMMdd")

                                If SHIP_DATE > SUPPLY_DATE Then
                                    ' imax = i
                                    Exit For
                                End If
                                'End If

                            Next

                            For i As Integer = imax_new To 1 Step -1
                                imax_i.Add(i)
                            Next
                            If imax_new < imax Then
                                For i As Integer = imax To imax_new + 1 Step -1
                                    imax_i.Add(i)
                                Next
                            End If

                            'imax = imax_new
                        Else
                            For i As Integer = imax To 1 Step -1
                                imax_i.Add(i)
                            Next
                        End If




                        For Each i As Integer In imax_i ' For i As Integer = imax To 1 Step -1
                            Dim rowSOTSUPPI As DataRow = frmASFBASE0.dst.Tables("SOTSUPPI").Rows.Find(i)
                            Dim SUPPLY_DATE As String = rowSOTSUPPI.Item("SUPPLY_DATE") & ""
                            If ATONCE = "1" And SUPPLY_DATE <> "00000000" Then
                                Dim SUPPLY_DATE_PLUS As Date = CDate(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                                SUPPLY_DATE_PLUS = SUPPLY_DATE_PLUS.AddDays(SO_PARM_ARRIVAL_BUFFER_DAYS)
                                SUPPLY_DATE = Format(SUPPLY_DATE_PLUS, "yyyyMMdd")
                            End If
                            Dim SHIP_DATE As String = rowSOTSUPPI.Item("SHIP_DATE") & ""
                            If SUPPLY_DATE = "" Then SUPPLY_DATE = Format(Now, "yyyyMMdd")
                            'If SUPPLY_DATE = "00000000" Then SUPPLY_DATE = Format(Now, "yyyyMMdd")

                            Dim IQ3 As Integer = i
                            If IQ3 > 4 Then IQ3 = 4
                            Dim IQ3DATE As Date = Now.Date
                            If SUPPLY_DATE <> "00000000" Then
                                IQ3DATE = DateValue(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                                If ATONCE = "1" Then
                                    IQ3DATE = IQ3DATE.AddDays(-1 * SO_PARM_ARRIVAL_BUFFER_DAYS)
                                End If
                            End If

                            rowICTSTDQ3.Item("DATE_" & CStr(IQ3)) = IQ3DATE

                            If SUPPLY_DATE <= DD Then
                                ' i=1 is for on hand
                                ORDR_LAST_UNIT = SQ(0, i) + SQ(1, i)
                                If Format(ORDR_RELEASE_AVAIL, "MM/dd/yyyy") = "01/01/0001" And i > 1 Then
                                    ORDR_RELEASE_AVAIL = DateValue(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                                    If ATONCE = "1" And SUPPLY_DATE <> "00000000" Then
                                        ORDR_RELEASE_AVAIL_ADJ = ORDR_RELEASE_AVAIL.AddDays(SO_PARM_DAYS_ADJ)
                                    Else
                                        ORDR_RELEASE_AVAIL_ADJ = ORDR_RELEASE_AVAIL
                                    End If

                                    If SHIP_DATE <> "" Then
                                        ORDR_RELEASE_SHIP = DateValue(Mid(SHIP_DATE, 5, 2) & "/" & Mid(SHIP_DATE, 7, 2) & "/" & Mid(SHIP_DATE, 1, 4))
                                        WIP_IND = rowSOTSUPPI.Item("WIP_IND") & ""
                                    End If
                                    'ORDR_LAST_UNIT = SQ(0, i) + SQ(1, i)
                                End If
                                If SQ(1, i) >= BALANCE Then
                                    If i = 1 Then ORDR_ALLO_CUR = BALANCE
                                    SQ(1, i) = SQ(1, i) - BALANCE
                                    rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) = Val(rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) & "") + BALANCE
                                    BALANCE = 0
                                    Exit For
                                    'ElseIf SQ(1, i) > 0 Then
                                    '    If i = 1 Then ORDR_ALLO_CUR = SQ(1, i)
                                    '    BALANCE = BALANCE - SQ(1, i)
                                    '    rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) = Val(rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) & "") + SQ(1, i)
                                    '    SQ(1, i) = 0
                                    '    If i = imax Then
                                    '        imax = imax - 1
                                    '    End If
                                ElseIf SQ(1, i) > 0 Then
                                    If i = 1 Then ORDR_ALLO_CUR = SQ(1, i)
                                    ' PREV LINE SAYS TO ALLOCATE WHAT YOU CAN FROM THE SUPPLY ONLY IF YOU ARE LOOKING AT ON HAND (i=1)
                                    If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Then
                                        ' FOR VAN AND RGI, THE FEELING IS THAT WE SHOULD ALLOCATE WHATEVER WE HAVE LEFT, REGARDLESS IF IT IS A PO, XIT, OR OH
                                        ORDR_ALLO_CUR = SQ(1, i)
                                    End If
                                    If ASCMAIN1.CLIENT = "RGIx" Then
                                        If INNER_PACK_QTY > 1 Then
                                            Dim LOOSE_QTY As Integer = (ORDR_ALLO_CUR Mod INNER_PACK_QTY)
                                            If LOOSE_QTY <> 0 And QTY_OPEN <> ORDR_ALLO_CUR Then
                                                ' BOTH EXPRESSIONS WORK, BECAUSE INTEGER DIVISION TRUNCATES, IT DOES NOT ROUND
                                                ORDR_ALLO_CUR = INNER_PACK_QTY * (ORDR_ALLO_CUR \ INNER_PACK_QTY)
                                                ' ORDR_ALLO_CUR = ORDR_ALLO_CUR - LOOSE_QTY
                                            End If
                                        End If
                                    End If
                                    BALANCE = BALANCE - ORDR_ALLO_CUR
                                    rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) = Val(rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) & "") + ORDR_ALLO_CUR
                                    SQ(1, i) = SQ(1, i) - ORDR_ALLO_CUR
                                    If SQ(1, i) = 0 And i = imax Then
                                        imax = imax - 1
                                    End If
                                End If

                            End If
                        Next i
                    End If
                    'ORDR_ALLO_CUR = QTY_OPEN - BALANCE
                    ORDR_ALLO_FUT = QTY_OPEN - BALANCE - ORDR_ALLO_CUR


                    If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" _
                    Or ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
                        ' WE MAY BE RELEASING EVEN IF SHORT, AND IF WE DO, IF ANYTHING IS COVERED BY A PO AND WE ARE NOT RELEASING AGAINST FUT AVAIL, THEN WE NEED TO SET THE BO FLAG
                        If BALANCE = 0 Then

                            ' SAME BLOCK AS BELOW
                            If edi850cust.Contains(rowSOTDEMDX.Item("CUST_CODE")) Then
                                ORDR_BACKORDER = "N"
                            Else
                                Dim rowARTCUST1 As DataRow = frmASFBASE0.dst.Tables("ARTCUST1").Rows.Find(rowSOTDEMDX.Item("CUST_CODE"))
                                If rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & "" = "1" _
                                Or (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
                                    ORDR_BACKORDER = "Y"
                                Else
                                    ORDR_BACKORDER = "N"
                                End If
                            End If

                        End If
                    End If

                    ' If we can't allocate the entire qty using supply records <= the demand date then ...
                    If BALANCE <> 0 Then
                        If edi850cust.Contains(rowSOTDEMDX.Item("CUST_CODE")) Then
                            ORDR_BACKORDER = "N"
                        Else
                            Dim rowARTCUST1 As DataRow = frmASFBASE0.dst.Tables("ARTCUST1").Rows.Find(rowSOTDEMDX.Item("CUST_CODE"))
                            If rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & "" = "1" _
                            Or (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
                                ORDR_BACKORDER = "Y"
                            Else
                                ORDR_BACKORDER = "N"
                            End If
                        End If

                        ' ORDR_RELEASE
                        ' H = Hold if Short (Shown as H but stored as Null)
                        ' X = Cancel regardless of whether allocable
                        ' C = Cancel the entire line if Short
                        ' S = Ship Short if necessary

                        ' If it's an Order
                        If DEMAND_TYPE = "O" Then
                            ' Check to see if Order Release Flag says to Allocate Nothing if we can't ship
                            Dim rowSOTORDR7 As DataRow = frmASFBASE0.dst.Tables("SOTORDR7").Rows.Find _
                                                         (New String() {ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE})
                            If rowSOTORDR7 IsNot Nothing Then
                                If rowSOTORDR7.Item("ORDR_BACKORDER") & "" <> "" Then
                                    ORDR_BACKORDER = rowSOTORDR7.Item("ORDR_BACKORDER")
                                End If
                                ORDR_RELEASE = rowSOTORDR7.Item("ORDR_RELEASE") & ""
                                If ORDR_RELEASE = "C" Then
                                    BALANCE = QTY_OPEN
                                    ORDR_LAST_UNIT = 0
                                    ORDR_RELEASE_AVAIL = Nothing
                                    ORDR_RELEASE_AVAIL_ADJ = Nothing
                                    ORDR_ALLO_CUR = 0
                                    ORDR_ALLO_FUT = 0

                                    ' Restore the Supply Array
                                    imax = imax_save
                                    For i As Integer = 1 To imax
                                        SQ(1, i) = SQ(2, i)
                                    Next i
                                End If
                            End If
                        End If

                        ' Try to Allocate from Futures

                        If (ORDR_BACKORDER = "Y" Or (ORDR_RELEASE <> "C" And ORDR_RELEASE <> "S")) And imax > 1 Then
                            ' ALTHOUGH THE NEXT LINE APPEARS TO FIX A BUG, IT MESSED UP BURLINGTON WHERE WE DID NOT LOOK TO THE FUTURE TO SEE A PO LATER THAN THE CANCEL DATE
                            'If (ORDR_BACKORDER = "Y" And (ORDR_RELEASE <> "C" And ORDR_RELEASE <> "S")) And imax > 1 Then
                            For i As Integer = 2 To imax
                                Dim rowSOTSUPPI As DataRow = frmASFBASE0.dst.Tables("SOTSUPPI").Rows.Find(i)
                                Dim SUPPLY_DATE As String = rowSOTSUPPI.Item("SUPPLY_DATE") & ""
                                Dim SHIP_DATE As String = rowSOTSUPPI.Item("SHIP_DATE") & ""
                                If ATONCE = "1" And SUPPLY_DATE <> "00000000" Then
                                    Dim SUPPLY_DATE_PLUS As Date = CDate(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                                    SUPPLY_DATE_PLUS = SUPPLY_DATE_PLUS.AddDays(SO_PARM_ARRIVAL_BUFFER_DAYS)
                                    SUPPLY_DATE = Format(SUPPLY_DATE_PLUS, "yyyyMMdd")
                                End If

                                Dim IQ3 As Integer = i
                                If IQ3 > 4 Then IQ3 = 4

                                If SUPPLY_DATE > DD Then
                                    'If ORDR_RELEASE <> "S" Or allocation_only Then
                                    If ORDR_RELEASE <> "S" Or allocation_only Then
                                        ORDR_RELEASE_AVAIL = DateValue(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                                        If ATONCE = "1" And SUPPLY_DATE <> "00000000" Then
                                            ORDR_RELEASE_AVAIL_ADJ = ORDR_RELEASE_AVAIL.AddDays(SO_PARM_DAYS_ADJ)
                                        Else
                                            ORDR_RELEASE_AVAIL_ADJ = ORDR_RELEASE_AVAIL
                                        End If
                                        If SHIP_DATE <> "" Then
                                            ORDR_RELEASE_SHIP = DateValue(Mid(SHIP_DATE, 5, 2) & "/" & Mid(SHIP_DATE, 7, 2) & "/" & Mid(SHIP_DATE, 1, 4))
                                            WIP_IND = rowSOTSUPPI.Item("WIP_IND") & ""
                                        End If
                                        ORDR_LAST_UNIT = SQ(0, i) + SQ(1, i)
                                    End If
                                    If SQ(1, i) >= BALANCE Then
                                        SQ(1, i) = SQ(1, i) - BALANCE
                                        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then ' ,MAYBE NYA TOO
                                            ORDR_ALLO_FUT += BALANCE
                                            rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) = Val(rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) & "") + BALANCE
                                        End If
                                        BALANCE = 0
                                        Exit For
                                    ElseIf SQ(1, i) > 0 Then
                                        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then ' ,MAYBE NYA TOO
                                            ORDR_ALLO_FUT += SQ(1, i)
                                            rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) = Val(rowICTSTDQ3.Item("QTY_" & CStr(IQ3)) & "") + SQ(1, i)
                                        End If
                                        BALANCE = BALANCE - SQ(1, i)
                                        SQ(1, i) = 0
                                        If i = imax Then
                                            imax = imax - 1
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                        'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        '    ORDR_ALLO_FUT = QTY_OPEN - BALANCE - ORDR_ALLO_CUR
                        'Else
                        '    ORDR_ALLO_CXL = QTY_OPEN - BALANCE - ORDR_ALLO_CUR - ORDR_ALLO_FUT
                        'End If
                        'ORDR_ALLO_FUT = QTY_OPEN - BALANCE - ORDR_ALLO_CUR

                        ORDR_ALLO_CXL = QTY_OPEN - BALANCE - ORDR_ALLO_CUR - ORDR_ALLO_FUT
                        ' THE ABOVE CALC IS NOT TRUE - BECAUSE OF CONFUSION OVER ORDR_ALLO_FUT

                    End If
                End If

                If (BALANCE <> QTY_OPEN Or ORDR_RELEASE <> "") Then
                    If DEMAND_TYPE = "O" Then
                        If (BALANCE = 0 Or BALANCE = QTY_OPEN) And ORDR_ALLO_CXL = 0 Then
                            ASCMAIN1.sql = "Update " & SOTORDR2 & " SOTORDR2" & vbCrLf
                            Dim mixed_allocation As Boolean = False
                            If BALANCE = 0 Then
                                mixed_allocation = (ORDR_ALLO_CUR <> 0 And (ORDR_ALLO_FUT <> 0 Or ORDR_ALLO_CXL <> 0)) Or
                                                   (ORDR_ALLO_FUT <> 0 And (ORDR_ALLO_CUR <> 0 Or ORDR_ALLO_CXL <> 0))
                                '   mixed_allocation = True
                                'Or (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") _ pulled RGI out, but that double allocated, so putting it back in
                                If (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") _
                                Or ((ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI") And Not mixed_allocation) _
                                Or (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") Then

                                    If (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
                                        ASCMAIN1.sql &= "" & vbCrLf _
                                            & " Set ORDR_QTY_ALLO = ORDR_QTY_OPEN" & vbCrLf _
                                            & ", ORDR_QTY_ALLO_CUR = ORDR_QTY_OPEN" & vbCrLf _
                                            & ", ORDR_QTY_ALLO_FUT = 0" & vbCrLf _
                                            & ", ORDR_QTY_ALLO_CXL = 0" & vbCrLf
                                    Else
                                        ASCMAIN1.sql &= "" & vbCrLf _
                                            & " Set ORDR_QTY_ALLO = ORDR_QTY_OPEN" & vbCrLf _
                                            & ", ORDR_QTY_ALLO_CUR = " & IIf(Format(ORDR_RELEASE_AVAIL, "MM/dd/yyyy") = "01/01/0001", "ORDR_QTY_OPEN", "0") & vbCrLf _
                                            & ", ORDR_QTY_ALLO_FUT = " & IIf(Format(ORDR_RELEASE_AVAIL, "MM/dd/yyyy") = "01/01/0001", "0", "ORDR_QTY_OPEN") & vbCrLf _
                                            & ", ORDR_QTY_ALLO_CXL = 0" & vbCrLf
                                    End If
                                Else

                                    ' THIS CODE WANTS TO UPDATE SOTORDR2 FOR ALL ROWS WITHIN AN ORDER GROUP FOR THE CURRENT STYLE-COLOR
                                    ' THIS MEANS THAT WE CANNOT USE MEMORY VARIABLES WITH TOTAL CUR/FUT ALLOC QTYS, 
                                    ' BECAUSE IF WE DID, THEN WE WOULD ALLOCATE THE ENTIRE GROUP QTY TO EACH SOTORDR2 IN THE GROUP
                                    ' AND THAT WOULD BE BAD FOR MULTI-ORDER GROUPS (LIKE BURLING@VAN) 
                                    ' AND FOR SAME STYLE-COLOR ON MULTIPLE LINES ON A SINGLE ORDER (LIKE BARBEDWIRE@RGI AND WAYFAIR@RGI)

                                    ' Q: NOW THAT ALL 3 ACCTS ARE USING THE ABOVE CODE, WHEN WOULD WE EVER WANT TO DO THIS?
                                    ' A: REGENCY WANTS THIS - SEE MTX59907 ON ORDER 442325 TESTED ON 03/18/2019
                                    ' IE - FOR ORDERS THAT HAVE A CURRENT ALLOCATION AND A FUTURE ALLOCATION ALL ON A SINGLE LINE

                                    ' THIS CODE WILL SCREW UP FOR MULTI-STORE ORDERS 
                                    ' - SEE VAN CODE ABOVE, AND REMEMBER THE BURLINGTON THAT BLEW SOCKETS
                                    ' THIS CODE WILL ALSO SCREW UP FOR ORDERS WITH MULTIPLE LINES, SAME STYLE-COLOR LIKE WAYFAIR & BARBEDWIRE

                                    ' now, the if statement above has us in this section only if RGI and mixed_allocation

                                    ASCMAIN1.sql &= "" & vbCrLf _
                                        & " Set ORDR_QTY_ALLO = " & CStr(QTY_OPEN) & vbCrLf _
                                        & ", ORDR_QTY_ALLO_CUR = " & CStr(ORDR_ALLO_CUR) & vbCrLf _
                                        & ", ORDR_QTY_ALLO_FUT = " & CStr(ORDR_ALLO_FUT) & vbCrLf _
                                        & ", ORDR_QTY_ALLO_CXL = " & CStr(ORDR_ALLO_CXL) & vbCrLf
                                End If
                            Else
                                ASCMAIN1.sql &= "" _
                                    & " Set ORDR_QTY_ALLO = 0" & vbCrLf _
                                    & ", ORDR_QTY_ALLO_CUR = 0" & vbCrLf _
                                    & ", ORDR_QTY_ALLO_FUT = 0" & vbCrLf _
                                    & ", ORDR_QTY_ALLO_CXL = 0" & vbCrLf
                            End If

                            ASCMAIN1.sql &= IIf(ORDR_RELEASE = "",
                                                ",    ORDR_RELEASE = Null",
                                                ",    ORDR_RELEASE = '" & ORDR_RELEASE & "'") & vbCrLf
                            ASCMAIN1.sql &= IIf(Format(ORDR_RELEASE_AVAIL, "MM/dd/yyyy") = "01/01/0001",
                                               ",    ORDR_RELEASE_AVAIL = Null, ORDR_RELEASE_SHIP = Null, WIP_IND = NULL",
                                               ",    ORDR_RELEASE_AVAIL = '" & Format(ORDR_RELEASE_AVAIL_ADJ, "dd-MMM-yyyy") & "', ORDR_RELEASE_SHIP = '" & Format(ORDR_RELEASE_SHIP, "dd-MMM-yyyy") & "', WIP_IND = '" & WIP_IND & "'") & vbCrLf
                            ' If ASCMAIN1.Running_in_VS And WIP_IND <> "" Then Stop
                            ASCMAIN1.sql &= ", ORDR_BACKORDER = '" & ORDR_BACKORDER & "'" & vbCrLf
                            ASCMAIN1.sql &= ", ORDR_LAST_UNIT = " & CStr(ORDR_LAST_UNIT) & vbCrLf
                            ASCMAIN1.sql &= "" _
                                & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                                & "   and STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                                & "   and COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf

                            ' UPDATE SOTORDR2 - ALL LINES IN THE GROUP FOR THIS STYLE-COLOR

                            If mixed_allocation Then
                                ASCMAIN1.sql = "" _
                                    & "Begin " & vbCrLf _
                                    & " Declare Cursor C1 is" & vbCrLf _
                                    & "  Select * from " & SOTORDR2 & " SOTORDR2" & vbCrLf _
                                    & "   where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                                    & "     and STYLE_CODE = '" & STYLE_CODE & "'" _
                                    & "     and COLOR_CODE = '" & COLOR_CODE & "'" _
                                    & "     For Update;" & vbCrLf _
                                    & "   ORDR_ALLO_CUR Number (8,0);" & vbCrLf _
                                    & "   ORDR_ALLO_FUT Number (8,0);" & vbCrLf _
                                    & "   ORDR_ALLO_CXL Number (8,0);" & vbCrLf _
                                    & "   QTY_OPEN Number (8,0);" & vbCrLf _
                                    & "   QTY_ALLO Number (8,0);" & vbCrLf _
                                    & "   QTY_ALLO_CUR Number (8,0);" & vbCrLf _
                                    & "   QTY_ALLO_FUT Number (8,0);" & vbCrLf _
                                    & "   QTY_ALLO_CXL Number (8,0);" & vbCrLf _
                                    & " Begin" & vbCrLf _
                                    & "  ORDR_ALLO_CUR := " & CStr(ORDR_ALLO_CUR) & ";" & vbCrLf _
                                    & "  ORDR_ALLO_FUT := " & CStr(ORDR_ALLO_FUT) & ";" & vbCrLf _
                                    & "  ORDR_ALLO_CXL := " & CStr(ORDR_ALLO_CXL) & ";" & vbCrLf _
                                    & "  For R1 in C1 Loop" & vbCrLf _
                                    & "   QTY_OPEN := NVL(R1.ORDR_QTY_OPEN,0);" & vbCrLf _
                                    & "   QTY_ALLO := 0;" & vbCrLf _
                                    & "   QTY_ALLO_CUR := 0;" & vbCrLf _
                                    & "   QTY_ALLO_FUT := 0;" & vbCrLf _
                                    & "   QTY_ALLO_CXL := 0;" & vbCrLf _
                                    & "   IF QTY_OPEN > 0 AND ORDR_ALLO_CUR > 0 THEN" & vbCrLf _
                                    & "    QTY_ALLO_CUR := LEAST(QTY_OPEN, ORDR_ALLO_CUR);" & vbCrLf _
                                    & "    QTY_OPEN := QTY_OPEN - QTY_ALLO_CUR;" & vbCrLf _
                                    & "    ORDR_ALLO_CUR := ORDR_ALLO_CUR - QTY_ALLO_CUR;" & vbCrLf _
                                    & "   END IF;" & vbCrLf _
                                    & "   IF QTY_OPEN > 0 AND ORDR_ALLO_FUT > 0 THEN" & vbCrLf _
                                    & "    QTY_ALLO_FUT := LEAST(QTY_OPEN, ORDR_ALLO_FUT);" & vbCrLf _
                                    & "    QTY_OPEN := QTY_OPEN - QTY_ALLO_FUT;" & vbCrLf _
                                    & "    ORDR_ALLO_FUT := ORDR_ALLO_FUT - QTY_ALLO_FUT;" & vbCrLf _
                                    & "   END IF;" & vbCrLf _
                                    & "   IF QTY_OPEN > 0 AND ORDR_ALLO_CXL > 0 THEN" & vbCrLf _
                                    & "    QTY_ALLO_CXL := LEAST(QTY_OPEN, ORDR_ALLO_CXL);" & vbCrLf _
                                    & "    QTY_OPEN := QTY_OPEN - QTY_ALLO_CXL;" & vbCrLf _
                                    & "    ORDR_ALLO_CXL := ORDR_ALLO_CXL - QTY_ALLO_CXL;" & vbCrLf _
                                    & "   END IF;" & vbCrLf _
                                    & "   QTY_ALLO := QTY_ALLO_CUR + QTY_ALLO_FUT;" & vbCrLf _
                                    & "   Update " & SOTORDR2 & " SOTORDR2" & vbCrLf _
                                    & "    Set ORDR_QTY_ALLO = QTY_ALLO" & vbCrLf _
                                    & "   , ORDR_QTY_ALLO_CUR = QTY_ALLO_CUR" & vbCrLf _
                                    & "   , ORDR_QTY_ALLO_FUT = QTY_ALLO_FUT" & vbCrLf _
                                    & "   , ORDR_QTY_ALLO_CXL = QTY_ALLO_CXL" & vbCrLf _
                                    & "   ,    ORDR_RELEASE = Null" & vbCrLf _
                                    & "   ,    ORDR_RELEASE_AVAIL = Null, ORDR_RELEASE_SHIP = Null, WIP_IND = NULL" & vbCrLf _
                                    & "   , ORDR_BACKORDER = 'Y'" & vbCrLf _
                                    & "   , ORDR_LAST_UNIT = " & CStr(ORDR_LAST_UNIT) & vbCrLf _
                                    & "    where Current of C1;" & vbCrLf _
                                    & "  End Loop;" & vbCrLf _
                                    & " End;" & vbCrLf _
                                    & "End;"

                                ASCDATA1.ExecuteSQL()

                            Else

                                ASCDATA1.ExecuteSQL()

                            End If


                        Else

                            Dim qOPEN As Int64 = QTY_OPEN - BALANCE
                            Dim sqlw As String = "" _
                            & "     ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                            & " and STYLE_CODE = '" & STYLE_CODE & "'" _
                            & " and COLOR_CODE = '" & COLOR_CODE & "'"

                            ASCMAIN1.sql = "Select * from " & SOTORDR2 & " where " & sqlw
                            frmASFBASE0.Fill_Records("SOTORDR2", "", , ASCMAIN1.sql)

                            For Each rowSOTORDR2 As DataRow In frmASFBASE0.dst.Tables("SOTORDR2").Select("")
                                Dim qALLO_CUR As Int64 = 0
                                Dim qALLO_FUT As Int64 = 0
                                Dim qALLO_CXL As Int64 = 0

                                Dim qTO_ALLOCATE As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                                If qTO_ALLOCATE <= qOPEN Then
                                    qOPEN = qOPEN - qTO_ALLOCATE
                                Else
                                    qTO_ALLOCATE = qOPEN
                                    qOPEN = 0
                                End If

                                If qTO_ALLOCATE <= ORDR_ALLO_CUR + ORDR_ALLO_FUT Then
                                    If qTO_ALLOCATE <= ORDR_ALLO_CUR Then
                                        qALLO_CUR = qTO_ALLOCATE
                                        qALLO_FUT = 0
                                    Else
                                        qALLO_CUR = ORDR_ALLO_CUR
                                        qALLO_FUT = qTO_ALLOCATE - ORDR_ALLO_CUR
                                    End If
                                Else
                                    qALLO_CUR = ORDR_ALLO_CUR
                                    qALLO_FUT = ORDR_ALLO_FUT
                                    qALLO_CXL = qTO_ALLOCATE - (ORDR_ALLO_CUR + ORDR_ALLO_FUT)
                                End If
                                ORDR_ALLO_CUR = ORDR_ALLO_CUR - qALLO_CUR
                                ORDR_ALLO_FUT = ORDR_ALLO_FUT - qALLO_FUT
                                ORDR_ALLO_CXL = ORDR_ALLO_CXL - qALLO_CXL
                                With rowSOTORDR2
                                    .Item("ORDR_QTY_ALLO") = qTO_ALLOCATE
                                    If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                                        qALLO_CUR = qALLO_CUR + qALLO_FUT
                                        qALLO_FUT = qALLO_CXL
                                        qALLO_CXL = 0
                                    End If
                                    .Item("ORDR_QTY_ALLO_CUR") = qALLO_CUR
                                    .Item("ORDR_QTY_ALLO_FUT") = qALLO_FUT
                                    .Item("ORDR_QTY_ALLO_CXL") = qALLO_CXL
                                    .Item("ORDR_RELEASE") = ORDR_RELEASE
                                    .Item("ORDR_RELEASE_AVAIL") = ORDR_RELEASE_AVAIL_ADJ
                                    .Item("ORDR_RELEASE_SHIP") = ORDR_RELEASE_SHIP
                                    .Item("ORDR_LAST_UNIT") = ORDR_LAST_UNIT
                                    .Item("ORDR_BACKORDER") = ORDR_BACKORDER
                                    .Item("WIP_IND") = WIP_IND
                                End With
                            Next
                            frmASFBASE0.Update_Record_TDA("SOTORDR2")
                        End If
                    Else
                        ' THIS SECTION NEEDS WORK - NOT DOING ANYTHING RIGHT
                        Dim qALLO_CUR As Int64 = 0
                        Dim qALLO_FUT As Int64 = 0
                        Dim qALLO_CXL As Int64 = 0

                        ASCMAIN1.sql = "Update " & SOTRSRV2 & " SOTRSRV2" _
                            & " Set RSRV_QTY_ALLO = " & CStr(QTY_OPEN - BALANCE) _
                            & ", ORDR_RELEASE_AVAIL = " & IIf(Format(ORDR_RELEASE_AVAIL, "MM/dd/yyyy") = "01/01/0001", "Null",
                                                              "'" & Format(ORDR_RELEASE_AVAIL_ADJ, "dd-MMM-yyyy") & "'") _
                            & ", ORDR_QTY_ALLO_CUR = " & CStr(qALLO_CUR) _
                            & ", ORDR_QTY_ALLO_FUT = " & CStr(qALLO_FUT) _
                            & ", ORDR_QTY_ALLO_CXL = " & CStr(qALLO_CXL) _
                            & " where RSRV_NO = '" & ORDR_GROUP_NO & "'" _
                            & "   and RSRV_LNO = " & CStr(ORDR_LNO)

                        ASCDATA1.ExecuteSQL()

                        'Dim rowSOTRSRV2 As DataRow = frmASFBASE0.dst.Tables("SOTRSRV2").Rows.Find(New Object() {ORDR_GROUP_NO, ORDR_LNO})

                        'rowSOTRSRV2.Item("RSRV_QTY_ALLO") = QTY_OPEN - BALANCE
                        'rowSOTRSRV2.Item("ORDR_RELEASE_AVAIL") = ORDR_RELEASE_AVAIL
                        'rowSOTRSRV2.Item("ORDR_QTY_ALLO_CUR") = qALLO_CUR ' qAC
                        'rowSOTRSRV2.Item("ORDR_QTY_ALLO_FUT") = qALLO_FUT ' qAF
                        'rowSOTRSRV2.Item("ORDR_QTY_ALLO_CXL") = qALLO_CXL
                    End If
                Else
                    If DEMAND_TYPE = "O" Then
                        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" _
                        Or ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                            ' TESTING FOR RGI ONLY
                            ASCMAIN1.sql = "Update " & SOTORDR2 & " SOTORDR2" & vbCrLf

                            ASCMAIN1.sql &= "" _
                                & " Set ORDR_QTY_ALLO = Null" & vbCrLf _
                                & ", ORDR_QTY_ALLO_CUR = Null" & vbCrLf _
                                & ", ORDR_QTY_ALLO_FUT = Null" & vbCrLf _
                                & ", ORDR_QTY_ALLO_CXL = Null" & vbCrLf _
                                & ", ORDR_RELEASE = Null" & vbCrLf _
                                & ", ORDR_RELEASE_AVAIL = Null, ORDR_RELEASE_SHIP = Null, WIP_IND = NULL" & vbCrLf _
                                & ", ORDR_BACKORDER = '" & ORDR_BACKORDER & "'" & vbCrLf _
                                & ", ORDR_LAST_UNIT = " & CStr(ORDR_LAST_UNIT) & vbCrLf

                            ASCMAIN1.sql &= "" _
                                & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                                & "   and STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                                & "   and COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf

                            ASCDATA1.ExecuteSQL()

                        End If
                    End If
                End If

                Dim QTY_ATS2 As Int64 = 0
                If ORDR_RELEASE = "" Then
                    QTY_ATS2 = QTY_OPEN
                Else
                    QTY_ATS2 = QTY_OPEN - BALANCE
                End If
                If imax_save <> 0 Then
                    Dim i As Integer = 1
                    Do While QTY_ATS2 <> 0
                        Dim qA As Int64 = 0
                        If QTY_ATS2 > SQ(5, i) - SQ(4, i) Then
                            qA = SQ(5, i) - SQ(4, i)
                        Else
                            qA = QTY_ATS2
                        End If
                        SQ(4, i) = SQ(4, i) + qA
                        QTY_ATS2 = QTY_ATS2 - qA
                        If i = imax_save Then
                            Exit Do
                        End If
                        i += 1
                    Loop
                End If
            Next

            If WSC <> "" Then
                Move_Allocations_to_Earliest_Supply_Date(frmASFBASE0, SOTORDR2, WHSE_CODE, STYLE_CODE, COLOR_CODE, SQ, read_only, (allocate_as_late_as_possible Or ATONCE = "1"), SO_PARM_DAYS_ADJ)
                WSC = ""
            End If
        End If


        ' frmASFBASE0.Update_Record_TDA("SOTRSRV2")
        'frmASFBASE0.dst.Tables("ICTSTDQ1").Rows.Clear()

        ASCMAIN1.sql = "Update " & SOTORDR2 & " Set ORDR_RELEASE_AVAIL = Null where ORDR_RELEASE_AVAIL = '01-JAN-0001'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & SOTORDR2 & " Set ORDR_RELEASE_SHIP = Null where ORDR_RELEASE_SHIP = '01-JAN-0001'"
        ASCDATA1.ExecuteSQL()

        If read_only Then
            If DEMAND_TYPE_O > 0 Then
                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is Select * from " & SOTORDR2 & " SOTORDR2;" & vbCrLf _
                    & " Begin For R1 in C1 Loop" & vbCrLf _
                    & "  Update " & SOTDEMD1 & " SOTDEMD1" & " Set " & vbCrLf _
                    & "  ORDR_QTY_ALLO = R1.ORDR_QTY_ALLO" & vbCrLf _
                    & ", ORDR_QTY_ALLO_CUR = R1.ORDR_QTY_ALLO_CUR" & vbCrLf _
                    & ", ORDR_QTY_ALLO_FUT = R1.ORDR_QTY_ALLO_FUT" & vbCrLf _
                    & ", ORDR_QTY_ALLO_CXL = R1.ORDR_QTY_ALLO_CXL" & vbCrLf _
                    & ", ORDR_RELEASE = R1.ORDR_RELEASE" & vbCrLf _
                    & ", ORDR_RELEASE_AVAIL = R1.ORDR_RELEASE_AVAIL" & vbCrLf _
                    & ", ORDR_RELEASE_SHIP = R1.ORDR_RELEASE_SHIP" & vbCrLf _
                    & ", ORDR_LAST_UNIT = R1.ORDR_LAST_UNIT" & vbCrLf _
                    & ", ORDR_BACKORDER = R1.ORDR_BACKORDER" & vbCrLf _
                    & " where ORDR_GROUP_NO = R1.ORDR_GROUP_NO" & vbCrLf _
                    & "   and ORDR_NO = R1.ORDR_NO" & vbCrLf _
                    & "   and ORDR_LNO = R1.ORDR_LNO" & vbCrLf _
                    & "   and STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
                    & "   and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
                    & "   and DEMAND_TYPE = 'O';" & vbCrLf _
                    & " End Loop; End; End;"
                ASCDATA1.ExecuteSQL()
            End If
            If DEMAND_TYPE_R > 0 Then
                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is Select * from " & SOTRSRV2 & " SOTRSRV2;" & vbCrLf _
                    & " Begin For R1 in C1 Loop" & vbCrLf _
                    & "  Update " & SOTDEMD1 & " SOTDEMD1" & " Set " & vbCrLf _
                    & "  ORDR_QTY_ALLO = R1.RSRV_QTY_ALLO" & vbCrLf _
                    & ", ORDR_RELEASE_AVAIL = R1.ORDR_RELEASE_AVAIL" & vbCrLf _
                    & ", ORDR_QTY_ALLO_CUR = R1.ORDR_QTY_ALLO_CUR" & vbCrLf _
                    & ", ORDR_QTY_ALLO_FUT = R1.ORDR_QTY_ALLO_FUT" & vbCrLf _
                    & ", ORDR_QTY_ALLO_CXL = R1.ORDR_QTY_ALLO_CXL" & vbCrLf _
                    & " where ORDR_GROUP_NO = R1.RSRV_NO" & vbCrLf _
                    & "   and ORDR_NO = R1.RSRV_NO" & vbCrLf _
                    & "   and ORDR_LNO = R1.RSRV_LNO" & vbCrLf _
                    & "   and STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
                    & "   and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
                    & "   and DEMAND_TYPE = 'R';" & vbCrLf _
                    & " End Loop; End; End;"
                ASCDATA1.ExecuteSQL()
            End If
        End If

        For Each rowSOTORDR7 As DataRow In frmASFBASE0.dst.Tables("SOTORDR7").Select("ORDR_RELEASE = 'X'")
            Dim ORDR_GROUP_NO As String = rowSOTORDR7.Item("ORDR_GROUP_NO")
            STYLE_CODE = rowSOTORDR7.Item("STYLE_CODE")
            COLOR_CODE = rowSOTORDR7.Item("COLOR_CODE")
            Dim sqlw As String = "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

            ASCMAIN1.sql = "Select * from " & SOTORDR2 & " where " & sqlw
            frmASFBASE0.Fill_Records("SOTORDR2", "", , ASCMAIN1.sql)

            For Each rowSOTORDR2 As DataRow In frmASFBASE0.dst.Tables("SOTORDR2").Select("")
                rowSOTORDR2.Item("ORDR_RELEASE") = "X"
                Dim ORDR_BACKORDER As String = rowSOTORDR7.Item("ORDR_BACKORDER") & ""
                If ORDR_BACKORDER = "" Or ORDR_BACKORDER = "Y" Then
                    rowSOTORDR2.Item("ORDR_BACKORDER") = "Y"
                End If
            Next

            'If Not read_only Then
            frmASFBASE0.Update_Record_TDA("SOTORDR2")
            'End If
        Next

        If allocation_only Then ' THIS IF HAS BEEN PUT IN PLACE TO MAKE SINGLE STYLE AND SINGLE ORDER ALLOCATIONS FASTER
            If showProgress Then
                ASCMAIN1.Progress("Styles w/No Demand", "")
            End If

            For Each rowSOTSUPP0 As DataRow In frmASFBASE0.dst.Tables("SOTSUPP0").Select("HAS_DEMAND = '0'")
                WHSE_CODE = rowSOTSUPP0.Item("WHSE_CODE")
                STYLE_CODE = rowSOTSUPP0.Item("STYLE_CODE")
                COLOR_CODE = rowSOTSUPP0.Item("COLOR_CODE")
                Dim HAS_DEMAND As String = rowSOTSUPP0.Item("HAS_DEMAND") & ""
                If HAS_DEMAND = "0" Then
                    SQ = Setup_WSC_Supply_by_Date(frmASFBASE0, WHSE_CODE, STYLE_CODE, COLOR_CODE, imax, WSC)
                End If
                If imax <> 0 Then
                    Update_ICTSTDQ1(frmASFBASE0, WHSE_CODE, STYLE_CODE, COLOR_CODE, SQ)
                End If
            Next
        End If

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & "  , SUM (ORDR_QTY_ALLO) ORDR_QTY_ALLO" & vbCrLf _
            & "   from " & SOTORDR2 & " SOTORDR2," & SOTORDR1 & " SOTORDR1" & vbCrLf _
            & "  where SOTORDR2.ORDR_STATUS = 'O' " & vbCrLf _
            & "    and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & IIf(STYLE_CODE_to_Allocate = "", "", "    and SOTORDR2.STYLE_CODE = '" & STYLE_CODE_to_Allocate & "'" & vbCrLf) _
            & IIf(COLOR_CODE_to_Allocate = "", "", "    and SOTORDR2.COLOR_CODE = '" & COLOR_CODE_to_Allocate & "'" & vbCrLf) _
            & "  group by SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  Update ICTSTAT2 Set WHSE_QTY_ALLO = 0" & vbCrLf _
            & "   where WHSE_QTY_ALLO <> 0" & vbCrLf _
            & IIf(STYLE_CODE_to_Allocate = "", "", "    and STYLE_CODE = '" & STYLE_CODE_to_Allocate & "'" & vbCrLf) _
            & IIf(COLOR_CODE_to_Allocate = "", "", "    and COLOR_CODE = '" & COLOR_CODE_to_Allocate & "'" & vbCrLf) _
            & ";" _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT2 Set WHSE_QTY_ALLO = R1.ORDR_QTY_ALLO" & vbCrLf _
            & "    where WHSE_CODE = R1.WHSE_CODE and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & SOTORDR1 & " Set ORDR_REL_HOLD_CODES = Null, ORDR_REL_BATCH_NO = Null"
        If Not read_only Then
            ASCDATA1.ExecuteSQL()
        End If

        If manual_release Then
            ' DO NOT TOUCH ICTSTDQX IF DOING A MANUAL RELEASE
        Else
            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" _
            Or ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then

                ' THIS BLOCK REMMED OUT BECAUSE IT HAPPENS INSIDE Update_Status_by_Date
                'ASCMAIN1.sql = "Delete from ICTSTDQ2" & vbCrLf _
                '    & ASCMAIN1.SQL_Add_WHERE( _
                '        IIf(STYLE_CODE_to_Allocate = "", "", " and STYLE_CODE = '" & STYLE_CODE_to_Allocate & "'" & vbCrLf) _
                '      & IIf(WHSE_CODE_to_allocate = "", "", " and WHSE_CODE = '" & WHSE_CODE_to_allocate & "'"))
                'ASCDATA1.ExecuteSQL()

                TAC.SOCMAIN1.Update_Status_by_Date(frmASFBASE0, ICTSTDQ1, ICTSTDQ2, ICTSTDQ3, WHSE_CODE_to_allocate, allocation_only, SOTORDR2, force_pick)
            End If
        End If
    End Sub

    Public Shared Function Setup_WSC_Supply_by_Date(frmASFBASE0 As ASFBASE0, _
        WHSE_CODE As String, STYLE_CODE As String, COLOR_CODE As String, ByRef imax As Integer, ByRef WSC As String) As Int64(,)

        Dim rowSOTSUPP0 As DataRow = frmASFBASE0.dst.Tables("SOTSUPP0").Rows.Find(New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
        rowSOTSUPP0.Item("HAS_DEMAND") = "1"

        frmASFBASE0.dst.Tables("SOTSUPPI").Rows.Clear()

        WSC = WHSE_CODE & "-" & STYLE_CODE & "-" & COLOR_CODE
        ASCMAIN1.Progress("-", WSC)

        Dim SQ(,) As Int64
        ReDim SQ(5, 100)
        ' sq(0,) = Cum Supply prior to this date
        ' sq(1,) = Remaining Supply Qty arriving on this date
        ' sq(2,) = Saved Supply Qty's, in case we have to de-allocate
        ' sq(3,) = Available to Sell from this Shipment
        ' sq(4,) = FIFO Demand
        ' sq(5,) = Absolute Supply Qty arriving on this date

        Dim i As Integer = 0
        Dim sqlw As String = "WHSE_CODE = '" & WHSE_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

        For Each rowSOTSUPP1 As DataRow In frmASFBASE0.dst.Tables("SOTSUPP1").Select(sqlw, "SUPPLY_DATE")
            i += 1
            If i = 1 And rowSOTSUPP1.Item("SUPPLY_DATE") & "" <> "00000000" Then
                i += 1
                frmASFBASE0.dst.Tables("SOTSUPPI").Rows.Add(New Object() {1, "00000000", "00000000", "0"})
            End If
            frmASFBASE0.dst.Tables("SOTSUPPI").Rows.Add(New Object() {i, _
                                                          rowSOTSUPP1.Item("SUPPLY_DATE"), _
                                                          rowSOTSUPP1.Item("SHIP_DATE"), _
                                                          rowSOTSUPP1.Item("WIP_IND")})
            SQ(1, i) = Val(rowSOTSUPP1.Item("SUPPLY_QTY") & "")
            SQ(5, i) = Val(rowSOTSUPP1.Item("SUPPLY_QTY") & "")
            SQ(0, i + 1) = SQ(0, i) + SQ(1, i)
        Next

        If i = 0 Then
            i = 1
            frmASFBASE0.dst.Tables("SOTSUPPI").Rows.Add(New Object() {i, "00000000", "00000000", "0"})
        End If
        imax = i

        Return SQ
    End Function

    Public Shared Sub Move_Allocations_to_Earliest_Supply_Date(
        frmASFBASE0 As ASFBASE0,
        SOTORDR2 As String,
        WHSE_CODE As String,
        STYLE_CODE As String,
        COLOR_CODE As String,
        SQ(,) As Int64,
        read_only As Boolean,
        allocate_as_late_as_possible As Boolean, SO_PARM_DAYS_ADJ As Integer)

        If frmASFBASE0.dst.Tables("SOTSUPPI").Rows.Count = 0 Then Exit Sub

        If Not allocate_as_late_as_possible Then
            For Each rowSOTSUPPI As DataRow In frmASFBASE0.dst.Tables("SOTSUPPI").Select("", "INDEX DESC")
                Dim I As Integer = Val(rowSOTSUPPI.Item("INDEX") & "")
                If SQ(1, I) <> 0 And SQ(1, I) > 0 Then
                    ASCMAIN1.sql = "Update " & SOTORDR2 _
                        & " Set ORDR_LAST_UNIT = ORDR_LAST_UNIT - " & CStr(SQ(1, I)) _
                        & " where ORDR_LAST_UNIT > " & CStr(SQ(0, I)) _
                        & "   and WHSE_CODE = '" & WHSE_CODE & "'" _
                        & "   and STYLE_CODE = '" & STYLE_CODE & "'" _
                        & "   and COLOR_CODE = '" & COLOR_CODE & "'"
                    ASCDATA1.ExecuteSQL()
                End If
            Next

            For Each rowSOTSUPPI As DataRow In frmASFBASE0.dst.Tables("SOTSUPPI").Select("", "INDEX")
                Dim I As Integer = Val(rowSOTSUPPI.Item("INDEX") & "")
                Dim SUPPLY_DATE As String = rowSOTSUPPI.Item("SUPPLY_DATE") & ""
                Dim SHIP_DATE As String = rowSOTSUPPI.Item("SHIP_DATE") & ""
                Dim WIP_IND As String = rowSOTSUPPI.Item("WIP_IND") & ""
                'If SQ(1, I) <> 0 Then
                '                sql = "Update SOWORDR2 "
                '                & " set ORDR_LAST_UNIT = ORDR_LAST_UNIT - " & CStr(SQ(1, i))
                '                & " where ORDR_LAST_UNIT > " & CStr(SQ(0, i))
                '                & " and STYLE_CODE = '" & STYLE_CODE & "'"
                '                & " and COLOR_CODE = '" & COLOR_CODE & "'"
                '                AccD.Execute sql

                'Dim sqlw As String = "" _
                '    & "     ORDR_LAST_UNIT > " & CStr(SQ(0, I)) _
                '    & " and WHSE_CODE = '" & WHSE_CODE & "'" _
                '    & " and STYLE_CODE = '" & STYLE_CODE & "'" _
                '    & " and COLOR_CODE = '" & COLOR_CODE & "'"
                '' SURPRISE - SOTORDR2 IN ADO.NET IS EMPTY
                'For Each rowSOTORDR2 As DataRow In frmASFBASE0.dst.Tables("SOTORDR2").Select(sqlw)
                '    rowSOTORDR2.Item("ORDR_LAST_UNIT") = Val(rowSOTORDR2.Item("ORDR_LAST_UNIT") & "") - SQ(1, I)
                'Next

                ASCMAIN1.sql = "Update " & SOTORDR2 & vbCrLf
                If I = 1 Then
                    ASCMAIN1.sql &= " Set ORDR_RELEASE_AVAIL = NULL, ORDR_RELEASE_SHIP = NULL, WIP_IND = NULL" & vbCrLf
                    ASCMAIN1.sql &= ", ORDR_QTY_ALLO_CUR = ORDR_QTY_ALLO, ORDR_QTY_ALLO_FUT = 0, ORDR_QTY_ALLO_CXL = 0" & vbCrLf
                Else
                    Dim SUPPLY_DATE_D As Date = CDate(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                    'ASCMAIN1.sql &= " Set ORDR_RELEASE_AVAIL = '" & Format(SUPPLY_DATE_D.AddDays(SO_PARM_DAYS_ADJ), "dd-MMM-yyyy") & "', WIP_IND = '" & WIP_IND & "'" & vbCrLf
                    ' WJZ 01/06 - SEE RITA EMAIL 01/05
                    ASCMAIN1.sql &= " Set ORDR_RELEASE_AVAIL = '" & Format(SUPPLY_DATE_D.AddDays(0), "dd-MMM-yyyy") & "', WIP_IND = '" & WIP_IND & "'" & vbCrLf
                    If SHIP_DATE <> "" Then
                        Dim SHIP_DATE_D As Date = CDate(Mid(SHIP_DATE, 5, 2) & "/" & Mid(SHIP_DATE, 7, 2) & "/" & Mid(SHIP_DATE, 1, 4))
                        ASCMAIN1.sql &= ", ORDR_RELEASE_SHIP = '" & Format(SHIP_DATE_D, "dd-MMM-yyyy") & "'" & vbCrLf
                    End If

                    If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                        ASCMAIN1.sql &= ", ORDR_QTY_ALLO_CXL = 0" & vbCrLf _
                            & ", ORDR_QTY_ALLO_CUR = CASE WHEN ORDR_CANCEL_DATE >= '" & Format(SUPPLY_DATE_D, "dd-MMM-yyyy") & "' THEN ORDR_QTY_ALLO ELSE 0 END" & vbCrLf _
                            & ", ORDR_QTY_ALLO_FUT = CASE WHEN ORDR_CANCEL_DATE <  '" & Format(SUPPLY_DATE_D, "dd-MMM-yyyy") & "' THEN ORDR_QTY_ALLO ELSE 0 END" & vbCrLf
                    Else
                        ASCMAIN1.sql &= ", ORDR_QTY_ALLO_CUR = 0" & vbCrLf _
                            & ", ORDR_QTY_ALLO_FUT = CASE WHEN ORDR_CANCEL_DATE >= '" & Format(SUPPLY_DATE_D, "dd-MMM-yyyy") & "' THEN ORDR_QTY_ALLO ELSE 0 END" & vbCrLf _
                            & ", ORDR_QTY_ALLO_CXL = CASE WHEN ORDR_CANCEL_DATE <  '" & Format(SUPPLY_DATE_D, "dd-MMM-yyyy") & "' THEN ORDR_QTY_ALLO ELSE 0 END" & vbCrLf
                    End If
                End If
                '2003A                & " where ORDR_LAST_UNIT >= " & CStr(SQ(0, i)) & " and ORDR_LAST_UNIT <= " & CStr(SQ(0, i + 1))

                'If STYLE_CODE = "MTF19228" And COLOR_CODE = "PURP" Then Stop

                ASCMAIN1.sql &= "" _
                    & " where ORDR_LAST_UNIT > " & CStr(SQ(0, I)) & " and ORDR_LAST_UNIT <= " & CStr(SQ(0, I + 1)) & vbCrLf _
                    & " and WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                    & " and STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                    & " and COLOR_CODE = '" & COLOR_CODE & "'"

                ' THIS MAY NEED TO BE VAN ONLY - GABE'S EXAMPLE OF A STYLE WITH 2376 ALLO TO CUR AND 24 TO FUT, WHICH ALL MOVED TO FUT
                ' this may have further complications if part of the allocation was split cur + fut
                'ASCMAIN1.sql &= " and not (ORDR_QTY_ALLO_CUR <> 0 and ORDR_QTY_ALLO_FUT <> 0)"
                ASCMAIN1.sql &= " and (ORDR_QTY_ALLO_CUR = 0)"


                'If Not read_only Then
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Select Sum (ORDR_QTY_ALLO) from " & SOTORDR2 _
                    & " where ORDR_LAST_UNIT > " & CStr(SQ(0, I)) & " and ORDR_LAST_UNIT <= " & CStr(SQ(0, I + 1)) _
                    & " and WHSE_CODE = '" & WHSE_CODE & "'" _
                    & " and STYLE_CODE = '" & STYLE_CODE & "'" _
                    & " and COLOR_CODE = '" & COLOR_CODE & "'"
                If ASCMAIN1.CLIENT = "RGI" Then
                    If I = 1 Then
                        ASCMAIN1.sql = "Select Sum (ORDR_QTY_ALLO_CUR) from " & SOTORDR2 _
                            & " where ORDR_QTY_ALLO_CUR <> 0" _
                            & " and WHSE_CODE = '" & WHSE_CODE & "'" _
                            & " and STYLE_CODE = '" & STYLE_CODE & "'" _
                            & " and COLOR_CODE = '" & COLOR_CODE & "'"
                    Else
                        ASCMAIN1.sql = "Select Sum (ORDR_QTY_ALLO_FUT) from " & SOTORDR2 _
                            & " where ORDR_LAST_UNIT > " & CStr(SQ(0, I)) & " and ORDR_LAST_UNIT <= " & CStr(SQ(0, I + 1)) _
                            & " and ORDR_QTY_ALLO_FUT <> 0" _
                            & " and WHSE_CODE = '" & WHSE_CODE & "'" _
                            & " and STYLE_CODE = '" & STYLE_CODE & "'" _
                            & " and COLOR_CODE = '" & COLOR_CODE & "'"
                    End If
                End If

                If ASCMAIN1.CLIENT = "RGI" Then
                    ' SKIP THIS - MTX31805 04/07/2019
                    ' PROB WANT TO SKIP THIS FOR ALL CUSTOMERS
                Else
                    Dim ORDR_QTY_ALLO As Int64 = Val(ASCDATA1.GetDataValue)
                    SQ(1, I) += (SQ(4, I) - ORDR_QTY_ALLO)
                    SQ(4, I) = ORDR_QTY_ALLO
                End If

                'End If
                'End If
            Next
        End If

        Update_ICTSTDQ1(frmASFBASE0, WHSE_CODE, STYLE_CODE, COLOR_CODE, SQ)

    End Sub

    Public Shared Sub Update_ICTSTDQ1(frmASFBASE0 As ASFBASE0, WHSE_CODE As String, STYLE_CODE As String, COLOR_CODE As String, SQ(,) As Int64)

        'If ASCMAIN1.Running_in_VS AndAlso STYLE_CODE = "409-856JFW" Then Stop

        Dim WasNowUsed As Boolean = False
        Dim LAST_SD As String = ""
        Dim QTY_ATS_CUM As Int64 = 0
        Dim QTY_ATS_NEG As Int64 = 0
        '    If COLOR_CODE = "GDGD" Then Stop
        For Each rowSOTSUPPI As DataRow In frmASFBASE0.dst.Tables("SOTSUPPI").Select("", "INDEX")
            Dim i As Integer = Val(rowSOTSUPPI.Item("INDEX") & "")
            Dim SUPPLY_DATE As String = rowSOTSUPPI.Item("SUPPLY_DATE") & ""
            If SUPPLY_DATE = "" Then SUPPLY_DATE = Format(Now, "yyyyMMdd") ' rgi plug - this should neever be
            If LAST_SD = SUPPLY_DATE And SUPPLY_DATE <> "" And (ASCMAIN1.CLIENT = "RGI") Then '  And 1 <> 1
                ' this routine is messing up on QTY_ATS and QTY_ATS_CUM when 2 shipments same date
                Dim STATUS_DATE As Date = CDate(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                Dim rowICTSTDQ1 As DataRow = frmASFBASE0.dst.Tables("ICTSTDQ1").Rows.Find _
                                             (New Object() {WHSE_CODE, STYLE_CODE, COLOR_CODE, STATUS_DATE})
                rowICTSTDQ1.Item("STATUS_QTY") = Val(rowICTSTDQ1.Item("STATUS_QTY") & "") + SQ(5, i) - SQ(4, i)
                rowICTSTDQ1.Item("SUPPLY_QTY") = Val(rowICTSTDQ1.Item("SUPPLY_QTY") & "") + SQ(1, i)
                rowICTSTDQ1.Item("QTY_ATS") = Val(rowICTSTDQ1.Item("QTY_ATS") & "") + SQ(5, i) - SQ(4, i)
                rowICTSTDQ1.Item("QTY_ATS_CUM") = Val(rowICTSTDQ1.Item("QTY_ATS_CUM") & "") + SQ(5, i) - SQ(4, i)
            Else
                Dim rowICTSTDQ1 As DataRow = frmASFBASE0.dst.Tables("ICTSTDQ1").NewRow
                rowICTSTDQ1.Item("WHSE_CODE") = WHSE_CODE
                rowICTSTDQ1.Item("STYLE_CODE") = STYLE_CODE
                rowICTSTDQ1.Item("COLOR_CODE") = COLOR_CODE
                If SUPPLY_DATE = "00000000" Then
                    rowICTSTDQ1.Item("STATUS_DATE") = Now.Date
                    WasNowUsed = True
                Else
                    If SUPPLY_DATE = Format(Now, "yyyyMMdd") And WasNowUsed Then
                        rowICTSTDQ1.Item("STATUS_DATE") = Now.Date.AddDays(1)
                    Else
                        rowICTSTDQ1.Item("STATUS_DATE") = CDate(Mid(SUPPLY_DATE, 5, 2) & "/" & Mid(SUPPLY_DATE, 7, 2) & "/" & Mid(SUPPLY_DATE, 1, 4))
                    End If
                End If
                rowICTSTDQ1.Item("STATUS_QTY") = SQ(5, i) - SQ(4, i)
                rowICTSTDQ1.Item("SUPPLY_QTY") = SQ(1, i)

                'QTY_ATS_CUM += Val(rowICTSTDQ1.Item("STATUS_QTY") & "")
                'rowICTSTDQ1.Item("QTY_ATS") = rowICTSTDQ1.Item("STATUS_QTY") ' I THINK
                'rowICTSTDQ1.Item("QTY_ATS_CUM") = QTY_ATS_CUM

                Dim QTY_ATS As Int64 = SQ(1, i) + QTY_ATS_NEG
                If QTY_ATS < 0 Then
                    QTY_ATS_NEG = QTY_ATS
                    QTY_ATS = 0
                Else
                    QTY_ATS_NEG = 0
                End If
                QTY_ATS_CUM += QTY_ATS
                rowICTSTDQ1.Item("QTY_ATS") = QTY_ATS
                rowICTSTDQ1.Item("QTY_ATS_CUM") = QTY_ATS_CUM

                Dim rowICTSTDQ1_dup As DataRow = frmASFBASE0.dst.Tables("ICTSTDQ1").Rows.Find _
                                                 (New Object() {WHSE_CODE, STYLE_CODE, COLOR_CODE, rowICTSTDQ1.Item("STATUS_DATE")})
                If rowICTSTDQ1_dup IsNot Nothing Then
                    rowICTSTDQ1_dup.Item("STATUS_QTY") = Val(rowICTSTDQ1_dup.Item("STATUS_QTY") & "") + SQ(5, i) - SQ(4, i)
                    rowICTSTDQ1_dup.Item("SUPPLY_QTY") = Val(rowICTSTDQ1_dup.Item("SUPPLY_QTY") & "") + SQ(1, i)
                    ' rowICTSTDQ1_dup.Item("QTY_ATS") = QTY_ATS '  rowICTSTDQ1_dup.Item("STATUS_QTY") ' I THINK
                    rowICTSTDQ1_dup.Item("QTY_ATS") = Val(rowICTSTDQ1_dup.Item("QTY_ATS") & "") + QTY_ATS '  ADDRESSING BUG AT NYA WHERE 2 SHIPMENTS SAME DATE DID NOT ACCUMULATE PROPERLY
                    rowICTSTDQ1_dup.Item("QTY_ATS_CUM") = QTY_ATS_CUM '  Val(rowICTSTDQ1_dup.Item("QTY_ATS_CUM") & "") + QTY_ATS_CUM
                Else
                    frmASFBASE0.dst.Tables("ICTSTDQ1").Rows.Add(rowICTSTDQ1)
                End If


                If ASCMAIN1.CLIENT = "RGI" Then

                    ' NOTE THAT THE BLOCK BELOW WAS NOT COMPLETELY CODED, SO I HAVE REVERTED THIS BLOCK AND RESTORED THE ONE UNDERNEATH
                    '' THESE ARE THE CHANGES TO MAKE ICTSTDQ1 RIGHT WITH RESPECT TO ATONCE - WJZ 03/14/21
                    'Dim SUPPLY As Int64 = SQ(1, i)
                    '' Dim USED As Int64 = frmASFBASE0.dst.Tables("ICTSTDQ3").Compute("SUM(QTY_" & CStr(i), $"WHSE_CODE = '{WHSE_CODE}' AND STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'")
                    '' WE NEED WHSE IN ICTSTDQ3 - FOR NOW, ASSUMING MS SINCE THIS IS AN RGI SPECIFIC ROUTINE
                    '' WE ARE USING ICTSTDQ3 SINCE THIS APPEARS TO BE A MORE RELIABLE VERSION OF ALLOCATION BY DATE THAN THE SQ ARRAY
                    'Dim USED As Int64 = Val(frmASFBASE0.dst.Tables("ICTSTDQ3").Compute("SUM(QTY_" & CStr(i) & ")", $"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'") & "")
                    ''rowICTSTDQ1.Item("STATUS_QTY") = SQ(5, i) - SQ(4, i)
                    ''rowICTSTDQ1.Item("SUPPLY_QTY") = SQ(1, i)
                    'rowICTSTDQ1.Item("QTY_ATS") = QTY_ATS
                    'rowICTSTDQ1.Item("QTY_ATS_CUM") = QTY_ATS_CUM
                    '' SIMULTANEOUS WITH THE ABOVE, I AM REMARKING OUT THE BELOW UNTIL I HAVE REASONS TO RESTORE THIS CODE - WJZ 03/14/21

                    ' IF THE SUPPLY DATE IS IN THE PAST (IE, A PAST DUE SHIPMENT) THEN ADD THE SUPPLY QTY TO THE STATUS QTY FIELDS FOR ALL DATES FORWARD OF THAT PD DATE
                    Dim STATUS_DATE_this_record As Date = rowICTSTDQ1.Item("STATUS_DATE")
                    If Format(STATUS_DATE_this_record, "MM/dd/yyyy") < Format(Now, "MM/dd/yyyy") Then
                        Dim sql_forward As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and STATUS_DATE > '" & Format(STATUS_DATE_this_record, "MM/dd/yyyy") & "'"
                        For Each rowCUM As DataRow In frmASFBASE0.dst.Tables("ICTSTDQ1").Select(sql_forward)
                            Dim add_to_status As Int64 = SQ(1, i) ' not sure of this in light of the comments in the section above - might need to be refined.
                            rowCUM.Item("STATUS_QTY") = Val(rowCUM.Item("STATUS_QTY") & "") + add_to_status
                            'rowCUM.Item("QTY_ATS") = Val(rowCUM.Item("QTY_ATS") & "") + add_to_status
                            rowCUM.Item("QTY_ATS_CUM") = Val(rowCUM.Item("QTY_ATS_CUM") & "") + add_to_status
                        Next
                    End If
                End If

                LAST_SD = SUPPLY_DATE
            End If
        Next
    End Sub

    Public Shared Function Price_Line(frm As ASFBASE0, _
                   CUST_CODE As String, rowARTCUST1 As DataRow, _
                   STYLE_CODE As String, COLOR_CODE As String, _
                   ORDR_QTY As Int32, ByRef ORDR_PRICE_SOURCE As String) As Decimal

        If Not frm.dst.Tables.Contains("ICTCLAS1") Then
            ASCMAIN1.sql = "Select * from ICTCLAS1"
            frm.Create_TDA(frm.dst.Tables.Add, "ICTCLAS1", "**", 0, False)
            frm.Fill_Records("ICTCLAS1")
        End If
        If Not frm.dst.Tables.Contains("ICTDISC1") Then
            ASCMAIN1.sql = "Select * from ICTDISC1"
            frm.Create_TDA(frm.dst.Tables.Add, "ICTDISC1", "**", 0, False)
            frm.Fill_Records("ICTDISC1")
        End If

        Dim rowICTSTYL1 As DataRow = frm.LookUp("ICTSTYL1", New String() {STYLE_CODE})
        Dim STYLE_STATUS As String = rowICTSTYL1.Item("STYLE_STATUS") & ""
        Dim STYLE_CLASS_CODE As String = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
        Dim STYLE_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PRICE") & "")
        Dim STYLE_PROMO_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PROMO_PRICE") & "")
        Dim CARTON_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
        Dim rowICTSTYC1 As DataRow = frm.LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})

        Dim STYLE_COLOR_STATUS As String = ""
        If rowICTSTYC1 IsNot Nothing Then
            STYLE_COLOR_STATUS = rowICTSTYC1.Item("STYLE_COLOR_STATUS") & ""
        End If

        Dim ORDR_UNIT_PRICE_CALC As Decimal = STYLE_PRICE

        ORDR_PRICE_SOURCE = "Q" ' Qty Break using Price Discount Schedule

        If STYLE_STATUS = "D" Or STYLE_COLOR_STATUS = "D" Then
            ORDR_UNIT_PRICE_CALC = 0.3 * STYLE_PRICE
            ORDR_PRICE_SOURCE = "D" ' Discontinued

        ElseIf STYLE_PROMO_PRICE <> 0 Then ' And ORDR_QTY >= CARTON_PACK_QTY Then
            ORDR_UNIT_PRICE_CALC = STYLE_PROMO_PRICE
            ORDR_PRICE_SOURCE = "P" ' Promo

        Else
            Dim CUST_DISC_PCT_EXTRA As String = rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") & ""
            Dim CUST_DISC_CASES As Decimal = 0
            Dim rowICTCLAS1 As DataRow = frm.dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)

            If rowICTCLAS1 IsNot Nothing Then
                Dim DISC_CODE As String = rowICTCLAS1.Item("DISC_CODE") & ""

                If DISC_CODE = "NONPVC" And rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "SP" Then
                    Dim CUST_DISC_PCT As Decimal = Val(rowARTCUST1.Item("CUST_DISC_PCT") & "")
                    ORDR_UNIT_PRICE_CALC = STYLE_PRICE * (100 - CUST_DISC_PCT) / 100
                    ORDR_PRICE_SOURCE = "S" ' Special Price

                Else
                    If DISC_CODE = "PVC" Then
                        If rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & "" = "5C" Then CUST_DISC_CASES = 5
                        If rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & "" = "FC" Then CUST_DISC_CASES = 1
                    Else
                        If rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "FC" Then CUST_DISC_CASES = 1
                        If rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "HC" Then CUST_DISC_CASES = 0.5
                    End If

                    Dim rowICTDISC1 As DataRow = frm.dst.Tables("ICTDISC1").Rows.Find(DISC_CODE)
                    If rowICTDISC1 IsNot Nothing Then
                        For I As Integer = 1 To 4
                            Dim CASES As Decimal = Val(rowICTDISC1.Item("DISC" & CStr(I) & "_CASES") & "")
                            If CUST_DISC_CASES <> 0 And CUST_DISC_CASES < CASES And CUST_DISC_CASES * CARTON_PACK_QTY > ORDR_QTY Then
                            Else
                                If ORDR_QTY >= CASES * CARTON_PACK_QTY Or CASES = 0 Or CUST_DISC_CASES = CASES Then
                                    Dim PCT As Decimal = Val(rowICTDISC1.Item("DISC" & CStr(I) & "_PCT") & "")
                                    ORDR_UNIT_PRICE_CALC = STYLE_PRICE * (100 - PCT) / 100
                                    If ORDR_UNIT_PRICE_CALC < 0 Then ORDR_UNIT_PRICE_CALC = 0
                                    ' ORDR_PRICE_SOURCE = rowICTDISC1.Item("ABBR" & CStr(I)) & "" ' Qty Break
                                    ' THIS IS THE PRICE TIER
                                    ORDR_PRICE_SOURCE = "Q" & CStr(I)
                                    Exit For
                                End If
                            End If
                        Next

                        If DISC_CODE = "NONPVC" Then
                            Dim CUST_DISC_PCT_EXTRA_PCT As Decimal = 0
                            If CUST_DISC_PCT_EXTRA = "1" Then CUST_DISC_PCT_EXTRA_PCT = 5
                            If CUST_DISC_PCT_EXTRA = "2" Then CUST_DISC_PCT_EXTRA_PCT = 10
                            If CUST_DISC_PCT_EXTRA_PCT <> 0 And CUST_DISC_CASES = 0 Then

                                ORDR_UNIT_PRICE_CALC = (100 - CUST_DISC_PCT_EXTRA_PCT) * ORDR_UNIT_PRICE_CALC / 100
                                ORDR_PRICE_SOURCE &= Format(CUST_DISC_PCT_EXTRA_PCT, "00")
                                If CUST_DISC_PCT_EXTRA_PCT = 0 Then
                                    ORDR_PRICE_SOURCE &= "XX" ' SHOULD NEVER HAPPEN SINCE THIS BLOCK IS ONLY IF CUST_DISC_PCT_EXTRA_PCT <> 0
                                Else
                                    ORDR_PRICE_SOURCE &= "X" & CUST_DISC_PCT_EXTRA
                                End If
                                ' THIS IS THE VOL DISC
                            End If
                        End If

                    End If
                End If

            End If
        End If

        ORDR_UNIT_PRICE_CALC = System.Math.Round(ORDR_UNIT_PRICE_CALC + 0.001, 2)
        Return ORDR_UNIT_PRICE_CALC
    End Function

    Public Shared Function Get_CTL_NOs( _
        USER_ID As String, _
        COMPUTER_NAME As String, _
        TABLE_NAME As String, _
        COLUMN_NAME As String, _
        How_Many As Integer) As String

        Dim CTL_NO As String = ""

        Dim using_sco As Boolean = False

        If using_sco Then
            ' to be determined with Dana tomorrow
        Else
            'CTL_NO = ASCMAIN1.Next_Control_No(TABLE_NAME & "." & COLUMN_NAME, How_Many)
            CTL_NO = ASCDATA1.ExecuteSF("TAPCTLN1", _
                                        New String() {"CTL_NO_TYPE_IN", "HOW_MANY_IN"}, _
                                        New Object() {TABLE_NAME & "." & COLUMN_NAME, How_Many})
            ASCMAIN1.sql = "Insert into TATCTLN2 " _
                & " (CTL_NO_TYPE,CTL_NO_LAST,CTL_NO_KEY,HOW_MANY,INIT_DATE,INIT_OPER)" _
                & " Values (:PARM1,:PARM2,:PARM3,:PARM4,SYSDATE,:PARM5)"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VNVNV", New Object() {TABLE_NAME & "." & COLUMN_NAME, Val(CTL_NO), CTL_NO, How_Many, USER_ID})
        End If

        'Dim SQL As String = "Insert into LOGTABLE Values ('" & USER_ID & "','" & MACHINE_NAME & "','" & TABLE_NAME & "','" & COLUMN_NAME & "'," & CStr(How_Many)

        'ASCDATA1.ExecuteSQL("Insert into LOGTABLE Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5)", _
        '                    "VVVVN", _
        '                    New Object() {USER_ID, MACHINE_NAME, TABLE_NAME, COLUMN_NAME, How_Many})

        Return CTL_NO
    End Function


    Public Shared Sub Track_Shipment(ByVal SHIP_VIA_CODE As String, ByVal SHIP_REF As String)

        ASCMAIN1.ActiveForm.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Locating POD")
        ASCMAIN1.sql = "Select CARRIER_URL_TRACKING, CARRIER_TRACKING_IND" _
        & " from SOTCARR1,SOTSVIA1 " _
        & " where SOTSVIA1.SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'" _
        & "   and SOTCARR1.CARRIER_CODE = SOTSVIA1.CARRIER_CODE"

        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, True)
        Dim CARRIER_URL_TRACKING As String = rowSOTCARR1.Item("CARRIER_URL_TRACKING") & String.Empty
        Dim CARRIER_TRACKING_IND As String = rowSOTCARR1.Item("CARRIER_TRACKING_IND") & String.Empty

        If CARRIER_TRACKING_IND = "I" Then
            ASCMAIN1.sql = "SELECT NVL(INV_NO_RESHIP, INV_NO) FROM SOTINVH1 WHERE SHIP_REF = :PARM1 AND SHIP_VIA_CODE = :PARM2"
            SHIP_REF = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {SHIP_REF, SHIP_VIA_CODE}) & String.Empty
        End If

        If CARRIER_URL_TRACKING = "" Then
            MsgBox("Cannot Determine Carrier Tracking Website URL for Ship Via " & SHIP_VIA_CODE, MsgBoxStyle.OkOnly, "Unable to Perform Requested Action")
        ElseIf SHIP_REF.Length = 0 AndAlso CARRIER_TRACKING_IND = "I" Then
            MsgBox("Cannot Locate tracking information from Tracking Number " & SHIP_REF, MsgBoxStyle.OkOnly, "Unable to Perform Requested Action")
        Else
            System.Diagnostics.Process.Start(CARRIER_URL_TRACKING & SHIP_REF)
        End If
        ASCMAIN1.ActiveForm.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


    'Public Shared Function Get_Raw_EDI(EDI_DOC_SEQ_NO As String) As String
    '    ASCMAIN1.sql = "Select GD.^DocumentBlobKEY^ RAW_DATA_FILE " _
    '        & " from GEN.^Document_tb^ GD, EDT850T1 " _
    '        & " where EDT850T1.GEN_DOC_NO = GD.^AppField1^" _
    '        & "   and GD.^TransactionSetID^ = '850'" _
    '        & "   and EDT850T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
    '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "^", Chr(34))

    '    'If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then
    '    '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GENAHA.")
    '    'End If

    '    Dim RAW_DATA As String = ""
    '    Dim RAW_DATA_FILE As String = ASCDATA1.GetDataValue
    '    If RAW_DATA_FILE <> "" Then
    '        'Dim FILENAME As String = "V:\Documents\" & RAW_DATA_FILE & ".DOC"
    '        Dim FILENAME As String = "\\192.168.170.103\gensrvnt\Documents\" & RAW_DATA_FILE & ".DOC"
    '        If My.Computer.FileSystem.FileExists(FILENAME) Then
    '            RAW_DATA = My.Computer.FileSystem.ReadAllText(FILENAME)
    '        End If
    '    End If
    '    Return RAW_DATA
    'End Function


    Public Shared Function Get_EDI_row(EDI_DOC_SEQ_NO As String, _
                                       Optional EDI_DOC_NO As String = "850", _
                                       Optional DocumentName As String = "") As DataRow
        Dim TABLE_NAME As String = "EDT850T1"
        If EDI_DOC_NO = "852" Then TABLE_NAME = "EDT852T1"
        If EDI_DOC_NO = "940" Then TABLE_NAME = "EDT940O1"
        If EDI_DOC_NO = "945" Then TABLE_NAME = "EDT945T1"
        If EDI_DOC_NO = "855" Then TABLE_NAME = "EDT855O1"

        'ASCMAIN1.sql = "Select GD.^DocumentBlobKEY^ RAW_DATA_FILE " _
        '    & " from GEN.^Document_tb^ GD, EDT850T1 " _
        '    & " where EDT850T1.GEN_DOC_NO = GD.^AppField1^" _
        '    & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'" _
        '    & "   and EDT850T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        ASCMAIN1.sql = "Select GD.*" _
            & " from GEN.^Document_tb^ GD, EDT850T1 " _
            & " where EDT850T1.GEN_DOC_NO = GD.^AppField1^" _
            & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'" _
            & "   and EDT850T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDT850T1", TABLE_NAME)
        If EDI_DOC_NO = "940" Then
            'ASCMAIN1.sql = "Select GD.^DocumentBlobKEY^ RAW_DATA_FILE " _
            '    & " from GEN.^Document_tb^ GD, EDT940O1 " _
            '    & " where GD.^DocumentName^ like '" & EDI_DOC_SEQ_NO & "%'" _
            '    & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'"
            ASCMAIN1.sql = "Select GD.* " _
                & " from GEN.^Document_tb^ GD " _
                & " where GD.^DocumentName^ like '" & EDI_DOC_SEQ_NO & "%'" _
                & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'"
        End If

        If EDI_DOC_NO = "855" Then
            ASCMAIN1.sql = "Select GD.* " _
                & " from GEN.^Document_tb^ GD " _
                & " where GD.^DocumentName^ = '" & DocumentName & "'" _
                & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'"
            ' WHY CANT WE USE APPFIELD 1?, or store THE RESPONSE IN A TABLE THAT WE CAN ACCESS BY GROUP NO?
            ' ALSO - WHAT IF THERE WERE MULTIPLE REQUESTS?  NOT REALLY CONNECTING THE DOTS PROPERLY HERE
        End If
        'If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then
        '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GENAHA.")
        'End If

        If EDI_DOC_NO = "856" Or EDI_DOC_NO = "810" Then
            ASCMAIN1.sql = "Select GD.* " _
                & " from GEN.^Document_tb^ GD " _
                & " where GD.^DocumentName^ = '" & DocumentName & "'" _
                & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'"
        End If

        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "^", Chr(34))

        If ASCMAIN1.CLIENT = "VAN" And EDI_DOC_NO = "850" Then
            ASCMAIN1.sql = ASCMAIN1.sql.Replace("EDT850T1.GEN_DOC_NO", "EDT850T1.EDI_DOC_NO")
        End If
        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.DBS_SERVER = "" Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GEN" & ASCMAIN1.CLIENT & ".")
        End If

        If EDI_DOC_NO = "850" Then
            ASCMAIN1.sql = ASCMAIN1.sql.Replace(" = '850'", " in ('850','875')")
        End If

        Return ASCDATA1.GetDataRow
    End Function

    Public Shared Function Get_Raw_EDI(EDI_DOC_SEQ_NO As String, _
                                       Optional ED_PARM_RAW_ARCHIVE As String = "", _
                                       Optional EDI_DOC_NO As String = "850", _
                                       Optional DocumentName As String = "") As String

        Dim row As DataRow = Get_EDI_row(EDI_DOC_SEQ_NO, EDI_DOC_NO, DocumentName)
        If row Is Nothing Then
            Return ""
        End If

        If ED_PARM_RAW_ARCHIVE = "" Then
            'Dim rowEDTPARM1 As DataRow = Lookup("EDTPARM1", "Z")
            'ED_PARM_RAW_ARCHIVE = rowEDTPARM1.Item("ED_PARM_RAW_ARCHIVE") & ""
            If ASCMAIN1.CLIENT = "NYA" Then ED_PARM_RAW_ARCHIVE = "\\192.168.170.103\gensrvnt\Documents\"
            If ASCMAIN1.CLIENT = "RGI" Then ED_PARM_RAW_ARCHIVE = "\\192.168.110.224\gensrvnt\Documents\"
        End If

        Dim RAW_DATA As String = ""
        Dim RAW_DATA_FILE As String = row.Item("DocumentBlobKEY") ' ASCDATA1.GetDataValue
        If RAW_DATA_FILE <> "" Then
            'Dim FILENAME As String = "V:\Documents\" & RAW_DATA_FILE & ".DOC"
            Dim FILENAME As String = ED_PARM_RAW_ARCHIVE & "\" & RAW_DATA_FILE & ".DOC"
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                RAW_DATA = My.Computer.FileSystem.ReadAllText(FILENAME)
            End If
        End If
        Return RAW_DATA
    End Function

    Public Shared Sub Log_Changes( _
                                frmASFBASE0 As ASFBASE0, _
                                ORDR_NO As String, _
                                row As DataRow, _
                                TABLE_NAME As String, _
                                ByRef Check_Changed_Fields As Boolean, _
                                REV_NO As Integer,
                                ByRef REV_LNO As Integer, _
                                LAST_DATE As Date)

        For i As Integer = 0 To row.Table.Columns.Count - 1
            Dim COLUMN_NAME As String = frmASFBASE0.dst.Tables(TABLE_NAME).Columns(i).ColumnName
            If row.Item(COLUMN_NAME) & "" _
            <> row.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                Check_Changed_Fields = True
                ' ASCMAIN1.Progress("-", COLUMN_NAME)
                Dim rowSOTORDXR As DataRow = frmASFBASE0.dst.Tables("SOTORDXR").NewRow
                With rowSOTORDXR
                    .Item("REV_NO") = REV_NO
                    REV_LNO += 1
                    .Item("REV_LNO") = REV_LNO
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = 0
                    .Item("INIT_DATE") = LAST_DATE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("COLUMN_NAME") = COLUMN_NAME
                    .Item("OLD_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Original)
                    .Item("NEW_VALUE") = row.Item(COLUMN_NAME)
                    .Item("EMODE") = frmASFBASE0.EntryMode
                End With
                frmASFBASE0.dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                Check_Changed_Fields = True
            End If
        Next i
    End Sub

    Public Shared Function Credit_Request(TERM_CODE As String, rowSOTORDR0 As DataRow) As String
        '
        Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO")
        Dim CUST_CODE As String = rowSOTORDR0.Item("CUST_CODE")
        Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
        Dim ORDR_AMT As Decimal = Val(rowSOTORDR0.Item("ORDR_AMT") & "")
        Dim CURR_CODE As String = "USD"

        ' Added 01/22/2019
        Dim rowARTCUST1 As DataRow = Nothing
        rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = (SELECT CUST_CODE FROM SOTORDR0 WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')")
        Dim rowGLTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM1 WHERE GL_PARM_KEY = 'Z'")
        Dim SEG4_CODE As String = rowGLTPARM1.Item("GL_PARM_DEF_SEG4") & String.Empty

        If rowARTCUST1 IsNot Nothing AndAlso rowARTCUST1.Item("SEG4_CODE") & String.Empty <> String.Empty Then
            SEG4_CODE = rowARTCUST1.Item("SEG4_CODE")
        End If

        ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO, SOTORDR1.CURR_CODE" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE_CURR) ORDR_AMT_CURR " & vbCrLf _
            & " from SOTORDR1,SOTORDR2" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "  and SOTORDR1.ORDR_GROUP_NO = :PARM1 group by SOTORDR1.ORDR_GROUP_NO, SOTORDR1.CURR_CODE"
        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {ORDR_GROUP_NO})

        If row IsNot Nothing Then
            CURR_CODE = row.Item("CURR_CODE") & ""
            ORDR_AMT = Val(row.Item("ORDR_AMT_CURR") & "")
        End If

        ' Modified 01/22/2019
        ASCMAIN1.sql = "Insert into EDT855O1 (" & vbCrLf _
            & "COMPANY_CODE,EDI_OUTBOUND_DOC_NO,ORDR_CUST_PO,ORDR_PO_DATE,ORDR_GROUP_NO," & vbCrLf _
            & "REQUEST_DATE,TERM_CODE,ORDR_SHIP_DATE,ORDR_CANCEL_DATE,AS_OF_DATE," & vbCrLf _
            & "ORDR_AMT,INIT_DATE,INIT_OPER,CURR_CODE,SEG4_CODE) Select " & vbCrLf _
            & "  :PARM1 COMPANY_CODE" & vbCrLf _
            & ", :PARM2 EDI_OUTBOUND_DOC_NO" & vbCrLf _
            & ", SOTORDR0.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR0.ORDR_DATE ORDR_PO_DATE" & vbCrLf _
            & ", SOTORDR0.ORDR_GROUP_NO ORDR_NO" & vbCrLf _
            & ", SYSDATE REQUEST_DTE" & vbCrLf _
            & ", :PARM3 TERM_CODE" & vbCrLf _
            & ", SOTORDR0.ORDR_SHIP_DATE" & vbCrLf _
            & ", SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
            & ", NULL AS_OF_DATE" & vbCrLf _
            & ", :PARM4 ORDR_AMT" & vbCrLf _
            & ", SYSDATE INIT_DATE" & vbCrLf _
            & ", :PARM5 INIT_OPER" & vbCrLf _
            & ", :PARM6 CURR_CODE" & vbCrLf _
            & ", :PARM7 SEG4_CODE" & vbCrLf _
            & " from SOTORDR0 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVNVVV", New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, TERM_CODE, ORDR_AMT, ASCMAIN1.USER_ID, CURR_CODE, SEG4_CODE})

        ASCMAIN1.sql = "Insert into EDT855O5 (" & vbCrLf _
            & "COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_ADR_SEQ,EDI_ADDR_TYPE," & vbCrLf _
            & "EDI_CUST_NAME_ADR,EDI_ADDRESS1,EDI_ADDRESS2,EDI_ADDRESS3," & vbCrLf _
            & "EDI_CITY,EDI_STATE,EDI_ZIPCODE,EDI_COUNTRY," & vbCrLf _
            & "EDI_ADDR_CODE,EDI_ADDR_CODE_QUAL) Select " & vbCrLf _
            & "'" & ASCMAIN1.DBS_COMPANY & "' COMPANY_CODE" & vbCrLf _
            & ", '" & EDI_OUTBOUND_DOC_NO & "' EDI_OUTBOUND_DOC_NO" & vbCrLf _
            & ", 1 EDI_ADR_SEQ, 'BY' EDI_ADDR_TYPE" & vbCrLf _
            & ", ARTCUST1.CUST_NAME EDI_CUST_NAME_ADR" & vbCrLf _
            & ", ARTCUST1.CUST_ADDR1 EDI_ADDRESS1" & vbCrLf _
            & ", ARTCUST1.CUST_ADDR2 EDI_ADDRESS2" & vbCrLf _
            & ", ARTCUST1.CUST_ADDR3 EDI_ADDRESS3" & vbCrLf _
            & ", ARTCUST1.CUST_CITY EDI_CITY" & vbCrLf _
            & ", ARTCUST1.CUST_STATE EDI_STATE" & vbCrLf _
            & ", ARTCUST1.CUST_ZIP_CODE EDI_ZIPCODE" & vbCrLf _
            & ", ARTCUST1.CUST_COUNTRY EDI_COUNTRY" & vbCrLf _
            & ", ARTCUST1.CUST_CODE EDI_ADDR_CODE" & vbCrLf _
            & ", 'BY' EDI_ADDR_CODE_QUAL" & vbCrLf _
            & " from ARTCUST1 where CUST_CODE = '" & CUST_CODE & "'"
        ASCDATA1.ExecuteSQL()

        ' THIS SHOULD BE PARAMETERIZED
        Dim EDI_TP_QUAL As String = "01"
        Dim EDI_TP_ID As String = "001921360"
        Dim EDI_DOC_NO As String = "855"

        'Dim rowEDTTRPM1 As DataRow = clsASCBASE1.LookUp("EDTTRPM1", New String() {EDI_TP_QUAL, EDI_TP_ID, EDI_DOC_NO})
        Dim sqlEDTTRPM1 = "Select * from EDTTRPM1 where EDI_TP_QUAL = :PARM1 and EDI_TP_ID = :PARM2 and EDI_DOC_NO = :PARM3"
        Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow(sqlEDTTRPM1, "VVV", New String() {EDI_TP_QUAL, EDI_TP_ID, EDI_DOC_NO})

        ASCMAIN1.sql = "Insert into EDTSYSIH (COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_APPLICATION_ID,EDI_PROCESS_IND," _
            & "EDI_OUR_ID,EDI_TP_ID,INIT_DATE,INIT_OPER)" _
            & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,SYSDATE,'" & ASCMAIN1.USER_ID & "')"
        Dim EDI_APPLICATION_ID As String = "PR"
        Dim EDI_PROCESS_IND As String = "1"
        '  EDI_PROCESS_IND = "T"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVVV", _
                New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, EDI_APPLICATION_ID, EDI_PROCESS_IND, _
                              rowEDTTRPM1.Item("EDI_OUR_ID"), EDI_TP_ID})

        Dim EVENT_DESC As String = String.Format("Request for Credit for {0} {1}", Format(ORDR_AMT, "#,##0.00"), CURR_CODE)
        ASCDATA1.ExecuteSQL("Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) " & vbCrLf _
                       & " Select 'SOTORDR1' TABLE_NAME, ORDR_NO TABLE_KEY, SYSDATE INIT_DATE, :PARM1 INIT_OPER, 'CR-REQ' EVENT_TYPE" & vbCrLf _
                       & ", :PARM2 EVENT_DESC, :PARM3 EVENT_KEY" & vbCrLf _
                       & "  from SOTORDR1 where ORDR_GROUP_NO = :PARM4", "VVVV", _
                       New Object() {ASCMAIN1.USER_ID, EVENT_DESC, EDI_OUTBOUND_DOC_NO, ORDR_GROUP_NO})

        Return EDI_OUTBOUND_DOC_NO
    End Function

    Public Shared Function Validate_Invoice_Date(DT As Date, MOS_BACK As Integer, MOS_FWD As Integer, ByRef EMsg As String) As Date()
        ASCMAIN1.sql = String.Format("SELECT DTE1, DTE2 FROM " _
            & "(SELECT PRD_END_DATE+1 DTE1 FROM GLTPARM2 WHERE OPS_YYYYPP = {0})," _
            & "(SELECT PRD_END_DATE DTE2 FROM GLTPARM2 WHERE OPS_YYYYPP = {1})", _
            ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 - MOS_BACK), ASCMAIN1.Period_Calc(ASCMAIN1.CYP, MOS_FWD))

        Dim row As DataRow = ASCDATA1.GetDataRow
        Dim dte1 As Date = row.Item("DTE1")
        Dim dte2 As Date = row.Item("DTE2")

        If DT & "" <> "" Then
            If Format(DT, "yyyyMMdd") < Format(dte1, "yyyyMMdd") _
            Or Format(DT, "yyyyMMdd") > Format(dte2, "yyyyMMdd") Then
                EMsg &= vbCr & "Valid Date Range is " & Format(dte1, "MM/dd/yyyy") & " thru " & Format(dte2, "MM/dd/yyyy")
            End If
        End If

        Return New Date() {dte1, dte2}
    End Function

    Public Shared Function IssueCredit(ByVal INV_NO As String, ByRef ErrorMessage As String) As Boolean

        Try

            IssueCredit = False


            Dim CCPA_NO As String = String.Empty

            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_NO = :PARM1", "V", INV_NO)

            If rowSOTINVH1 Is Nothing Then
                ErrorMessage = "Cannot Locate the supplied Credit Invoice No: " & INV_NO
                Exit Function
            End If

            If rowSOTINVH1.Item("CC_CRED_TRANS_ID") & String.Empty <> String.Empty Then
                ErrorMessage = "The credit was already issued a refund using a credit card."
                Exit Function
            End If

            Dim CreditAmount As Decimal = Math.Abs(Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty))
            If CreditAmount = 0 Then
                ErrorMessage = "Cannot process Credit for $0.00 credit"
                Exit Function
            End If

            Dim Transaction_ID As String = (rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty).ToString.Trim
            If Transaction_ID.Length = 0 Then
                ErrorMessage = "Invalid or Missing Transaction ID"
                Exit Function
            End If

            Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE TRANS_ID = :PARM1", "V", New Object() {Transaction_ID})
            If rowARTCCPA1 Is Nothing Then
                ErrorMessage = "Invalid or Missing Transaction ID"
                Exit Function
            End If

            Dim creditCard As String = (rowARTCCPA1.Item("CUST_CREDIT_CARD_LAST4") & String.Empty).ToString.Trim
            If creditCard.Length <> 4 Then
                creditCard = (rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty).ToString.Trim
            End If

            ' Get the last 4
            If creditCard.Length > 0 Then
                creditCard = StrReverse(StrReverse(creditCard).Substring(0, 4))
            End If

            If creditCard.Length <> 4 Then
                ErrorMessage = "Invalid or Missing Credit Card Number"
                Exit Function
            End If

            Dim rowSOTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPARM1 WHERE SO_PARM_KEY = :PARM1", "V", "Z")
            Dim ff As New ABSolution.ASFBASE1
            Dim CreditCardProcessor As New TAC.TAFCARDF(ff)

            Try
                CreditCardProcessor.test_mode = rowSOTPARM1.Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
                CreditCardProcessor.CUST_CODE = rowSOTINVH1.Item("CUST_CODE") & String.Empty
                CreditCardProcessor.CCPA_REASON = "M"
                CreditCardProcessor.TRAN_TYPE = "C"
                CreditCardProcessor.MerchantSetup()
                CreditCardProcessor.rowARTCCPA1 = rowARTCCPA1
                CCPA_NO = CreditCardProcessor.CC_Credit(Transaction_ID, CreditAmount, creditCard)
            Catch ex As Exception
                ErrorMessage = "Error Processing Credit Card Refund: " & ex.Message
            End Try

            If CCPA_NO.Length > 0 Then
                ASCMAIN1.sql = "Update SOTINVH1 set CCPA_NO = '" & CCPA_NO & "', CC_CRED_TRANS_ID = '" & CreditCardProcessor.MerchantTransID & "' WHERE INV_TYPE = 'C' AND INV_NO = '" & INV_NO & "'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update ARTCCPA1 set INV_NO = '" & INV_NO & "' WHERE CCPA_NO = '" & CCPA_NO & "'"
                ASCDATA1.ExecuteSQL()

                ErrorMessage = CCPA_NO
                IssueCredit = True
            Else
                ErrorMessage = "Could not process Credit Card Refund for the following reason: " & CreditCardProcessor.responseErrorMessage
            End If

        Catch ex As Exception
            ErrorMessage = "Error Processing Credit Card Refund: " & ex.Message
        End Try

    End Function
    Public Shared Sub GET_STYLE_COLOR_LOCATIONS(WHSE_CODE As String, STYLE_CODE As String, COLOR_CODE As String,
                                                ByRef LOCATION_CODE As String, ByRef LOCATION_ROUTE_SEQ As Int32, ByRef PICK_QTY As Int32)

        ASCMAIN1.sql = "SELECT X.LOCATION_CODE, WHTLOCM1.LOCATION_ROUTE_SEQ, X.LOCATION_QTY FROM WHTLOCM1,  " & vbCrLf _
                        & "(SELECT WHSE_CODE, LOCATION_CODE, LOCATION_QTY FROM WHTLOCB1 " & vbCrLf _
                        & "WHERE WHSE_CODE = '" & WHSE_CODE & "' AND  STYLE_CODE = '" & STYLE_CODE & "' AND " & vbCrLf _
                        & " COLOR_CODE = '" & COLOR_CODE & "' AND  LOCATION_QTY > 0  " & vbCrLf _
                        & " ORDER BY LOCATION_QTY,   LOCATION_CODE) X  " & vbCrLf _
                        & "WHERE WHTLOCM1.WHSE_CODE = X.WHSE_CODE AND WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE and NVL(WHTLOCM1.LOCATION_USE,'A') in ('A','E')" & vbCrLf _
                        & "ORDER BY X.LOCATION_QTY, WHTLOCM1.LOCATION_ROUTE_SEQ, X.LOCATION_CODE"

        Dim TBL As DataTable = ASCDATA1.GetDataTable
        'If TBL.Rows.Count > 0 Then
        '    Dim ROW As DataRow = TBL.Select("", "LOCATION_QTY DESC")(0)
        '    LOCATION_CODE = ROW.Item("LOCATION_CODE") & ""
        '    LOCATION_ROUTE_SEQ = Val(ROW.Item("LOCATION_ROUTE_SEQ") & "")
        'Else
        '    Dim dana As String = ""
        'End If
        'Dim LOCATION_QTY As Int64
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            LOCATION_CODE = row.Item("LOCATION_CODE") & ""
            LOCATION_ROUTE_SEQ = Val(row.Item("LOCATION_ROUTE_SEQ") & "")
            Dim DANAQTY As Int64 = Val(row.Item("LOCATION_QTY") & "")
            If Val(row.Item("LOCATION_QTY") & "") >= PICK_QTY Then
                Exit For
            End If
        Next
        'COLOR_CODE = "X"

    End Sub
    Public Shared Sub Prepare_Pick_Ticket_Locations(frmASFBASE0 As ASFBASE0, SOTPLOC1 As String, PICK_NO As String)
        ' HOWDO WE HANDLE THE CONSOLIDATED TICKETS?
        ' PICK TICKET OR CONSOLIDATED PICK TICKET?

        ASCMAIN1.sql = "Insert into SOTPLOC1 SELECT SPOTICK2.PICK_NO, SPOTICK2.PICK_LNO, WHTLOCB1.LOCATION_CODE, WHTLOCB1.LOCATION_QTY FROM WHTLOCB1, SOTPICK2  " _
                  & " where WHTLOCB1.STYLE_CODE = SOTPICK2.STYLE_CODE AND WHTLOCB1.COLOR_CODE = SOTPICK2.COLOR_CODE AND WHTLOCB1.LOCATION_QTY <> 0 " _
                  & " SOTPICK2.PICK_NO = " & PICK_NO
        ASCDATA1.ExecuteSQL()
        Dim SEQUENCE As Int64 = 0
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        If Not frmASFBASE0.dst.Tables.Contains("SOTPLOC1") Then
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "SOTPLOC1", "*")
        End If
        For Each rowSOTPLOC1 As DataRow In frmASFBASE0.dst.Tables("SOTPLOC1").Select(" ORDER BY STYLE_CODE, COLOR_CODE, LOCATION_QTY DESC ")
            If rowSOTPLOC1.Item("STYLE_CODE") & "" = STYLE_CODE And rowSOTPLOC1.Item("COLOR_CODE") & "" = COLOR_CODE Then
                SEQUENCE += 1
            Else
                SEQUENCE = 1
            End If
            rowSOTPLOC1.Item("SEQUENCE") = SEQUENCE
            STYLE_CODE = rowSOTPLOC1.Item("STYLE_CODE")
            COLOR_CODE = rowSOTPLOC1.Item("COLOR_CODE")
        Next

    End Sub

    Public Shared Function Create_Invoice(
            frmASFBASE0 As ASFBASE0,
            ByVal CTL_NO As String,
             Optional ByVal return_PDF As Boolean = True,
             Optional ByVal pro_forma As Boolean = False,
             Optional ByVal CTL_NO_type As String = "INV_NO",
             Optional ByVal chkCONS_INV As String = "Not SPECIFIED",
             Optional ByVal chkEXPORT_INFO As String = "") As String

        frmASFBASE0.Cursor = Cursors.WaitCursor

        ASCMAIN1.Progress("Now Preparing Invoice For Printing")

        If Not frmASFBASE0.ROWs.ContainsKey("ARTPARM1") Then
            frmASFBASE0.Get_PARM("ARTPARM1")
        End If

        Dim REPORT_NAME As String = "SORINVP1"
        If Not frmASFBASE0.REPORTS.ContainsKey(REPORT_NAME) Then
            frmASFBASE0.REPORTS.Add(REPORT_NAME, frmASFBASE0.Load_rptClass(REPORT_NAME))
            frmASFBASE0.REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        Dim RPT As String = frmASFBASE0.ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If RPT = "" Then RPT = REPORT_NAME

        ' If the user opens the same Invoice more than once it causes an error and ABSolution closes down.
        Dim tempFileName As String = CTL_NO & "_" & DateTime.Now.ToString("yyyyMMddHHmmss")

        Dim sql As String = "" ' " And SOTINVH1.INV_TYPE = 'I'"

        If InStr(CTL_NO, ",") = 0 Then
            sql &= " and SOTINVH1." & CTL_NO_type & " = '" & CTL_NO & "'"
        Else
            sql &= " and SOTINVH1." & CTL_NO_type & " in ('" & Replace(CTL_NO, ",", "','") & "')"
            tempFileName = Split(CTL_NO, ",")(0) & "_" & DateTime.Now.ToString("yyyyMMddHHmmss")
        End If

        'If INV_NO <> "" Then
        '    If InStr(INV_NO, ",") = 0 Then
        '        sql &= " and SOTINVH1.INV_NO = '" & INV_NO & "'"
        '    Else
        '        sql &= " and SOTINVH1.INV_NO in ('" & Replace(INV_NO, ",", "','") & "')"
        '        tempFileName = Split(INV_NO, ",")(0) & "_" & DateTime.Now.ToString("yyyyMMddHHmmss")
        '    End If
        'ElseIf SHIP_BOL_NO <> "" Then
        '    If InStr(SHIP_BOL_NO, ",") = 0 Then
        '        sql &= " and SOTINVH1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        '    Else
        '        sql &= " and SOTINVH1.SHIP_BOL_NO in ('" & Replace(SHIP_BOL_NO, ",", "','") & "')"
        '        tempFileName = Split(SHIP_BOL_NO, ",")(0) & "_" & DateTime.Now.ToString("yyyyMMddHHmmss")
        '    End If
        'ElseIf PICK_NO <> "" Then
        '    If InStr(PICK_NO, ",") = 0 Then
        '        sql &= " and SOTINVH1.PICK_NO = '" & PICK_NO & "'"
        '    Else
        '        sql &= " and SOTINVH1.PICK_NO in ('" & Replace(PICK_NO, ",", "','") & "')"
        '        tempFileName = Split(PICK_NO, ",")(0) & "_" & DateTime.Now.ToString("yyyyMMddHHmmss")
        '    End If
        'End If

        If chkCONS_INV = "NOT SPECIFIED" Then
            ASCMAIN1.sql = "Select * from SOTINVH1 " & ASCMAIN1.SQL_Add_WHERE(sql)
            Dim rowSOTINVH1_sample As DataRow = ASCDATA1.GetDataRow
            If rowSOTINVH1_sample IsNot Nothing AndAlso rowSOTINVH1_sample.Item("INV_NO_CONS") & "" <> "" Then
                chkCONS_INV = "1"
                sql = " and SOTINVH1.INV_NO in (Select Distinct INV_NO from SOTINVH1 where SOTINVH1.INV_NO_CONS in (Select Distinct INV_NO_CONS from SOTINVH1 " & ASCMAIN1.SQL_Add_WHERE(sql) & "))"
            Else
                chkCONS_INV = ""
            End If
        End If

        ' frmASFBASE0.REPORTS(REPORT_NAME).Fill_Records_RPT(sql)
        frmASFBASE0.REPORTS(REPORT_NAME).Fill_Records_RPT(New String() {sql, IIf(pro_forma, "1", "0"), "", "", chkCONS_INV})

        Dim FILENAME As String = ""
        With frmASFBASE0.REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("CONS_INV", chkCONS_INV)
            .CR_params.Add("EXPORT_INFO", chkEXPORT_INFO)
            Dim REPORT_NO As String = ""
            If return_PDF Then
                REPORT_NO = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
                FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
                .Print_Report_End(, True)
            Else
                REPORT_NO = .Generate_Report(RPT, "Sales Invoice", , True, , , , , False)
                FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
                .Print_Report_End()
            End If
        End With

        frmASFBASE0.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return FILENAME

    End Function

    Public Shared Function email_Invoice(frmASFBASE0 As ASFBASE0,
                      ByVal CUST_CODE As String,
                      ByVal CUST_NAME As String,
                      ByVal CUST_EMAIL As String,
                      ByVal CUST_CONTACT As String,
                      ByVal FILENAME As String,
                      ByVal ATTACHMENT As String,
                      ByVal SUBJECT As String,
                      ByVal INV_NO As String,
                      Optional ByVal ORDR_NO As String = "") As String

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        If CUST_EMAIL <> "" Then
            EMAIL_ADDRESSs.Add(CUST_EMAIL, IIf(CUST_CONTACT = "", CUST_EMAIL, CUST_CONTACT))
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        ATTACHMENTs.Add(ATTACHMENT, ATTACHMENT)

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, "INV", False, True, CUST_CODE, CUST_NAME, "Customer")

        If SEND_NO <> "" Then
            If ORDR_NO = "" Then
                ORDR_NO = ASCDATA1.GetDataValue("Select ORDR_NO from SOTINVH1 where INV_NO = '" & INV_NO & "'")
            End If
            If ORDR_NO <> "" Then
                TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO, frmASFBASE0.DATETIME_STAMP, ASCMAIN1.USER_ID, "EML", "email Invoice " & INV_NO, SEND_NO)
            End If
        End If

        Return SEND_NO
    End Function

    Public Shared Function Get_PO_Cost(frmASFBASE0 As ASFBASE0, STYLE_CODE As String, VEND_CODE As String, rowSOTORDR1 As DataRow) As Decimal

        Dim ORDR_SHIP_DATE As Date
        If rowSOTORDR1.Item("ORDR_SHIP_DATE") & "" = "" Then
            ORDR_SHIP_DATE = Now.Date
        Else
            ORDR_SHIP_DATE = rowSOTORDR1.Item("ORDR_SHIP_DATE")
        End If

        Dim PO_COST As Decimal = 0
        Dim rowICTSTYV1 As DataRow = frmASFBASE0.LookUp("ICTSTYV1", New String() {STYLE_CODE, VEND_CODE})
        If rowICTSTYV1 IsNot Nothing Then
            If rowICTSTYV1.Item("NEW_PO_COST_DATE") & "" <> "" AndAlso Format(rowICTSTYV1.Item("NEW_PO_COST_DATE"), "yyyyMMdd") <= Format(ORDR_SHIP_DATE, "yyyyMMdd") Then
                PO_COST = Val(rowICTSTYV1.Item("NEW_PO_COST") & "")
            Else
                PO_COST = Val(rowICTSTYV1.Item("PO_COST") & "")
            End If
        End If
        Return PO_COST
    End Function

    Public Shared Sub Update_Credit_Authorizations()

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select EDT855T1.ORDR_NO ORDR_GROUP_NO, EDT855T1.APPROVAL_NO, EDT855T1.EDI_TP_ID" & vbCrLf _
            & ", EDT855T1.DECISION_AMT, EDT855T1.DECISION_CODE, EDT855T1.EDI_DOC_SEQ_NO" _
            & ", NVL(SOTAUTH1.ORDR_GROUP_NO,'0000000000') ORDR_GROUP_NO_AUTH" & vbCrLf _
            & " from SOTAUTH1,EDT855T1 where EDI_DOC_SEQ_NO IN (" & vbCrLf _
            & "Select MAX (EDI_DOC_SEQ_NO) EDI_DOC_SEQ_NO" & vbCrLf _
            & " from EDT855T1 where ORDR_NO IN (Select Distinct SOTORDR1.ORDR_GROUP_NO from SOTORDR1 where ORDR_STATUS = 'O')" & vbCrLf _
            & " group by ORDR_NO)" & vbCrLf _
            & "   and SOTAUTH1.ORDR_GROUP_NO (+) = EDT855T1.ORDR_NO" & vbCrLf _
            & "   and NVL(SOTAUTH1.ORDR_CRED_CLR_AUTH,'X') <> 'A'" & vbCrLf _
            & "   and NVL(SOTAUTH1.ORDR_CRED_CLR_AUTH_CTL_NO,'0000000000') < EDT855T1.EDI_DOC_SEQ_NO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   IF R1.ORDR_GROUP_NO_AUTH = '0000000000' THEN" & vbCrLf _
            & "    Insert into SOTAUTH1 (ORDR_GROUP_NO, INIT_OPER, INIT_DATE)" & vbCrLf _
            & "     Values (R1.ORDR_GROUP_NO, 'XXX', SYSDATE);" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "   Update SOTAUTH1 SET " & vbCrLf _
            & "    ORDR_CRED_CLR_BY = R1.EDI_TP_ID," & vbCrLf _
            & "    ORDR_CRED_CLR_AUTH = DECODE(R1.DECISION_CODE,'A7','A',R1.DECISION_CODE)," & vbCrLf _
            & "    ORDR_CRED_CLR_DATE = SYSDATE," & vbCrLf _
            & "    ORDR_CRED_CLR_AUTH_TYPE = 'F'," & vbCrLf _
            & "    ORDR_CRED_CLR_AUTH_NO = R1.APPROVAL_NO," & vbCrLf _
            & "    ORDR_CRED_CLR_AUTH_CTL_NO = R1.EDI_DOC_SEQ_NO," & vbCrLf _
            & "    ORDR_CRED_CLR_AUTH_AMT = R1.DECISION_AMT," & vbCrLf _
            & "    LAST_OPER = 'XXX'," & vbCrLf _
            & "    LAST_DATE = SYSDATE" & vbCrLf _
            & "   where ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"

        ASCDATA1.ExecuteSQL()
    End Sub

    Public Shared Sub Create_943_for_Transfer_Order_Receipt(frmASFBASE0 As ASFBASE0, rowSOTINVH1 As DataRow)

        ' Conditions for calling - rowSOTINVH1.Item("ORDR_TYPE_CODE") = "XFR"

        Dim rowICTWHSE1 As DataRow = frmASFBASE0.LookUp("ICTWHSE1", rowSOTINVH1.Item("WHSE_CODE_TO"))

        If rowICTWHSE1.Item("LP_CODE") & "" = "TSI" Then
            Dim rowEDTTRPM1 As DataRow = frmASFBASE0.LookUp("EDTTRPM1",
                                                New String() {rowICTWHSE1.Item("WHSE_EDI_QUAL"), rowICTWHSE1.Item("WHSE_EDI_ID"), "943"})

            If Not frmASFBASE0.dst.Tables.Contains("EDT943O1") Then
                frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "EDT943O1", "*")
                frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "EDT943O2", "*")
            End If

            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim SHIP_BOL_NO As String = rowSOTINVH1.Item("SHIP_BOL_NO")
            Dim PICK_NO As String = rowSOTINVH1.Item("PICK_NO")
            Dim rowSOTSHIP1 As DataRow = frmASFBASE0.LookUp("SOTSHIP1", SHIP_BOL_NO)

            Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
            Dim CONTAINER_NO As String = "XFR_" & Mid(SHIP_BOL_NO, 5, 6)

            Dim rowEDT943O1 As DataRow = frmASFBASE0.dst.Tables("EDT943O1").NewRow
            With rowEDT943O1
                .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                .Item("EDI_REPORTING_CODE") = "J"
                .Item("EDI_DEPOSITOR_ORDER_NO") = ""
                .Item("EDI_SHIPMENT_DATE") = rowSOTSHIP1.Item("SHIPPED_ACTUAL")
                .Item("EDI_PO_SHIPMENT_NO") = PICK_NO ' SHIP_BOL_NO - CHG MADE A REQ OF HUE
                .Item("EDI_NAME") = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME")
                .Item("EDI_WH_ID_CODE") = ""
                .Item("EDI_DIVISION_CODE") = rowICTWHSE1.Item("LP_WHSE_ID")
                .Item("EDI_ARRIVAL_DATE") = CDate(rowSOTSHIP1.Item("SHIPPED_ACTUAL")).AddDays(1)
                .Item("EDI_CARRIER_SCAC") = ""
                .Item("EDI_PALLET_QTY") = 0
                .Item("EDI_SEAL_NUMBER_CONTAINER") = CONTAINER_NO
                .Item("EDI_SEAL_NUMBER") = ""
                .Item("INIT_DATE") = frmASFBASE0.DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
            End With
            frmASFBASE0.dst.Tables("EDT943O1").Rows.Add(rowEDT943O1)

            ASCMAIN1.sql = "Select SOTINVH2.*" & vbCrLf _
                & ",ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.CUST_CODE CUST_CODE_STYLE" & vbCrLf _
                & ",ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY" & vbCrLf _
                & ",ICTSTYC1.UPC_CODE, ICTSTYC1.HIDE_COLOR_3PL" & vbCrLf _
                & " from SOTINVH2,ICTSTYL1,ICTSTYC1" & vbCrLf _
                & " where SOTINVH2.INV_TYPE = 'I' AND SOTINVH2.INV_NO = '" & INV_NO & "'" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = SOTINVH2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = SOTINVH2.COLOR_CODE" & vbCrLf

            Dim lead_item_processed As Boolean = False

            Dim EDI_DOC_LNO As Int64 = 0
            For Each rowSOTINVH2 As DataRow In ASCDATA1.GetDataTable.Select("", "STYLE_CODE,COLOR_CODE")

                Dim STYLE_CODE As String = rowSOTINVH2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTINVH2.Item("COLOR_CODE")
                Dim CUST_CODE_STYLE As String = rowSOTINVH2.Item("CUST_CODE_STYLE") & ""
                Dim STYLE_GROUP_CODE As String = rowSOTINVH2.Item("STYLE_GROUP_CODE") & ""
                Dim HIDE_COLOR_3PL As String = rowSOTINVH2.Item("HIDE_COLOR_3PL") & ""

                If Not lead_item_processed Then
                    Dim EDI_DIVISION_CODE As String = rowEDT943O1.Item("EDI_DIVISION_CODE")
                    If CUST_CODE_STYLE = "DOLGEN" Then
                        EDI_DIVISION_CODE = "NYDG"
                    ElseIf CUST_CODE_STYLE = "WALMART" Then
                        If STYLE_GROUP_CODE = "07" Then
                            EDI_DIVISION_CODE = "NYWB"
                        Else
                            EDI_DIVISION_CODE = "NYWM"
                        End If
                    End If
                    rowEDT943O1.Item("EDI_DIVISION_CODE") = EDI_DIVISION_CODE
                    lead_item_processed = True
                End If

                Dim rowEDT943O2 As DataRow = frmASFBASE0.dst.Tables("EDT943O2").NewRow
                With rowEDT943O2
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    EDI_DOC_LNO += 1
                    .Item("EDI_DOC_LNO") = EDI_DOC_LNO
                    .Item("EDI_UNITS_SHIPPED") = rowSOTINVH2.Item("ORDR_QTY_SHIP")
                    .Item("EDI_UPC_CASE_CODE") = rowSOTINVH2.Item("UPC_CODE")
                    .Item("EDI_STYLE_NO") = IIf(HIDE_COLOR_3PL = "1", STYLE_CODE, STYLE_CODE & "" & COLOR_CODE)
                    .Item("EDI_CUST_STYLE_NO") = ""
                    .Item("EDI_ITEM_DESC") = rowSOTINVH2.Item("STYLE_DESC")

                    Dim EDI_SUB_INNER_QTY As Int64 = 0
                    Dim CARTON_PACK_QTY As Int64 = Val(rowSOTINVH2.Item("CARTON_PACK_QTY") & "")
                    Dim INNER_PACK_QTY As Int64 = Val(rowSOTINVH2.Item("INNER_PACK_QTY") & "")
                    If INNER_PACK_QTY <> 0 And CARTON_PACK_QTY <> 0 AndAlso CARTON_PACK_QTY Mod INNER_PACK_QTY = 0 Then
                        EDI_SUB_INNER_QTY = CARTON_PACK_QTY / INNER_PACK_QTY
                    End If
                    .Item("EDI_SUB_INNER_QTY") = EDI_SUB_INNER_QTY

                    .Item("EDI_PACK_QTY") = CARTON_PACK_QTY
                    '.Item("EDI_SIZE") = CARTON_VOLUME
                    '.Item("EDI_WEIGHT") = rowPOTSHIPD.Item("CASE_WEIGHT_GRS")
                    .Item("EDI_PO_ORDER_NO") = ""
                End With
                frmASFBASE0.dst.Tables("EDT943O2").Rows.Add(rowEDT943O2)
            Next

            '4012453780      TAYLORED

            ASCMAIN1.sql = "Insert into EDTSYSIH (COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_APPLICATION_ID,EDI_PROCESS_IND," _
                & "EDI_OUR_ID,EDI_TP_ID,INIT_DATE,INIT_OPER)" _
                & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,SYSDATE,'" & ASCMAIN1.USER_ID & "')"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVVV",
                    New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, "AR", "1",
                                  rowEDTTRPM1.Item("EDI_OUR_ID"), rowICTWHSE1.Item("WHSE_EDI_ID")})

            frmASFBASE0.Update_Record_TDA("EDT943O1")
            frmASFBASE0.Update_Record_TDA("EDT943O2")
        End If

    End Sub

    Public Shared Sub Update_Status_by_Date(
                                           frmASFBASE0 As ASFBASE0,
                                           ICTSTDQ1 As String,
                                           ICTSTDQ2 As String,
                                           ICTSTDQ3 As String,
                                           WHSE_CODE_to_allocate As String,
                                           allocation_only As Boolean,
                                           SOTORDR2 As String,
                                           force_pick As Boolean)

        Dim sqlx As String = IIf(WHSE_CODE_to_allocate = "", "", " and WHSE_CODE = '" & WHSE_CODE_to_allocate & "'")
        Dim sqlx2 As String = ""
        If Not allocation_only Then
            sqlx2 = " and (STYLE_CODE,COLOR_CODE) in (Select Distinct STYLE_CODE,COLOR_CODE from " & SOTORDR2 & ")"
        End If

        ASCMAIN1.Progress("Updating Status/Date", "")

        If frmASFBASE0.Name = "SOROREL1" Then
            frmASFBASE0.Create_BAs("ICTSTDQ1")
            frmASFBASE0.Update_BAs("ICTSTDQ1")
        Else
            frmASFBASE0.Update_Record_TDA("ICTSTDQ1")
        End If

        ASCDATA1.ExecuteSQL("Delete from ICTSTDQ1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTSTDQ1 & ")" & sqlx & sqlx2)
        ASCDATA1.ExecuteSQL("Insert into ICTSTDQ1 Select * from " & ICTSTDQ1 & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTSTDQ1 & ")" & sqlx & sqlx2)

        Dim WSC As String = ""
        Dim WSCi As Integer = 0
        Dim rowICTSTDQ2 As DataRow = Nothing
        For Each rowICTSTDQ1 As DataRow In frmASFBASE0.dst.Tables("ICTSTDQ1").Select("", "WHSE_CODE,STYLE_CODE,COLOR_CODE,STATUS_DATE")
            Dim WHSE_CODE = rowICTSTDQ1.Item("WHSE_CODE")
            Dim STYLE_CODE = rowICTSTDQ1.Item("STYLE_CODE")
            Dim COLOR_CODE = rowICTSTDQ1.Item("COLOR_CODE")
            If WSC <> WHSE_CODE & ":" & STYLE_CODE & ":" & COLOR_CODE Then
                WSC = WHSE_CODE & ":" & STYLE_CODE & ":" & COLOR_CODE
                WSCi = 0
            End If
            WSCi += 1

            If WSCi = 1 Then
                rowICTSTDQ2 = frmASFBASE0.dst.Tables("ICTSTDQ2").Rows.Add _
                    (New Object() {WHSE_CODE,
                                   STYLE_CODE,
                                   COLOR_CODE,
                                   rowICTSTDQ1.Item("STATUS_DATE"),
                                   rowICTSTDQ1.Item("QTY_ATS_CUM"),
                                   DBNull.Value, DBNull.Value,
                                   DBNull.Value, DBNull.Value,
                                   DBNull.Value, DBNull.Value,
                                   rowICTSTDQ1.Item("QTY_ATS")})
            Else
                Dim maxBuckets As Integer = 4
                If ASCMAIN1.CLIENT = "NYA" Then maxBuckets = 9

                If WSCi > maxBuckets Then WSCi = maxBuckets
                rowICTSTDQ2.Item("DATE_" & CStr(WSCi)) = rowICTSTDQ1.Item("STATUS_DATE")
                rowICTSTDQ2.Item("QTY_" & CStr(WSCi)) = rowICTSTDQ1.Item("QTY_ATS_CUM")
                rowICTSTDQ2.Item("ADD_" & CStr(WSCi)) = rowICTSTDQ1.Item("QTY_ATS")
            End If
        Next

        If frmASFBASE0.Name = "SOROREL1" Then
            frmASFBASE0.Create_BAs("ICTSTDQ2")
            frmASFBASE0.Update_BAs("ICTSTDQ2")
        Else
            frmASFBASE0.Update_Record_TDA("ICTSTDQ2", "1=1")
        End If

        'frmASFBASE0.Update_Record_TDA("ICTSTDQ2", "1=1")

        ASCDATA1.ExecuteSQL("Delete from ICTSTDQ2 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTSTDQ2 & ")" & sqlx & sqlx2)
        ASCDATA1.ExecuteSQL("Insert into ICTSTDQ2 Select * from " & ICTSTDQ2 & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTSTDQ2 & ")" & sqlx & sqlx2)


        If ASCMAIN1.CLIENT = "RGI" Then

            If Not force_pick Then
                If frmASFBASE0.Name = "SOROREL1" Then
                    frmASFBASE0.Create_BAs("ICTSTDQ3")
                    frmASFBASE0.Update_BAs("ICTSTDQ3")
                Else
                    frmASFBASE0.Update_Record_TDA("ICTSTDQ3", "1=1")
                End If

                ASCDATA1.ExecuteSQL("Delete from ICTSTDQ3 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTSTDQ2 & ")" & sqlx2)
                ASCDATA1.ExecuteSQL("Insert into ICTSTDQ3 Select * from " & ICTSTDQ3 & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTSTDQ2 & ")" & sqlx2)
            End If
        End If

    End Sub


    Public Shared Sub Generate_855(ByVal clsASCBASE1 As ASCBASE1,
                            ByVal ORDR_GROUP_NO As String)

        Dim EDI_OUR_ID As String = ""
        Dim EDI_DOC_SEQ_NO As String = ""
        Dim EDI_PURP_CODE As String = ""
        Dim Err_Msg As String = ""
        Dim dt As Date = Now + ASCMAIN1.NowTSD

        ASCMAIN1.sql = " Select Distinct R0.ORDR_SHIP_DATE, R0.ORDR_CANCEL_DATE, R0.ORDR_AMT," & vbCrLf _
            & " T1.EDI_DOC_SEQ_NO, T1.EDI_TP_QUAL, T1.EDI_TP_ID,  " & vbCrLf _
            & " T1.EDI_OUR_ID, T1.EDI_PO_NO, T1.EDI_PO_DATE, T1.EDI_SUPPLIER_NO" & vbCrLf _
            & " from SOTORDR0 R0, SOTORDR1 R1, EDT850T1 T1 " & vbCrLf _
            & " Where R0.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            & " And R0.ORDR_GROUP_NO = R1.ORDR_GROUP_NO " & vbCrLf _
            & " And R1.EDI_DOC_SEQ_NO is Not Null " & vbCrLf _
            & " And R1.EDI_DOC_SEQ_NO = T1.EDI_DOC_SEQ_NO"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        If rowSOTORDR1 IsNot Nothing Then
            EDI_DOC_SEQ_NO = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""
            Dim EDI_Outbound_Doc_No As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
            EDI_OUR_ID = Replace(rowSOTORDR1.Item("EDI_OUR_ID"), " ", "")


            ASCMAIN1.sql = "Select M1.* from EDTTRPM1 M1, EDTSLSP1 P1" & vbCrLf _
            & " Where M1.EDI_DOC_NO = '855'" & vbCrLf _
            & " And M1.EDI_TP_QUAL = '" & rowSOTORDR1.Item("EDI_TP_QUAL") & "'" & vbCrLf _
            & " And M1.EDI_TP_ID = rtrim('" & rowSOTORDR1.Item("EDI_TP_ID") & "')" & vbCrLf _
            & " And M1.CUST_CODE = P1.CUST_CODE" & vbCrLf
            Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow

            If rowEDTTRPM1 IsNot Nothing Then
                Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID")

                Dim rowEDT855O1 As DataRow = clsASCBASE1.dst.Tables("EDT855O1").NewRow  ' TBLs("EDT810O1").NewRow
                With rowEDT855O1
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                    .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("EDI_PO_NO")
                    .Item("ORDR_PO_DATE") = rowSOTORDR1.Item("EDI_PO_DATE")
                    .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                    .Item("EDI_PURPOSE_CODE") = "00"
                    .Item("REQUEST_DATE") = dt
                    .Item("ORDR_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                    .Item("ORDR_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
                    .Item("ORDR_AMT") = Val(rowSOTORDR1.Item("ORDR_AMT") & "")
                    .Item("INIT_DATE") = dt
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("EDI_ACK_TYPE") = "AC"
                    .Item("EDI_SUPPLIER_NO") = rowSOTORDR1.Item("EDI_SUPPLIER_NO") & ""
                End With
                clsASCBASE1.dst.Tables("EDT855O1").Rows.Add(rowEDT855O1)


                ASCMAIN1.sql = "Select * from EDT850T2" & vbCrLf _
                & " Where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                For Each rowEDT850T2 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim rowEDT855O2 As DataRow = clsASCBASE1.dst.Tables("EDT855O2").NewRow  ' TBLs("EDT810O1").NewRow
                    With rowEDT855O2
                        .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                        .Item("EDI_DTL_SEQ") = rowEDT850T2.Item("EDI_DTL_SEQ")
                        .Item("EDI_TOTAL_QTY") = rowEDT850T2.Item("EDI_TOTAL_QTY")
                        .Item("EDI_PRICE") = rowEDT850T2.Item("EDI_PRICE")
                        .Item("EDI_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
                        .Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY")
                        .Item("EDI_PO4_INNER") = rowEDT850T2.Item("EDI_PO4_INNER")
                        .Item("EDI_PO4_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
                        .Item("EDI_ITEM") = rowEDT850T2.Item("EDI_ITEM")
                        .Item("EDI_UPC") = rowEDT850T2.Item("EDI_UPC")
                        .Item("EDI_SKU") = rowEDT850T2.Item("EDI_SKU")
                        .Item("EDI_GTIN") = rowEDT850T2.Item("EDI_GTIN")
                        .Item("EDI_STYLE_DESC") = rowEDT850T2.Item("EDI_STYLE_DESC")
                    End With
                    clsASCBASE1.dst.Tables("EDT855O2").Rows.Add(rowEDT855O2)
                Next


                ASCMAIN1.sql = "Select * from EDT850T3" & vbCrLf _
                & " Where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                For Each rowEDT850T3 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim rowEDT855O3 As DataRow = clsASCBASE1.dst.Tables("EDT855O3").NewRow  ' TBLs("EDT810O1").NewRow
                    With rowEDT855O3
                        .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                        .Item("EDI_DTL_SEQ") = rowEDT850T3.Item("EDI_DTL_SEQ")
                        .Item("EDI_SDQ_SEQ") = rowEDT850T3.Item("EDI_SDQ_SEQ")
                        .Item("EDI_STORE_01") = rowEDT850T3.Item("EDI_STORE_01")
                        .Item("EDI_QTY_01") = rowEDT850T3.Item("EDI_QTY_01")
                        .Item("EDI_STORE_02") = rowEDT850T3.Item("EDI_STORE_02")
                        .Item("EDI_QTY_02") = rowEDT850T3.Item("EDI_QTY_02")
                        .Item("EDI_STORE_03") = rowEDT850T3.Item("EDI_STORE_03")
                        .Item("EDI_QTY_03") = rowEDT850T3.Item("EDI_QTY_03")
                        .Item("EDI_STORE_04") = rowEDT850T3.Item("EDI_STORE_04")
                        .Item("EDI_QTY_04") = rowEDT850T3.Item("EDI_QTY_04")
                        .Item("EDI_STORE_05") = rowEDT850T3.Item("EDI_STORE_05")
                        .Item("EDI_QTY_05") = rowEDT850T3.Item("EDI_QTY_05")
                        .Item("EDI_STORE_06") = rowEDT850T3.Item("EDI_STORE_06")
                        .Item("EDI_QTY_06") = rowEDT850T3.Item("EDI_QTY_06")
                        .Item("EDI_STORE_07") = rowEDT850T3.Item("EDI_STORE_07")
                        .Item("EDI_QTY_07") = rowEDT850T3.Item("EDI_QTY_07")
                        .Item("EDI_STORE_08") = rowEDT850T3.Item("EDI_STORE_08")
                        .Item("EDI_QTY_08") = rowEDT850T3.Item("EDI_QTY_08")
                        .Item("EDI_STORE_09") = rowEDT850T3.Item("EDI_STORE_09")
                        .Item("EDI_QTY_09") = rowEDT850T3.Item("EDI_QTY_09")
                        .Item("EDI_STORE_10") = rowEDT850T3.Item("EDI_STORE_10")
                        .Item("EDI_QTY_10") = rowEDT850T3.Item("EDI_QTY_10")
                        .Item("EDI_SDQ_UOM") = rowEDT850T3.Item("EDI_SDQ_UOM")
                        .Item("EDI_SDQ_QUAL") = rowEDT850T3.Item("EDI_SDQ_QUAL")
                    End With
                    clsASCBASE1.dst.Tables("EDT855O3").Rows.Add(rowEDT855O3)
                Next

                ASCMAIN1.sql = "Select * from EDT850T5 " & vbCrLf _
                & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
                & " And EDI_ADDR_TYPE <> 'SF' "
                For Each rowEDT850T5 As DataRow In ASCDATA1.GetDataTable.Rows

                    Dim rowEDT855O5 As DataRow = clsASCBASE1.dst.Tables("EDT855O5").NewRow  ' TBLs("EDT810O1").NewRow
                    With rowEDT855O5
                        .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                        .Item("EDI_ADR_SEQ") = rowEDT850T5.Item("EDI_ADR_SEQ") & ""
                        .Item("EDI_ADDR_TYPE") = rowEDT850T5.Item("EDI_ADDR_TYPE") & ""
                        .Item("EDI_CUST_NAME_ADR") = rowEDT850T5.Item("EDI_CUST_NAME_ADR") & ""
                        .Item("EDI_ADDRESS1") = rowEDT850T5.Item("EDI_ADDRESS1") & ""
                        .Item("EDI_ADDRESS2") = rowEDT850T5.Item("EDI_ADDRESS2") & ""
                        .Item("EDI_CITY") = rowEDT850T5.Item("EDI_CITY") & ""
                        .Item("EDI_STATE") = rowEDT850T5.Item("EDI_STATE") & ""
                        .Item("EDI_ZIPCODE") = rowEDT850T5.Item("EDI_ZIPCODE") & ""
                        .Item("EDI_COUNTRY") = rowEDT850T5.Item("EDI_COUNTRY") & ""
                        .Item("EDI_ADDR_CODE") = rowEDT850T5.Item("EDI_ADDR_CODE") & ""
                        .Item("EDI_ADDR_CODE_QUAL") = rowEDT850T5.Item("EDI_ADDR_CODE_QUAL") & ""
                    End With
                    clsASCBASE1.dst.Tables("EDT855O5").Rows.Add(rowEDT855O5)
                Next

                clsASCBASE1.Update_Record_TDA("EDT855O1")
                clsASCBASE1.Update_Record_TDA("EDT855O2")
                clsASCBASE1.Update_Record_TDA("EDT855O3")
                clsASCBASE1.Update_Record_TDA("EDT855O5")


                ASCMAIN1.sql = " Insert into EDTSYSIH values (" _
                & "'" & ASCMAIN1.DBS_COMPANY & "'," _
                & "'" & EDI_Outbound_Doc_No _
                & "','PR','T','" _
                & EDI_OUR_ID & "','" _
                & EDI_TP_ID & "','" _
                & Format(Now, "dd-MMM-yy") & "','" _
                & ASCMAIN1.USER_ID & "')"
                '                ASCMAIN1.sql = " Insert into EDTSYSIH values (" _
                '& "'" & ASCMAIN1.DBS_COMPANY & "'," _
                '& "'" & EDI_Outbound_Doc_No _
                '& "','OPR','" & TAC.TACMAIN1.EDI_PROCESS_IND & "','" _
                '& EDI_TP_ID & "')"
                ASCDATA1.ExecuteSQL()
            Else
                Err_Msg = Err_Msg & "No 855 Partner Setup"
            End If
        Else
            Err_Msg = Err_Msg & "Could Not Link Sales Order to 850"
        End If
    End Sub

    Public Shared Sub CreateWebInvoice(ByRef f As ABSolution.ASFBASE0, ByVal InvoiceType As String, ByVal InvoiceNo As String)

        Try

            If Not (ASCMAIN1.DBS_SERVER = "RGI" AndAlso ASCMAIN1.DBS_COMPANY = "RGI") Then
                Exit Sub
            End If

            ' Only permit this to work in Production for regency.
            If Not (ASCMAIN1.DBS_SERVER = ASCMAIN1.DBS_COMPANY) Then
                Exit Sub
            End If

            Dim rowSOTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPARM1 WHERE SO_PARM_KEY = 'Z'")
            If Not rowSOTPARM1.Table.Columns.Contains("SO_PARM_WEB_INVOICES") Then
                Exit Sub
            End If

            Dim SO_PARM_WEB_INVOICES As String = (rowSOTPARM1.Item("SO_PARM_WEB_INVOICES") & String.Empty).ToString.Trim
            If SO_PARM_WEB_INVOICES.Length = 0 Then
                Exit Sub
            End If

            If Not My.Computer.FileSystem.DirectoryExists(SO_PARM_WEB_INVOICES) Then
                Exit Sub
            End If

            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_TYPE = :PARM1 AND INV_NO = :PARM2", "VV", New Object() {InvoiceType, InvoiceNo})
            If rowSOTINVH1 Is Nothing Then
                Exit Sub
            End If

            If Not SO_PARM_WEB_INVOICES.EndsWith("\") Then
                SO_PARM_WEB_INVOICES &= "\"
            End If

            ' May need to create the directory for the customer
            SO_PARM_WEB_INVOICES &= rowSOTINVH1.Item("CUST_CODE")
            If Not My.Computer.FileSystem.DirectoryExists(SO_PARM_WEB_INVOICES) Then
                My.Computer.FileSystem.CreateDirectory(SO_PARM_WEB_INVOICES)
            End If
            SO_PARM_WEB_INVOICES &= "\"

            Dim RPT As String = "SORINVP1"
            Dim webFileName As String = InvoiceNo
            Dim CONS_INV As String = "0"
            If rowSOTINVH1.Item("INV_NO_CONS") & String.Empty <> String.Empty Then
                CONS_INV = "1"
            End If

            If Not f.REPORTS.ContainsKey(RPT) Then
                f.REPORTS.Add(RPT, f.Load_rptClass(RPT))
                f.REPORTS(RPT).Prepare_dst(False, "")
            End If

            f.REPORTS(RPT).Fill_Records_RPT(New String() {" and SOTINVH1.INV_TYPE = '" & InvoiceType & "' and SOTINVH1.INV_NO = '" & InvoiceNo & "'"})

            Dim REPORT_NO As String = String.Empty
            With f.REPORTS(RPT).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", CONS_INV)
                .CR_params.Add("EXPORT_INFO", "0")

                ' Set the customers Invoice
                Select Case ASCMAIN1.DBS_COMPANY
                    Case "RGI"
                        RPT = "SORINVPR"
                End Select

                REPORT_NO = .Generate_Report(RPT, "Invoice", , True, , , "PDF", webFileName, False)
                .Print_Report_End(True, True)
            End With

            My.Computer.FileSystem.MoveFile(ASCMAIN1.Folders("Temp") & webFileName & ".pdf", SO_PARM_WEB_INVOICES & webFileName & ".pdf", True)

        Catch ex As Exception
            MessageBox.Show("PDF of Invoice could not be placed in the expected directory. The following error occurred: " & ex.Message, "Web Invoice", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Shared Sub Build_ARTCSUMC(
        frmASFBASE0 As ASFBASE0,
        CUST_CODE As String,
        YP0 As String,
        YP1 As String,
        ARTCSUMC As String,
        Optional clear_ARTCSUMB As Boolean = True,
        Optional SALES_DIVISION_CODE As String = "")

        If clear_ARTCSUMB Then
            ASCMAIN1.sql = "Truncate Table " & ARTCSUMC
            ASCDATA1.ExecuteSQL()
        End If

        Dim Xs As String = ""
        For I As Integer = 1 To 12
            Xs &= ", X.AMT" & Format(I, "00")
        Next

        '        Dim sqls As New Dictionary(Of String, String)

        Dim sqlSOTINVH2where As String = "" _
            & " where SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & YP0 & "'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & YP1 & "'" & vbCrLf _
            & IIf(SALES_DIVISION_CODE <> "", "   and SOTINVH1.SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'" & vbCrLf, "")

        Dim sqlARTPYMT5where As String = "" _
            & " where ARTPYMT2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & YP0 & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP <= '" & YP1 & "'" & vbCrLf _
            & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf

        Dim ACCT_CODE_ARX As String = frmASFBASE0.ROWs("ARTPARM1").Item("AR_PARM_EXCH_ACCT_CODE") & ""

        Dim sqlARTPYMT4where As String = "" _
            & " where ARTPYMT2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & YP0 & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP <= '" & YP1 & "'" & vbCrLf _
            & IIf(ACCT_CODE_ARX = "", "", "   and ARTPYMT4.ACCT_CODE <> '" & ACCT_CODE_ARX & "'" & vbCrLf) _
            & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf

        Dim sqlSOTINVH1where As String = "" _
            & " where SOTINVH1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & YP0 & "'" & vbCrLf _
            & "   and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & YP1 & "'" & vbCrLf _
            & "   and SOTINVH1.REASON_CODE is Not Null" & vbCrLf _
            & "   and (SOTINVH1.ORDR_TYPE_CODE = 'TOP' or SOTINVH1.ORDR_TYPE_CODE = 'DIF')" & vbCrLf _
            & IIf(SALES_DIVISION_CODE <> "", "   and SOTINVH1.SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'" & vbCrLf, "")

        Dim SHP As String = ""
        Dim DED As String = ""
        Dim CRM As String = ""
        Dim GLD As String = ""
        For I As Integer = 1 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(YP0, I - 1)
            SHP &= ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) AMT" & Format(I, "00")
            DED &= ", SUM (DECODE(ARTPYMT1.OPS_YYYYPP,'" & YP & "',DECODE(NVL(ARTPYMT5.CHARGEBACK_IND,'0'),'0',ARTPYMT5.GL_DIST_AMT,0),0)) AMT" & Format(I, "00")
            CRM &= ", SUM (DECODE(SOTINVH1.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH1.INV_TOTAL_AMOUNT,0),0)) AMT" & Format(I, "00")
            GLD &= ", SUM (DECODE(ARTPYMT1.OPS_YYYYPP,'" & YP & "',ARTPYMT4.GL_DIST_AMT,0)) AMT" & Format(I, "00")
        Next


        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, 10 LINE, X.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ICTSTYL1, (Select SOTINVH2.CUST_CODE, SOTINVH2.STYLE_CODE" & vbCrLf _
            & SHP & vbCrLf _
            & " from SOTINVH2,SOTINVH1" & vbCrLf _
            & sqlSOTINVH2where _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH2.ORDR_UNIT_PRICE > 0" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.STYLE_CODE) X" & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE (+) = X.STYLE_CODE"
        ASCDATA1.ExecuteSQL()

        'Dim LINE As String = "DECODE(SOTINVH1.ORDR_TYPE_CODE,'DIF',21,20)"
        'ASCMAIN1.sql = "Insert into " & ARTCSUMC _
        '    & " Select X.LINE, X.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
        '    & Xs & vbCrLf _
        '    & " from ICTSTYL1, (Select " & LINE & " LINE, SOTINVH2.STYLE_CODE" & vbCrLf _
        '    & SHP & vbCrLf _
        '    & " from SOTINVH2,SOTINVH1" & vbCrLf _
        '    & sqlSOTINVH2where _
        '    & "   and SOTINVH2.INV_TYPE = 'C'" & vbCrLf _
        '    & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
        '    & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
        '    & "   and SOTINVH2.ORDR_UNIT_PRICE > 0" & vbCrLf _
        '    & " group by " & LINE & ",SOTINVH2.STYLE_CODE) X" & vbCrLf _
        '    & " where ICTSTYL1.STYLE_CODE (+) = X.STYLE_CODE"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, 20 LINE, X.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ICTSTYL1, (Select SOTINVH2.CUST_CODE, SOTINVH2.STYLE_CODE" & vbCrLf _
            & SHP & vbCrLf _
            & " from SOTINVH2,SOTINVH1" & vbCrLf _
            & sqlSOTINVH2where _
            & "   and SOTINVH2.INV_TYPE = 'C'" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.ORDR_UNIT_PRICE > 0" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.STYLE_CODE) X" & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE (+) = X.STYLE_CODE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE" & vbCrLf _
            & ", 40 LINE" & vbCrLf _
            & ", X.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ICTSTYL1" & vbCrLf _
            & ", (Select SOTINVH2.CUST_CODE, SOTINVH2.STYLE_CODE" & vbCrLf _
            & Replace(SHP, "ORDR_UNIT_PRICE", "ORDR_UNIT_COST") & vbCrLf _
            & " from SOTINVH2,SOTINVH1" & vbCrLf _
            & sqlSOTINVH2where _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.STYLE_CODE) X" & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE (+) = X.STYLE_CODE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, 80 LINE, X.REASON_CODE, ARTREAS1.REASON_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ARTREAS1, (Select ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE" & vbCrLf _
            & DED & vbCrLf _
            & " from ARTPYMT1,ARTPYMT2,ARTPYMT5" & vbCrLf _
            & sqlARTPYMT5where _
            & " group by ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE) X" & vbCrLf _
            & " where ARTREAS1.REASON_CODE (+) = X.REASON_CODE"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, DECODE(X.INV_TYPE,'I',71,70) LINE, X.REASON_CODE, ARTREAS1.REASON_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ARTREAS1, (Select SOTINVH1.CUST_CODE, SOTINVH1.INV_TYPE, SOTINVH1.REASON_CODE" & vbCrLf _
            & CRM & vbCrLf _
            & " from SOTINVH1" & vbCrLf _
            & sqlSOTINVH1where _
            & " group by SOTINVH1.CUST_CODE, SOTINVH1.INV_TYPE, SOTINVH1.REASON_CODE) X" & vbCrLf _
            & " where ARTREAS1.REASON_CODE (+) = X.REASON_CODE"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, 81 LINE, X.ACCT_CODE, GLTACCT1.ACCT_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from GLTACCT1, (Select ARTPYMT2.CUST_CODE, ARTPYMT4.ACCT_CODE" & vbCrLf _
            & GLD & vbCrLf _
            & " from ARTPYMT1,ARTPYMT2,ARTPYMT4" & vbCrLf _
            & sqlARTPYMT4where _
            & " group by ARTPYMT2.CUST_CODE, ARTPYMT4.ACCT_CODE) X" & vbCrLf _
            & " where GLTACCT1.ACCT_CODE (+) = X.ACCT_CODE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from " & ARTCSUMC & " where CUST_CODE = '" & CUST_CODE & "'"
        frmASFBASE0.Fill_Records("ARTCSUMB", "", clear_ARTCSUMB, ASCMAIN1.sql)

        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 30, 10, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 30, 20, 1)

        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 50, 30, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 50, 40, -1)

        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 50, 1)
        'Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 60, -1)
        'Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 61, -1)
        'Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 64, -1)
        ' intentionally leaving out PPP 65 because we get re-imbursed for it = 
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 70, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 71, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 80, -1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 81, -1)
    End Sub

    Public Shared Sub Build_ARTCSUMA_Totals(
        frmASFBASE0 As ASFBASE0,
        CUST_CODE As String,
        LINE As Integer, LINE_to_add As Integer, S As Integer)

        Dim rowARTCSUMA As DataRow = frmASFBASE0.dst.Tables("ARTCSUMA").Rows.Find(New Object() {CUST_CODE, LINE_to_add})
        Dim rowARTCSUMB As DataRow = frmASFBASE0.dst.Tables("ARTCSUMB").NewRow
        rowARTCSUMB.Item("CUST_CODE") = CUST_CODE
        rowARTCSUMB.Item("LINE") = LINE
        rowARTCSUMB.Item("CODE_VALUE") = rowARTCSUMA.Item("LINE_ABBR")
        rowARTCSUMB.Item("DESC_VALUE") = rowARTCSUMA.Item("LINE_DESC")

        For I As Integer = 1 To 12
            Dim C As String = "AMT" & Format(I, "00")
            rowARTCSUMB.Item(C) = Val(rowARTCSUMA.Item(C) & "") * S
        Next
        frmASFBASE0.dst.Tables("ARTCSUMB").Rows.Add(rowARTCSUMB)
    End Sub

    Public Shared Sub Create_ARTCSUMA(
        frmASFBASE0 As ASFBASE0,
        CUST_CODE As String)

        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 10, "Gross Shipments", "GRS")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 20, "Returns", "RTN")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 30, "Net Sales", "NET")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 40, "Cost of Goods", "CGS")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 50, "$GP on Net Sales", "GP")
        'Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 60, "Testers/Samples/Misc", "ST")
        'Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 61, "Gift w/Purchase", "GWP")
        'Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 64, "Displays", "DSP")
        ' Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 65, "Pre-Paid Promo", "PP")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 70, "Misc Credits", "CR")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 71, "Misc Charges", "DR")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 80, "Deductions", "DED")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 81, "GL Write-Offs", "GL")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 99, "Net Profit", "NP")
    End Sub

    Public Shared Sub Add_ARTCSUMA(
        frmASFBASE0 As ASFBASE0,
        CUST_CODE As String,
        LINE As Integer, LINE_DESC As String, LINE_ABBR As String)
        Dim rowARTCSUMA As DataRow = frmASFBASE0.dst.Tables("ARTCSUMA").NewRow
        rowARTCSUMA.Item("CUST_CODE") = CUST_CODE
        rowARTCSUMA.Item("LINE") = LINE
        rowARTCSUMA.Item("LINE_DESC") = LINE_DESC
        rowARTCSUMA.Item("LINE_ABBR") = LINE_ABBR
        frmASFBASE0.dst.Tables("ARTCSUMA").Rows.Add(rowARTCSUMA)
    End Sub

    Public Shared Function Create_ARTCSUMC(frmASFBASE0 As ASFBASE0) As String

        Dim ARTCSUMC As String = ""

        With frmASFBASE0.dst
            If .Relations.Contains("ARTCSUMA_ARTCSUMB") Then
                For I As Integer = 0 To 13
                    Dim C As String = "AMT" & Format(I, "00")
                    .Tables("ARTCSUMA").Columns.Remove(C)
                Next
                .Relations.Remove("ARTCSUMA_ARTCSUMB")
            End If

            ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, 0 LINE, ARTCUST1.CUST_NAME CODE_VALUE, ARTCUST1.CUST_NAME DESC_VALUE"
            For I As Integer = 1 To 12
                ASCMAIN1.sql &= ",0.01 AMT" & Format(I, "00")
            Next
            ASCMAIN1.sql &= " from ARTCUST1 where ROWNUM < 1"
            ARTCSUMC = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Alter Table " & ARTCSUMC & " Modify DESC_VALUE VARCHAR2(300)"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select CUST_CODE, LINE"
            For I As Integer = 1 To 12
                ASCMAIN1.sql &= ", AMT" & Format(I, "00")
            Next
            ASCMAIN1.sql &= " from " & ARTCSUMC & " ARTCSUMC where CUST_CODE = :PARM1 and LINE = :PARM2"
            frmASFBASE0.Create_TDA(.Tables.Add, "ARTCSUMA", "**", 0, False, "VN", 0)
            With .Tables("ARTCSUMA")
                .Columns.Add("LINE_DESC")
                .Columns.Add("LINE_ABBR")
                .Columns.Add("AMT00", GetType(System.Decimal))
                .Columns.Add("AMT13", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("CUST_CODE"), .Columns("LINE")}
            End With

            ASCMAIN1.sql = "Select CUST_CODE, LINE, CODE_VALUE, DESC_VALUE"
            Dim T As String = ""
            For I As Integer = 1 To 12
                ASCMAIN1.sql &= ", AMT" & Format(I, "00")
                T &= "+ISNULL(AMT" & Format(I, "00") & ",0)"
            Next
            ASCMAIN1.sql &= " from " & ARTCSUMC & " ARTCSUMC where CUST_CODE = :PARM1 and LINE = :PARM2"
            frmASFBASE0.Create_TDA(.Tables.Add, "ARTCSUMB", "**", 0, False, "VN", 0)
            .Tables("ARTCSUMB").Columns.Add("AMT00", GetType(System.Decimal), Mid(T, 2))
            .Tables("ARTCSUMB").Columns.Add("AMT13", GetType(System.Decimal), Mid(T, 2))

            frmASFBASE0.Create_Relation("ARTCSUMA", "ARTCSUMB", "CUST_CODE,LINE")
            For I As Integer = 0 To 13
                Dim C As String = "AMT" & Format(I, "00")
                .Tables("ARTCSUMA").Columns(C).Expression = "SUM(CHILD(ARTCSUMA_ARTCSUMB)." & C & ")"
            Next
        End With

        Return ARTCSUMC
    End Function

    Public Shared Sub DeRelease(SOTPICK1 As String, Optional show_progress As Boolean = True)

        ' Update Pick Ticket, Shipment Control & Carton Tables

        Dim sql_pick_D As String = " and SOTPICK1.PICK_NO in (Select PICK_NO from " & SOTPICK1 & ")"

        If show_progress Then ASCMAIN1.Progress("-", "Status Qtys")

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select SOTORDR1.WHSE_CODE" & vbCrLf _
            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY,0)) QTY" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY_CANC_REL,0)) QTY_CANC" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY_BACK_REL,0)) QTY_BACK" & vbCrLf _
            & " from SOTORDR2,SOTPICK2,SOTPICK1,SOTORDR1 " & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & sql_pick_D & vbCrLf _
            & " group by SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTSTAT2 " & vbCrLf _
            & " Set WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) - R1.QTY, " & vbCrLf _
            & "     WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) + R1.QTY + R1.QTY_CANC" & vbCrLf _
            & " where STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
            & "   and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
            & "   and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "   Insert into ICTSTAT2 (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_PICK, WHSE_QTY_OPEN)" & vbCrLf _
            & "   Values (R1.STYLE_CODE, R1.COLOR_CODE, R1.WHSE_CODE, -1 * R1.QTY, R1.QTY + R1.QTY_CANC);" & vbCrLf _
            & " End If;" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.Progress("-", "Cartons")
        'ASCMAIN1.sql = "Delete FROM SOTCART2" & vbCrLf _
        '    & " where CART_NO in (Select CART_NO from SOTCART1 where PICK_NO in" & vbCrLf _
        '    & " (Select PICK_NO from SOTPICK1 " & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
        '    & "))"
        'ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Delete from SOTCART1 where PICK_NO in" & vbCrLf _
        '    & " (Select PICK_NO from SOTPICK1 " & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
        '    & ")"
        'ASCDATA1.ExecuteSQL()

        If show_progress Then ASCMAIN1.Progress("-", "Order Status")

        ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_STATUS = 'O'" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_STATUS = 'O'" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        If ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "Delete from SOTCONF2" & vbCrLf _
                & " where ORDR_NO in " & vbCrLf _
                & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
                & ")"
            ASCDATA1.ExecuteSQL()
        End If

        If show_progress Then ASCMAIN1.Progress("-", "Tickets")
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select SOTPICK2.* from SOTPICK1,SOTPICK2" & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & sql_pick_D & ";" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update SOTORDR2 " & vbCrLf _
            & " Set ORDR_QTY_PICK = NVL(ORDR_QTY_PICK,0) - NVL(R1.PICK_QTY,0)," & vbCrLf _
            & "     ORDR_QTY_OPEN = NVL(ORDR_QTY_OPEN,0) + NVL(R1.PICK_QTY,0) + NVL(R1.PICK_QTY_CANC_REL,0)," & vbCrLf _
            & "     ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) - NVL(R1.PICK_QTY_CANC_REL,0)" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = R1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ' not sure why this is done so soon - added it redundantly to the bottom of the routine
        ' problem is that the pick tickets are nor marked as deleted yet, so the ordr_cnt_pick is not set to 0
        ' REMMING OUT THIS SECTION IN FAVOR OF DOING IT AT THE END

        'ASCMAIN1.sql = "" _
        '    & "Begin Declare Cursor C1 is" & vbCrLf _
        '    & " Select DISTINCT ORDR_GROUP_NO from SOTSHIP1 " & vbCrLf _
        '    & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from SOTPICK1" & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & ");" & vbCrLf _
        '    & " Begin For R1 in C1 Loop" & vbCrLf _
        '    & " SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
        '    & " End Loop; End; End;"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTPICK1 set " & vbCrLf _
            & " PICK_STATUS = 'D', LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D)
        ASCDATA1.ExecuteSQL()

        If show_progress Then ASCMAIN1.Progress("-", "Shipments")
        ASCMAIN1.sql = "Select SHIP_BOL_NO" & vbCrLf _
            & ", Sum (Decode (PICK_STATUS,'P',1,0)) PICK" & vbCrLf _
            & ", Sum (Decode (PICK_STATUS,'F',1,0)) SHIP" & vbCrLf _
            & ", Count (*) TOTAL" & vbCrLf _
            & " from SOTPICK1 " & vbCrLf _
            & " where SHIP_BOL_NO in " & vbCrLf _
            & " (Select Distinct SHIP_BOL_NO from " & SOTPICK1 & ")" & vbCrLf _
            & " group by SHIP_BOL_NO"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")

            If Val(row.Item("PICK") & "") = 0 Then

                Dim SHIP_STATUS As String = ""

                If Val(row.Item("SHIP") & "") = 0 Then
                    SHIP_STATUS = "D"
                Else
                    SHIP_STATUS = "F" ' SHOULDNT SET F WITHOUT OTHER FIELDS WHICH GET THEIR VALUE VIA DATA ENTRY IN SHIPMENTS CONF
                    Throw New Exception("Uncertain status for Shipment " & SHIP_BOL_NO & " during De-Release - please call ABS")
                    Stop ' MUST RESEARCH HOW THIS IS POSSIBLE, IF IT EVER HAPPENS
                    SHIP_STATUS = ""
                End If

                If SHIP_STATUS <> "" Then
                    ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_STATUS = '" & SHIP_STATUS & "'" _
                        & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
                        & ", LP_STATUS = NULL" _
                        & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                    ASCDATA1.ExecuteSQL()
                End If
            End If
        Next

        'ASCMAIN1.sql = "" _
        '    & "Begin Declare Cursor C1 is" & vbCrLf _
        '    & " Select Distinct ORDR_GROUP_NO from SOTSHIP1 " & vbCrLf _
        '    & " where SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from SOTPICK1" & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & ");" & vbCrLf _
        '    & " Begin For R1 in C1 Loop" & vbCrLf _
        '    & " SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
        '    & " End Loop; End; End;"
        'ASCDATA1.ExecuteSQL()

        ' SHIP_BOL_NOs should be in temp table SOTPICK1

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select Distinct ORDR_GROUP_NO from SOTSHIP1 " & vbCrLf _
            & " where SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTPICK1 & ");" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

    End Sub

    Public Shared Function Create_Carton(
       F As ASFBASE1,
       PKG_CODE As String,
       PKG_CUBE As Decimal,
       CART_NO_seq As Integer,
       CART_SEQ As Integer,
       PICK_NO As String,
       single_carton As Boolean) As Decimal

        Dim ROWSOTPICK1 As DataRow = F.dst.Tables("SOTPICK1").Rows.Find(New Object() {PICK_NO})

        Dim rowSOTCART1 As DataRow = F.dst.Tables("SOTCART1").NewRow
        With rowSOTCART1
            .Item("CART_NO") = "TEMP" & Format(CART_NO_seq, "000000")

            .Item("CART_FREIGHT") = 0
            .Item("PICK_NO") = PICK_NO
            .Item("CART_TOTAL_WGT_ACTUAL") = 0
            .Item("CART_TOTAL_WGT_CALC") = 0
            .Item("CART_SEQ") = CART_SEQ
            .Item("PACKAGING_TYPE") = "31"
            .Item("PKG_L") = 0
            .Item("PKG_W") = 0
            .Item("PKG_H") = 0

            .Item("PKG_CODE") = PKG_CODE
            .Item("PKG_CUBE") = PKG_CUBE
        End With

        F.dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

        Dim PKG_CUBE_PACK_CUM As Decimal = 0
        Dim CART_LNO As Integer = 0
        Dim CART_TOTAL_UNITS_REL As Integer = 0
        Dim iterations As Integer = 0

        For Each rowSOTPICK2 As DataRow In F.dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' and CART_NO Is NULL", "CUBE_REQD DESC")
            Dim STYLE_CODE As String = rowSOTPICK2.Item("STYLE_CODE") & ""
            Dim CUBE_REQD As Decimal = Val(rowSOTPICK2.Item("CUBE_REQD") & "")

            If CUBE_REQD = 0 Then
                Throw New Exception($"Volumetric Cartonization with 0 Cube for Style {STYLE_CODE} in Pick No {PICK_NO}")
            End If

            iterations += 1
            If iterations > 1000 Then
                Throw New Exception($"Volumetric Cartonization over 1000 iterations (Carton Details) for Pick No {PICK_NO}")
            End If

            If PKG_CUBE_PACK_CUM + CUBE_REQD > PKG_CUBE Then
                Exit For
            Else

                Dim rowICTSTYL1 As DataRow = F.dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)

                Dim PICK_QTY As Integer = Val(rowSOTPICK2.Item("PICK_QTY") & "")
                CART_TOTAL_UNITS_REL += PICK_QTY
                CART_LNO += 1
                Dim rowSOTCART2 As DataRow = F.dst.Tables("SOTCART2").NewRow
                With rowSOTCART2
                    .Item("CART_NO") = rowSOTCART1.Item("CART_NO")
                    .Item("CART_LNO") = CART_LNO
                    .Item("ORDR_NO") = rowSOTPICK2.Item("ORDR_NO")
                    .Item("ORDR_LNO") = rowSOTPICK2.Item("ORDR_LNO")
                    .Item("QTY_PACKED") = PICK_QTY
                    .Item("STYLE_CODE") = rowSOTPICK2.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = rowSOTPICK2.Item("COLOR_CODE")
                    .Item("STYLE_PREPACK") = 0
                    .Item("QTY_REL") = PICK_QTY
                    .Item("STYLE_WEIGHT") = rowICTSTYL1.Item("STYLE_WEIGHT")
                End With

                F.dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                PKG_CUBE_PACK_CUM += CUBE_REQD
                rowSOTPICK2.Item("CART_NO") = rowSOTCART1.Item("CART_NO")
            End If
        Next
        rowSOTCART1.Item("PKG_CUBE_PACK") = PKG_CUBE_PACK_CUM
        rowSOTCART1.Item("CART_TOTAL_UNITS_REL") = CART_TOTAL_UNITS_REL
        rowSOTCART1.Item("CART_TOTAL_UNITS") = CART_TOTAL_UNITS_REL

        Return PKG_CUBE_PACK_CUM

    End Function

    Public Shared Sub Create_Cartons_For_PICK_NO(F As ASFBASE1, PICK_NO As String, ByRef CART_NO_seq As Integer, Optional single_carton As Boolean = False)

        ' RUNAWAY LOOP (IF NO SOTCART2S)
        ' WHAT ABOUT  IF 0 CUBE REQD

        Dim numPKGBuffer As Decimal = 0 '  0.05 ' Hard Coded to use only 95% capacity of a carton
        Dim CUBE_REQD_remaining As Decimal = Val(F.dst.Tables("SOTPICK2").Compute("SUM(CUBE_REQD)", $"PICK_NO = '{PICK_NO}'") & "")

        Dim rows() As DataRow = F.dst.Tables("WHTPKGM1").Select("", "INNER_CUBE DESC")
        Dim PKG_CODE_largest As String = rows(0).Item("PKG_CODE")
        Dim INNER_CUBE_largest As Decimal = Val(rows(0).Item("INNER_CUBE"))
        Dim INNER_CUBE_largest_net As Decimal = INNER_CUBE_largest * (1 - numPKGBuffer)

        Dim CART_SEQ As Integer = 0
        Dim iterations As Integer = 0

        Do While CUBE_REQD_remaining > 0
            iterations += 1
            If iterations > 1000 Then
                Throw New Exception($"Volumetric Cartonization over 1000 iterations (Cartons) for Pick No {PICK_NO}")
            End If

            If CUBE_REQD_remaining >= INNER_CUBE_largest_net Then
                CART_NO_seq += 1
                CART_SEQ += 1
                Dim CUBE_ACTUAL As Decimal = TAC.SOCMAIN1.Create_Carton(F, PKG_CODE_largest, INNER_CUBE_largest_net, CART_NO_seq, CART_SEQ, PICK_NO, single_carton)
                CUBE_REQD_remaining -= CUBE_ACTUAL
            Else
                For Each rowWHTPKGM1 As DataRow In F.dst.Tables("WHTPKGM1").Select("", "INNER_CUBE")
                    Dim INNER_CUBE As Decimal = Val(rowWHTPKGM1.Item("INNER_CUBE") & "")
                    Dim INNER_CUBE_net As Decimal = INNER_CUBE * (1 - numPKGBuffer)
                    If INNER_CUBE_net >= CUBE_REQD_remaining Then
                        CART_NO_seq += 1
                        CART_SEQ += 1
                        Dim CUBE_ACTUAL As Decimal = TAC.SOCMAIN1.Create_Carton(F, rowWHTPKGM1.Item("PKG_CODE"), INNER_CUBE_net, CART_NO_seq, CART_SEQ, PICK_NO, single_carton)
                        CUBE_REQD_remaining -= CUBE_ACTUAL
                        Exit For
                    End If
                Next
            End If

            If single_carton Then CUBE_REQD_remaining = 0
        Loop

    End Sub

End Class
