Imports System.Math

Public Class APR1099E
    Dim Report_Subt As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("APTPARM1")

        ' Range_Events(grpCHK_DATE_RANGE)
        optSHOW.Value = "W"
        Absx1.txtFor("TIN").Text = ROWs("APTPARM1").Item("AP_PARM_1099_TAX_ID") & ""
        Absx1.txtFor("TCC").Text = "" ' "19K37"
        Absx1.numFor("CUTOFF").Value = Val(ROWs("APTPARM1").Item("AP_PARM_1099_LIMIT") & "")

        Dim YYYY = Now.Date.AddMonths(-6).Year
        Absx1.dteFor("CHK_DATE_F").Value = "01/01/" & YYYY
        Absx1.dteFor("CHK_DATE_L").Value = "12/31/" & YYYY
    End Sub

    Protected Overrides Sub Build_Workfile()
        With dst
            Dim SQLX As String = ""
            SQLX = " FROM APTCHCK1, APTCHCK2, APTINVH1, APTVEND1" _
            & "  WHERE APTCHCK1.CHECK_STATUS = 'I'" _
            & " AND APTCHCK1.CHECK_DATE >= '" & Format(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") & "'" _
            & " AND APTCHCK1.CHECK_DATE <= '" & Format(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") & "'" _
            & IIf(optSHOW.Value = "O", " AND APTVEND1.VEND_TAX_ID IS NULL", "") _
            & IIf(optSHOW.Value = "W", " AND APTVEND1.VEND_TAX_ID IS NOT NULL", "") _
            & IIf(chkDTL.Checked = False, "    AND APTINVH1.INV_1099_AMT <> 0", "") _
            & "    AND APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" _
            & "    AND APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" _
            & "    AND APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" _
            & "    AND APTVEND1.VEND_CODE = APTCHCK1.VEND_CODE_AP"

            ' DGJ'

            If ASCMAIN1.CLIENT = "VAN" Then
                SQLX = " FROM APTCHCK1, APTCHCK2, APTINVH1, APTVEND1, ANNA_1099" _
                & "  WHERE APTCHCK1.CHECK_STATUS = 'I'" _
                & " AND APTCHCK1.CHECK_DATE >= '" & Format(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") & "'" _
                & " AND APTCHCK1.CHECK_DATE <= '" & Format(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") & "'" _
                & "    AND APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" _
                & "    AND APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" _
                & "    AND APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" _
                & "    AND APTVEND1.VEND_CODE = APTCHCK1.VEND_CODE_AP" _
                & "    AND ANNA_1099.VEND_CODE = APTCHCK2.VEND_CODE" _
                & "    AND ANNA_1099.CHECK_NUM = APTCHCK2.CHECK_NUM"
            End If



            ASCMAIN1.Progress("Compiling Check data", "")

            ASCMAIN1.sql = "SELECT DISTINCT APTCHCK1.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTCHCK1", 2))

            ASCMAIN1.sql = "SELECT DISTINCT APTCHCK2.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTCHCK2", 3))

            ASCMAIN1.Progress("Evaluating Invoice data", "")
            ASCMAIN1.sql = "SELECT APTINVH1.*, DECODE(APTINVH1.INV_AMT,0,0, " _
            & " APTINVH1.INV_1099_AMT * APTCHCK2.INV_AMT_APPLIED / APTINVH1.INV_AMT) PMT_1099 " & SQLX

            ''' DGJ'
            ''ASCMAIN1.Progress("Evaluating Invoice data", "")
            ''ASCMAIN1.sql = "SELECT APTINVH1.*, DECODE(APTINVH1.INV_AMT,0,0, " _
            ''& " APTINVH1.INV_AMT * APTCHCK2.INV_AMT_APPLIED / APTINVH1.INV_AMT) PMT_1099 " & SQLX

            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTINVH1", 1))

            ASCMAIN1.Progress("Evaluating Vendor data", "")
            ASCMAIN1.sql = "SELECT DISTINCT APTVEND1.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTVEND1", 1))

            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql = "Select APTVEND1.VEND_CODE,APTVEND1.VEND_NAME,APTVEND1.VEND_NAME,APTVEND1.VEND_ADDR1,APTVEND1.VEND_ADDR2,APTVEND1.VEND_CITY,APTVEND1.VEND_STATE," _
                & " APTVEND1.VEND_ZIP_CODE, SUM(TO_NUMBER(ANNA_1099.CHECK_AMOUNT)) AMT_ANNA_1099, SUM(APTCHCK1.CHECK_AMT) CHECK_AMT_AP FROM  ANNA_1099, APTCHCK1, APTVEND1 WHERE  ANNA_1099.CHECK_NUM Is Not NULL" _
                & " And ANNA_1099.CHECK_STATUS = 'I'" _
                & " And APTCHCK1.VEND_CODE = ANNA_1099.VEND_CODE" _
                & " And APTCHCK1.CHECK_NUM = ANNA_1099.CHECK_NUM" _
                & " And APTCHCK1.CHECK_STATUS = ANNA_1099.CHECK_STATUS" _
                & " And APTVEND1.VEND_CODE = ANNA_1099.VEND_CODE " _
                & " GROUP BY APTVEND1.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.VEND_NAME, APTVEND1.VEND_ADDR1, APTVEND1.VEND_ADDR2, APTVEND1.VEND_CITY, APTVEND1.VEND_STATE," _
                & " APTVEND1.VEND_ZIP_CODE" _
                & " ORDER BY APTVEND1.VEND_CODE"

                .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTFIRE1", 1))

            End If



            ASCMAIN1.Progress("Merging Check, Invoice and Vendor data", "")
            ASCMAIN1.sql = " SELECT APTCHCK1.VEND_CODE_AP, " _
            & " SUM(DECODE(APTINVH1.INV_AMT,0,0,APTINVH1.INV_1099_AMT * APTCHCK2.INV_AMT_APPLIED / APTINVH1.INV_AMT)) AS PMT_1099, " _
            & " '1' AS PRINT_IND" & SQLX _
            & " GROUP BY APTCHCK1.VEND_CODE_AP"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APT1099V", 1))
            .Tables("APT1099V").Columns("PRINT_IND").ReadOnly = False

            Dim CUTOFF As Decimal = Val(Absx1.numFor("CUTOFF").Value & "")
            For Each rowAPT1900V As DataRow In dst.Tables("APT1099V").Select("PMT_1099 < " & CStr(CUTOFF))
                rowAPT1900V.Item("PRINT_IND") = "0"
            Next

            ASCMAIN1.sql = "Select * from APTPARM1 Where AP_PARM_KEY = 'Z'"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTPARM1", 1))
        End With

        Check_if_Empty("APTCHCK1")

    End Sub

    Public Overrides Sub Print_Report()
        Report_Subt = "1099 Details for Payments Made from " & Format$(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") _
        & " to " & Format$(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") _
        & IIf(Absx1.numFor("CUTOFF").Value > 0, " over " & Format$(Absx1.numFor("CUTOFF").Value, "$##,###.00"), "")
        Generate_Report(RPT, , Report_Subt)

        '' 1099 Form
        Report_Subt = ""
        RPT = "APR1099F"
        ''RPT = "APR1099N"
        ''Dim YEAR As String = "22"
        ''CR_params.Add("YEAR", YEAR)
        Generate_Report(RPT, "1099 Form", Report_Subt)

        '' Payment Review
        Report_Subt = "Summary of Payments Made from " & Format$(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") _
        & " to " & Format$(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") _
        & IIf(Absx1.numFor("CUTOFF").Value > 0, " over " & Format$(Absx1.numFor("CUTOFF").Value, "$##,###.00"), "")
        RPT = "APR1099G"
        Generate_Report(RPT, "Payment Review", Report_Subt)
        ''If chkFIRE.Checked Then
        ''    UPDATE_FIRE_IRS

        ''End If


    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If Absx1.cmbFor("RYP").Text = "" Then
            '    EMsg &= "You Must Select a Period"
            'End If
        End If

    End Sub
    Sub UPDATE_FIRE_IRS()
        Dim YEAR As String = "2023"

        Dim TIN As String = Absx1.txtFor("TIN").Text
        Dim TCC As String = Absx1.txtFor("TCC").Text
        Dim ENTITY_NAME As String = "VANDALE INDUSTRIES"
        Dim TRANSMITTER_NAME As String = "VANDALE INDUSTRIES"         '  "APPLIED BUSINESS SYSTEMS"
        Dim CONTACT As String = "ANNAMARIA SIEGEL"
        Dim CONTACT_PHONE As String = ""
        Dim CONTACT_EMAIL As String = "asiegel@vandale.com"

        If ASCMAIN1.Running_in_VS Then Stop
        Dim FILENAME As String = ASCMAIN1.Folders("Temp") & YEAR & "_IRS_1099_FIE"

        Dim voided_checks As Integer = 0
        Dim TOTAL_VENDORS As Integer = 0
        Dim TOTAL_PAYEES As Integer = 0 ' RIP THROUGH TO GET COUNT OF PAYEES
        Dim REC_COUNT As Integer = 0

        Dim T As String = "".PadLeft(200)

        Using sw As New System.IO.StreamWriter(FILENAME)

            'Sections
            'Transmitter "T" Record : Identifies the Transmitter of electronic file.
            T = "".PadLeft(200)
            Mid(T, 1, 1) = "T"
            Mid(T, 2, 5) = YEAR 'CStr(rowGLTBANK1.Item("BANK_ACCT_ID") & "").PadLeft(20, "0")    ' Bank Account No
            Mid(T, 7, 9) = TIN
            Mid(T, 16, 5) = TCC
            Mid(T, 28, 1) = "T" ' "T" = Test, 
            Mid(T, 30, 40) = TRANSMITTER_NAME
            Mid(T, 110, 40) = ENTITY_NAME
            Mid(T, 190, 40) = "" 'ENTITY ADRESS_1
            Mid(T, 230, 40) = "" 'CITY
            Mid(T, 270, 2) = "" 'STATE
            Mid(T, 272, 9) = "" 'ZIPCODE
            Mid(T, 296, 8) = Format(TOTAL_PAYEES, "00000000")
            Mid(T, 304, 40) = CONTACT
            Mid(T, 344, 15) = CONTACT_PHONE
            Mid(T, 359, 50) = CONTACT_EMAIL
            REC_COUNT += 1
            Mid(T, 500, 8) = Format(REC_COUNT, "00000000")
            Mid(T, 518, 1) = "I" 'Vendor Indicator 'V Softwar purchsed from Vendor ' or 'I Inhouse'
            ' 519- 740 Vendor Information

            sw.Write(T & vbLf)



            'Payer "A" Record : Identifies the Payer (the institution Or person making payments) the type of document being reported, And other miscellaneous information.
            T = "".PadLeft(200)
            Mid(T, 1, 1) = "A"
            Mid(T, 2, 5) = YEAR
            Mid(T, 12, 9) = TIN


            Mid(T, 26, 2) = "NE"

            Mid(T, 28, 1) = "1" 'Amount Codes

            Mid(T, 53, 40) = ENTITY_NAME
            Mid(T, 133, 1) = "0"
            Mid(T, 134, 40) = "" 'ENTITY ADRESS_1
            Mid(T, 174, 40) = "" 'CITY
            Mid(T, 214, 2) = "" 'STATE
            Mid(T, 216, 9) = "" 'ZIPCODE
            Mid(T, 225, 15) = CONTACT_PHONE
            REC_COUNT += 1
            Mid(T, 500, 8) = Format(REC_COUNT, "00000000")

            sw.Write(T & vbLf)

            'Payee "B" Record : Identifies the Payee, the specific payment amounts And information pertinent to the form.

            For Each rowAPTFIRE1 As DataRow In dst.Tables("APTFIRE1").Rows
                T = "".PadLeft(200)
                Mid(T, 1, 1) = "B"
                ''        Mid(T, 3, 20) = CStr(rowGLTBANK1.Item("BANK_ACCT_ID") & "").PadLeft(20, "0")    ' Bank Account No
                ''        Mid(T, 23, 1) = ""      ' Space
                sw.Write(T & vbLf)


                TOTAL_VENDORS = TOTAL_VENDORS + 1
            Next


            'End Of Payer "C" Record : Summary of B records for the payees And money amounts by payer And type of return.
            T = "".PadLeft(200)
            Mid(T, 1, 1) = "C"
            ''        Mid(T, 3, 20) = CStr(rowGLTBANK1.Item("BANK_ACCT_ID") & "").PadLeft(20, "0")    ' Bank Account No
            ''        Mid(T, 23, 1) = TOTAL_VENDORS
            sw.Write(T & vbLf)

            'State Totals "K" Record : Summary of State(s) Totals (for Combined Federal/ State files). Each state will have a separate K record.
            Mid(T, 1, 1) = "K"
            ''        Mid(T, 3, 20) = CStr(rowGLTBANK1.Item("BANK_ACCT_ID") & "").PadLeft(20, "0")    ' Bank Account No
            ''        Mid(T, 23, 1) = TOTAL_VENDORS
            sw.Write(T & vbLf)

            'End Of Transmission "F" Record : End Of Transmission.
            Mid(T, 1, 1) = "F"
            ''        Mid(T, 3, 20) = CStr(rowGLTBANK1.Item("BANK_ACCT_ID") & "").PadLeft(20, "0")    ' Bank Account No
            ''        Mid(T, 23, 1) = TOTAL_VENDORS
            sw.Write(T & vbLf)
        End Using



        '  If ASCMAIN1.DBS_COMPANY <> "EXP" Then Stop

        ''    ASCMAIN1.sql = "Select * from " & APTCHCK1

        ''    For Each row As DataRow In ASCDATA1.GetDataTable.Select("CHECK_AMT > 0", "CHECK_NUM")

        ''        If Val(row.Item("CHECK_AMT") & "") <= 0 Then
        ''            Throw New Exception("Negative Amount: " & Format(Val(row.Item("CHECK_AMT") & ""), "#,##0.00") & " in Check " & row.Item("CHECK_NUM"))
        ''            Stop ' only positive dollar amounts permitted
        ''        End If

        ''        Dim T As String = "".PadLeft(200)
        ''        If row.Item("CHECK_STATUS") = "I" Then
        ''            Mid(T, 1, 1) = "I"      ' I=Issued, V=Voided, S=Stop
        ''        Else
        ''            Mid(T, 1, 1) = "V"      ' I=Issued, V=Voided, S=Stop
        ''            voided_checks += 1
        ''        End If
        ''        Mid(T, 2, 1) = ""       ' Space
        ''        Mid(T, 3, 20) = CStr(rowGLTBANK1.Item("BANK_ACCT_ID") & "").PadLeft(20, "0")    ' Bank Account No
        ''        Mid(T, 23, 1) = ""      ' Space
        ''        Mid(T, 24, 18) = CStr(row.Item("CHECK_NUM") & "").PadLeft(18, "0")              ' Check No
        ''        Mid(T, 42, 1) = ""      ' Space
        ''        Mid(T, 43, 18) = Format(100 * Val(row.Item("CHECK_AMT") & ""), "000000000000000000")  ' positive amounts only, no decimal point
        ''        Mid(T, 61, 1) = ""      ' Space
        ''        Mid(T, 62, 8) = Format(row.Item("CHECK_DATE"), "yyyyMMdd")                      ' Check Date - no slashes
        ''        Mid(T, 70, 1) = ""      ' Space
        ''        Mid(T, 71, 8) = ""      ' Paid Date
        ''        Mid(T, 79, 1) = ""      ' Space
        ''        Mid(T, 80, 15) = ""     ' Additional Information pertaining to Check
        ''        Mid(T, 95, 50) = row.Item("VEND_NAME") & "" ' Expanded Additional Information - Payee Name if desired
        ''        Mid(T, 145, 50) = ""    ' 2nd Payee Name (only if we subscribe to Payee Name Verification Service)
        ''        Mid(T, 195, 6) = ""       ' Spaces

        ''        sw.Write(T & vbLf) ' for unix style
        ''        'sw.WriteLine(T) ' for windows style
        ''    Next

        ''If voided_checks > 0 Then
        ''    MsgBox("There are " & CStr(voided_checks) & " Voided Checks in this batch", MsgBoxStyle.OkOnly, "Verfication")
        ''End If

        ' TRANSMIT IN CODE. NOT POSSIBLE YET
        ''FILENAME_SIGNED = FILENAME & "S"
        '''Sign_File(Me, SSH_APP_CODE, FILENAME)
        ''Sign_File_nSoftware(Me, SSH_APP_CODE, FILENAME)
    End Sub
End Class