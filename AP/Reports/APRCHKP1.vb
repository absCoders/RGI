Imports nsoftware.InPay
Imports nsoftware.IPWorksEncrypt

Public Class APRCHKP1
    Dim BANK_LAST_CHECK_NO As Int64
    Dim BANK_LAST_CHECK_NO_orig As Int64
    Dim rowGLTBANK1 As DataRow
    Dim rowAPTPYMT1 As DataRow
    Dim APTINVH1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")
    End Sub

    Overrides Sub Clear_Record()
        Load_Drop_Down("BATCH_NO_PYMT")
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim BATCH_NO_PYMT As String = Absx1.cmbFor("BATCH_NO_PYMT").Text

        Prepare_dst(True, BATCH_NO_PYMT)

        Check_if_Empty("APTPYMT1")
    End Sub

    Public Overrides Sub Print_Report()

        Dim BATCH_NO_PYMT As String = rowAPTPYMT1.Item("BATCH_NO_PYMT")
        Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")
        Dim PYMT_METHOD As String = rowAPTPYMT1.Item("PYMT_METHOD")

        If rowGLTBANK1.Item("CHECK_REPORT") & "" <> "" Then
            RPT = rowGLTBANK1.Item("CHECK_REPORT")
        End If
        CR_params.Add("COPY", "")
        Generate_Report(RPT)

        If PYMT_METHOD = "DBAUTH" Then ' If rowGLTBANK1.Item("BANK_PYMT_METHOD") & "" = "DBAUTH" Then
            Send_DBAUTHs()
        End If

        CR_params.Add("PAYMENT_SELECTION", "0")
        Generate_Report("APRPYMT1", "Printed Checks Report", "")

        rowGLTBANK1 = Fill_Record("GLTBANK1", BANK_CODE)
        rowGLTBANK1.Item("BATCH_NO_PYMT") = BATCH_NO_PYMT
        Update_Record_TDA("GLTBANK1")

        Write_Event_Log_Batch("APTINVH1", "Select APTINVH1.VOUCHER_NO, 'Check ' || TRIM(TO_CHAR(TO_NUMBER(SUBSTR(APTINVH1.CHECK_NUM,2)) + " & CStr(BANK_LAST_CHECK_NO_orig) & ",'0000000000')) || ' Printed' from APTINVH1 where APTINVH1.BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")
        Write_Event_Log("GLTBANK1", BANK_CODE, "Checks Printed (" & CStr(dst.Tables("APTPYMT2").Rows.Count) & ")")
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)

        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("BATCH_NO_PYMT").Text = "" Then
                EMsg &= vbCr & "You must pick a Batch"
            Else
                Dim BATCH_NO_PYMT As String = Absx1.cmbFor("BATCH_NO_PYMT").Text
                rowAPTPYMT1 = LookUp("APTPYMT1", BATCH_NO_PYMT)
                rowGLTBANK1 = LookUp("GLTBANK1", rowAPTPYMT1.Item("BANK_CODE"))
                If rowGLTBANK1.Item("BATCH_NO_PYMT") & "" <> "" _
                And rowGLTBANK1.Item("BATCH_NO_PYMT") & "" <> BATCH_NO_PYMT Then
                    EMsg &= vbCr & "Bank " & "" & " is in Process with Batch " & rowGLTBANK1.Item("BATCH_NO_PYMT")
                End If

                ASCMAIN1.sql = "Select VEND_CODE, VEND_NAME from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE from APTPYMT2 where BATCH_NO_PYMT = '" & Absx1.cmbFor("BATCH_NO_PYMT").Text & "' union Select Distinct VEND_CODE from APTPYMT2 where BATCH_NO_PYMT = '" & Absx1.cmbFor("BATCH_NO_PYMT").Text & "') and VEND_ON_HOLD = '1'"
                Dim tblAPTVEND1_hold As DataTable = ASCDATA1.GetDataTable
                If tblAPTVEND1_hold.Rows.Count <> 0 Then
                    For Each row As DataRow In tblAPTVEND1_hold.Rows
                        EMsg &= vbCr & "Vendor on Pymt Hold: " & row.Item("VEND_CODE") & ":" & row.Item("VEND_NAME")
                    Next
                End If

                If ASCMAIN1.CLIENT = "VAN" Then
                    ASCMAIN1.sql = "SELECT * FROM APTCHCK1 WHERE REGISTER_IND = '0' or (OPS_YYYYPP_F is Not Null and REGISTER_IND_F = '0')"
                    If ASCDATA1.GetDataTable.Rows.Count <> 0 Then
                        If MsgBox("There are records in the Check Register from a Prior Check run." _
                                  & vbCrLf & vbCrLf & "Do you want to Continue anyway?",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            EMsg &= vbCr & "Print and Update the Check Register"
                        End If
                    End If
                End If
            End If

        ElseIf eItemKey = "Update" Then
            Dim PYMT_METHOD As String = rowAPTPYMT1.Item("PYMT_METHOD")
            If PYMT_METHOD = "ECHECK" Then
                Dim errorList As New List(Of String)
                Dim clsAPCCHECK = New TAC.APCCHECK
                errorList = clsAPCCHECK.ValidateEntry(dst.Tables("APTINVH1"))

                For Each errorMsg As String In errorList
                    If errorMsg.Length = 0 Then Continue For
                    EMsg &= vbCr & errorMsg
                Next
            End If
        End If
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst
            Create_TDA(.Tables.Add, "APTCHCK1", "*")
            Create_TDA(.Tables.Add, "APTCHCK2", "*")
            Create_TDA(.Tables.Add, "APTVEND5", "*")
            Create_TDA(.Tables.Add, "APTPYMT1", "*", 1)
            Create_TDA(.Tables.Add, "APTPYMT2", "*", 1)

            ASCMAIN1.sql = "Select * from APTINVH1 where ROWNUM < 1"
            APTINVH1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & APTINVH1 & " Add Primary Key (VOUCHER_NO)")

            ASCMAIN1.sql = "Select * from " & APTINVH1
            Create_TDA(.Tables.Add, "APTINVH1", "**", 0, Update_COLUMN_NAMEs:="CHECK_NUM,CHECK_DATE,INV_STATUS,INV_PAYMENTS,INV_DISC_TAKEN,INV_LAST_PMT_DATE,BANK_CODE,INV_BALANCE,BATCH_PYMT,BATCH_DISC")

            ASCMAIN1.sql = "Select * from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE from " & APTINVH1 & " union Select Distinct VEND_CODE_AP from " & APTINVH1 & ")"
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, 1)

            ASCMAIN1.sql = "Select * from APTVEND2 where VEND_CODE in (Select Distinct VEND_CODE from " & APTINVH1 & " union Select Distinct VEND_CODE_AP from " & APTINVH1 & ")"
            Create_TDA(.Tables.Add, "APTVEND2", "**", 0, False, 2)

            Create_TDA(.Tables.Add, "GLTBANK1", "*")
        End With

        If perform_fill Then
            Fill_Records_RPT(New String() {sqlw})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        'sqlw = parms(0)
        Dim BATCH_NO_PYMT As String = parms(0)
        EnforceConstraints(False)

        If BATCH_NO_PYMT <> "" Then
            rowAPTPYMT1 = Fill_Record("APTPYMT1", BATCH_NO_PYMT)
        Else
            ASCMAIN1.sql = "Select NVL(BATCH_NO_PYMT,'000000') BATCH_NO_PYMT, CHECK_DATE, BANK_CODE from APTCHCK1 where BANK_CODE = '" & parms(1) & "' and CHECK_NUM = '" & parms(2) & "'"
            Fill_Records("APTPYMT1", "", True, ASCMAIN1.sql)
            rowAPTPYMT1 = dst.Tables("APTPYMT1").Rows(0)
        End If

        Dim CHECK_DATE As Date = rowAPTPYMT1.Item("CHECK_DATE")
        Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")

        If BATCH_NO_PYMT <> "" Then
            Fill_Records("APTPYMT2", BATCH_NO_PYMT)
        Else
            ASCMAIN1.sql = "Select NVL(BATCH_NO_PYMT,'000000') BATCH_NO_PYMT, CHECK_NUM, VEND_CODE_AP, NULL VOUCHER_NO, CHECK_AMT BATCH_PYMT, 0 BATCH_DISC, VEND_ALT_CODE, VEND_CODE, VEND_NAME from APTCHCK1 where BANK_CODE = '" & parms(1) & "' and CHECK_NUM = '" & parms(2) & "'"
            Fill_Records("APTPYMT2", "", True, ASCMAIN1.sql)
        End If

        rowGLTBANK1 = Fill_Record("GLTBANK1", BANK_CODE)

        ASCDATA1.ExecuteSQL("Delete from " & APTINVH1)

        If BATCH_NO_PYMT <> "" Then
            ASCDATA1.ExecuteSQL("Insert into " & APTINVH1 & " Select APTINVH1.* from APTINVH1 where APTINVH1.BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")
        Else
            Dim sql2 As String = "APTCHCK2.BANK_CODE = '" & parms(1) & "' and APTCHCK2.CHECK_NUM = '" & parms(2) & "'"
            ASCDATA1.ExecuteSQL("Insert into " & APTINVH1 & " Select APTINVH1.* from APTINVH1,APTCHCK2 where APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO and " & sql2)
            ASCMAIN1.sql = "Begin Declare Cursor C1 is Select * from APTCHCK2 where " & sql2 & "; Begin for R1 in C1 Loop Update " & APTINVH1 & " Set BATCH_NO_PYMT = NVL(BATCH_NO_PYMT,'000000'), BATCH_PYMT = NVL(R1.INV_AMT_APPLIED,0) - NVL(R1.INV_DISC_TAKEN,0), BATCH_DISC = R1.INV_DISC_TAKEN where VOUCHER_NO = R1.VOUCHER_NO; End Loop; End; End;"
            ASCDATA1.ExecuteSQL()
        End If
        Fill_Records("APTINVH1")

        Fill_Records("APTVEND1")
        Fill_Records("APTVEND2")

        If BATCH_NO_PYMT <> "" Then
            Generate_Checks(BANK_CODE, CHECK_DATE)
        End If

        EnforceConstraints(True)
    End Sub

    Sub Generate_Checks(BANK_CODE As String, CHECK_DATE As Date)
        sql = "Select FIELD_LENGTH from ASTFFMT1 where COLUMN_NAME = 'CHECK_NUM'"
        Dim CHECK_NUM_length As Integer = Val(ASCDATA1.GetDataValue(sql))
        If CHECK_NUM_length = 0 Then
            CHECK_NUM_length = 10
        End If

        BANK_LAST_CHECK_NO = Val(rowGLTBANK1.Item("BANK_LAST_CHECK_NO") & "")
        BANK_LAST_CHECK_NO_orig = BANK_LAST_CHECK_NO
        Dim CHECK_NUM As String = ""

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Select("", "CHECK_NUM")
            Do
                BANK_LAST_CHECK_NO = BANK_LAST_CHECK_NO + 1
                CHECK_NUM = Format(BANK_LAST_CHECK_NO, "".PadLeft(CHECK_NUM_length, "0"))
                LookUp("APTCHCK1", New String() {BANK_CODE, CHECK_NUM})
            Loop While cdr IsNot Nothing
            Dim CHECK_NUM_X As String = rowAPTPYMT2.Item("CHECK_NUM")
            rowAPTPYMT2.Item("CHECK_NUM") = CHECK_NUM

            For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("CHECK_NUM = '" & CHECK_NUM_X & "'", "")
                rowAPTINVH1.Item("CHECK_NUM") = CHECK_NUM
            Next
        Next
    End Sub

    Overrides Sub Update_Record()

        Dim SEQ_NUM As Integer = 0
        Dim CHECK_NUM As String = ""

        Dim CHECK_DATE As Date = rowAPTPYMT1.Item("CHECK_DATE")
        Dim BANK_CODE As String = rowAPTPYMT1.Item("BANK_CODE")
        Dim BATCH_NO_PYMT As String = rowAPTPYMT1.Item("BATCH_NO_PYMT")
        Dim PYMT_METHOD As String = rowAPTPYMT1.Item("PYMT_METHOD")

        Dim SSH_APP_CODE As String = rowGLTBANK1.Item("SSH_APP_CODE") & ""
        Dim BANK_PYMT_METHOD As String = "CHECK"
        If rowGLTBANK1.Item("BANK_PYMT_METHOD") & "" <> "" Then
            BANK_PYMT_METHOD = rowGLTBANK1.Item("BANK_PYMT_METHOD")
        End If

        ASCMAIN1.Progress("Now Processing AP Items")

        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("", "CHECK_NUM, VOUCHER_NO")

            Dim INV_AMT_APPLIED As Double = Val(rowAPTINVH1.Item("BATCH_PYMT") & "") + Val(rowAPTINVH1.Item("BATCH_DISC") & "")
            Dim INV_DISC_TAKEN As Double = Val(rowAPTINVH1.Item("BATCH_DISC") & "")

            rowAPTINVH1.Item("CHECK_DATE") = CHECK_DATE
            rowAPTINVH1.Item("INV_STATUS") = "P"
            rowAPTINVH1.Item("INV_PAYMENTS") = Val(rowAPTINVH1.Item("INV_PAYMENTS") & "") + Val(rowAPTINVH1.Item("BATCH_PYMT") & "")
            rowAPTINVH1.Item("INV_DISC_TAKEN") = Val(rowAPTINVH1.Item("INV_DISC_TAKEN") & "") + Val(rowAPTINVH1.Item("BATCH_DISC") & "")
            rowAPTINVH1.Item("INV_LAST_PMT_DATE") = CHECK_DATE
            rowAPTINVH1.Item("BANK_CODE") = BANK_CODE
            rowAPTINVH1.Item("INV_BALANCE") = Val(rowAPTINVH1.Item("INV_BALANCE") & "") - INV_AMT_APPLIED
            rowAPTINVH1.Item("BATCH_PYMT") = 0
            rowAPTINVH1.Item("BATCH_DISC") = 0

            If CHECK_NUM <> rowAPTINVH1.Item("CHECK_NUM") Then
                SEQ_NUM = 0
                CHECK_NUM = rowAPTINVH1.Item("CHECK_NUM")
            End If

            Dim rowAPTCHCK2 As DataRow = dst.Tables("APTCHCK2").NewRow
            rowAPTCHCK2.Item("BANK_CODE") = BANK_CODE
            rowAPTCHCK2.Item("CHECK_NUM") = CHECK_NUM
            SEQ_NUM = SEQ_NUM + 1
            rowAPTCHCK2.Item("SEQ_NUM") = SEQ_NUM
            rowAPTCHCK2.Item("VEND_CODE") = rowAPTINVH1.Item("VEND_CODE")
            rowAPTCHCK2.Item("INV_NUM") = rowAPTINVH1.Item("INV_NUM")
            rowAPTCHCK2.Item("INV_DATE") = rowAPTINVH1.Item("INV_DATE")
            rowAPTCHCK2.Item("VOUCHER_NO") = rowAPTINVH1.Item("VOUCHER_NO")
            rowAPTCHCK2.Item("INV_AMT_APPLIED") = INV_AMT_APPLIED
            rowAPTCHCK2.Item("INV_DISC_TAKEN") = INV_DISC_TAKEN
            dst.Tables("APTCHCK2").Rows.Add(rowAPTCHCK2)
        Next
        Update_Record_TDA("APTCHCK2")
        Update_Record_TDA("APTINVH1")

        ASCMAIN1.Progress("Now Processing Checks")

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Rows
            Call ASCMAIN1.Progress("", rowAPTPYMT2.Item("CHECK_NUM"))

            Dim rowAPTVEND5 = Fill_Record("APTVEND5", rowAPTPYMT2.Item("VEND_CODE"), True, False)
            rowAPTVEND5.ITEM("VEND_PAYMENTS_MTD") = Val(rowAPTVEND5.ITEM("VEND_PAYMENTS_MTD") & "") + Val(rowAPTPYMT2.Item("BATCH_PYMT") & "")
            rowAPTVEND5.ITEM("VEND_PAYMENTS_YTD") = Val(rowAPTVEND5.ITEM("VEND_PAYMENTS_YTD") & "") + Val(rowAPTPYMT2.Item("BATCH_PYMT") & "")
            rowAPTVEND5.ITEM("VEND_DISC_TAKEN_MTD") = Val(rowAPTVEND5.ITEM("VEND_DISC_TAKEN_MTD") & "") + Val(rowAPTPYMT2.Item("BATCH_DISC") & "")
            rowAPTVEND5.ITEM("VEND_DISC_TAKEN_YTD") = Val(rowAPTVEND5.ITEM("VEND_DISC_TAKEN_YTD") & "") + Val(rowAPTPYMT2.Item("BATCH_DISC") & "")
            rowAPTVEND5.ITEM("VEND_NUM_CHKS_MTD") = Val(rowAPTVEND5.ITEM("VEND_NUM_CHKS_MTD") & "") + 1
            rowAPTVEND5.ITEM("VEND_NUM_CHKS_YTD") = Val(rowAPTVEND5.ITEM("VEND_NUM_CHKS_YTD") & "") + 1
            rowAPTVEND5.ITEM("VEND_LAST_PMT_DATE") = CHECK_DATE
            rowAPTVEND5.ITEM("VEND_LAST_PMT_AMT") = Val(rowAPTPYMT2.Item("BATCH_PYMT") & "")

            Dim rowAPTCHCK1 As DataRow = dst.Tables("APTCHCK1").NewRow
            rowAPTCHCK1.Item("BANK_CODE") = BANK_CODE
            rowAPTCHCK1.Item("CHECK_NUM") = rowAPTPYMT2.Item("CHECK_NUM")
            rowAPTCHCK1.Item("CHECK_DATE") = CHECK_DATE
            rowAPTCHCK1.Item("CHECK_AMT") = rowAPTPYMT2.Item("BATCH_PYMT")
            rowAPTCHCK1.Item("PYMT_METHOD") = PYMT_METHOD
            rowAPTCHCK1.Item("VEND_CODE") = rowAPTPYMT2.Item("VEND_CODE")
            rowAPTCHCK1.Item("VEND_CODE_AP") = rowAPTPYMT2.Item("VEND_CODE_AP")
            rowAPTCHCK1.Item("VEND_ALT_CODE") = rowAPTPYMT2.Item("VEND_ALT_CODE")
            rowAPTCHCK1.Item("VEND_NAME") = rowAPTPYMT2.Item("VEND_NAME")
            rowAPTCHCK1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowAPTCHCK1.Item("CHECK_STATUS") = "I"
            rowAPTCHCK1.Item("REGISTER_IND") = "0"
            rowAPTCHCK1.Item("BATCH_NO_PYMT") = rowAPTPYMT2.Item("BATCH_NO_PYMT")
            rowAPTCHCK1.Item("INIT_DATE") = DATETIME_STAMP
            rowAPTCHCK1.Item("INIT_OPER") = ASCMAIN1.USER_ID

            If rowGLTBANK1.Item("BANK_PP_IND") & "" = "1" Then
                If SSH_APP_CODE = "" Then
                    Throw New Exception("No SSH record to use for PP Transmit")
                End If
                rowAPTCHCK1.Item("POS_PAY_STATUS_IND") = "P" ' Pending Transmission
            End If

            If PYMT_METHOD = "ACH" Then
                ' note - we may have to move this update to APRCHKR1 to send Voids as well as Issued Payments
                ' or - we may need to add code to the void check routine to set this flag on a void where applicable
                If SSH_APP_CODE = "" Then
                    Throw New Exception("No SSH record to use for ACH Transmit")
                End If
                rowAPTCHCK1.Item("ACH_PAY_STATUS_IND") = "P" ' Pending Transmission
            End If

            dst.Tables("APTCHCK1").Rows.Add(rowAPTCHCK1)
        Next
        Update_Record_TDA("APTCHCK1")
        Update_Record_TDA("APTVEND5")

        Write_Event_Log_Batch("APTINVH1", "Select APTCHCK2.VOUCHER_NO, 'Check ' || APTCHCK2.CHECK_NUM || ' Updated' from APTCHCK2,APTCHCK1 where APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM and APTCHCK1.BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")

        ASCDATA1.ExecuteSQL("Delete from APTPYMT1 where BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")
        ASCDATA1.ExecuteSQL("Delete from APTPYMT2 where BATCH_NO_PYMT = '" & BATCH_NO_PYMT & "'")

        rowGLTBANK1 = Fill_Record("GLTBANK1", BANK_CODE)
        If Val(rowGLTBANK1.Item("BANK_LAST_CHECK_NO") & "") <> BANK_LAST_CHECK_NO_orig Then Stop ' NEED TO ROLLBACK
        rowGLTBANK1.Item("BANK_LAST_CHECK_NO") = BANK_LAST_CHECK_NO
        rowGLTBANK1.Item("BATCH_NO_PYMT") = ""
        Update_Record_TDA("GLTBANK1")

        If PYMT_METHOD = "ECHECK" Then
            Dim clsAPCCHECK = New TAC.APCCHECK
            For Each rowAPTCHCK1 As DataRow In dst.Tables("APTCHCK1").Select("CHECK_AMT <> 0", "VEND_CODE")
                If clsAPCCHECK.Send_eChecks(rowAPTCHCK1, TAC.APCCHECK.eCheckTypes.Authorize) Then
                    ' The process worked

                Else
                    ' The process failed

                End If
            Next

            ' NEED TO DISPLAY PYMT METHOD AFTER A BATCH IS SELECTED FOR CHECK PRINTING
            ' NEED TO PROVIDE A WARNING TO THE USER THAT THE ECHECK IS GOING TO BE POSTED TO THE BANK UPON UPDATE
            ' ALTERNATE IS TO CREATE ANOTHER SCREEN FOR ECHECKS MODELED AFTER ACH/PP
            ' HOW DO WE CONTROL THIS - DO WE DO IT AFTER THE COMMIT?
            ' IS THERE PROTECTION AGAINST DUP
        End If

    End Sub

    Sub Send_DBAUTHs()
        Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", ASCMAIN1.USER_ID)
        Dim USER_SIGNATURE As String =
          rowASTUSER1.Item("USER_NAME") & vbCrLf _
        & rowASTUSER1.Item("USER_TITLE") & vbCrLf _
        & rowASTUSER1.Item("USER_COMPANY") & vbCrLf _
        & "Tel: " & rowASTUSER1.Item("USER_TELEPHONE") & vbCrLf _
        & "Fax: " & rowASTUSER1.Item("USER_FAX") & vbCrLf _
        & rowASTUSER1.Item("USER_EMAIL") & vbCrLf

        Dim VEND_CODE As String
        Dim VEND_CODE_AP As String
        Dim VEND_ALT_CODE As String

        Dim VEND_EMAIL As String
        Dim VEND_CONTACT As String
        Dim VEND_PHONE As String

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Rows
            ASCMAIN1.Progress("Now Generating emails")

            VEND_CODE = rowAPTPYMT2.Item("VEND_CODE")
            VEND_CODE_AP = rowAPTPYMT2.Item("VEND_CODE_AP")
            VEND_ALT_CODE = rowAPTPYMT2.Item("VEND_ALT_CODE") & ""

            If VEND_CODE <> VEND_CODE_AP Then
                LookUp("APTVEND1", VEND_CODE_AP)
                VEND_EMAIL = cdr.Item("VEND_EMAIL") & ""
                VEND_CONTACT = cdr.Item("VEND_CONTACT") & ""
                VEND_PHONE = cdr.Item("VEND_PHONE") & ""
            Else
                If VEND_ALT_CODE <> "VENDOR" And VEND_ALT_CODE <> "" Then
                    LookUp("APTVEND2", New String() {VEND_CODE_AP, VEND_ALT_CODE})
                    VEND_EMAIL = cdr.Item("VEND_ALT_EMAIL") & ""
                    VEND_CONTACT = cdr.Item("VEND_ALT_CONTACT") & ""
                    VEND_PHONE = cdr.Item("VEND_ALT_PHONE") & ""
                Else
                    LookUp("APTVEND1", VEND_CODE_AP)
                    VEND_EMAIL = cdr.Item("VEND_EMAIL") & ""
                    VEND_CONTACT = cdr.Item("VEND_CONTACT") & ""
                    VEND_PHONE = cdr.Item("VEND_PHONE") & ""
                End If
            End If
            Dim RecordSelectionFormula As String
            RecordSelectionFormula = "{APTINVH1.CHECK_NUM} = '" & rowAPTPYMT2.Item("CHECK_NUM") & "'"
            CR_params.Add("COPY", "1")
            Generate_Report(RPT, , , RecordSelectionFormula)

            Dim MailCCList As String = rowGLTBANK1.Item("ACCT_EMAIL") & ""
            If rowGLTBANK1.Item("BANK_EMAIL") & "" <> "" Then
                If MailCCList <> "" Then
                    MailCCList &= ";"
                End If
                MailCCList &= rowGLTBANK1.Item("BANK_EMAIL") & ""
            End If

            If VEND_EMAIL <> "" Then
                F.Mail_PDF_Report(
                "Payment Instructions Attached",
                VEND_EMAIL,
                VEND_CONTACT & vbCrLf & "Please open the attached file for Payment Instructions." & vbCrLf & vbCrLf & "Thank You," & USER_SIGNATURE,
                MailCCList,
                rowAPTPYMT2.Item("CHECK_NUM"),
                RecordSelectionFormula)
            End If
        Next
    End Sub

    Private Sub cmbBATCH_NO_PYMT_ValueChanged(sender As Object, e As EventArgs) Handles cmbBATCH_NO_PYMT.ValueChanged

        txtBankCode.Clear()
        txtPaymentMethod.Clear()

        Dim BATCH_NO_PYMT As String = cmbBATCH_NO_PYMT.Text
        Dim rowAPTPYMT1 As DataRow = LookUp("APTPYMT1", BATCH_NO_PYMT)
        If rowAPTPYMT1 IsNot Nothing Then
            txtBankCode.Text = rowAPTPYMT1.Item("BANK_CODE") & String.Empty
            txtPaymentMethod.Text = rowAPTPYMT1.Item("PYMT_METHOD") & String.Empty
        End If

    End Sub

End Class