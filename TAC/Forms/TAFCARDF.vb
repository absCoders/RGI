Imports Infragistics.Win.UltraWinEditors

Public Class TAFCARDF

    Public CUST_CODE As String = String.Empty
    Public CCPA_NO_CREDIT As String = String.Empty
    Public CCPA_REASON As String = String.Empty
    Public CCPA_NO As String = String.Empty
    Public MerchantTransID As String = String.Empty

    Public rowARTCCPA1 As DataRow = Nothing
    Public rowARTCUSTC As DataRow = Nothing

    Public ShowAgedTotals As Boolean = False
    Public AGE0 As Decimal = 0
    Public AGE1 As Decimal = 0
    Public AGE2 As Decimal = 0
    Public AGE3 As Decimal = 0
    Public AGE4 As Decimal = 0
    Public TOTAL_DUE As Decimal = 0
    Public TOTAL_W_FUTURES As Decimal = 0

    Public TRAN_TYPE As String = String.Empty
    Public AgingDescription As String = String.Empty
    Public AgingHeadings As New Dictionary(Of String, String)

    Public ORDR_NO As String = String.Empty
    Public INV_NO As String = String.Empty
    Public LENS_BANK_INV_NOs As List(Of String)
    Public STMT_NO As String = String.Empty

    Public test_mode As Boolean = False
    Private rowARTCUST1 As DataRow = Nothing
    Private CUST_NAME As String = String.Empty
    Private CCPA_AMT_orig As Decimal = 0

    Private CUST_CREDIT_CARD_NO As String = String.Empty

    Dim COLs() As String = {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_LAST4", "CUST_CREDIT_CARD_KEY" _
    , "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE" _
    , "CUST_CREDIT_CARD_NAME", "CUST_CREDIT_CARD_ADDR1" _
    , "CUST_CREDIT_CARD_CITY", "CUST_CREDIT_CARD_STATE", "CUST_CREDIT_CARD_ZIP_CODE", "CUST_CREDIT_CARD_COUNTRY"}

    Public objCCProcessor As TAC.ARCCCARD = New TAC.ARCCCARD()

    Public responseErrorMessage As String = String.Empty
    Private CC_PROC_FOLDER As String = String.Empty

    Public LockAmountField As Boolean = False
    Public maxAmount As Decimal = 999999

    Public overrideSaleWithCapture As Boolean = False
    Public overrideSaleTransactionID As String = String.Empty
    Public overrideSaleCreditCardFullName As String = String.Empty
    Public overrideSaleApprovalCode As String = String.Empty

    Private clsTACENCRY As TAC.ASCENCRY
    Private tblTATCNTRY As DataTable

    Public Sub New(ByVal FF As ASFBASE1)

        ' Not needed since the form / class is independent
        'frmASFBASE1 = FF
        InitializeComponent()

        SetProcessingType()

        If Not ROWs.ContainsKey("ARTPARM1") Then
            Get_PARM("ARTPARM1")
        End If

        rowARTCCPA1 = ASCDATA1.GetDataTable("SELECT * FROM ARTCCPA1 WHERE ROWNUM < 1").NewRow
        maxAmount = numCCPA_AMT.MaxValue

        tblTATCNTRY = ASCDATA1.GetDataTable("SELECT COUNTRY_CODE2 COUNTRY_CODE, COUNTRY_NAME FROM TATCNTRY ORDER BY COUNTRY_NAME")
        cbeCountry.DataSource = tblTATCNTRY
        cbeCountry.DisplayMember = "COUNTRY_NAME"
        cbeCountry.ValueMember = "COUNTRY_CODE"
        cbeCountry.DropDownListWidth = cbeCountry.Width * 3
    End Sub

    Private Sub TAFCARDF_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If Val(MyBase.Absx1.numFor("CCPA_AMT").Value & String.Empty) = 0 Then
            MyBase.Absx1.numFor("CCPA_AMT").Focus()
            MyBase.Absx1.numFor("CCPA_AMT").SelectAll()
        End If

        lblLeaves.Top = numLeaves.Top + 5
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Create_TDA(dst.Tables.Add, "ARTCCPA1", "*")
        Create_TDA(dst.Tables.Add, "ARTCCPA2", "*")
        Create_TDA(dst.Tables.Add, "ARTCCPDA", "*", 1)
        Create_TDA(dst.Tables.Add, "ARTCUSTC", "*", 1)
        dst.Tables("ARTCUSTC").Columns.Add("EXP", GetType(System.Int16))

        Initialize_DataLayer()

        CUST_CREDIT_CARD_NO = CUST_CREDIT_CARD_NO.Trim

        If CCPA_NO.Length > 0 Then
            Fill_Records("ARTCCPA1", CCPA_NO)
            DecryptARTCCPA1()
            If dst.Tables("ARTCCPA1").Rows.Count = 1 Then
                rowARTCCPA1 = dst.Tables("ARTCCPA1").Rows(0)
                CUST_CODE = rowARTCCPA1.Item("CUST_CODE") & String.Empty
                CUST_CREDIT_CARD_NO = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty
            End If

            Fill_Records("ARTCCPA2", CCPA_NO)
            Fill_Records("ARTCCPDA", CCPA_NO)
        End If

        Fill_Records("ARTCUSTC", CUST_CODE)
        DecryptARTCUSTC()

        For Each row As DataRow In dst.Tables("ARTCUSTC").Select()

            If (row.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty).ToString.Length >= 4 Then
                row.Item("EXP") = Val(row.Item("CUST_CREDIT_CARD_EXP_DATE").ToString.Substring(2, 2) _
                    & row.Item("CUST_CREDIT_CARD_EXP_DATE").ToString.Substring(1, 2))
            Else
                row.Item("EXP") = 0
            End If

            If row.Item("CUST_CREDIT_CARD_PREFERRED") & String.Empty = String.Empty Then
                row.Item("CUST_CREDIT_CARD_PREFERRED") = "0"
            End If
        Next

        Dim rowARTCUSTC As DataRow = Nothing
        If rowARTCCPA1 Is Nothing OrElse rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty <> "" Then
            ' If we pass in credit card data then use it
            CUST_CREDIT_CARD_NO = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty

            ' Update card data
            If dst.Tables("ARTCUSTC").Select("CUST_CREDIT_CARD_NO = '" & CUST_CREDIT_CARD_NO & "'").Length > 0 Then
                rowARTCUSTC = dst.Tables("ARTCUSTC").Select("CUST_CREDIT_CARD_NO = '" & CUST_CREDIT_CARD_NO & "'")(0)
                rowARTCUSTC.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                rowARTCUSTC.Item("LAST_OPER") = ASCMAIN1.USER_ID
            Else
                rowARTCUSTC = dst.Tables("ARTCUSTC").NewRow
                rowARTCUSTC.Item("CUST_CODE") = CUST_CODE
                rowARTCUSTC.Item("CUST_CREDIT_CARD_STATUS") = "A"
                rowARTCUSTC.Item("CUST_CREDIT_CARD_KEY") = ASCMAIN1.Next_Control_No("ARTCUSTC.CUST_CREDIT_CARD_KEY")
                rowARTCUSTC.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                rowARTCUSTC.Item("INIT_OPER") = ASCMAIN1.USER_ID
                dst.Tables("ARTCUSTC").Rows.Add(rowARTCUSTC)
            End If

            ' As per Mario and per Danny - Update the CC Data and Save New Entries
            For Each COLUMN_NAME As String In COLs
                If COLUMN_NAME <> "CUST_CREDIT_CARD_KEY" Then
                    rowARTCUSTC.Item(COLUMN_NAME) = rowARTCCPA1.Item(COLUMN_NAME)
                End If
            Next

            Try
                EncryptARTCUSTC()
                Update_Record_TDA("ARTCUSTC")
                DecryptARTCUSTC()
            Catch ex As Exception

            End Try

        Else
            For Each row As DataRow In dst.Tables("ARTCUSTC").Select("", "CUST_CREDIT_CARD_PREFERRED DESC, EXP DESC")
                CUST_CREDIT_CARD_NO = row.Item("CUST_CREDIT_CARD_NO") & String.Empty
                Exit For
            Next
        End If

        If CUST_CREDIT_CARD_NO.Length > 0 Then
            If dst.Tables("ARTCUSTC").Select("CUST_CREDIT_CARD_NO = '" & CUST_CREDIT_CARD_NO & "'").Length > 0 Then
                rowARTCUSTC = dst.Tables("ARTCUSTC").Select("CUST_CREDIT_CARD_NO = '" & CUST_CREDIT_CARD_NO & "'")(0)
                For Each COLUMN_NAME As String In COLs
                    rowARTCCPA1.Item(COLUMN_NAME) = rowARTCUSTC.Item(COLUMN_NAME)
                Next
                optCC.Value = "X"
                MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO").Text = CUST_CREDIT_CARD_NO
                LoadCCDataIntoRow(rowARTCUSTC)
                Try
                    txt_ValueChanged(MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO"), Nothing)
                Catch ex As Exception

                End Try
            End If
        Else
            If ORDR_NO.Length > 0 Then
                Dim rowARTCCPA1_CUST As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = (SELECT MAX(CCPA_NO) FROM ARTCCPA1 WHERE ORDR_NO = :PARM1)", "V", New Object() {ORDR_NO})
                If rowARTCCPA1_CUST IsNot Nothing Then
                    For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE", "CUST_CREDIT_CARD_EXP_DATE"}
                        ' this line was used to temporarily fix the incorrectly encrypted CC values in an ARTCCPA1 record that we think came in over the old API - but not sure how it encrypted anything
                        'rowARTCCPA1.Item(field & "_E") = clsTACENCRY.EncryptString(rowARTCCPA1.Item(field) & String.Empty)

                        Select Case field
                            Case "CUST_CREDIT_CARD_EXP_DATE"
                                If rowARTCCPA1_CUST.Item(field & "_E") & String.Empty = String.Empty Then
                                    Continue For
                                End If

                                If rowARTCCPA1_CUST.Item(field) & String.Empty <> String.Empty Then
                                    Continue For
                                End If

                        End Select

                        rowARTCCPA1_CUST.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1_CUST.Item(field & "_E") & String.Empty)
                        rowARTCCPA1_CUST.Item(field & "_E") = DBNull.Value
                    Next
                End If
                CUST_CREDIT_CARD_NO = rowARTCCPA1_CUST.Item("CUST_CREDIT_CARD_NO") & String.Empty
                For Each COLUMN_NAME As String In COLs
                    rowARTCCPA1.Item(COLUMN_NAME) = rowARTCCPA1_CUST.Item(COLUMN_NAME)
                Next
                optCC.Value = "X"
                MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO").Text = CUST_CREDIT_CARD_NO
                LoadCCDataIntoRow(rowARTCCPA1_CUST)
                Try
                    txt_ValueChanged(MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO"), Nothing)
                Catch ex As Exception

                End Try
            End If
        End If

        CC_Start()
        CCPA_AMT_orig = Val(rowARTCCPA1.Item("CCPA_AMT") & "")
        lblLeaves.Visible = False
        numLeaves.Visible = False
        optType.Value = TRAN_TYPE

        Absx1.numFor("CCPA_AMT").MaskInput = "nnn,nnn,nnn.nn"
        If CCPA_REASON = "B" OrElse LockAmountField Then
            Absx1.numFor("CCPA_AMT").ReadOnly = True
        End If
        grpReasonVoid.Visible = (TRAN_TYPE = "V" Or TRAN_TYPE = "C")

        If TRAN_TYPE = "V" Or TRAN_TYPE = "C" Then
            optCC.Visible = False
            Set_Read_Only(grpCCData, True)
            grpAgedTotals.Visible = False
            cmdCCSubmit.Text = optType.Items(optType.CheckedIndex).DisplayText
            'lblResponseText.Text = "Click Submit to " & optType.Items(optType.CheckedIndex).DisplayText
            lblResponseText.Appearance.ForeColor = Color.Red
            ShowAgedTotals = False
            grpReasonVoid.Text = "Reason to " & optType.Items(optType.CheckedIndex).DisplayText
        End If

        'optCC.Visible = True
        grpAgedTotals.Visible = ShowAgedTotals
        Set_Read_Only(grpAgedTotals, True)
        If ShowAgedTotals Then
            Absx1.numFor("AGE0").Value = AGE0
            Absx1.numFor("AGE1").Value = AGE1
            Absx1.numFor("AGE2").Value = AGE2
            Absx1.numFor("AGE3").Value = AGE3
            Absx1.numFor("AGE4").Value = AGE4
            Absx1.numFor("TOTAL_DUE").Value = TOTAL_DUE
            Absx1.numFor("TOTAL_W_FUTURES").Value = TOTAL_W_FUTURES

            If AgingDescription <> "" Then
                grpAgedTotals.Text = AgingDescription
            End If
            If AgingHeadings.Count <> 0 Then
                lbl1.Text = AgingHeadings("AGE1")
                lbl2.Text = AgingHeadings("AGE2")
                lbl3.Text = AgingHeadings("AGE3")
                lbl4.Text = AgingHeadings("AGE4")
            End If
        Else
            Me.Height = Me.Height - grpAgedTotals.Height
        End If

        Set_Read_Only(grpType, True)

        If ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & "" = "1" Then
            test_mode = True
            lblTestMode.Visible = True
            MsgBox("Credit Card Processing is operating in Test Mode", MsgBoxStyle.OkOnly, "Please Contact ABS")
        Else
            lblLiveMode.Visible = True
        End If

        MerchantSetup()

        If TRAN_TYPE <> "C" And TRAN_TYPE <> "V" And TRAN_TYPE <> "A" Then
            'If Format(ROWs("ARTPARM1").Item("AR_PARM_CC_CUTOFF_TIME"), "HHmm") _
            ' < Format(Now + ASCMAIN1.NowTSD, "HHmm") And Not test_mode Then
            '    lblResponseText.Text = "Disabled past " & Format(ROWs("ARTPARM1").Item("AR_PARM_CC_CUTOFF_TIME"), "HH:mm")
            '    lblResponseText.Visible = True
            '    cmdCCSubmit.Enabled = False
            'End If
        End If

        Me.Text &= " - " & optType.Items(optType.CheckedIndex).DisplayText

        If ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & "" = "1" Then
            test_mode = True
        End If

        lblCVV2.Visible = True
        lblCVV2.BringToFront()
        txtCUST_CREDIT_CARD_VER_CODE.BringToFront()

        'If ASCMAIN1.USER_ID = "edz" Then
        '    UltraButton1.Visible = True
        'End If

    End Sub

    Private Sub cmdCCSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCCSubmit.Click

        If objCCProcessor Is Nothing Then
            MessageBox.Show("The software is not setup to process credit cards.", "Process Credit Card")
            Exit Sub
        End If

        Dim CCPA_AMT As Decimal = Val(Absx1.numFor("CCPA_AMT").Value & "")
        Dim new_record As Boolean = False

        If CCPA_AMT > maxAmount Then
            MessageBox.Show("Maximum transaction is " & Format(maxAmount, "#,##0.00") & ".", "Process Credit Card")
            Absx1.numFor("CCPA_AMT").Value = maxAmount
            Exit Sub
        End If

        EMsg = String.Empty

        Absx1.txtFor("CUST_CREDIT_CARD_NO").Text = Absx1.txtFor("CUST_CREDIT_CARD_NO").Text.Trim
        Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text = Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text.Trim
        Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").Text = Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").Text.Trim

        Absx1.txtFor("CUST_CREDIT_CARD_NAME").Text = Absx1.txtFor("CUST_CREDIT_CARD_NAME").Text.Trim
        Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text = Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text.Trim
        Absx1.txtFor("CUST_CREDIT_CARD_CITY").Text = Absx1.txtFor("CUST_CREDIT_CARD_CITY").Text.Trim
        Absx1.txtFor("CUST_CREDIT_CARD_STATE").Text = Absx1.txtFor("CUST_CREDIT_CARD_STATE").Text.Trim
        Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").Text = Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").Text.Trim

        Absx1.txtFor("CCPA_REASON_VOID").Text = Absx1.txtFor("CCPA_REASON_VOID").Text.Trim
        If (TRAN_TYPE = "V" Or TRAN_TYPE = "C") And Absx1.txtFor("CCPA_REASON_VOID").Text = "" Then
            EMsg &= vbCr & "You Must enter a Reason for the " & optType.Items(optType.CheckedIndex).DisplayText
        End If

        CUST_CREDIT_CARD_NO = MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO").Text

        Try
            txt_ValueChanged(MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO"), Nothing)
        Catch ex As Exception

        End Try

        Try
            objCCProcessor.CustomerCreditCard.CardNumber = Absx1.txtFor("CUST_CREDIT_CARD_NO").Text
            objCCProcessor.CustomerCreditCard.CardExpMonth = Mid$(Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text, 1, 2) & ""
            objCCProcessor.CustomerCreditCard.CardExpYear = Mid$(Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text, 3, 2) & ""
            objCCProcessor.ValidateCard()
        Catch ex As Exception
            If optType.Value = "C" Then ' WE NEED TO PROCESS THE CREDIT ON THE SAME EXACT CARD, AND SOMETIMES THE CARD IS NOW EXPIRED
            Else
                EMsg &= vbCr & ex.Message
            End If
        End Try

        If Not objCCProcessor.DateCheckPassed Then
            If optType.Value = "C" Then
            Else
                EMsg &= vbCr & "Expiration Date is Not Valid"
            End If
        End If

        If Not objCCProcessor.DigitCheckPassed Then
            EMsg &= vbCr & "Credit Card No is not Valid"
        End If

        If Absx1.txtFor("CUST_CREDIT_CARD_NAME").TextLength = 0 _
            OrElse Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").TextLength = 0 _
            OrElse Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").TextLength = 0 Then
            EMsg &= vbCr & "Credit Card Name, Street Address & Zip Code is Mandatory"
        End If

        With Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE")
            If .TextLength = 0 Then
                EMsg &= vbCr & "Credit Card Expiration Date is Mandatory"
            Else
                Dim EXP As String = Format$(Val(.Text), "0000")
                Dim YY As Integer = Val(Mid(ASCMAIN1.CYM, 3, 2))
                Dim Ys As String = Format$(YY, "00") & "x" & Format$(YY + 1, "00") & "x" & Format$(YY + 2, "00") & "x" & Format$(YY + 3, "00") & "x" & Format$(YY + 4, "00") & "x" & Format$(YY + 5, "00") & "x" & Format$(YY + 6, "00") & "x" & Format$(YY + 7, "00") & "x"
                Dim Ms As String = "01x" & "02x" & "03x" & "04x" & "05x" & "06x" & "07x" & "08x" & "09x" & "10x" & "11x" & "12x"
                If .Text <> EXP _
                Or InStr(Ms, Mid(EXP, 1, 2) & "x") = 0 _
                Or InStr(Ys, Mid(EXP, 3, 2) & "x") = 0 _
                Or Mid(EXP, 3, 2) & Mid(EXP, 1, 2) < Mid(ASCMAIN1.CYM, 3, 4) Then
                    If optType.Value = "C" Then
                        MsgBox("FYI: Invalid Expiration Date" & vbCr & "Credit Process will still be attempted", MsgBoxStyle.OkOnly, "Processing Credit")
                    Else
                        EMsg &= vbCr & "Invalid Expiration Date (" & .Text & ")"
                    End If
                End If
            End If
        End With

        With Absx1.txtFor("CUST_CREDIT_CARD_NO")
            If .TextLength = 0 Then
                EMsg &= vbCr & "Credit Card No is Mandatory"
            Else
                Dim CUST_CREDIT_CARD_NO_Valid As String = Format$(CLng(.Text), "".PadLeft(Len(.Text), "0"))
                If .Text <> CUST_CREDIT_CARD_NO_Valid Then
                    EMsg &= vbCr & "Invalid Credit Card No (" & .Text & ")"
                End If
            End If
        End With

        If Val(Absx1.numFor("CCPA_AMT").Value & "") <= 0 Then
            EMsg &= vbCr & "Credit Card Amount must be > 0"
        End If

        If optCC.Value = "N" Then
            If dst.Tables("ARTCUSTC").Select("CUST_CREDIT_CARD_NO = '" & CUST_CREDIT_CARD_NO & "'").Length > 0 Then
                EMsg &= vbCr & "Credit Card is Already on file - use Edit and Select from Lookup"
            End If
        End If

        Dim CUST_CREDIT_CARD_COUNTRY As String = String.Empty
        If cbeCountry.SelectedItem Is Nothing Then
            EMsg &= vbCr & "Credit Card Country is required"
        Else
            CUST_CREDIT_CARD_COUNTRY = cbeCountry.SelectedItem.DataValue & String.Empty
            If CUST_CREDIT_CARD_COUNTRY = "USA" Then
                CUST_CREDIT_CARD_COUNTRY = "US"
                cbeCountry.Value = CUST_CREDIT_CARD_COUNTRY
            End If

            If CUST_CREDIT_CARD_COUNTRY.Length = 0 Then
                EMsg &= vbCr & "Credit Card Country is required"
            ElseIf tblTATCNTRY.Select($"COUNTRY_CODE = '{CUST_CREDIT_CARD_COUNTRY}'").Length = 0 Then
                EMsg &= vbCr & "Credit Card Country is Invalid"
            End If
        End If

        Dim ZIP_CODE As String = Replace(Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").Text, "-", "") & String.Empty
        If CUST_CREDIT_CARD_COUNTRY = "US" Then
            If ZIP_CODE.Length <> 5 And ZIP_CODE.Length <> 9 Then
                ZIP_CODE = ""
            End If

            If ZIP_CODE.Length < 5 Then
                EMsg &= vbCr & "US Transactions require a 5 character Credit Card Zip Code."
            End If
        Else
            If ZIP_CODE.Length = 0 Then
                If MessageBox.Show("International Credit Card Zip Code is missing. Do you want to continue?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    EMsg &= vbCr & "Credit Card Zip Code is required"
                End If
            End If
        End If

        If Absx1.txtFor("CUST_CREDIT_CARD_CITY").TextLength = 0 Then
            EMsg &= vbCr & "Credit Card City is required"
        End If

        If Absx1.txtFor("CUST_CREDIT_CARD_STATE").TextLength = 0 Then
            EMsg &= vbCr & "Credit Card State is required"
        End If

        If Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").TextLength = 0 Then
            EMsg &= vbCr & "Credit Card CVV2 is required"
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Submit Credit Card for Processing")
            Exit Sub
        End If

        Dim warningMessage As String = String.Empty
        Try
            Select Case objCCProcessor.GetCreditCardType()
                Case Nothing
                            ' Nothing 
                Case TAC.ARCCCARD.CreditCardTypes.vctAmex
                    If Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").TextLength <> 4 Then
                        warningMessage = "Americam Express cards usually have a 4 digit CVV2. Do you want to continue?"
                    End If
                Case Else
                    If Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").TextLength <> 3 Then
                        warningMessage = "This type of credit card usually has a 3 digit CVV2. Do you want to continue?"
                    End If
            End Select
        Catch ex As Exception

        End Try

        If warningMessage.Length > 0 Then
            If MessageBox.Show(warningMessage, "Submit", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                Exit Sub
            End If
        End If

        Dim TransactionNumber As String = ASCMAIN1.Next_Control_No("ARTCCPA1.TRANS_NUM")
        Dim CUST_CREDIT_CARD_KEY As String = ""
        If optCC.Value = "N" Then
            CUST_CREDIT_CARD_KEY = ASCMAIN1.Next_Control_No("ARTCUSTC.CUST_CREDIT_CARD_KEY")
            If MyBase.Absx1.txtFor("CUST_CREDIT_CARD_KEY").TextLength = 0 Then
                MyBase.Absx1.txtFor("CUST_CREDIT_CARD_KEY").Text = CUST_CREDIT_CARD_KEY
            End If
        End If

        BeginTrans()

        CCPA_NO = rowARTCCPA1.Item("CCPA_NO") & ""
        Dim CCPA_NO_orig As String = CCPA_NO

        Dim LAST_DATE As Date = Now + ASCMAIN1.NowTSD

        If CCPA_NO <> "" Then
            Dim rowARTCCPA1_orig As DataRow = dst.Tables("ARTCCPA1").NewRow
            For i As Int32 = 0 To rowARTCCPA1_orig.Table.Columns.Count - 1
                rowARTCCPA1_orig.Item(i) = rowARTCCPA1.Item(i)
            Next
            dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1_orig)
            rowARTCCPA1_orig.AcceptChanges()
            rowARTCCPA1 = dst.Tables("ARTCCPA1").Rows.Find(CCPA_NO)
        End If

        If optType.Value <> "V" Then
            If optType.Value = "C" Then
                CCPA_NO = ""
            End If
            If CCPA_NO = "" Then
                CCPA_NO = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
                new_record = True
                rowARTCCPA1 = dst.Tables("ARTCCPA1").NewRow
                rowARTCCPA1.Item("CCPA_NO") = CCPA_NO
                rowARTCCPA1.Item("CUST_CODE") = CUST_CODE
                ' rowARTCCPA1.Item("CUST_NAME") = CUST_NAME
                rowARTCCPA1.Item("CCPA_REASON") = CCPA_REASON
                rowARTCCPA1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowARTCCPA1.Item("INIT_DATE") = LAST_DATE
                rowARTCCPA1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                rowARTCCPA1.Item("CUST_CREDIT_CARD_KEY") = MyBase.Absx1.txtFor("CUST_CREDIT_CARD_KEY").Text
            End If

            If optType.Value = "A" Then
                rowARTCCPA1.Item("CCPA_DATE_AUTH") = LAST_DATE
            ElseIf optType.Value = "S" Then
                rowARTCCPA1.Item("CCPA_DATE_SALE") = LAST_DATE
            End If

            If optType.Value = "C" Then
                rowARTCCPA1.Item("CCPA_AMT") = -1 * Val(Absx1.numFor("CCPA_AMT").Value & "")
                rowARTCCPA1.Item("CCPA_NOTE") = Absx1.txtFor("CCPA_REASON_VOID").Text
                rowARTCCPA1.Item("CCPA_NO_CREDITED") = CCPA_NO_orig
            Else
                rowARTCCPA1.Item("CCPA_AMT") = Val(Absx1.numFor("CCPA_AMT").Value & "")
                rowARTCCPA1.Item("CCPA_NOTE") = Absx1.txtFor("CCPA_NOTE").Text
                rowARTCCPA1.Item("CCPA_REASON_VOID") = Absx1.txtFor("CCPA_REASON_VOID").Text
            End If

            LoadCCDataIntoRow(rowARTCCPA1)

            rowARTCCPA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowARTCCPA1.Item("LAST_DATE") = LAST_DATE

            rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text

            If rowARTCCPA1.RowState = DataRowState.Detached Then
                dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1)
            Else
                rowARTCCPA1.EndEdit()
            End If

            If optType.Value <> "C" Then
                dst.Tables("ARTCUSTC").AcceptChanges()

                If optCC.Value = "E" Then

                    rowARTCUSTC = dst.Tables("ARTCUSTC").Rows.Find(New Object() {CUST_CODE, CUST_CREDIT_CARD_NO})
                    If rowARTCUSTC IsNot Nothing Then
                        LoadCCDataIntoRow(rowARTCUSTC)
                        rowARTCUSTC.Item("CUST_CREDIT_CARD_STATUS") = "A"
                        rowARTCUSTC.Item("LAST_DATE") = LAST_DATE
                        rowARTCUSTC.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        DecryptARTCUSTC()
                        Update_Record_TDA("ARTCUSTC")
                        EncryptARTCUSTC()
                    End If
                ElseIf optCC.Value = "N" Then
                    rowARTCUSTC = dst.Tables("ARTCUSTC").NewRow
                    rowARTCUSTC.Item("CUST_CODE") = CUST_CODE
                    LoadCCDataIntoRow(rowARTCUSTC)
                    rowARTCUSTC.Item("CUST_CREDIT_CARD_STATUS") = "A"
                    rowARTCUSTC.Item("CUST_CREDIT_CARD_KEY") = CUST_CREDIT_CARD_KEY
                    dst.Tables("ARTCUSTC").Rows.Add(rowARTCUSTC)
                    rowARTCUSTC.Item("INIT_DATE") = LAST_DATE
                    rowARTCUSTC.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    EncryptARTCUSTC()
                    Update_Record_TDA("ARTCUSTC")
                    DecryptARTCUSTC()
                End If
            End If
        End If

        '** the code below is replicated in individual procedures below, such as CC_Capture, CC_Sale, CC_Auth
        objCCProcessor.CreditCardProcessingNo = CCPA_NO
        objCCProcessor.InternalReference = "Customer: " & rowARTCCPA1.Item("CUST_CODE") & ", TransType: " & TRAN_TYPE
        objCCProcessor.TransactionNumber = TransactionNumber

        objCCProcessor.CustomerCreditCard.CardHolderFirstName = Absx1.txtFor("CUST_CREDIT_CARD_NAME").Text
        objCCProcessor.CustomerCreditCard.CardHolderLastName = ""
        objCCProcessor.CustomerCreditCard.CardHolderAddress = Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text
        objCCProcessor.CustomerCreditCard.CardHolderCity = Absx1.txtFor("CUST_CREDIT_CARD_CITY").Text
        objCCProcessor.CustomerCreditCard.CardHolderState = Absx1.txtFor("CUST_CREDIT_CARD_STATE").Text
        objCCProcessor.CustomerCreditCard.CardHolderZipCode = ZIP_CODE
        objCCProcessor.CustomerCreditCard.CardHolderCountry = cbeCountry.Value
        objCCProcessor.CustomerCreditCard.CardHolderTelephone = ""
        objCCProcessor.CustomerCreditCard.CardCVVData = Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").Text
        objCCProcessor.CustomerCreditCard.CustomerID = CUST_CODE
        objCCProcessor.XmlDirectory = CC_PROC_FOLDER

        If rowARTCCPA1.Item("CUST_CREDIT_CARD_KEY") & String.Empty = String.Empty Then
            If CUST_CREDIT_CARD_KEY.Length > 0 Then
                rowARTCCPA1.Item("CUST_CREDIT_CARD_KEY") = CUST_CREDIT_CARD_KEY
            ElseIf Absx1.txtFor("CUST_CREDIT_CARD_KEY").TextLength > 0 Then
                rowARTCCPA1.Item("CUST_CREDIT_CARD_KEY") = Absx1.txtFor("CUST_CREDIT_CARD_KEY").Text
            End If
        End If

        Try

            Select Case optType.Value
                Case "A"
                    objCCProcessor.TransactionAmount = Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "#.00")
                    objCCProcessor.TransactionNumber = TransactionNumber

                    CheckTestMode()
                    objCCProcessor.AuthOnly()

                Case "S"
                    objCCProcessor.TransactionAmount = Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "#.00")
                    objCCProcessor.TransactionNumber = TransactionNumber
                    CheckTestMode()

                    If overrideSaleWithCapture Then
                        objCCProcessor.Capture(overrideSaleTransactionID, overrideSaleCreditCardFullName, overrideSaleApprovalCode)
                    Else
                        objCCProcessor.Sale()
                    End If

                Case "C"
                    CheckTestMode()
                    objCCProcessor.TransactionAmount = Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "#.00")
                    objCCProcessor.TransactionNumber = TransactionNumber
                    objCCProcessor.Credit()

                Case "V"
                    CheckTestMode()

                    Dim CreditCardInfo As New TAC.ARCCCARD.CreditCard
                    CreditCardInfo = CreateCreditCardInfo(rowARTCCPA1)

                    With CreditCardInfo
                        .InvoiceNumber = String.Empty
                        .TransArmorToken = String.Empty
                        .RefundAmount = Math.Abs(Val(rowARTCCPA1.Item("CCPA_AMT") & String.Empty))
                    End With

                    objCCProcessor.Refund(CreditCardInfo)

                Case Else
                    MessageBox.Show("Unknow Processing type.", "Process", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
            End Select

            If objCCProcessor.NetworkResponse Is Nothing Then
                If new_record Then
                    rowARTCCPA1.Delete()
                    dst.Tables("ARTCCPA1").AcceptChanges()
                    CCPA_NO = String.Empty
                    rowARTCCPA1 = dst.Tables("ARTCCPA1").NewRow
                End If
                Rollback(objCCProcessor.LastError & IIf(objCCProcessor.RawResponseText.Length > 0, " " & objCCProcessor.RawResponseText, ""))
                Exit Sub
            End If

            If objCCProcessor.NetworkResponse.Approved Then
                Dim responseDollarAmount As Decimal = Val(objCCProcessor.NetworkResponse.ExtractNodeFromResponse("DollarAmount"))
                Dim responseCTR As String = objCCProcessor.NetworkResponse.ExtractNodeFromResponse("CTR")
                If responseDollarAmount > 0 AndAlso (responseDollarAmount + 0.05) < CCPA_AMT Then
                    MessageBox.Show($"Amount requested does not match the amount approved: {responseDollarAmount}. Amount will be updated.", "Process Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    MessageBox.Show(responseCTR, "Process Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Absx1.numFor("CCPA_AMT").Value = responseDollarAmount
                    rowARTCCPA1.Item("CCPA_AMT") = responseDollarAmount
                End If
            End If

            responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)
            lblResponseText.Text = responseErrorMessage

            If objCCProcessor.NetworkResponse.Approved Then
                lblResponseText.Appearance.ForeColor = Drawing.Color.Green
            Else
                lblResponseText.Appearance.ForeColor = Drawing.Color.Red
                ' 12/1/2022 - The bank may approve it but Payeezy may not approve it.
                If lblResponseText.Text.ToUpper = "APPROVED" Then
                    lblResponseText.Text = "Declined"
                End If
            End If

            dst.Tables("ARTCCPA2").Rows.Clear()
            Dim rowARTCCPA2 As DataRow = dst.Tables("ARTCCPA2").NewRow
            rowARTCCPA2.Item("CCPA_NO") = CCPA_NO
            rowARTCCPA2.Item("RESPONSE_TEXT") = objCCProcessor.NetworkResponse.Text
            responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)

            If objCCProcessor.NetworkResponse.Approved Then
                rowARTCCPA2.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.Code
                rowARTCCPA2.Item("RESPONSE_APPROVAL_CODE") = objCCProcessor.NetworkResponse.ApprovalCode
            Else
                rowARTCCPA2.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.ErrorCode
            End If
            AdjustResponseCode(rowARTCCPA2)

            Dim RESPONSE_DATA As String = objCCProcessor.NetworkResponse.Data
            If RESPONSE_DATA.Length > rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength Then
                RESPONSE_DATA = RESPONSE_DATA.Substring(0, rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength).Trim
            End If
            rowARTCCPA2.Item("RESPONSE_DATA") = RESPONSE_DATA

            rowARTCCPA2.Item("RESPONSE_AVS") = objCCProcessor.NetworkResponse.AVSResult
            rowARTCCPA2.Item("INIT_DATE") = LAST_DATE
            rowARTCCPA2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowARTCCPA2.Item("CCPA_TYPE") = optType.Value
            dst.Tables("ARTCCPA2").Rows.Add(rowARTCCPA2)

            Dim rowARTCCPDA As DataRow
            Fill_Records("ARTCCPDA", CCPA_NO)
            If dst.Tables("ARTCCPDA").Rows.Count = 0 Then
                rowARTCCPDA = dst.Tables("ARTCCPDA").NewRow

                rowARTCCPDA.Item("CCPA_NO") = CCPA_NO
                dst.Tables("ARTCCPDA").Rows.Add(rowARTCCPDA)
            Else
                rowARTCCPDA = dst.Tables("ARTCCPDA").Rows(0)
            End If

            If objCCProcessor.NetworkResponse.DetailAggregate.Length > 0 AndAlso (rowARTCCPDA.Item("DETAIL_AGGREGATE") & String.Empty).ToString.Length = 0 Then
                rowARTCCPDA.Item("DETAIL_AGGREGATE") = objCCProcessor.NetworkResponse.DetailAggregate
                rowARTCCPDA.Item("COMM_CARD_TYPE") = String.Empty ' objCCProcessor.NetworkResponse.CommercialCardType
            End If

            If optType.Value <> "V" Then
                Record_Response(False)
            End If

            If optType.Value = "S" Then
                Record_Level_2_3_Aggregates(CCPA_NO, CCPA_NO)
            End If

            If objCCProcessor.NetworkResponse.Approved Then
                If optType.Value = "V" Then
                    rowARTCCPA1.Item("CCPA_STATUS") = "V"
                Else
                    If optType.Value = "A" Then
                        rowARTCCPA1.Item("CCPA_STATUS") = "T"
                    Else
                        rowARTCCPA1.Item("CCPA_STATUS") = "A"
                    End If
                End If
                If optType.Value = "C" Then
                    ASCMAIN1.sql = "Update ARTPYMT2 Set CCPA_NO_CREDIT = :PARM1 where CCPA_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {CCPA_NO, CCPA_NO_orig})
                    CCPA_NO_CREDIT = CCPA_NO
                End If
            Else
                If optType.Value <> "V" Then
                    rowARTCCPA1.Item("CCPA_STATUS") = "E"
                End If
            End If

            If optType.Value <> "V" Then
                rowARTCCPA1.Item("ORDR_NO") = ORDR_NO
                rowARTCCPA1.Item("INV_NO") = INV_NO
                rowARTCCPA1.Item("STMT_NO") = STMT_NO
            End If

            rowARTCCPA1.Item("CCPA_TYPE") = optType.Value
            EncryptARTCCPA1()
            Update_Record_TDA("ARTCCPA1")
            Update_Record_TDA("ARTCCPA2")
            Update_Record_TDA("ARTCCPDA")

            Try
                ASCMAIN1.sql = $"BEGIN ARTCCPA1_ARTCUSTC('{CCPA_NO}'); End;"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Catch ex As Exception

            End Try

            Dim MSG As String = "Credit Card Payment Submitted"

            If optType.Value = "A" Then
                If objCCProcessor.NetworkResponse.Approved Then
                    MSG = "Credit Card Auth Approved For " & Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "$#,###.00")
                Else
                    MSG = "Credit Card Auth Declined"
                End If
            End If

            If optType.Value = "V" Then
                MSG = "Credit Card Payment Voided"
            ElseIf optType.Value = "C" Then
                MSG = "Credit Card Credit Processed"
            End If

            If Not objCCProcessor.NetworkResponse.Approved Then
                MSG = "Credit Card Declined." '  & objCCProcessor.NetworkResponse.Text
            End If

            Dim EVENT_DESC As String = MSG
            EVENT_DESC = EVENT_DESC.Replace("'", " ")
            If EVENT_DESC.Length > 150 Then
                EVENT_DESC = EVENT_DESC.Substring(0, 150)
            End If

            Try
                ' Write an Order event
                If ORDR_NO.Length > 0 AndAlso EVENT_DESC.Length > 0 Then
                    TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "CCP", EVENT_DESC)
                End If
            Catch ex As Exception
                ' Nothing
            End Try

            CommitTrans(MSG)

        Catch ex As Exception
            Rollback("Please call ABS - Error processing Card Payment: " & ex.Message)
        End Try

        Me.Close()

    End Sub

    Private Function GetTaxInfo(ByVal CUST_CODE As String, ByVal CUST_ADDR_TYPE As String, ByVal CUST_ADDR_CODE As String) As Double

        Dim taxPercent As Double = 0
        Try
            Dim rowARTCUST2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_ADDR_TYPE = :PARM2 AND CUST_ADDR_CODE = :PARM3", "VVV", New Object() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})

            If rowARTCUST2 Is Nothing Then
                Return 0
            End If

            Dim STAX_CODE As String = rowARTCUST2.Item("STAX_CODE") & String.Empty

            If STAX_CODE.Length = 0 Then
                Return 0
            End If

            Dim rowARTSTAX1 As DataRow = LookUp("ARTSTAX1", STAX_CODE)
            If rowARTSTAX1 Is Nothing Then
                Return 0
            End If

            Return Val(rowARTSTAX1.Item("STAX_RATE") & "")
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Sub LoadCCDataIntoRow(ByVal row As DataRow)
        For Each COL As String In COLs
            Select Case COL
                Case "CUST_CREDIT_CARD_COUNTRY"
                    If row.Item(COL) & "" = "USA" Or row.Item(COL) & "" = "US" Then
                        row.Item(COL) = "US"
                        cbeCountry.Value = "US"
                    Else
                        If cbeCountry.Value & "" <> "" AndAlso row.Item(COL) & "" <> cbeCountry.Value Then
                            row.Item("CUST_CREDIT_CARD_COUNTRY") = cbeCountry.SelectedItem.DataValue
                        Else
                            If Absx1.cbeFor(COL).Value = "USA" Then
                                row.Item(COL) = "US"
                            Else
                                row.Item(COL) = Absx1.cbeFor(COL).Value
                            End If

                        End If
                    End If

                Case "CUST_CREDIT_CARD_KEY"
                    If row.Item(COL) & "" <> Absx1.txtFor(COL).Text Then
                        If Absx1.txtFor(COL).TextLength > 0 Then
                            row.Item(COL) = Absx1.txtFor(COL).Text
                        End If
                    End If

                Case Else
                    If row.Item(COL) & "" <> Absx1.txtFor(COL).Text Then
                        row.Item(COL) = Absx1.txtFor(COL).Text
                    End If
            End Select

        Next
    End Sub

    Sub CC_Clear()
        For Each COL As String In COLs
            Select Case COL
                Case "CUST_CREDIT_CARD_COUNTRY"
                    cbeCountry.Value = String.Empty

                Case Else
                    Absx1.txtFor(COL).Clear()
            End Select
        Next
        Call CC_EnableControls(False)
    End Sub

    Sub CC_Start()

        CUST_CODE = rowARTCCPA1.Item("CUST_CODE")
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
        CUST_NAME = rowARTCUST1.Item("CUST_NAME") & ""
        grpCCData.Text = "Customer " & CUST_CODE
        Absx1.txtFor("CUST_NAME").Text = CUST_NAME
        Absx1.txtFor("CUST_CONTACT").Text = rowARTCUST1.Item("CUST_CONTACT") & ""
        Absx1.medFor("CUST_PHONE").Value = rowARTCUST1.Item("CUST_PHONE") & ""
        lblResponseText.Text = ""
        optCC.Value = "X"

        ASCMAIN1.sql = "Select Sum (CUST_PYMT_AMT) from ARTPYMT1,ARTPYMT2 where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO and ARTPYMT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "' and CUST_CODE = '" & CUST_CODE & "' and NVL(PYMT_DELETED,'0') <> '1'"
        Dim PAID_MTD As Decimal = Val(ASCDATA1.GetDataValue & "")
        Absx1.numFor("PAID_MTD").Value = PAID_MTD
    End Sub

    Sub CC_EnableControls(ByVal tf As Boolean)

        Absx1.txtFor("CUST_CREDIT_CARD_NO").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_NAME").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_CITY").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_STATE").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").ReadOnly = Not tf
        cbeCountry.ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").ReadOnly = Not tf

        Absx1.txtFor("CUST_CREDIT_CARD_NO").PasswordChar = "*"
        cmdUseCustomerAddress.Visible = tf
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "CCPA_AMT"
                If CCPA_AMT_orig <> 0 Then
                    Dim CCPA_AMT = Val(Absx1.numFor("CCPA_AMT").Value & "")
                    If CCPA_AMT_orig - CCPA_AMT <= 0 Then
                        numLeaves.Value = 0
                    Else
                        numLeaves.Value = CCPA_AMT_orig - CCPA_AMT
                    End If
                    ' numLeaves.Value = CCPA_AMT_orig - CCPA_AMT
                    If CCPA_AMT <> CCPA_AMT_orig Then
                        lblLeaves.Visible = True
                        numLeaves.Visible = True
                        lblLeaves.Top = numLeaves.Top + 5
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CREDIT_CARD_NO"
                Dim txtctl As UltraWinEditors.UltraTextEditor
                txtctl = DirectCast(sender, UltraWinEditors.UltraTextEditor)
                If txtctl.ReadOnly Then
                    Exit Sub
                End If
                If txtctl.PasswordChar = "*" Then
                    e.Handled = True
                    e.SuppressKeyPress = True
                End If
        End Select
    End Sub

    Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CREDIT_CARD_NO"
                Dim CUST_CREDIT_CARD_NO As String = Absx1.txtFor("CUST_CREDIT_CARD_NO").Text
                Absx1.txtFor("CUST_CREDIT_CARD_LAST4").Clear()
                Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Clear()
                picCC.Image = Nothing
                Dim IMAGE_FILE As String = String.Empty

                If CUST_CREDIT_CARD_NO.Length > 4 Then
                    Absx1.txtFor("CUST_CREDIT_CARD_LAST4").Text = CUST_CREDIT_CARD_NO.Substring(CUST_CREDIT_CARD_NO.Length - 4)
                End If

                Try
                    objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
                    Select Case objCCProcessor.GetCreditCardType()
                        Case TAC.ARCCCARD.CreditCardTypes.vctAmex
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "AMEX"
                            IMAGE_FILE = "AMEX.GIF"
                        Case TAC.ARCCCARD.CreditCardTypes.vctMasterCard
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "MSTR"
                            IMAGE_FILE = "MSTR.GIF"
                        Case TAC.ARCCCARD.CreditCardTypes.vctDiscover
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "DISC"
                            IMAGE_FILE = "DISC.GIF"
                        Case TAC.ARCCCARD.CreditCardTypes.vctVisa, ARCCCARD.CreditCardTypes.vctVisaElectron
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "VISA"
                            IMAGE_FILE = "VISA.GIF"
                        Case TAC.ARCCCARD.CreditCardTypes.vctDiners
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "DC"
                        Case Else
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = String.Empty
                            IMAGE_FILE = String.Empty
                    End Select

                Catch ex As Exception
                    Exit Sub
                End Try

                If IMAGE_FILE <> "" Then
                    picCC.Visible = True
                    picCC.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "\ABS\CC\", IMAGE_FILE)
                Else
                    picCC.Image = Nothing
                End If
        End Select

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(
    ByVal ctl As Control,
    ByVal COLUMN_NAME As String,
    Optional ByRef sql_where As String = "",
    Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "CUST_CREDIT_CARD_NO"
                sql_where = "CUST_CODE = '" & CUST_CODE & "'"
                If optCC.Value = "N" Then
                    Cancel = True
                    'SelectPreviouslyUsed()
                    Cancel = True
                End If

            Case "CUST_CREDIT_CARD_LAST4"
                sql_where = "CUST_CODE = '" & CUST_CODE & "'"
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CREDIT_CARD_NO"

                Dim rowARTCUSTC As DataRow = dst.Tables("ARTCUSTC").Rows.Find(New Object() {CUST_CODE, txtctl.Text})
                If rowARTCUSTC IsNot Nothing Then
                    MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO").Text = rowARTCUSTC.Item("CUST_CREDIT_CARD_NO") & String.Empty
                Else
                    rowARTCUSTC = dst.Tables("ARTCUSTC").NewRow
                End If

                LoadDataFromARTCUSTC(rowARTCUSTC)
                Try
                    txt_ValueChanged(MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO"), Nothing)
                Catch ex As Exception

                End Try
        End Select
    End Sub

    Sub LoadDataFromARTCUSTC(ByVal rowARTCUSTC As DataRow)

        With rowARTCUSTC
            Absx1.txtFor("CUST_CREDIT_CARD_NAME").Text = .Item("CUST_CREDIT_CARD_NAME") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text = .Item("CUST_CREDIT_CARD_ADDR1") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_CITY").Text = .Item("CUST_CREDIT_CARD_CITY") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_STATE").Text = .Item("CUST_CREDIT_CARD_STATE") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").Text = .Item("CUST_CREDIT_CARD_ZIP_CODE") & ""
            cbeCountry.Value = .Item("CUST_CREDIT_CARD_COUNTRY") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text = .Item("CUST_CREDIT_CARD_EXP_DATE") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").Text = .Item("CUST_CREDIT_CARD_VER_CODE") & ""
        End With
    End Sub

    Public Sub MerchantSetup()

        Dim rowARTCCPRC As DataRow = ASCDATA1.GetDataRow($"Select * from ARTCCPRC WHERE CC_PROC_CODE = '{ROWs("SOTPARM1").Item("SO_PARM_CC_PROC_CODE") & String.Empty}'")

        objCCProcessor.MerchantAccount.ProcessingServer = rowARTCCPRC.Item("CC_PROC_SERVER") & String.Empty
        objCCProcessor.MerchantAccount.UserID = rowARTCCPRC.Item("CC_PROC_USER_ID") & String.Empty
        objCCProcessor.MerchantAccount.Password = rowARTCCPRC.Item("CC_PROC_PASSWORD") & String.Empty
        objCCProcessor.MerchantAccount.KeyId = rowARTCCPRC.Item("CC_PROC_DATAWIRE_ID") & String.Empty
        objCCProcessor.MerchantAccount.HMACKey = rowARTCCPRC.Item("CC_PROC_AUTH_TRANS_KEY") & String.Empty
        objCCProcessor.Gateway = Val(rowARTCCPRC.Item("CC_PROC_CODE") & String.Empty)
    End Sub

    Private Sub optCC_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCC.ValueChanged

        Absx1.txtFor("CUST_CREDIT_CARD_NO").PasswordChar = "*"
        Absx1.txtFor("CUST_CREDIT_CARD_NO").Appearance.BackColor = Drawing.Color.Empty

        Select Case optCC.Value
            Case "N"
                CC_Clear()
                Call CC_EnableControls(True)
                Absx1.txtFor("CUST_CREDIT_CARD_NO").PasswordChar = ""
                Absx1.txtFor("CUST_CREDIT_CARD_NO").Appearance.BackColor = Drawing.Color.AliceBlue
                Absx1.txtFor("CUST_CREDIT_CARD_NO").Focus()

                If ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & "" = "1" Then
                    Absx1.txtFor("CUST_CREDIT_CARD_NO").Text = "4444333322221111"
                    Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text = DateAdd(DateInterval.Year, 1, DateTime.Now).ToString("MMyy")
                    Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").Text = "123"
                End If

            Case "E"
                LoadCCDataFromInitialCCPA()
                Call CC_EnableControls(True)
                Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Focus()

            Case "X"
                LoadCCDataFromInitialCCPA()
        End Select
    End Sub

    Sub LoadCCDataFromInitialCCPA()
        For Each COLUMN_NAME As String In New String() _
        {"CUST_CREDIT_CARD_NO" _
        , "CUST_CREDIT_CARD_NAME", "CUST_CREDIT_CARD_ADDR1" _
        , "CUST_CREDIT_CARD_CITY", "CUST_CREDIT_CARD_STATE" _
        , "CUST_CREDIT_CARD_ZIP_CODE", "CUST_CREDIT_CARD_COUNTRY" _
        , "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE" _
        , "CCPA_AMT", "CCPA_NOTE", "CCPA_REASON_VOID"}
            If COLUMN_NAME = "CCPA_AMT" AndAlso Val(rowARTCCPA1.Item(COLUMN_NAME) & "") <= 0 Then
                Absx1.CtlFor(COLUMN_NAME).Text = 0
            ElseIf COLUMN_NAME = "CCPA_AMT" Then
                Absx1.CtlFor(COLUMN_NAME).Text = Val(rowARTCCPA1.Item(COLUMN_NAME) & "")
            ElseIf COLUMN_NAME = "CUST_CREDIT_CARD_COUNTRY" Then
                Absx1.cbeFor(COLUMN_NAME).Value = rowARTCCPA1.Item(COLUMN_NAME) & ""
            Else
                Absx1.CtlFor(COLUMN_NAME).Text = rowARTCCPA1.Item(COLUMN_NAME) & ""
            End If
        Next

        Call CC_EnableControls(False)
    End Sub

    Sub CheckTestMode()
        If test_mode Or optCC.Value = "A" Then
            objCCProcessor.CustomerCreditCard.CardNumber = "4444333322221111"
            objCCProcessor.CustomerCreditCard.CardExpMonth = System.DateTime.Now.ToString("MM")
            objCCProcessor.CustomerCreditCard.CardExpYear = DateAdd(DateInterval.Year, 1, DateTime.Now).ToString("yy")
            objCCProcessor.CustomerCreditCard.CardCVVData = "123"
            objCCProcessor.ValidateCard()
        End If
    End Sub

    Private Sub cmdUseCustomerAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUseCustomerAddress.Click
        For Each COL As String In New String() {"ADDR1", "CITY", "STATE", "ZIP_CODE", "COUNTRY"}
            Select Case COL
                Case "COUNTRY"
                    Try
                        Dim CUST_COUNTRY As String = rowARTCUST1.Item("CUST_" & COL) & String.Empty
                        If CUST_COUNTRY.Length = 0 Then
                            CUST_COUNTRY = "US"
                        End If
                        CUST_COUNTRY = CUST_COUNTRY.ToUpper.Trim

                        Select Case CUST_COUNTRY.Length
                            Case 2
                                If tblTATCNTRY.Select($"COUNTRY_CODE2 = '{CUST_COUNTRY}'").Length > 0 Then
                                    cbeCountry.Value = CUST_COUNTRY
                                Else
                                    cbeCountry.Value = String.Empty
                                End If
                            Case 3
                                Dim rowTATCNTRY As DataRow = LookUp("TATCNTRY", CUST_COUNTRY)
                                If rowTATCNTRY IsNot Nothing Then
                                    CUST_COUNTRY = rowTATCNTRY.Item("COUNTRY_CODE2") & String.Empty
                                    cbeCountry.Value = CUST_COUNTRY
                                Else
                                    cbeCountry.Value = String.Empty
                                End If
                            Case Else
                                cbeCountry.Value = String.Empty
                        End Select
                    Catch ex As Exception
                        cbeCountry.Value = String.Empty
                    End Try

                Case Else
                    Absx1.txtFor("CUST_CREDIT_CARD_" & COL).Text = rowARTCUST1.Item("CUST_" & COL) & ""
            End Select
        Next


        If Absx1.txtFor("CUST_CREDIT_CARD_NAME").Text.Trim.Length = 0 Then
            Absx1.txtFor("CUST_CREDIT_CARD_NAME").Text = rowARTCUST1.Item("CUST_NAME") & ""
        End If
    End Sub

    Public Function CC_Capture(ByVal CCPA_AMT As Decimal) As String

        ' CC_Capture, CC_Authorize & CC_Sale are called by SH modules which instantiate this form as a class
        ' when this happens, Form Load does not fire 
        '  (and it is good that it does not because there is a lot of UI going on in Form Load)
        ' A notable difference between CC_Capture and the others is that CC_Capture does not update Oracle,
        '  whereas CC_Authorize and CC_Sale do
        ' However, the datalayer that is used needs to (still) be the calling form's datalayer,
        '  since Form Load is not called - which is where we Create_TDA
        ' So these routines will continue to use the calling form's datalayer 
        '  until we come up with a more elegant solution

        Initialize_DataLayer() ' Instantiation of this form does not call Form_Load

        ' This code used to be in the Create_CC_Capture routine
        ' however, after the credit card is processed: ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
        ' causes / in a lock the credit card trancasction does not get recorded
        ' This happened on 7/23/2010 causing issues settling credit cards and payments against the customer

        If Not dst.Tables.Contains("ARTCCPA1") Then
            dst = clsASCBASE1.dst
            Create_TDA(dst.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPDA", "*", 1)
            Create_TDA(dst.Tables.Add, "ARTCUSTC", "*", 1)
            dst.Tables("ARTCUSTC").Columns.Add("EXP", GetType(System.Int16))
            If rowARTCCPA1 Is Nothing Then
                rowARTCCPA1 = dst.Tables("ARTCCPA1").NewRow
            Else
                dst.Tables("ARTCCPA1").ImportRow(rowARTCCPA1)
            End If
        End If

        Dim CCPA_NO As String = rowARTCCPA1.Item("CCPA_NO") & String.Empty
        Dim CUST_CODE As String = rowARTCCPA1.Item("CUST_CODE") & String.Empty

        Fill_Records("ARTCCPA1", CCPA_NO)
        DecryptARTCCPA1()
        Fill_Records("ARTCCPDA", CCPA_NO)

        Fill_Records("ARTCUSTC", CUST_CODE)
        DecryptARTCUSTC()

        ' Sync the row with what was passed in
        If rowARTCCPA1 IsNot Nothing AndAlso rowARTCCPA1.Item("CCPA_NO") = CCPA_NO Then
            If dst.Tables("ARTCCPA1").Select("CCPA_NO = '" & CCPA_NO & "'").Length > 0 Then
                Dim rowARTCCPA1_CURR As DataRow = dst.Tables("ARTCCPA1").Select("CCPA_NO = '" & CCPA_NO & "'")(0)
                For Each col As DataColumn In rowARTCCPA1.Table.Columns
                    If rowARTCCPA1_CURR.Table.Columns.Contains(col.ColumnName) Then
                        rowARTCCPA1_CURR.Item(col.ColumnName) = rowARTCCPA1.Item(col.ColumnName)
                    End If
                Next
            End If
        End If

        MerchantSetup()

        If ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & "" = "1" Then
            test_mode = True
        End If

        Prepare_Component(CCPA_AMT)
        CheckTestMode()

        Dim CreditCardInfo As New TAC.ARCCCARD.CreditCard
        CreditCardInfo = CreateCreditCardInfo(rowARTCCPA1)

        With CreditCardInfo
            .TransArmorToken = String.Empty
            .RefundAmount = 0
            .InvoiceNumber = INV_NO
            .CaptureAmount = CCPA_AMT
            .invoiceHeaderRow = objCCProcessor.CustomerCreditCard.invoiceHeaderRow
            .invoiceDetailsTable = objCCProcessor.CustomerCreditCard.invoiceDetailsTable

            .Level2Data = objCCProcessor.CustomerCreditCard.Level2Data
            .Level3Data = objCCProcessor.CustomerCreditCard.Level3Data
        End With

        objCCProcessor.Capture(CreditCardInfo)

        Dim CCPA_NO_CAPTURE As String = String.Empty

        If objCCProcessor.NetworkResponse.Approved Then
            rowARTCCPA1.Item("CCPA_STATUS") = "C"
            ' Pass in a new CCPA_NO for the ARTCCPA* records
            Dim CCPA_NO_NEW As String = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
            CCPA_NO_CAPTURE = Create_CC_Capture_Entry(rowARTCCPA1, "", CCPA_NO_NEW)
        Else
            rowARTCCPA1.Item("CCPA_STATUS") = "B"

            ' Record Error trying to Capture Previously Authorized Sale
            Dim rowARTCCPA2 As DataRow = dst.Tables("ARTCCPA2").NewRow
            rowARTCCPA2.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
            rowARTCCPA2.Item("RESPONSE_TEXT") = objCCProcessor.NetworkResponse.Text
            If objCCProcessor.NetworkResponse.Approved Then
                rowARTCCPA2.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.Code
            Else
                rowARTCCPA2.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.ErrorCode
            End If
            rowARTCCPA2.Item("RESPONSE_APPROVAL_CODE") = objCCProcessor.NetworkResponse.ApprovalCode

            AdjustResponseCode(rowARTCCPA2)

            Dim RESPONSE_DATA As String = objCCProcessor.NetworkResponse.Data
            If RESPONSE_DATA.Length > rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength Then
                RESPONSE_DATA = RESPONSE_DATA.Substring(0, rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength).Trim
            End If
            rowARTCCPA2.Item("RESPONSE_DATA") = RESPONSE_DATA
            rowARTCCPA2.Item("RESPONSE_AVS") = objCCProcessor.NetworkResponse.AVSResult
            rowARTCCPA2.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            rowARTCCPA2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowARTCCPA2.Item("CCPA_TYPE") = "E"
            responseErrorMessage = objCCProcessor.NetworkResponse.Text
            dst.Tables("ARTCCPA2").Rows.Add(rowARTCCPA2)
        End If

        responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)

        BeginTrans()
        EncryptARTCCPA1()
        Update_Record_TDA("ARTCCPA1")
        Update_Record_TDA("ARTCCPA2")
        Update_Record_TDA("ARTCCPDA")
        CommitTrans()

        Return CCPA_NO_CAPTURE

    End Function

    Public Function Create_CC_Capture_Entry(ByRef rowARTCCPA1_AUTH As DataRow, ByVal EntryNote As String, ByVal CCPA_NO_CAPTURED As String) As String

        Dim LAST_DATE As Date = Now + ASCMAIN1.NowTSD

        Dim rowARTCCPA1_Capture As DataRow = dst.Tables("ARTCCPA1").NewRow
        rowARTCCPA1_Capture.Item("CCPA_NO") = CCPA_NO_CAPTURED

        rowARTCCPA1_AUTH.Item("CCPA_NO_CAPTURE") = CCPA_NO_CAPTURED
        rowARTCCPA1_AUTH.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_AUTH.Item("LAST_DATE") = LAST_DATE

        rowARTCCPA1_Capture.Item("CCPA_NO_AUTH") = rowARTCCPA1_AUTH.Item("CCPA_NO")

        rowARTCCPA1_Capture.Item("CUST_CODE") = rowARTCCPA1_AUTH.Item("CUST_CODE")

        If objCCProcessor.NetworkResponse.Approved Then
            rowARTCCPA1_Capture.Item("CCPA_STATUS") = "A"
        Else
            rowARTCCPA1_Capture.Item("CCPA_STATUS") = objCCProcessor.NetworkResponse.Code
        End If

        rowARTCCPA1_Capture.Item("CCPA_REASON") = "C"
        rowARTCCPA1_Capture.Item("CCPA_NOTE") = EntryNote
        rowARTCCPA1_Capture.Item("CCPA_AMT") = objCCProcessor.TransactionAmount
        rowARTCCPA1_Capture.Item("CCPA_DATE_AUTH") = rowARTCCPA1_AUTH.Item("CCPA_DATE_AUTH")
        rowARTCCPA1_Capture.Item("CCPA_DATE_SALE") = LAST_DATE
        'rowARTCCPA1.Item("CCPA_AUTH") = ""
        rowARTCCPA1_Capture.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_Capture.Item("INIT_DATE") = LAST_DATE
        rowARTCCPA1_Capture.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_Capture.Item("LAST_DATE") = LAST_DATE

        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_NO") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NO")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_EXP_DATE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_EXP_DATE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_VER_CODE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_VER_CODE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_NAME") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NAME")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_ADDR1") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ADDR1")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_CITY") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_CITY")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_STATE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_STATE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_ZIP_CODE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ZIP_CODE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_COUNTRY") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_COUNTRY")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_LAST4") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_LAST4")

        'rowARTCCPA1_Capture.Item("RESPONSE_RETRIEVAL_NO") = objCCProcessor.NetworkResponse.RetrievalNumber
        rowARTCCPA1_Capture.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.Code
        AdjustResponseCode(rowARTCCPA1_Capture)

        'rowARTCCPA1_Capture.Item("RESPONSE_BATCH_NO") = objCCProcessor.NetworkResponse.BatchNumber
        rowARTCCPA1_Capture.Item("RESPONSE_APPROVAL_CODE") = objCCProcessor.NetworkResponse.ApprovalCode
        rowARTCCPA1_Capture.Item("RESPONSE_TEXT") = objCCProcessor.NetworkResponse.Text
        responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)

        rowARTCCPA1_Capture.Item("CARD_LEVEL_RESULT") = rowARTCCPA1_AUTH.Item("CARD_LEVEL_RESULT")
        rowARTCCPA1_Capture.Item("DATAWIRE_RETURN_CODE") = rowARTCCPA1_AUTH.Item("DATAWIRE_RETURN_CODE")
        rowARTCCPA1_Capture.Item("DATAWIRE_STATUS") = rowARTCCPA1_AUTH.Item("DATAWIRE_STATUS")
        rowARTCCPA1_Capture.Item("ACI_CODE") = rowARTCCPA1_AUTH.Item("ACI_CODE")
        rowARTCCPA1_Capture.Item("TRANSACTION_DATE") = rowARTCCPA1_AUTH.Item("TRANSACTION_DATE")
        rowARTCCPA1_Capture.Item("TRANS_ID") = rowARTCCPA1_AUTH.Item("TRANS_ID")
        rowARTCCPA1_Capture.Item("TRANS_NUM") = objCCProcessor.TransactionNumber & String.Empty
        rowARTCCPA1_Capture.Item("VALIDATION_CODE") = rowARTCCPA1_AUTH.Item("VALIDATION_CODE")

        rowARTCCPA1_Capture.Item("CCPA_NOTE") = "Pre-Auth Sale"
        ' this needs to go into a separate column

        rowARTCCPA1_Capture.Item("CCPA_TYPE") = "S"
        rowARTCCPA1_Capture.Item("ORDR_NO") = rowARTCCPA1_AUTH.Item("ORDR_NO")
        rowARTCCPA1_Capture.Item("INV_NO") = rowARTCCPA1_AUTH.Item("INV_NO")

        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_TYPE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_TYPE")
        rowARTCCPA1_Capture.Item("OPS_YYYYPP") = ASCMAIN1.CYP

        dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1_Capture)

        Try
            Dim responseDollarAmount As Decimal = Val(objCCProcessor.NetworkResponse.ExtractNodeFromResponse("DollarAmount"))
            If responseDollarAmount > 0 AndAlso responseDollarAmount + 0.05 < Val(rowARTCCPA1_Capture.Item("CCPA_AMT") & String.Empty) Then
                rowARTCCPA1_Capture.Item("CCPA_AMT") = responseDollarAmount
            End If
        Catch ex As Exception

        End Try

        If rowARTCCPA1_AUTH.Item("ORDR_NO") & String.Empty <> String.Empty Then
            TAC.TACMAIN1.Record_Event("SOTORDR1", rowARTCCPA1_AUTH.Item("ORDR_NO"),
                                  rowARTCCPA1_AUTH.Item("LAST_DATE"),
                                  rowARTCCPA1_AUTH.Item("LAST_OPER"),
                                  "", "Sale Captured Inv " & rowARTCCPA1_AUTH.Item("INV_NO") & " " & Format(Val(rowARTCCPA1_Capture.Item("CCPA_AMT") & ""), "$#,##0.00"))
        End If


        Record_Level_2_3_Aggregates(rowARTCCPA1_AUTH.Item("CCPA_NO") & String.Empty, CCPA_NO_CAPTURED)
        Record_Response2(rowARTCCPA1_Capture)

        Return CCPA_NO_CAPTURED
    End Function

    Sub Record_Level_2_3_Aggregates(ByVal CCPA_NO_ORIG As String, ByVal CCPA_NO_CAPTURED As String)

        Dim rowARTCCPDA As DataRow = Nothing

        If CCPA_NO_ORIG <> CCPA_NO_CAPTURED Then
            rowARTCCPDA = dst.Tables("ARTCCPDA").NewRow()
            rowARTCCPDA.Item("CCPA_NO") = CCPA_NO_CAPTURED
            If dst.Tables("ARTCCPDA").Select("CCPA_NO = '" & CCPA_NO_ORIG & "'", "").Length > 0 Then
                rowARTCCPDA.Item("DETAIL_AGGREGATE") = dst.Tables("ARTCCPDA").Select("CCPA_NO = '" & CCPA_NO_ORIG & "'", "")(0).Item("DETAIL_AGGREGATE") & String.Empty
                rowARTCCPDA.Item("COMM_CARD_TYPE") = dst.Tables("ARTCCPDA").Select("CCPA_NO = '" & CCPA_NO_ORIG & "'", "")(0).Item("COMM_CARD_TYPE")
            Else
                Dim row As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPDA WHERE CCPA_NO = :PARM1", "V", New Object() {CCPA_NO_ORIG})
                If row IsNot Nothing Then
                    rowARTCCPDA.Item("DETAIL_AGGREGATE") = row.Item("DETAIL_AGGREGATE") & String.Empty
                    'rowARTCCPDA.Item("COMM_CARD_TYPE") = row.Item("COMM_CARD_TYPE") & String.Empty
                End If
            End If
            dst.Tables("ARTCCPDA").Rows.Add(rowARTCCPDA)

        ElseIf dst.Tables("ARTCCPDA").Select("CCPA_NO = '" & CCPA_NO_ORIG & "'", "").Length > 0 Then
            rowARTCCPDA = dst.Tables("ARTCCPDA").Select("CCPA_NO = '" & CCPA_NO_ORIG & "'", "")(0)
        Else
            Exit Sub
        End If

        If objCCProcessor.Level2Data IsNot Nothing Then
            If objCCProcessor.Level2Data.PONumber.Length = 0 Then
                objCCProcessor.Level2Data.PONumber = CCPA_NO_CAPTURED
            End If
            rowARTCCPDA.Item("LEVEL2_AGGREGATE") = objCCProcessor.GetLevel2Aggregate() & String.Empty
        End If

        If objCCProcessor.Level3Data IsNot Nothing AndAlso objCCProcessor.Level3Data.Count > 0 Then
            Dim LEVEL3_AGGREGATE As String = objCCProcessor.GetLevel3Aggregate & String.Empty

            Select Case LEVEL3_AGGREGATE.Length

                Case 0 To 3000
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE") = LEVEL3_AGGREGATE
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_2") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_3") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_4") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_5") = String.Empty

                Case 3001 To 6000
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE") = LEVEL3_AGGREGATE.Substring(0, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_2") = LEVEL3_AGGREGATE.Substring(3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_3") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_4") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_5") = String.Empty

                Case 6001 To 9000
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE") = LEVEL3_AGGREGATE.Substring(0, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_2") = LEVEL3_AGGREGATE.Substring(3000, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_3") = LEVEL3_AGGREGATE.Substring(6000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_4") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_5") = String.Empty

                Case 9001 To 12000
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE") = LEVEL3_AGGREGATE.Substring(0, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_2") = LEVEL3_AGGREGATE.Substring(3000, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_3") = LEVEL3_AGGREGATE.Substring(6000, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_4") = LEVEL3_AGGREGATE.Substring(9000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_5") = String.Empty

                Case 12001 To 15000
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE") = LEVEL3_AGGREGATE.Substring(0, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_2") = LEVEL3_AGGREGATE.Substring(3000, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_3") = LEVEL3_AGGREGATE.Substring(6000, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_4") = LEVEL3_AGGREGATE.Substring(9000, 3000)
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_5") = LEVEL3_AGGREGATE.Substring(12000)

                Case Else ' String to big. Forfeit Level 3 Pricing. Hopefully will never execute
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_2") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_3") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_4") = String.Empty
                    rowARTCCPDA.Item("LEVEL3_AGGREGATE_5") = String.Empty

            End Select

        End If

        ' Need to update the Settlement amount
        If rowARTCCPDA.Item("DETAIL_AGGREGATE") & String.Empty <> String.Empty Then
            Try
                Dim detailXml As New XmlDocument()
                detailXml.LoadXml(rowARTCCPDA.Item("DETAIL_AGGREGATE") & String.Empty)
                Dim detailXmlNode As XmlNode = detailXml.SelectSingleNode("/FDMSDetailAggregate/SettlementAmount")
                If detailXmlNode IsNot Nothing Then
                    detailXmlNode.ChildNodes(0).InnerText = Val(objCCProcessor.TransactionAmount * 100)
                    rowARTCCPDA.Item("DETAIL_AGGREGATE") = detailXml.InnerXml
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Sub Initalize_CCPA()

        Initialize_DataLayer()

        MerchantSetup()

        If Not dst.Tables.Contains("ARTCCPA1") Then
            dst = clsASCBASE1.dst
            Create_TDA(dst.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPDA", "*", 1)
            Create_TDA(dst.Tables.Add, "ARTCUSTC", "*", 1)
            dst.Tables("ARTCUSTC").Columns.Add("EXP", GetType(System.Int16))
            If rowARTCCPA1 Is Nothing Then
                rowARTCCPA1 = dst.Tables("ARTCCPA1").NewRow
            End If
        End If

        Dim CCPA_AMT As Decimal = Val(rowARTCCPA1.Item("CCPA_AMT") & "")

        Dim CCPA_NO As String = String.Empty

        ' Monthly Auto Queue, do not create a new Credit Card Entry
        If rowARTCCPA1.Item("CCPA_STATUS") & String.Empty = "2" Then
            CCPA_NO = rowARTCCPA1.Item("CCPA_NO") & String.Empty
        Else
            CCPA_NO = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
            rowARTCCPA1.Item("CCPA_NO") = CCPA_NO
        End If
        Dim INIT_DATE As Date = Now + ASCMAIN1.NowTSD

        rowARTCCPA1.Item("ORDR_NO") = ORDR_NO
        rowARTCCPA1.Item("CCPA_TYPE") = TRAN_TYPE

        If Not IsDate(rowARTCCPA1.Item("INIT_DATE") & String.Empty) Then
            rowARTCCPA1.Item("INIT_DATE") = INIT_DATE
        End If

        If rowARTCCPA1.Item("INIT_OPER") & String.Empty = String.Empty Then
            rowARTCCPA1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        End If

        rowARTCCPA1.Item("LAST_DATE") = INIT_DATE
        rowARTCCPA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1.Item("OPS_YYYYPP") = ASCMAIN1.CYP

        Dim CUST_CREDIT_CARD_NO As String = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty
        If CUST_CREDIT_CARD_NO.Length >= 4 Then
            rowARTCCPA1.Item("CUST_CREDIT_CARD_LAST4") = StrReverse(StrReverse(CUST_CREDIT_CARD_NO).Substring(0, 4))
        End If

        Try
            objCCProcessor.CustomerCreditCard = CreateCreditCardInfo(rowARTCCPA1)
            'objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
            'objCCProcessor.CustomerCreditCard.CardExpMonth = Mid(MMYY, 1, 2)
            'objCCProcessor.CustomerCreditCard.CardExpYear = Mid(MMYY, 3, 2)
            objCCProcessor.ValidateCard()
        Catch ex As Exception
            'Stop
        End Try

        Dim CUST_CREDIT_CARD_TYPE As String = String.Empty
        Select Case objCCProcessor.CustomerCreditCard.CardType
            Case TAC.ARCCCARD.CreditCardTypes.vctAmex
                CUST_CREDIT_CARD_TYPE = "AMEX"
            Case TAC.ARCCCARD.CreditCardTypes.vctMasterCard
                CUST_CREDIT_CARD_TYPE = "MSTR"
            Case TAC.ARCCCARD.CreditCardTypes.vctDiscover
                CUST_CREDIT_CARD_TYPE = "DISC"
            Case TAC.ARCCCARD.CreditCardTypes.vctVisa, ARCCCARD.CreditCardTypes.vctVisaElectron
                CUST_CREDIT_CARD_TYPE = "VISA"
            Case TAC.ARCCCARD.CreditCardTypes.vctDiners
                CUST_CREDIT_CARD_TYPE = "DC"
            Case Else
                CUST_CREDIT_CARD_TYPE = String.Empty
        End Select

        rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = CUST_CREDIT_CARD_TYPE

        If dst.Tables("ARTCCPA1").Select("CCPA_NO = '" & rowARTCCPA1.Item("CCPA_NO") & "'").Length = 0 Then
            If rowARTCCPA1.Table.TableName = "ARTCCPA1" Then
                dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1)
            Else
                dst.Tables("ARTCCPA1").ImportRow(rowARTCCPA1)
            End If
        End If

        Prepare_Component(CCPA_AMT)
    End Sub

    Sub Prepare_Component(ByVal CCPA_AMT As Decimal)

        objCCProcessor.CustomerCreditCard.CardNumber = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO")
        Dim MMYY As String = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE")
        objCCProcessor.CustomerCreditCard.CardExpMonth = Mid(MMYY, 1, 2)
        objCCProcessor.CustomerCreditCard.CardExpYear = Mid(MMYY, 3, 2)
        objCCProcessor.TransactionAmount = Format(CCPA_AMT, "#.00")
        objCCProcessor.CustomerCreditCard.CardHolderAddress = rowARTCCPA1.Item("CUST_CREDIT_CARD_ADDR1") & ""

        Dim ZIP_CODE As String = Replace(rowARTCCPA1.Item("CUST_CREDIT_CARD_ZIP_CODE") & "", "-", "") & String.Empty
        'If ZIP_CODE.Length <> 5 And ZIP_CODE.Length <> 9 Then
        '    ZIP_CODE = ""
        'End If
        objCCProcessor.CustomerCreditCard.CardHolderZipCode = ZIP_CODE
    End Sub

    Public Function CC_Authorize(ByVal BeginCommit) As String

        Initialize_DataLayer() ' Instantiation of this form does not call Form_Load

        If Not dst.Tables.Contains("ARTCCPA1") Then
            dst = clsASCBASE1.dst
            Create_TDA(dst.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPDA", "*", 1)
            Create_TDA(dst.Tables.Add, "ARTCUSTC", "*", 1)
            dst.Tables("ARTCUSTC").Columns.Add("EXP", GetType(System.Int16))
            If rowARTCCPA1 Is Nothing Then
                rowARTCCPA1 = dst.Tables("ARTCCPA1").NewRow
            End If
        End If

        Initalize_CCPA()
        CheckTestMode()
        If objCCProcessor.TransactionNumber & String.Empty = String.Empty Then
            objCCProcessor.TransactionNumber = ASCMAIN1.Next_Control_No("ARTCCPA1.TRANS_NUM")
        End If
        objCCProcessor.AuthOnly()
        Record_Response()
        Record_DetailAggregate()
        Record_Audit(BeginCommit, True)

        CCPA_NO = rowARTCCPA1.Item("CCPA_NO")
        If BeginCommit Then
            Try
                BeginTrans()
                If dst.Tables("ARTCCPA1").Rows.Count = 0 Then
                    If rowARTCCPA1 IsNot Nothing Then
                        Dim rowartccpa1x As DataRow = dst.Tables("ARTCCPA1").NewRow
                        For Each dataCol As DataColumn In dst.Tables("ARTCCPA1").Columns
                            rowartccpa1x.Item(dataCol.ColumnName) = rowARTCCPA1.Item(dataCol.ColumnName)
                        Next
                        dst.Tables("ARTCCPA1").Rows.Add(rowartccpa1x)
                    End If
                End If
                EncryptARTCCPA1()
                Update_Record_TDA("ARTCCPA1")
                Update_Record_TDA("ARTCCPA2")
                Update_Record_TDA("ARTCCPDA")
                CommitTrans()
            Catch ex As Exception
                Rollback()
            End Try
        End If

        Return objCCProcessor.NetworkResponse.ApprovalCode

    End Function

    Public Function CC_Sale(ByVal CCPA_AMT As Decimal) As String

        ' CC_Capture, CC_Authorize & CC_Sale are called by SH modules which instantiate this form as a class
        ' when this happens, Form Load does not fire 
        '  (and it is good that it does not because there is a lot of UI going on in Form Load)
        ' A notable difference between CC_Capture and the others is that CC_Capture does not update Oracle,
        '  whereas CC_Authorize and CC_Sale do
        ' However, the datalayer that is used needs to (still) be the calling form's datalayer,
        '  since Form Load is not called - which is where we Create_TDA
        ' So these routines will continue to use the calling form's datalayer 
        '  until we come up with a more elegant solution

        Initialize_DataLayer() ' Instantiation of this form does not call Form_Load

        ' This code used to be in the Create_CC_Capture routine
        ' however, after the credit card is processed: ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
        ' causes / in a lock the credit card trancasction does not get recorded
        ' This happened on 7/23/2010 causing issues settling credit cards and payments against the customer

        If Not dst.Tables.Contains("ARTCCPA1") Then
            dst = clsASCBASE1.dst
            Create_TDA(dst.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPDA", "*", 1)
            Create_TDA(dst.Tables.Add, "ARTCUSTC", "*", 1)
            dst.Tables("ARTCUSTC").Columns.Add("EXP", GetType(System.Int16))
            If rowARTCCPA1 Is Nothing Then
                rowARTCCPA1 = dst.Tables("ARTCCPA1").NewRow
            End If
        End If

        MerchantSetup()

        If ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & "" = "1" Then
            test_mode = True
        End If

        Prepare_Component(CCPA_AMT)
        CheckTestMode()

        Dim CreditCardInfo As New TAC.ARCCCARD.CreditCard
        CreditCardInfo = CreateCreditCardInfo(rowARTCCPA1)

        With CreditCardInfo
            .TransArmorToken = String.Empty
            .RefundAmount = 0
            .InvoiceNumber = INV_NO
            .CaptureAmount = CCPA_AMT
            .invoiceHeaderRow = objCCProcessor.CustomerCreditCard.invoiceHeaderRow
            .invoiceDetailsTable = objCCProcessor.CustomerCreditCard.invoiceDetailsTable

            .Level2Data = objCCProcessor.CustomerCreditCard.Level2Data
            .Level3Data = objCCProcessor.CustomerCreditCard.Level3Data
        End With

        objCCProcessor.CustomerCreditCard = CreditCardInfo

        objCCProcessor.Sale()

        Dim CCPA_NO_SALE As String = String.Empty

        If objCCProcessor.NetworkResponse Is Nothing Then
            CCPA_NO_SALE = String.Empty
            rowARTCCPA1.Item("CCPA_STATUS") = "B"

            ' Record Error trying to Capture Previously Authorized Sale
            Dim rowARTCCPA2 As DataRow = dst.Tables("ARTCCPA2").NewRow
            rowARTCCPA2.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
            rowARTCCPA2.Item("RESPONSE_TEXT") = objCCProcessor.LastError
            rowARTCCPA2.Item("RESPONSE_CODE") = "E"
            rowARTCCPA2.Item("RESPONSE_APPROVAL_CODE") = ""

            Dim RESPONSE_DATA As String = objCCProcessor.LastError
            If RESPONSE_DATA.Length > rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength Then
                RESPONSE_DATA = RESPONSE_DATA.Substring(0, rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength).Trim
            End If
            rowARTCCPA2.Item("RESPONSE_DATA") = RESPONSE_DATA

            rowARTCCPA2.Item("RESPONSE_AVS") = String.Empty
            rowARTCCPA2.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            rowARTCCPA2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowARTCCPA2.Item("CCPA_TYPE") = "E"
            dst.Tables("ARTCCPA2").Rows.Add(rowARTCCPA2)

        ElseIf objCCProcessor.NetworkResponse.Approved Then
            rowARTCCPA1.Item("CCPA_STATUS") = "S"
            ' Pass in a new CCPA_NO for the ARTCCPA* records
            Dim CCPA_NO_NEW As String = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
            CCPA_NO_SALE = Create_CC_Capture_Entry(rowARTCCPA1, "CC Payment", CCPA_NO_NEW)

            ' Update the fields for a sale
            Dim rowARTCCPA1_SALE As DataRow = dst.Tables("ARTCCPA1").Select("CCPA_NO = '" & CCPA_NO_SALE & "'")(0)
            rowARTCCPA1_SALE.Item("CCPA_NO_AUTH") = String.Empty
            rowARTCCPA1_SALE.Item("CCPA_AUTH") = String.Empty
            rowARTCCPA1_SALE.Item("CCPA_REASON") = "C"
            rowARTCCPA1_SALE.Item("CCPA_DATE_AUTH") = DBNull.Value
            rowARTCCPA1_SALE.Item("CCPA_NOTE") = "CC Payment"
        Else
            rowARTCCPA1.Item("CCPA_STATUS") = "B"

            ' Record Error trying to Capture Previously Authorized Sale
            Dim rowARTCCPA2 As DataRow = dst.Tables("ARTCCPA2").NewRow
            rowARTCCPA2.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
            rowARTCCPA2.Item("RESPONSE_TEXT") = objCCProcessor.NetworkResponse.Text
            If objCCProcessor.NetworkResponse.Approved Then
                rowARTCCPA2.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.Code
            Else
                rowARTCCPA2.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.ErrorCode
            End If
            rowARTCCPA2.Item("RESPONSE_APPROVAL_CODE") = objCCProcessor.NetworkResponse.ApprovalCode

            AdjustResponseCode(rowARTCCPA2)

            Dim RESPONSE_DATA As String = objCCProcessor.NetworkResponse.Data
            If RESPONSE_DATA.Length > rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength Then
                RESPONSE_DATA = RESPONSE_DATA.Substring(0, rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength).Trim
            End If
            rowARTCCPA2.Item("RESPONSE_DATA") = RESPONSE_DATA

            rowARTCCPA2.Item("RESPONSE_AVS") = objCCProcessor.NetworkResponse.AVSResult
            rowARTCCPA2.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            rowARTCCPA2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowARTCCPA2.Item("CCPA_TYPE") = "E"
            dst.Tables("ARTCCPA2").Rows.Add(rowARTCCPA2)
        End If

        responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)

        BeginTrans()
        EncryptARTCCPA1()
        Update_Record_TDA("ARTCCPA1")
        Update_Record_TDA("ARTCCPA2")
        Update_Record_TDA("ARTCCPDA")
        CommitTrans()

        Return CCPA_NO_SALE

    End Function


    ''' <summary>
    ''' Issue Credit Card refund/Credit
    ''' </summary>
    ''' <param name="CreditAmount">Credit Amount</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CC_Credit(ByVal CreditAmount As Decimal) As String

        Initialize_DataLayer() ' Instantiation of this form does not call Form_Load

        If Not dst.Tables.Contains("ARTCCPA1") Then
            dst = clsASCBASE1.dst
            Create_TDA(dst.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(dst.Tables.Add, "ARTCCPDA", "*", 1)
            Create_TDA(dst.Tables.Add, "ARTCUSTC", "*", 1)
            dst.Tables("ARTCUSTC").Columns.Add("EXP", GetType(System.Int16))
            If rowARTCCPA1 Is Nothing Then
                rowARTCCPA1 = dst.Tables("ARTCCPA1").NewRow
            End If

            If rowARTCCPA1.Item("CUST_CODE") & String.Empty = String.Empty Then
                rowARTCCPA1.Item("CUST_CODE") = CUST_CODE
            End If
        End If

        MerchantSetup()

        If ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & "" = "1" Then
            test_mode = True
        End If

        Dim CreditCardInfo As New TAC.ARCCCARD.CreditCard
        CreditCardInfo = CreateCreditCardInfo(rowARTCCPA1)

        With CreditCardInfo
            .InvoiceNumber = String.Empty
            .TransArmorToken = String.Empty
            .RefundAmount = CreditAmount
        End With

        CheckTestMode()
        objCCProcessor.Refund(CreditCardInfo)

        Dim CCPA_NO_CREDIT As String = String.Empty

        If objCCProcessor.NetworkResponse.Approved Then
            rowARTCCPA1.Item("CCPA_STATUS") = "S"
            ' Pass in a new CCPA_NO for the ARTCCPA* records
            Dim CCPA_NO_NEW As String = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
            CCPA_NO_CREDIT = Create_CC_Credit_Entry(rowARTCCPA1, "CC Credit", CCPA_NO_NEW)
            MerchantTransID = objCCProcessor.NetworkResponse.TransactionId
        Else
            responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)
            Return String.Empty
        End If

        BeginTrans()
        EncryptARTCCPA1()
        Update_Record_TDA("ARTCCPA1")
        Update_Record_TDA("ARTCCPA2")
        Update_Record_TDA("ARTCCPDA")
        CommitTrans()

        Return CCPA_NO_CREDIT

    End Function

    Public Function Create_CC_Credit_Entry(ByRef rowARTCCPA1_AUTH As DataRow, ByVal EntryNote As String, ByVal CCPA_NO_CREDIT As String) As String

        Dim LAST_DATE As Date = Now + ASCMAIN1.NowTSD

        Dim rowARTCCPA1_Capture As DataRow = dst.Tables("ARTCCPA1").NewRow
        rowARTCCPA1_Capture.Item("CCPA_NO") = CCPA_NO_CREDIT

        rowARTCCPA1_AUTH.Item("CCPA_NO_CAPTURE") = CCPA_NO_CREDIT
        rowARTCCPA1_AUTH.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_AUTH.Item("LAST_DATE") = LAST_DATE

        rowARTCCPA1_Capture.Item("CCPA_NO_AUTH") = String.Empty

        rowARTCCPA1_Capture.Item("CUST_CODE") = rowARTCCPA1_AUTH.Item("CUST_CODE")

        If objCCProcessor.NetworkResponse.Approved Then
            rowARTCCPA1_Capture.Item("CCPA_STATUS") = "A"
        Else
            rowARTCCPA1_Capture.Item("CCPA_STATUS") = objCCProcessor.NetworkResponse.Code
        End If

        rowARTCCPA1_Capture.Item("CCPA_REASON") = "M"
        rowARTCCPA1_Capture.Item("CCPA_NOTE") = EntryNote
        rowARTCCPA1_Capture.Item("CCPA_AMT") = objCCProcessor.TransactionAmount
        rowARTCCPA1_Capture.Item("CCPA_DATE_AUTH") = DBNull.Value
        'rowARTCCPA1_Capture.Item("CCPA_DATE_SALE") = LAST_DATE
        rowARTCCPA1_Capture.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_Capture.Item("INIT_DATE") = LAST_DATE
        rowARTCCPA1_Capture.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_Capture.Item("LAST_DATE") = LAST_DATE

        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_NO") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NO")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_EXP_DATE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_EXP_DATE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_VER_CODE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_VER_CODE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_NAME") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NAME")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_ADDR1") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ADDR1")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_CITY") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_CITY")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_STATE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_STATE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_ZIP_CODE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ZIP_CODE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_COUNTRY") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_COUNTRY")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_LAST4") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_LAST4")

        If objCCProcessor.NetworkResponse.Approved Then
            rowARTCCPA1.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.Code
        Else
            rowARTCCPA1.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.ErrorCode
        End If

        AdjustResponseCode(rowARTCCPA1)

        'rowARTCCPA1_Capture.Item("RESPONSE_BATCH_NO") = objCCProcessor.NetworkResponse.BatchNumber
        rowARTCCPA1_Capture.Item("RESPONSE_APPROVAL_CODE") = objCCProcessor.NetworkResponse.ApprovalCode
        rowARTCCPA1_Capture.Item("RESPONSE_TEXT") = objCCProcessor.NetworkResponse.Text
        responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)

        rowARTCCPA1_Capture.Item("CARD_LEVEL_RESULT") = rowARTCCPA1_AUTH.Item("CARD_LEVEL_RESULT")
        rowARTCCPA1_Capture.Item("DATAWIRE_RETURN_CODE") = rowARTCCPA1_AUTH.Item("DATAWIRE_RETURN_CODE")
        rowARTCCPA1_Capture.Item("DATAWIRE_STATUS") = rowARTCCPA1_AUTH.Item("DATAWIRE_STATUS")
        rowARTCCPA1_Capture.Item("ACI_CODE") = rowARTCCPA1_AUTH.Item("ACI_CODE")
        rowARTCCPA1_Capture.Item("TRANSACTION_DATE") = DateTime.Now.ToString("MMddyy")
        rowARTCCPA1_Capture.Item("TRANS_ID") = objCCProcessor.NetworkResponse.TransactionId
        rowARTCCPA1_Capture.Item("TRANS_NUM") = objCCProcessor.TransactionNumber & String.Empty
        rowARTCCPA1_Capture.Item("VALIDATION_CODE") = rowARTCCPA1_AUTH.Item("VALIDATION_CODE")

        rowARTCCPA1_Capture.Item("CCPA_TYPE") = "C"
        rowARTCCPA1_Capture.Item("ORDR_NO") = rowARTCCPA1_AUTH.Item("ORDR_NO")
        rowARTCCPA1_Capture.Item("INV_NO") = rowARTCCPA1_AUTH.Item("INV_NO")

        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_TYPE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_TYPE")
        rowARTCCPA1_Capture.Item("OPS_YYYYPP") = ASCMAIN1.CYP

        dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1_Capture)


        If rowARTCCPA1_AUTH.Item("ORDR_NO") & String.Empty <> String.Empty Then
            TAC.TACMAIN1.Record_Event("SOTORDR1", rowARTCCPA1_AUTH.Item("ORDR_NO"),
                                  rowARTCCPA1_AUTH.Item("LAST_DATE"),
                                  rowARTCCPA1_AUTH.Item("LAST_OPER"),
                                  "", "Credit Issued (" & rowARTCCPA1_AUTH.Item("INV_NO") & ") for amount of " & Format(Val(rowARTCCPA1_Capture.Item("CCPA_AMT") & ""), "$#,##0.00"))
        End If


        'Record_Level_2_3_Aggregates(rowARTCCPA1_AUTH.Item("CCPA_NO") & String.Empty, CCPA_NO_CREDIT)
        Record_Response2(rowARTCCPA1_Capture)

        Return CCPA_NO_CREDIT
    End Function


    Sub Record_Response(Optional ByVal RecordResponse2 As Boolean = True)

        If objCCProcessor.NetworkResponse.Approved Then
            rowARTCCPA1.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.Code
        Else
            rowARTCCPA1.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.ErrorCode
        End If

        AdjustResponseCode(rowARTCCPA1)

        rowARTCCPA1.Item("RESPONSE_APPROVAL_CODE") = objCCProcessor.NetworkResponse.ApprovalCode
        rowARTCCPA1.Item("RESPONSE_TEXT") = objCCProcessor.NetworkResponse.Text
        responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)
        rowARTCCPA1.Item("TRANSACTION_DATE") = DateTime.Now.ToString("MMddyy")
        rowARTCCPA1.Item("TRANS_ID") = objCCProcessor.NetworkResponse.TransactionId
        MerchantTransID = objCCProcessor.NetworkResponse.TransactionId
        rowARTCCPA1.Item("TRANS_NUM") = objCCProcessor.TransactionNumber

        If RecordResponse2 Then Record_Response2(rowARTCCPA1)
    End Sub

    Sub Record_Response2(ByVal rowARTCCPA1 As DataRow)
        Dim rowARTCCPA2 As DataRow = dst.Tables("ARTCCPA2").NewRow
        rowARTCCPA2.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
        rowARTCCPA2.Item("RESPONSE_TEXT") = objCCProcessor.NetworkResponse.Text
        responseErrorMessage = IIf(objCCProcessor.NetworkResponse.Text.Length > 0, objCCProcessor.NetworkResponse.Text, objCCProcessor.NetworkResponse.Data)
        If objCCProcessor.NetworkResponse.Approved Then
            rowARTCCPA1.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.Code
            rowARTCCPA2.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.Code
        Else
            rowARTCCPA1.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.ErrorCode
            rowARTCCPA2.Item("RESPONSE_CODE") = objCCProcessor.NetworkResponse.ErrorCode
        End If
        rowARTCCPA2.Item("RESPONSE_APPROVAL_CODE") = objCCProcessor.NetworkResponse.ApprovalCode

        AdjustResponseCode(rowARTCCPA1)
        AdjustResponseCode(rowARTCCPA2)

        Dim RESPONSE_DATA As String = objCCProcessor.NetworkResponse.Data
        If RESPONSE_DATA.Length > rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength Then
            RESPONSE_DATA = RESPONSE_DATA.Substring(0, rowARTCCPA2.Table.Columns("RESPONSE_DATA").MaxLength).Trim
        End If
        rowARTCCPA2.Item("RESPONSE_DATA") = RESPONSE_DATA

        rowARTCCPA2.Item("RESPONSE_AVS") = objCCProcessor.NetworkResponse.AVSResult
        rowARTCCPA2.Item("INIT_DATE") = rowARTCCPA1.Item("INIT_DATE")
        rowARTCCPA2.Item("INIT_OPER") = rowARTCCPA1.Item("INIT_OPER")
        rowARTCCPA2.Item("CCPA_TYPE") = rowARTCCPA1.Item("CCPA_TYPE")
        dst.Tables("ARTCCPA2").Rows.Add(rowARTCCPA2)
        MerchantTransID = objCCProcessor.NetworkResponse.TransactionId
    End Sub

    Sub Record_DetailAggregate()
        Dim CCPA_NO As String = rowARTCCPA1.Item("CCPA_NO") & String.Empty
        Dim rowARTCCPDA As DataRow

        If dst.Tables("ARTCCPDA").Select("CCPA_NO = '" & CCPA_NO & "'").Length = 0 Then
            rowARTCCPDA = dst.Tables("ARTCCPDA").NewRow

            rowARTCCPDA.Item("CCPA_NO") = CCPA_NO
            dst.Tables("ARTCCPDA").Rows.Add(rowARTCCPDA)
        Else
            rowARTCCPDA = dst.Tables("ARTCCPDA").Select("CCPA_NO = '" & CCPA_NO & "'")(0)
        End If

        If objCCProcessor.NetworkResponse.DetailAggregate.Length > 0 AndAlso (rowARTCCPDA.Item("DETAIL_AGGREGATE") & String.Empty).ToString.Length = 0 Then
            rowARTCCPDA.Item("DETAIL_AGGREGATE") = objCCProcessor.NetworkResponse.DetailAggregate
            rowARTCCPDA.Item("COMM_CARD_TYPE") = "" 'objCCProcessor.NetworkResponse.CommercialCardType
        End If

    End Sub

    Sub Record_Audit(ByVal BeginCommit As Boolean, Optional ByVal UpdateDetailAggregate As Boolean = False)

        If BeginCommit Then
            BeginTrans()
        End If

        Dim EVENT_DESC As String = ""
        If objCCProcessor.NetworkResponse.Approved Then

            If rowARTCCPA1.Item("CCPA_TYPE") = "S" Then
                rowARTCCPA1.Item("CCPA_DATE_SALE") = rowARTCCPA1.Item("LAST_DATE")
                rowARTCCPA1.Item("CCPA_STATUS") = "A"
                EVENT_DESC = "CC Sale Appr "
            ElseIf rowARTCCPA1.Item("CCPA_TYPE") = "A" Then
                rowARTCCPA1.Item("CCPA_DATE_AUTH") = rowARTCCPA1.Item("LAST_DATE")
                rowARTCCPA1.Item("CCPA_STATUS") = "T"
                EVENT_DESC = "CC Auth Appr "
            End If

        Else
            If rowARTCCPA1.Item("CCPA_TYPE") = "S" Then
                rowARTCCPA1.Item("CCPA_STATUS") = "E"
                EVENT_DESC = "CC Sale Decl "
            ElseIf rowARTCCPA1.Item("CCPA_TYPE") = "A" Then
                rowARTCCPA1.Item("CCPA_STATUS") = "E"
                EVENT_DESC = "CC Auth Decl "
            End If
        End If

        If ORDR_NO <> "" Then
            TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_NO,
                                  rowARTCCPA1.Item("LAST_DATE"),
                                  rowARTCCPA1.Item("LAST_OPER"),
                                  "", EVENT_DESC & Format(Val(rowARTCCPA1.Item("CCPA_AMT") & ""), "$#,##0.00"))
        End If

        If BeginCommit Then
            CommitTrans()
        End If

    End Sub

    Sub Initialize_DataLayer()

        ' this will point dst and ROWs back to the calling form's datalayer

        ' This is necessary because:
        ' 1) sometimes the calling form needs to look at the data created in ARTCCPA1 (see SHFORDRX, which updates oracle in its own update routines)
        ' 2) sometimed the calling form handles the update to oracle (see Web Order import, which checks for Auth Declines)

        'ROWs = frmASFBASE1.ROWs
        'dst = frmASFBASE1.dst

        Dim rowASTPARMP As DataRow = ASCDATA1.GetDataRow("Select * from ASTPARMP WHERE AS_PARM_KEY = 'Z'")

        clsTACENCRY = New TAC.ASCENCRY()
        If rowASTPARMP Is Nothing OrElse Not rowASTPARMP.Table.Columns.Contains("AS_PARM_USE_ENCRYPTION") OrElse rowASTPARMP.Item("AS_PARM_USE_ENCRYPTION") & String.Empty <> "1" Then
            clsTACENCRY.UseEncryption = False
        Else
            clsTACENCRY.UseEncryption = True
        End If

    End Sub

    Private Sub SetProcessingType()

        If ROWs Is Nothing Then
            ROWs = New Dictionary(Of String, DataRow)
        End If

        If Not ROWs.ContainsKey("SOTPARM1") Then
            Get_PARM("SOTPARM1")
        End If

        Dim CC_PROC_FOLDER As String = String.Empty
        Dim SO_PARM_CC_PROC_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_CC_PROC_CODE") & String.Empty

        Dim rowARTCCPRC As DataRow = LookUp("ARTCCPRC", ROWs("SOTPARM1").Item("SO_PARM_CC_PROC_CODE") & String.Empty)

        If rowARTCCPRC IsNot Nothing Then
            CC_PROC_FOLDER = rowARTCCPRC.Item("CC_PROC_FOLDER") & String.Empty
        End If

        objCCProcessor = New TAC.ARCCCARD(CC_PROC_FOLDER, Val(SO_PARM_CC_PROC_CODE))

    End Sub

    Private Sub numCCPA_AMT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles numCCPA_AMT.GotFocus
        numCCPA_AMT.Editor.SelectAll()
    End Sub

    Private Sub SelectPreviouslyUsed()

        Try

            ASCMAIN1.CodeSelector.SQL = $"SELECT CUST_CREDIT_CARD_LAST4 LAST_4, CUST_CREDIT_CARD_EXP_DATE EXP_DATE, CUST_CREDIT_CARD_NAME CARD_NAME, MAX(CCPA_NO) CCPA_NO
                                            FROM ARTCCPA1
                                            WHERE CUST_CODE = '{CUST_CODE}'
                                            GROUP BY CUST_CREDIT_CARD_LAST4, CUST_CREDIT_CARD_EXP_DATE, CUST_CREDIT_CARD_NAME
                                            ORDER BY CUST_CREDIT_CARD_LAST4"


            Using F As New ASFCODE1
                ASCMAIN1.CodeSelector.VIEW_NAME = String.Empty
                ASCMAIN1.CodeSelector.TABLE_NAME = String.Empty
                ASCMAIN1.CodeSelector.Custom_sql_where = String.Empty
                ASCMAIN1.CodeSelector.Custom_sqlkey = String.Empty
                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                ASCMAIN1.CodeSelector.VIEW_DESC = "Previously Used Credit Cards"
                F.ShowDialog()
            End Using

            If ASCMAIN1.CodeSelector.SelectedCodes.Count = 0 Then
                Exit Sub
            End If

            Dim sql As String = $"Select * From ARTCCPA1 WHERE CCPA_NO = '{ASCMAIN1.CodeSelector.SelectedRows(0).Item("CCPA_NO") & String.Empty}'"
            Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow(sql)
            Dim rowARTCUSTC As DataRow = dst.Tables("ARTCUSTC").NewRow
            rowARTCUSTC.Item("CUST_CODE") = rowARTCCPA1.Item("CUST_CODE")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_KEY") = rowARTCCPA1.Item("CUST_CREDIT_CARD_KEY")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_EXP_DATE") = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_VER_CODE") = rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_NAME") = rowARTCCPA1.Item("CUST_CREDIT_CARD_NAME")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_ADDR1") = rowARTCCPA1.Item("CUST_CREDIT_CARD_ADDR1")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_CITY") = rowARTCCPA1.Item("CUST_CREDIT_CARD_CITY")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_STATE") = rowARTCCPA1.Item("CUST_CREDIT_CARD_STATE")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_ZIP_CODE") = rowARTCCPA1.Item("CUST_CREDIT_CARD_ZIP_CODE")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_COUNTRY") = rowARTCCPA1.Item("CUST_CREDIT_CARD_COUNTRY")
            'rowARTCUSTC.Item("CUST_CREDIT_CARD_STATUS") = rowARTCCPA1.Item("CUST_CREDIT_CARD_STATUS")
            'rowARTCUSTC.Item("INIT_OPER") = rowARTCCPA1.Item("INIT_OPER")
            'rowARTCUSTC.Item("INIT_DATE") = rowARTCCPA1.Item("INIT_DATE")
            'rowARTCUSTC.Item("LAST_OPER") = rowARTCCPA1.Item("LAST_OPER")
            'rowARTCUSTC.Item("LAST_DATE") = rowARTCCPA1.Item("LAST_DATE")
            'rowARTCUSTC.Item("CUST_CREDIT_CARD_PREFERRED") = rowARTCCPA1.Item("CUST_CREDIT_CARD_PREFERRED")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_LAST4") = rowARTCCPA1.Item("CUST_CREDIT_CARD_LAST4")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_NO") = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_NO_E") = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO_E")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_EXP_DATE_E") = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE_E")
            rowARTCUSTC.Item("CUST_CREDIT_CARD_VER_CODE_E") = rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE_E")
            rowARTCUSTC.Item("TRANSARMORTOKEN") = rowARTCCPA1.Item("TRANSARMORTOKEN")

            DecryptARTCUSTCDatarow(rowARTCUSTC)
            LoadDataFromARTCUSTC(rowARTCUSTC)
            Try
                Absx1.txtFor("CUST_CREDIT_CARD_NO").Text = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & ""
                txt_ValueChanged(MyBase.Absx1.txtFor("CUST_CREDIT_CARD_NO"), Nothing)
            Catch ex As Exception

            End Try

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Select Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DecryptARTCCPA1()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCCPA1 As DataRow In dst.Tables("ARTCCPA1").Rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                rowARTCCPA1.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1.Item(field & "_E") & String.Empty)
                rowARTCCPA1.Item(field & "_E") = DBNull.Value
            Next
        Next
    End Sub

    Private Sub DecryptARTCCPA1Datarow(ByRef rowARTCCPA1 As DataRow)
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} '  "CUST_CREDIT_CARD_EXP_DATE",
            rowARTCCPA1.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1.Item(field & "_E") & String.Empty)
            rowARTCCPA1.Item(field & "_E") = DBNull.Value
        Next
    End Sub

    Private Sub EncryptARTCCPA1()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCCPA1 As DataRow In dst.Tables("ARTCCPA1").Rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                rowARTCCPA1.Item(field & "_E") = clsTACENCRY.EncryptString(rowARTCCPA1.Item(field) & String.Empty)
                rowARTCCPA1.Item(field) = DBNull.Value
            Next
        Next
    End Sub

    Private Sub DecryptARTCUSTC()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCUSTC As DataRow In dst.Tables("ARTCUSTC").Rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} '  "CUST_CREDIT_CARD_EXP_DATE",
                rowARTCUSTC.Item(field) = clsTACENCRY.DecryptString(rowARTCUSTC.Item(field & "_E") & String.Empty)
                rowARTCUSTC.Item(field & "_E") = DBNull.Value
            Next
        Next
    End Sub

    Private Sub DecryptARTCUSTCDatarow(ByRef rowARTCUSTC As DataRow)
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} '  "CUST_CREDIT_CARD_EXP_DATE",
            rowARTCUSTC.Item(field) = clsTACENCRY.DecryptString(rowARTCUSTC.Item(field & "_E") & String.Empty)
            rowARTCUSTC.Item(field & "_E") = DBNull.Value
        Next
    End Sub

    Private Sub EncryptARTCUSTC()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCUSTC As DataRow In dst.Tables("ARTCUSTC").Rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                rowARTCUSTC.Item(field & "_E") = clsTACENCRY.EncryptString(rowARTCUSTC.Item(field) & String.Empty)
                rowARTCUSTC.Item(field) = DBNull.Value
            Next
        Next
    End Sub

    Private Sub UltraButton1_Click(sender As System.Object, e As System.EventArgs) Handles UltraButton1.Click

        Dim sql As String = String.Empty

        If Not ASCMAIN1.Running_in_VS Then
            Exit Sub
        End If

        Stop

        'If 1 = 1 Then Exit Sub

        Try
            BeginTrans()

            sql = "Select * from ARTCCPA1 WHERE CUST_CREDIT_CARD_NO IS NOT NULL"
            Fill_Records("ARTCCPA1", String.Empty, True, sql)
            EncryptARTCCPA1()

            sql = "Select * from ARTCUSTC WHERE CUST_CREDIT_CARD_NO IS NOT NULL"
            Fill_Records("ARTCUSTC", String.Empty, True, sql)
            EncryptARTCUSTC()

            Update_Record_TDA("ARTCCPA1")
            Update_Record_TDA("ARTCUSTC")

            CommitTrans("Data Converted")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

    Public Function CreateCreditCardInfo(ByRef rowARTCCPA1 As DataRow) As TAC.ARCCCARD.CreditCard
        Dim CreditCardInfo As New TAC.ARCCCARD.CreditCard
        With CreditCardInfo
            .CardCVVData = rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE") & String.Empty

            Dim CUST_CREDIT_CARD_EXP_DATE As String = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
            CUST_CREDIT_CARD_EXP_DATE = CUST_CREDIT_CARD_EXP_DATE.PadLeft(4, "0")
            .CardExpMonth = CUST_CREDIT_CARD_EXP_DATE.Substring(0, 2)
            .CardExpYear = CUST_CREDIT_CARD_EXP_DATE.Substring(2, 2)

            .CardHolderAddress = rowARTCCPA1.Item("CUST_CREDIT_CARD_ADDR1") & String.Empty
            .CardHolderCity = rowARTCCPA1.Item("CUST_CREDIT_CARD_CITY") & String.Empty
            .CardHolderCountry = rowARTCCPA1.Item("CUST_CREDIT_CARD_COUNTRY") & String.Empty
            .CardHolderEmail = ""
            .CardHolderFirstName = rowARTCCPA1.Item("CUST_CREDIT_CARD_NAME") & String.Empty
            .CardHolderLastName = ""
            .CardHolderState = rowARTCCPA1.Item("CUST_CREDIT_CARD_STATE") & String.Empty
            .CardHolderTelephone = ""
            .CardHolderZipCode = rowARTCCPA1.Item("CUST_CREDIT_CARD_ZIP_CODE") & String.Empty
            .CardNumber = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty
            .CustomerID = rowARTCCPA1.Item("CUST_CODE") & String.Empty
            .ResponseApprovalCode = rowARTCCPA1.Item("RESPONSE_APPROVAL_CODE") & String.Empty
            .TransactionID = rowARTCCPA1.Item("TRANS_ID") & String.Empty

            If IsDate(rowARTCCPA1.Item("CCPA_DATE_AUTH") & String.Empty) Then
                .AuthorizationDate = CDate(rowARTCCPA1.Item("CCPA_DATE_AUTH") & String.Empty)
            Else
                .AuthorizationDate = DateTime.Now
            End If

            ' Default values
            .CaptureAmount = 0
            .invoiceDetailsTable = Nothing
            .invoiceHeaderRow = Nothing
            .Level2Data = Nothing
            .Level3Data = Nothing
            .RefundAmount = 0
        End With

        Return CreditCardInfo

    End Function

    Private Sub AdjustResponseCode(ByRef row As DataRow)
        Try
            ' Done for Payeezy
            Select Case (row.Item("RESPONSE_CODE") & String.Empty).ToString.ToUpper
                Case "TRUE"
                    row.Item("RESPONSE_CODE") = "A"
                Case "FALSE"
                    row.Item("RESPONSE_CODE") = "E"
            End Select
        Catch ex As Exception

        End Try
    End Sub

End Class