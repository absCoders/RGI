Public Class TAFTEST1
    Private CreditCardProcessor As TAC.TAFCARD1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            'Create_TDA(.Tables.Add, "DETJOBM1", "*")

            'ASCMAIN1.sql = "Select * from DETJOBM1 where JOB_STATUS = 'O' or JOB_STATUS = 'H'"
            'Create_TDA(.Tables.Add, "DETJOBMX", "**", 0, False)

            'ASCMAIN1.sql = "SELECT * FROM ARTOPEN1 WHERE OPS_YYYYPP = '200812' AND INV_TYPE = 'I'"
            'Create_TDA(.Tables.Add, "ARTOPEN1", "**", 0, True, "", -1, "INV_DUE_DATE")

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            ASCMAIN1.sql = "Select * from SOTINVH1 where INV_NO in ('0007034928','0007035339','0007035388','0007035460')"
            Create_TDA(.Tables.Add, "SOTINVH1", "**", 0)
            Fill_Records("SOTINVH1")

            'ASCMAIN1.sql = "Select * from DETTEST1"
            'Create_TDA(.Tables.Add, "DETTEST1", "**", 0)
            'Fill_Records("DETTEST1")

            ' USE THIS ONE TO CAPTURE PREVIOUSLY AUTHORIZED CC TRANSACTIONS
            ASCMAIN1.sql = "SELECT * FROM ARTCCPA1 WHERE CCPA_STATUS = 'T' AND CCPA_TYPE = 'A'"
            Create_TDA(.Tables.Add, "ARTCCPA1", "**", 0)

            'Create_TDA(.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(.Tables.Add, "ARTCCPA2", "*")

            Create_TDA(.Tables.Add, "SOTSVIA1", "*")

        End With

        grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grd.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grd.DataSource = dst.Tables("ARTOPEN1")

        grdARTCCPA1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grdARTCCPA1.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdARTCCPA1.DataSource = dst.Tables("ARTCCPA1")

        grdARTCCPA1.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grdARTCCPA1, True)

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")

        'ASCMAIN1.sql = "Select * from DETJOBM4 where JOB_NO = :PARM1 and STATUS_CODE = :PARM2"
        'Create_ResultSet("DETJOBM4", "VV")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"

                'Validate_Code("JOB_NO")

            Case "Update"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Select"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Update"
                Stop
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Cancel"
                Call Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf, MODE_description)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Select").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        Absx1.txtFor("JOB_NO").Text = ""
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)
        Fill_Records("ARTOPEN1")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("")
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Call MyBase.txt_KeyDown(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "JOB_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("JOB_NO").Text <> "" Then
                        Call Click_Command("Select Job", e)
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "JOB_NO"
                Call Click_Command("Select Job")
        End Select
    End Sub

#End Region

    Private Sub cmdMagicButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMagicButton.Click
        For Each rowARTOPEN1 As DataRow In dst.Tables("ARTOPEN1").Rows
            Dim TERM_CODE As String = rowARTOPEN1.Item("TERM_CODE") & ""
            Dim INV_DATE As Date = rowARTOPEN1.Item("INV_DATE")
            rowARTOPEN1.Item("INV_DUE_DATE") = SOCMAIN1.Calculate_INV_DUE_DATE(Me, TERM_CODE, INV_DATE)
        Next
    End Sub

    Private Sub cmdTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest.Click

        'Stop

        'Dim ff As New ASFTTIP1
        'ff.ShowDialog()

        'Dim slocation As String = "C:\VS\ODG\ABS_HTML\bin\ABS_HTML.exe"
        'Dim sType As String = "ABS_HTML.ABS_HTML_Main"
        'Dim formAsm As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom(slocation)

        'Dim ClassType As Type = formAsm.GetType(sType)
        'Dim Classobj As New Object
        'Classobj = Activator.CreateInstance(ClassType)

        ''            Dim FormToShow As Form = CType(Classobj, Form)
        'Dim FormToShow As Form = CType(Classobj, Form)
        'FormToShow.Text = "abs test"
        'FormToShow.Tag = "lets see if we can pass a string"
        ''FormToShow.MdiParent = Me
        'FormToShow.WindowState = FormWindowState.Normal

        'FormToShow.ShowDialog()

        Exit Sub

        'Dim RCPControl As New nsoftware.IPWorks.Rcp

        'Dim s As String = ""
        's &= "Line 1" & vbCrLf
        's &= "Line 2" & vbCrLf
        's &= "Line 3" & vbCrLf
        's &= "Line 4" & vbCrLf
        's &= "Line 5" & vbCrLf
        's &= "Line 6"

        'Dim FILENAME As String = "okie03." & ASCMAIN1.Next_Control_No("SPOOLEDJOB")
        'Using jobWriter As New System.IO.StreamWriter _
        '(ASCMAIN1.Folders("Temp") & FILENAME)
        '    jobWriter.Write(s & vbCrLf)
        'End Using


        'RCPControl.RemoteHost = "192.168.130.3"
        'RCPControl.User = "mattinam"
        'RCPControl.Password = "m24bw"
        'RCPControl.Protocol = nsoftware.IPWorks.RcpProtocols.protRexec
        'RCPControl.RemoteFile = "./abs/spool/queue/" & FILENAME
        'RCPControl.LocalFile = ASCMAIN1.Folders("Temp") & FILENAME
        'RCPControl.PutFile()

        Stop
        'Stop
        'ASCMAIN1.sql = "begin test (1); end;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        'Dim A As Integer = 1
        'Dim B As Integer = 0
        'Dim s As String = ""
        'ASCMAIN1.sql = "Select * from x " & CStr(A / B)
        'For Each row As DataRow In ASCDATA1.GetDataTable.Rows
        '    s &= row.Item(0)
        'Next
        'MsgBox(s)
        'Stop
    End Sub

    Private Sub grd_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grd.InitializeLayout

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Fill_Records("ARTCCPA1")
        UltraButton2.Enabled = True
    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click

        Dim CCPA_NO As String = grdARTCCPA1.ActiveRow.Cells("CCPA_NO").Text
        Dim rowARTCCPA1 As DataRow = dst.Tables("ARTCCPA1").Rows.Find(CCPA_NO)
        ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_NO = '" & rowARTCCPA1.Item("ORDR_NO") & "'"
        Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow
        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
        Dim INV_TOTAL_AMOUNT As Decimal = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & "")


        ASCMAIN1.sql = "Select * from ARTOPEN1 where INV_TYPE = 'I' AND INV_NUM = '" & rowSOTINVH1.Item("INV_NO") & "'"
        Dim rowARTOPEN1 As DataRow = ASCDATA1.GetDataRow
        Dim INV_BALANCE As Decimal = Val(rowARTOPEN1.Item("INV_BALANCE") & "")


        If rowARTCCPA1.Item("CCPA_STATUS") & "" <> "T" Then
            MsgBox("Captured Already")
            Exit Sub
        End If

        If MsgBox("Customer " & rowARTCCPA1.Item("CUST_CODE") & " Amount = " & CStr(INV_TOTAL_AMOUNT) & " Balance Open = " & CStr(INV_BALANCE), MsgBoxStyle.YesNo, "OK to Charge") <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        rowARTCCPA1.Item("INV_NO") = INV_NO

        Dim CCPA_NO_AUTH As String = CCPA_NO ' rowSOTORDR1.Item("CCPA_NO") & ""
        Dim CCPA_NO_CAPTURE As String = ""
        Try
            Me.CreditCardProcessor = New TAC.TAFCARD1(Me)

            With Me.CreditCardProcessor
                .rowARTCCPA1 = rowARTCCPA1
                CCPA_NO_CAPTURE = .CC_Capture(INV_TOTAL_AMOUNT)
            End With

        Catch ex As Exception
            Stop
        End Try


        BeginTrans()

        Dim RA As Int32 = New DataView(dst.Tables("ARTCCPA1"), "", "", DataViewRowState.Added).ToTable.Rows.Count
        Dim RM As Int32 = New DataView(dst.Tables("ARTCCPA1"), "", "", DataViewRowState.ModifiedCurrent).ToTable.Rows.Count
        Dim RD As Int32 = New DataView(dst.Tables("ARTCCPA1"), "", "", DataViewRowState.Deleted).ToTable.Rows.Count
        If RA <> 1 Or RM <> 1 Or RD <> 0 Then Stop

        Update_Record_TDA("ARTCCPA1")
        Update_Record_TDA("ARTCCPA2")

        ASCDATA1.ExecuteSQL("UPDATE SOTINVH1 SET CCPA_NO = '" & CCPA_NO_CAPTURE & "' where INV_NO = '" & INV_NO & "'")
        CommitTrans()
        MsgBox("Credit Card Charged")
    End Sub


    Private Function ProcessCreditCardAuthorization(ByVal AUTH_CCPA_NO As String, ByVal ChargeAmount As Decimal) As String
        'ProcessCreditCard = False

        AUTH_CCPA_NO = AUTH_CCPA_NO.Trim
        If AUTH_CCPA_NO.Length = 0 Then Return String.Empty
        If ChargeAmount <= 0 Then Return String.Empty

        MyBase.Fill_Records("ARTCCPA1", AUTH_CCPA_NO)
        If dst.Tables("ARTCCPA1").Rows.Count <> 1 Then Return String.Empty

        Dim rowARTCCPA1_AUTH As DataRow = dst.Tables("ARTCCPA1").Rows(0)
        Dim AUTH_RESPONSE_APPROVAL_CODE As String = (rowARTCCPA1_AUTH.Item("RESPONSE_APPROVAL_CODE") & String.Empty).ToString.Trim

        If AUTH_RESPONSE_APPROVAL_CODE.Length = 0 Then Return String.Empty

        If Val(rowARTCCPA1_AUTH.Item("CCPA_AMT") & String.Empty) < ChargeAmount Then
            Return String.Empty
        End If

        Dim CCPA_NO As String = String.Empty

        Try
            Me.CreditCardProcessor = New TAC.TAFCARD1(Me)
            With Me.CreditCardProcessor
                .rowARTCCPA1 = rowARTCCPA1_AUTH
                .CC_Capture(ChargeAmount)
                CCPA_NO = .Create_CC_Capture_Entry(rowARTCCPA1_AUTH, "")
            End With

        Catch ex As Exception
            MessageBox.Show("The following error occurred processing a credit card: " & ex.Message, "Charge Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End Try

        Return CCPA_NO

    End Function

    Private Sub cmdCreateAR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreateAR.Click

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Rows
            Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
            rowARTOPEN1.Item("CUST_CODE") = rowSOTINVH1.Item("CUST_CODE")
            rowARTOPEN1.Item("INV_TOTAL_AMOUNT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
            rowARTOPEN1.Item("INV_TYPE") = rowSOTINVH1.Item("INV_TYPE")
            rowARTOPEN1.Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")
            rowARTOPEN1.Item("INV_DATE") = rowSOTINVH1.Item("INV_DATE")
            rowARTOPEN1.Item("CUST_SHIP_TO_NO") = rowSOTINVH1.Item("CUST_SHIP_TO_NO")
            rowARTOPEN1.Item("POST_CODE") = rowSOTINVH1.Item("POST_CODE")
            rowARTOPEN1.Item("TERM_CODE") = rowSOTINVH1.Item("TERM_CODE") 'set terms from customer onto SOTINVH1
            rowARTOPEN1.Item("INV_DUE_DATE") = rowSOTINVH1.Item("INV_DATE")
            'rowARTOPEN1.Item("INV_DISC_DATE")
            rowARTOPEN1.Item("SREP_CODE") = rowSOTINVH1.Item("SREP_CODE")
            rowARTOPEN1.Item("STAX_CODE") = rowSOTINVH1.Item("STAX_CODE")
            'rowARTOPEN1.Item("APPLY_TO_INV_NUM")
            'rowARTOPEN1.Item("APPLY_TO_INV_TYPE")
            rowARTOPEN1.Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")
            'rowARTOPEN1.Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO")
            rowARTOPEN1.Item("INV_SALES") = rowSOTINVH1.Item("INV_SALES")
            rowARTOPEN1.Item("INV_DISC") = 0
            rowARTOPEN1.Item("INV_FREIGHT") = rowSOTINVH1.Item("INV_FREIGHT")
            rowARTOPEN1.Item("INV_STAX") = rowSOTINVH1.Item("INV_STAX")
            rowARTOPEN1.Item("INV_TOTAL_AMOUNT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
            'rowARTOPEN1.Item("INV_LAST_PMT")
            'rowARTOPEN1.Item("PMT")
            'rowARTOPEN1.Item("INV_DISC_TAKEN")
            'rowARTOPEN1.Item("INV_WRITE_OFF")
            rowARTOPEN1.Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
            'rowARTOPEN1.Item("INV_LAST_PMT_REF")
            'rowARTOPEN1.Item("INV_LAST_PMT_REF_DT")
            'rowARTOPEN1.Item("CUST_CODE_SO")
            rowARTOPEN1.Item("REASON_CODE") = rowSOTINVH1.Item("REASON_CODE")
            rowARTOPEN1.Item("INIT_OPER") = rowSOTINVH1.Item("INIT_OPER")
            'rowARTOPEN1.Item("LAST_OPER")
            rowARTOPEN1.Item("INIT_DATE") = rowSOTINVH1.Item("INIT_DATE")
            'rowARTOPEN1.Item("LAST_DATE")
            rowARTOPEN1.Item("INV_MISC_CHG") = rowSOTINVH1.Item("INV_MISC_CHG")
            rowARTOPEN1.Item("SEG2_CODE") = "000"
            rowARTOPEN1.Item("SEG3_CODE") = "000"
            rowARTOPEN1.Item("SEG4_CODE") = "000"
            'rowARTOPEN1.Item("OPS_YYYYPP_F")
            rowARTOPEN1.Item("CURR_CODE") = "USD"
            rowARTOPEN1.Item("CURR_EXCH_RATE") = 1
            rowARTOPEN1.Item("INV_SALES_CURR") = 0
            rowARTOPEN1.Item("INV_DISC_CURR") = 0
            rowARTOPEN1.Item("INV_FREIGHT_CURR") = rowSOTINVH1.Item("INV_FREIGHT")
            rowARTOPEN1.Item("INV_STAX_CURR") = rowSOTINVH1.Item("INV_STAX")
            rowARTOPEN1.Item("INV_MISC_CHG_CURR") = rowSOTINVH1.Item("INV_MISC_CHG")
            rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
            'rowARTOPEN1.Item("INV_PMT_CURR")
            'rowARTOPEN1.Item("INV_DISC_TAKEN_CURR")
            'rowARTOPEN1.Item("INV_WRITE_OFF_CURR")
            rowARTOPEN1.Item("INV_BALANCE_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
            rowARTOPEN1.Item("INV_NOTES") = rowSOTINVH1.Item("INV_NOTES")
            'rowARTOPEN1.Item("ORDR_NO_WEB")
            'rowARTOPEN1.Item("INV_PROFIT_B2C")
            rowARTOPEN1.Item("ORDR_TYPE_CODE") = rowSOTINVH1.Item("ORDR_TYPE_CODE")
            rowARTOPEN1.Item("INV_REF") = rowSOTINVH1.Item("INV_REF")
            rowARTOPEN1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowARTOPEN1.Item("DIVISION_CODE") = rowSOTINVH1.Item("DIVISION_CODE")
            dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
        Next

        Stop
        Update_Record_TDA("ARTOPEN1")
    End Sub

    Private Sub cmdTestTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestTime.Click
        Dim row As DataRow = dst.Tables("DETTEST1").Rows(0)
        row.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
        row.Item("INIT_OPER") = "1"
        Update_Record_TDA("DETTEST1")

        ASCDATA1.ExecuteSQL("UPDATE DETTEST1 SET LAST_DATE = SYSDATE, LAST_OPER = '2'")
        MsgBox("Done")
    End Sub

    Private Sub cmdTestFreight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestFreight.Click

        Stop

        ASCMAIN1.sql = "SELECT * FROM SOTINVH1 WHERE INV_NO = '0007088225'"
        Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow

        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
        Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO")
        Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE")
        Dim CUST_SHIP_TO_NO As String = rowSOTINVH1.Item("CUST_SHIP_TO_NO")
        Dim SHIP_VIA_CODE As String = rowSOTINVH1.Item("SHIP_VIA_CODE")

        Dim INV_SALES As Decimal = Val(rowSOTINVH1.Item("INV_SALES") & "")

        ASCMAIN1.sql = "Select Sum (ORDR_QTY_SHIP) from SOTINVH2 " _
        & "  where INV_TYPE = 'I' and INV_NO = '" & INV_NO & "'"
        Dim INV_QTY As Int32 = Val(ASCDATA1.GetDataValue & "")

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow
        Dim ORDR_DATE As Date = rowSOTORDR1.Item("ORDR_DATE")
        Dim ORDR_SOURCE As String = rowSOTORDR1.Item("ORDR_SOURCE")

        Dim ORDR_DPD As Boolean = False
        For I As Integer = 1 To 5
            ORDR_DPD = ORDR_DPD Or ("" & String.Empty) = "1"
            MsgBox(ORDR_DPD)
        Next
        Stop

        Dim FRT As Decimal = SOCMAIN1.Get_INV_FREIGHT(Me, CUST_CODE, CUST_SHIP_TO_NO, SHIP_VIA_CODE, ORDR_DATE _
                                , INV_QTY, INV_SALES, rowSOTORDR1.Item("ORDR_DPD") & "" = "1", ORDR_SOURCE)

        MsgBox("Frt = " & FRT)

    End Sub

    Private Sub cmdTestCCAuth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestCCAuth.Click

        'ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & "0007100377" & "'"
        'Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow
        'Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
        'Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")
        'Dim CUST_CREDIT_CARD_NO As String = "371055358751001"
        'Dim CUST_CREDIT_CARD_EXP_DATE As String = "0310"
        'Dim ORDR_TOTAL_AMT As Decimal = rowSOTORDR1.Item("ORDR_TOTAL_AMT")

        '' ARTCCPA1,ARTCCPA2,ARTCUSTC
        ''Create_TDA(.Tables.Add, "ARTCCPA1", "*")
        ''Create_TDA(.Tables.Add, "ARTCCPA2", "*")
        ''Create_TDA(.Tables.Add, "ARTCUSTC", "*")

        'Dim rowARTCCPA1 As DataRow = SOCMAIN1.CC_Auth _
        '(Me, ORDR_NO, CUST_CODE, CUST_CREDIT_CARD_NO, CUST_CREDIT_CARD_EXP_DATE, ORDR_TOTAL_AMT, _
        ' "W", True)

        'If rowARTCCPA1.Item("CCPA_STATUS") = "T" Then
        '    rowSOTORDR1.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
        '    rowSOTORDR1.Item("ORDR_CC_AUTH_AMT") = ORDR_TOTAL_AMT
        'Else
        '    rowSOTORDR1.Item("ORDR_STATUS_WEB") = "C" ' RESPONSE_TEXT
        'End If

        Stop
        Dim OBAD As String = ""

        ASCMAIN1.sql = "SELECT WC.*, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_STATUS, SOTORDR1.ORDR_TOTAL_AMT " _
        & "  FROM ( " _
        & " SELECT SOTORDWC.*, SOTORDR1.ORDR_TYPE_CODE " _
        & " FROM SOTORDWC, SOTORDR1 " _
        & " WHERE SOTORDWC.ORDR_NO LIKE '0007%' " _
        & " AND SOTORDWC.ORDR_NO = SOTORDR1.ORDR_NO " _
        & " AND SOTORDR1.ORDR_NO <> '0007013478'" _
        & " AND SOTORDR1.ORDR_NO <> '0007073751'" _
        & " AND SOTORDR1.ORDR_NO <> '0007001206'" _
        & " AND SOTORDR1.CCPA_NO IS NULL" _
        & " AND SOTORDR1.ORDR_TYPE_CODE = 'B2B') WC, SOTORDR1 " _
        & " WHERE WC.ORDR_NO = SOTORDR1.ORDR_NO"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows

            dst.Tables("ARTCCPA1").Rows.Clear()
            dst.Tables("ARTCCPA2").Rows.Clear()

            Dim ORDR_NO As String = row.Item("ORDR_NO")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            ASCMAIN1.Progress("Now Authorizing " & ORDR_NO)

            'Stop
            Dim CUST_CREDIT_CARD_NO As String = row.Item("ORDR_CC_NUMBER") ' "371055358751001"
            Dim CUST_CREDIT_CARD_EXP_DATE As String = row.Item("ORDR_CC_EXP_DATE") ' "0310"
            Dim ORDR_TOTAL_AMT As Decimal = row.Item("ORDR_TOTAL_AMT")


            'Stop
            Dim rowARTCCPA1 As DataRow = SOCMAIN1.CC_Auth _
            (Me, ORDR_NO, CUST_CODE, CUST_CREDIT_CARD_NO, CUST_CREDIT_CARD_EXP_DATE, ORDR_TOTAL_AMT, _
             "W", True)

            If rowARTCCPA1.Item("CCPA_STATUS") = "T" Then
                ASCMAIN1.sql = "Update SOTORDR1 set CCPA_NO = '" & rowARTCCPA1.Item("CCPA_NO") & "'" _
                & ", ORDR_CC_AUTH_AMT = " & CStr(ORDR_TOTAL_AMT) _
                & " where ORDR_NO = '" & ORDR_NO & "'"
                ASCDATA1.ExecuteSQL()
            Else
                OBAD &= ",'" & ORDR_NO & "'"
                Stop
            End If

        Next

        ASCMAIN1.Progress("")

    End Sub

    Private Sub cmdTestCCSale_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestCCSale.Click

        ASCMAIN1.sql = "SELECT WC.*, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_STATUS" _
        & ", SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.INV_TOTAL_AMOUNT, ARTOPEN1.INV_BALANCE " _
        & "  FROM ( " _
        & " SELECT SOTORDWC.*, SOTORDR1.ORDR_TYPE_CODE " _
        & " FROM SOTORDWC, SOTORDR1 " _
        & " WHERE SOTORDWC.ORDR_NO LIKE '0007%' " _
        & " AND SOTORDWC.ORDR_NO = SOTORDR1.ORDR_NO " _
        & " AND SOTORDR1.ORDR_TYPE_CODE = 'B2B') WC, SOTORDR1, SOTINVH1, ARTOPEN1 " _
        & " WHERE WC.ORDR_NO = SOTORDR1.ORDR_NO " _
        & " AND SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO " _
        & " AND NVL(SOTINVH1.INV_TOTAL_AMOUNT,0) <> 0 " _
        & " AND ARTOPEN1.INV_NUM = SOTINVH1.INV_NO" _
        & " AND ARTOPEN1.INV_BALANCE = SOTINVH1.INV_TOTAL_AMOUNT" _
        & " AND SOTORDR1.ORDR_NO <> '0007001206'" _
        & " AND SOTORDR1.ORDR_NO <> '0007073751'" _
        & " AND SOTORDR1.ORDR_NO <> '0007001921'" _
        & " AND SOTORDR1.CCPA_NO IS NOT NULL" _
        & " AND SOTINVH1.CCPA_NO IS NULL " _
        & " ORDER BY WC.ORDR_NO,  SOTINVH1.INV_NO"

        ' for the 8 that had to be redone
        ASCMAIN1.sql = "SELECT SOTINVH1.INV_NO, SOTINVH1.ORDR_NO FROM SOTINVH1" _
        & " WHERE SOTINVH1.INV_NO IN ('0007130542','0007130110'," _
        & "'0007131050','0007129398','0007131243','0007129339','0007130779','0007007506')  "

        ' for the 33 from Week 1
        ASCMAIN1.sql = "SELECT INV_NO, ORDR_NO, INV_TOTAL_AMOUNT " _
        & "  FROM SOTINVH1  " _
        & "  WHERE INV_NO IN ('0007004146','0007024741','0007024748','0007020770','0007021596','0007009067','0007005532','0007024738','0007027598', " _
        & "  '0007008095','0007005360','0007028717','0007030666','0007027446','0007029861','0007026793','0007024765','0007003256','0007007984', " _
        & "  '0007014951','0007024753','0007026736','0007026720','0007026615','0007024390', " _
        & "  '0007024456','0007024958','0007009905','0007025006','0007007407','0007024761','0007027535','0007031049')"

        Stop

        For Each rowCharge As DataRow In ASCDATA1.GetDataTable.Rows

            dst.Tables("ARTCCPA1").Rows.Clear()
            dst.Tables("ARTCCPA2").Rows.Clear()

            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & rowCharge.Item("ORDR_NO") & "'"
            Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow
            ASCMAIN1.sql = "Select * from SOTINVH1 where INV_NO = '" & rowCharge.Item("INV_NO") & "'"
            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow
            'If rowSOTINVH1.Item("CCPA_NO") & "" <> "" Then Stop ' WAS THIS PAID ALREADY?

            Dim CCPA_NO As String = rowSOTORDR1.Item("CCPA_NO")
            Dim rowARTCCPA1 As DataRow = Fill_Record("ARTCCPA1", CCPA_NO)

            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim INV_TOTAL_AMOUNT As Decimal = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & "")

            ASCMAIN1.Progress("Now Capturing Sale for Invoice " & INV_NO)

            ASCMAIN1.sql = "Select * from ARTOPEN1 where INV_TYPE = 'I' AND INV_NUM = '" & rowSOTINVH1.Item("INV_NO") & "'"
            Dim rowARTOPEN1 As DataRow = ASCDATA1.GetDataRow
            Dim INV_BALANCE As Decimal = Val(rowARTOPEN1.Item("INV_BALANCE") & "")

            If INV_BALANCE <> INV_TOTAL_AMOUNT Then
                Stop
            End If

            'If rowARTCCPA1.Item("CCPA_STATUS") & "" <> "T" Then
            '    Stop
            'End If

            'If MsgBox("Customer " & rowARTCCPA1.Item("CUST_CODE") & " Amount = " & CStr(INV_TOTAL_AMOUNT) & " Balance Open = " & CStr(INV_BALANCE), MsgBoxStyle.YesNo, "OK to Charge") <> MsgBoxResult.Yes Then
            '    Stop
            'End If

            rowARTCCPA1.Item("INV_NO") = INV_NO

            Dim CCPA_NO_AUTH As String = CCPA_NO
            Dim CCPA_NO_CAPTURE As String = ""
            Try
                Me.CreditCardProcessor = New TAC.TAFCARD1(Me)

                With Me.CreditCardProcessor
                    .rowARTCCPA1 = rowARTCCPA1
                    CCPA_NO_CAPTURE = .CC_Capture(INV_TOTAL_AMOUNT)
                End With

            Catch ex As Exception
                Stop
            End Try


            BeginTrans()

            Update_Record_TDA("ARTCCPA1")
            Update_Record_TDA("ARTCCPA2")

            ASCDATA1.ExecuteSQL("UPDATE SOTINVH1 SET CCPA_NO = '" & CCPA_NO_CAPTURE & "' where INV_NO = '" & INV_NO & "'")
            CommitTrans()

        Next

    End Sub

    Private Sub cmdTestSocket_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestSocket.Click
        Dim ipp As New nsoftware.IPWorks.Ipport
        ipp.Connect("192.168.130.45", 9100)
        'ipp.Connect("192.168.100.81", 9100)



        'ipp.SendFile("C:\CRTEST")


        Dim BSR As New System.IO.BinaryReader( _
            System.IO.File.Open("C:\CRTEST", System.IO.FileMode.Open))

        'ipp.SendStream(BSR.BaseStream)
        'ipp.SetDataToSend(BSR.ReadBytes(BSR.BaseStream.Length), 0, BSR.BaseStream.Length)


        ipp.Send(BSR.ReadBytes(BSR.BaseStream.Length))
        BSR.Close()

        Using SW As New System.IO.StreamReader("C:\CRTEST")
            'ipp.SendStream(SW)

            '    Dim X As String = SW.ReadToEnd
            '    ipp.DataToSend = X & Chr(12)
        End Using

        'ipp.DataToSend = "hello world" & Chr(12)

        ipp.Disconnect()
        ipp.Dispose()


    End Sub
End Class