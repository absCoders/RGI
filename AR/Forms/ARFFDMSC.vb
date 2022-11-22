Imports ABSolution

Public Class ARFFDMSC

    Private RESPONSE_BATCH_NO As String = String.Empty
    Private tblARTCCPA1 As String = String.Empty
    Private SO_PARM_CC_PROC_CODE As String = String.Empty
    Private rowARTCCPRC As DataRow = Nothing
    Private CreditCardProcessingSetupForUse As Boolean = False
    Private BANK_CODE As String = String.Empty
    Private Const CURR_CODE As String = "USD"
    Private clsTACENCRY As TAC.ASCENCRY

    Private Const ACI_CODE As String = "RR"

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            Get_PARM("SOTPARM1")
            Get_PARM("ARTPARM1")

            tblARTCCPA1 = ASCMAIN1.Temp_Table("SELECT CCPA_NO, CCPA_AMT FROM ARTCCPA1 WHERE ROWNUM < 1")

            SO_PARM_CC_PROC_CODE = ROWs("SOTPARM1").Item("SO_PARM_CC_PROC_CODE") & String.Empty
            rowARTCCPRC = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPRC WHERE CC_PROC_CODE = '" & SO_PARM_CC_PROC_CODE & "'")

            Create_TDA(.Tables.Add, "ARTCCPA1", "*")
            .Tables("ARTCCPA1").Columns.Add("CUST_NAME", GetType(System.String))

            ' NEXT FEW LINES ARE TEMP - SEE InquiryBatch for More
            With .Tables("ARTCCPA1")
                .Columns.Add("TOTAL_DUE", GetType(System.Decimal))
                .Columns.Add("AGE_1", GetType(System.Decimal))
                .Columns.Add("AGE_2", GetType(System.Decimal))
                .Columns.Add("AGE_3", GetType(System.Decimal))
                .Columns.Add("AGE_4", GetType(System.Decimal))
            End With

            ASCMAIN1.sql = "SELECT * FROM ARTCCPA2 WHERE CCPA_NO IN (SELECT CCPA_NO FROM " & tblARTCCPA1 & ")"
            Create_TDA(.Tables.Add, "ARTCCPA2", "*")

            ASCMAIN1.sql = "SELECT * FROM ARTCCPDA WHERE CCPA_NO IN (SELECT CCPA_NO FROM " & tblARTCCPA1 & ")"
            Create_TDA(.Tables.Add, "ARTCCPDA", "**", 0, False)

            Create_TDA(.Tables.Add, "ARTCCPS2", "*")
            Create_TDA(.Tables.Add, "ARTCCPAM", "*")

            With .Tables.Add("ARTCCPA0")
                .Columns.Add("LINE", GetType(System.Int32))
                .Columns.Add("TEXT")
                .Columns.Add("TRANS", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("LINE")}
            End With

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")
            Create_TDA(.Tables.Add, "ARTPYMT3", "*")
            Create_TDA(.Tables.Add, "ARTPYMT4", "*")
            Create_TDA(.Tables.Add, "ARTPYMT5", "*")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")

        End With

        grdARTCCPA1.DataSource = dst.Tables("ARTCCPA1")
        Call Create_Summary(grdARTCCPA1, "CUST_CODE", "Count")
        Call Create_Summary(grdARTCCPA1, "CCPA_AMT")

        With grdARTCCPA1.DisplayLayout.Bands("ARTCCPA1")
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
            .Columns("CCPA_AMT").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdARTCCPA1, "CCPA_STATUS")
        ASCMAIN1.Add_Value_List(grdARTCCPA1, "CCPA_REASON")
        ASCMAIN1.Add_Value_List(grdARTCCPA1, "CCPA_TYPE")
        ASCMAIN1.Add_Value_List(grdARTCCPA1, "RESPONSE_CODE")


        grdARTCCPS2.DataSource = dst.Tables("ARTCCPS2")
        Call Create_Summary(grdARTCCPS2, "RESPONSE_PYMT_TYPE_NET_AMT")
        Call Create_Summary(grdARTCCPS2, "RESPONSE_PYMT_TYPE_TRANS")

        grdARTCCPA0.DataSource = dst.Tables("ARTCCPA0")
        Absx1.txtFor("BANK_CODE").Text = ROWs("ARTPARM1").Item("AR_PARM_BANK_CODE_CC") & String.Empty

        clsTACENCRY = New TAC.ASCENCRY()
        ValidateEncryption()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Inquiry Batch"

                If Not ASCMAIN1.Logical_Lock("ARTCCPA1", "F") Then Exit Sub

            Case "Settle Batch"

                If RESPONSE_BATCH_NO.Length = 0 Then
                    EMsg &= vbCr & "The response Batch No is not set"
                    Exit Select
                End If

                If dst.Tables("ARTCCPA1").Select($"RESPONSE_BATCH_NO = '{RESPONSE_BATCH_NO}'").Length = 0 Then
                    EMsg &= vbCr & "There are no Settled Paymets to Apply"
                    Exit Select
                End If

                If MessageBox.Show("Do you want to Settle the Batch?", "Settle Batch", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Inquiry Batch"
                EntryMode = "C"
                Inquiry_Batch()
                Mode_Settings(True)

            Case "Excel"
                If EntryMode = "C" Then
                    Export_to_Excel(New UltraWinGrid.UltraGrid() {grdARTCCPA0, grdARTCCPS2, grdARTCCPA1})
                    grdARTCCPA0.Tag = "*"
                End If

            Case "Settle Batch"
                Apply_CC_Payment()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Inquiry Batch").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Excel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Settle Batch").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

        splCC.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()

        EnforceConstraints(False)

        For Each tableName As String In New String() {"ARTCCPA1", "ARTCCPA2", "ARTCCPAM", "ARTCCPDA", "ARTPYMT1", "ARTPYMT2", _
                                                      "ARTPYMT3", "ARTPYMT4", "ARTPYMT5", "ARTOPEN1"}
            dst.Tables(tableName).Rows.Clear()
        Next

        RESPONSE_BATCH_NO = String.Empty
        grdARTCCPA1.Text = "Today's Payments"

        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Credit Card Transaction")

        MyBase.EnforceConstraints(False)


        MyBase.EnforceConstraints(True)

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()

            MyBase.CommitTrans("Update Complete")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdARTCCPA1, "BBB", "Replicate Transaction", "Void Transaction", "Remove From Queue")
        Load_Popup_Menu(grdARTCCPA1, "B", "Void Transaction")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name


            End Select
        End If

    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Replicate Transaction"

                If UltraExplorerBar1.Groups("Credit Card Options").Items("Settle Batch").Settings.Enabled = DefaultableBoolean.True Then
                    Dim zMsg As String = "If you Replicate this Transaction, you will have to Cancel and then Re-Inquiry the Batch before Settling."
                    zMsg &= vbCrLf & vbCrLf & "OK to Continue to Replicate Transaction?"
                    If MessageBox.Show(zMsg, "Replicate Transaction", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
                End If

                Dim row As DataRow = dst.Tables("ARTCCPA1").Rows.Find(grd.ActiveRow.Cells("CCPA_NO").Text)
                Dim rowARTCCPA1 As DataRow = dst.Tables("ARTCCPA1").NewRow
                rowARTCCPA1.ItemArray = row.ItemArray
                rowARTCCPA1.Item("CCPA_NO") = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
                rowARTCCPA1.Item("ORDR_NO") = ""
                rowARTCCPA1.Item("INV_NO") = ""
                rowARTCCPA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowARTCCPA1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1)
                Update_Record_TDA("ARTCCPA1")

                UltraExplorerBar1.Groups("Credit Card Options").Items("Settle Batch").Settings.Enabled = DefaultableBoolean.False

                MsgBox("Record for " & Format(Val(rowARTCCPA1.Item("CCPA_AMT") & ""), "$#,##0.00") & " has been Replicated", MsgBoxStyle.OkOnly, "Verification")

            Case "Remove From Queue"
                If grd.Selected.Rows.Count = 0 Then
                    MessageBox.Show("There are no selected transactions to remove.", "Remove From Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim lstCCPA_NOS As New List(Of String)
                Dim lstBadCCPA_NOS As New List(Of String)

                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    lstCCPA_NOS.Add(grdRow.Cells("CCPA_NO").Value & String.Empty)
                Next

                For Each CCPA_NO As String In lstCCPA_NOS
                    Dim rowARTCCPA1 As DataRow = dst.Tables("ARTCCPA1").Rows.Find(CCPA_NO)
                    Dim CUST_CODE As String = rowARTCCPA1.Item("CUST_CODE") & String.Empty
                    Dim CCPA_AMT As Decimal = Val(rowARTCCPA1.Item("CCPA_AMT") & String.Empty).ToString("#,##0.00")

                    If rowARTCCPA1.Item("CCPA_STATUS") & String.Empty <> "A" Then
                        MessageBox.Show("Charge for customer " & CUST_CODE & " for " & CCPA_AMT & " cannot be removed from the queue. It does not have a Status of 'A'.", "Remove From Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        lstBadCCPA_NOS.Add(CCPA_NO)
                        Continue For
                    End If

                    If grd.ActiveRow.Cells("CCPA_STATUS").Value <> "A" Then
                        MessageBox.Show($"Charge for customer {CUST_CODE} for {CCPA_AMT} cannot be Voided it has a Status of {grd.ActiveRow.Cells("CCPA_STATUS").Text}", "Void Transaction", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        lstBadCCPA_NOS.Add(CCPA_NO)
                        Continue For
                    End If

                    If grd.ActiveRow.Cells("RESPONSE_BATCH_NO").Text & String.Empty <> String.Empty Then
                        MessageBox.Show("Charge for customer " & CUST_CODE & " for " & CCPA_AMT & " cannot be removed from the queue. It was settled.", "Remove From Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        lstBadCCPA_NOS.Add(CCPA_NO)
                        Continue For
                    End If

                    Dim CCPA_DATE_AUTH As String = rowARTCCPA1.Item("CCPA_DATE_AUTH") & String.Empty

                    If IsDate(CCPA_DATE_AUTH) Then
                        If DateDiff(DateInterval.Day, CDate(CCPA_DATE_AUTH), DateTime.Now) < 90 Then
                            MessageBox.Show("Charge for customer " & CUST_CODE & " for " & CCPA_AMT & " removed from the queue. It is not over 90 days old.", "Remove From Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            lstBadCCPA_NOS.Add(CCPA_NO)
                            Continue For
                        End If
                    End If

                Next

                For Each CCPA_NO As String In lstBadCCPA_NOS
                    If lstCCPA_NOS.Contains(CCPA_NO) Then
                        lstCCPA_NOS.Remove(CCPA_NO)
                    End If
                Next

                If lstCCPA_NOS.Count = 0 Then
                    MessageBox.Show("There is nothing to remove from the queue.", "Remove From Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If UltraExplorerBar1.Groups("Screen Control").Items("Settle Batch").Settings.Enabled = DefaultableBoolean.True Then
                    Dim zMsg As String = "If you Remove the " & lstCCPA_NOS.Count & " Transactions from the Queue, you will have to Cancel and then Re-Inquiry the Batch before Settling."
                    zMsg &= vbCrLf & vbCrLf & "OK to Continue to Remove From Queue?"
                    If MessageBox.Show(zMsg, "Remove From Queue", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
                End If

                UltraExplorerBar1.Groups("Screen Control").Items("Settle Batch").Settings.Enabled = DefaultableBoolean.False

                For Each CCPA_NO As String In lstCCPA_NOS
                    ASCMAIN1.sql = $"Update ARTCCPA1 
                                        SET LAST_OPER = '{ASCMAIN1.USER_ID}', 
                                        LAST_DATE = SYSDATE, 
                                        RESPONSE_BATCH_NO = 'REMOVE' 
                                        WHERE RESPONSE_BATCH_NO IS NULL 
                                        and CCPA_NO = :PARM1"
                    Try
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {CCPA_NO})
                        Dim rowARTCCPA1 As DataRow = dst.Tables("ARTCCPA1").Rows.Find(CCPA_NO)
                        rowARTCCPA1.Delete()
                    Catch ex As Exception
                        MessageBox.Show("Error Removing From Queue: " & ex.Message, "Remove From Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Next

                MessageBox.Show(lstCCPA_NOS.Count & " entries Removed From Queue: ", "Remove From Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Mode_Settings(False)

            Case "Void Transaction"

                If grd.ActiveRow.Cells("CCPA_STATUS").Value <> "A" Then
                    MessageBox.Show($"This charge cannot be Voided it has a Status of {grd.ActiveRow.Cells("CCPA_STATUS").Text}", "Void Transaction", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If grd.ActiveRow.Cells("RESPONSE_BATCH_NO").Text & String.Empty <> String.Empty Then
                    MessageBox.Show($"This charge cannot be Voided it has a Sale/Capture date of Today.", "Void Transaction", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text

                If UltraExplorerBar1.Groups("Screen Control").Items("Settle Batch").Settings.Enabled = DefaultableBoolean.True Then
                    Dim zMsg As String = "If you Void this Transaction, you will have to Cancel and then Re-Inquiry the Batch before Settling."
                    zMsg &= vbCrLf & vbCrLf & "OK to Continue to Void this Transaction?"
                    If MessageBox.Show(zMsg, "Void Transaction", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
                End If

                Dim rowARTCCPA1 As DataRow = dst.Tables("ARTCCPA1").Rows.Find(grd.ActiveRow.Cells("CCPA_NO").Text)

                Dim frmCCProcessor As New TAFCARDF(Me)
                frmCCProcessor.CUST_CODE = ""
                frmCCProcessor.rowARTCCPA1 = rowARTCCPA1
                frmCCProcessor.TRAN_TYPE = "V"
                frmCCProcessor.ORDR_NO = rowARTCCPA1.Item("ORDR_NO") & ""
                frmCCProcessor.INV_NO = rowARTCCPA1.Item("INV_NO") & ""
                frmCCProcessor.STMT_NO = rowARTCCPA1.Item("STMT_NO") & ""

                frmCCProcessor.CCPA_REASON = rowARTCCPA1.Item("CCPA_REASON") & ""
                frmCCProcessor.ShowAgedTotals = False
                frmCCProcessor.ShowDialog()

                frmCCProcessor.Dispose()

                UltraExplorerBar1.Groups("Screen Control").Items("Settle Batch").Settings.Enabled = DefaultableBoolean.False

        End Select
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub ValidateEncryption()

        Dim rowASTPARMP As DataRow = ASCDATA1.GetDataRow("Select * from ASTPARMP WHERE AS_PARM_KEY = 'Z'")
        If rowASTPARMP Is Nothing OrElse Not rowASTPARMP.Table.Columns.Contains("AS_PARM_USE_ENCRYPTION") OrElse rowASTPARMP.Item("AS_PARM_USE_ENCRYPTION") & String.Empty <> "1" Then
            clsTACENCRY.UseEncryption = False
        Else
            clsTACENCRY.UseEncryption = True
        End If
    End Sub

    Sub Inquiry_Batch()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Checking Current Batch with CC Processor")

        ' Need temp table of CCPA_NOs since FDMS needs to limit the number of transactions per settlement
        Dim sql As String = String.Empty
        sql = "SELECT CCPA_NO, CCPA_AMT
                FROM ARTCCPA1
                where CCPA_STATUS = 'A'
                and CCPA_TYPE IN ('S','C')
                and RESPONSE_BATCH_NO IS NULL"

        If chkOverNinety.Checked Then
            sql &= " and NVL(ARTCCPA1.CCPA_DATE_AUTH, SYSDATE) > SYSDATE - 90"
        End If

        ASCDATA1.ExecuteSQL("Truncate table " & tblARTCCPA1)
        ASCDATA1.ExecuteSQL($"INSERT INTO {tblARTCCPA1} {sql}")

        Dim numToBeSettled As Int32 = ASCDATA1.GetDataValue($"SELECT COUNT(*) FROM {tblARTCCPA1}")
        Dim valToBeSettled As Decimal = Val(ASCDATA1.GetDataValue($"SELECT SUM(CCPA_AMT) FROM {tblARTCCPA1}") & String.Empty)

        ASCMAIN1.sql = "Select ARTCCPA1.*, ARTCUST1.CUST_NAME" _
                    & " from ARTCCPA1, ARTCUST1 " _
                    & " where ARTCCPA1.CUST_CODE = ARTCUST1.CUST_CODE (+)" _
                    & " and ARTCCPA1.CCPA_NO IN (SELECT CCPA_NO FROM " & tblARTCCPA1 & ")"

        Fill_Records("ARTCCPA1", String.Empty, True, ASCMAIN1.sql)
        If clsTACENCRY.UseEncryption = True Then
            For Each rowARTCCPA1 As DataRow In dst.Tables("ARTCCPA1").Rows
                For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                    rowARTCCPA1.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1.Item(field & "_E") & String.Empty)
                    rowARTCCPA1.Item(field & "_E") = DBNull.Value
                Next
            Next
        End If

        Fill_Records("ARTCCPDA")
        grdARTCCPA1.Text = "Today's Payments: " & dst.Tables("ARTCCPA1").Rows.Count & " of " & numToBeSettled & " - Total CC Amount: " & Math.Round(valToBeSettled, 2).ToString("#,##0.00")

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default

        ' NEXT BLOCK IS TEMP
        Create_Lookup("ARTSTMT1")
        Dim clsARCCARD As New TAC.ARCCCARD

        For Each rowARTCCPA1 As DataRow In dst.Tables("ARTCCPA1").Select()
            ' confusion in abs stds over whether the lookup has rights to use the CMDs dictionary, which is built by CreateTDA
            Dim rowARTSTMT1 = LookUp("ARTSTMT1", New String() _
            {ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), rowARTCCPA1.Item("CUST_CODE")}, True)
            rowARTCCPA1.Item("TOTAL_DUE") = rowARTSTMT1.ITEM("TOTAL_DUE")
            rowARTCCPA1.Item("AGE_1") = rowARTSTMT1.ITEM("AGE_1")
            rowARTCCPA1.Item("AGE_2") = rowARTSTMT1.ITEM("AGE_2")
            rowARTCCPA1.Item("AGE_3") = rowARTSTMT1.ITEM("AGE_3")
            rowARTCCPA1.Item("AGE_4") = rowARTSTMT1.ITEM("AGE_4")

            Dim CUST_CREDIT_CARD_TYPE As String = rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") & String.Empty
            If CUST_CREDIT_CARD_TYPE.Length = 0 Then
                Dim CUST_CREDIT_CARD_NO As String = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty
                Try
                    clsARCCARD.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
                    Select Case clsARCCARD.GetCreditCardType()
                        Case TAC.ARCCCARD.CreditCardTypes.vctAmex
                            rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = "AMEX"
                        Case TAC.ARCCCARD.CreditCardTypes.vctMasterCard
                            rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = "MSTR"
                        Case TAC.ARCCCARD.CreditCardTypes.vctDiscover
                            rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = "DISC"
                        Case TAC.ARCCCARD.CreditCardTypes.vctVisa, ARCCCARD.CreditCardTypes.vctVisaElectron
                            rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = "VISA"
                        Case TAC.ARCCCARD.CreditCardTypes.vctDiners
                            rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = "DC"
                        Case Else
                            rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = "UNK"
                    End Select
                Catch ex As Exception
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = "UNK"
                End Try
            End If
        Next
        With grdARTCCPA1.DisplayLayout.Bands(0)
            .Columns("TOTAL_DUE").Header.Caption = "Total Due"
            .Columns("AGE_1").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1")
            .Columns("AGE_2").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2")
            .Columns("AGE_3").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3")
            .Columns("AGE_4").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4")
        End With
        Create_Lookup("ARTSTMT1", "*", "STMT_NO = :PARM1", "V", False)

        dst.Tables("ARTCCPS2").Rows.Clear()

        Dim Totals(2, 2) As Decimal

        dst.Tables("ARTCCPA0").Rows.Clear()

        Totals(1, 1) = Val(dst.Tables("ARTCCPA1").Compute("COUNT(CCPA_NO)", "CCPA_STATUS = 'A'") & "")
        Totals(1, 2) = Val(dst.Tables("ARTCCPA1").Compute("SUM(CCPA_AMT)", "CCPA_STATUS = 'A'") & "")
        dst.Tables("ARTCCPA0").Rows.Add(New Object() {1, "Charges Entered", Totals(1, 1), Totals(1, 2)})

        Totals(2, 1) = Totals(1, 1) ' Val(dst.Tables("ARTCCPS2").Compute("SUM(RESPONSE_PYMT_TYPE_TRANS)", "") & "")
        Totals(2, 2) = Totals(1, 2) 'Val(dst.Tables("ARTCCPS2").Compute("SUM(RESPONSE_PYMT_TYPE_NET_AMT)", "") & "")
        dst.Tables("ARTCCPA0").Rows.Add(New Object() {2, "Polled Totals", Totals(2, 1), Totals(2, 2)})

        dst.Tables("ARTCCPA0").Rows.Add(New Object() {3, "Difference", Totals(1, 1) - Totals(2, 1), Totals(1, 2) - Totals(2, 2)})

        Sort_grdColumns(grdARTCCPA0, "LINE", True)

        splCC.Visible = True
        grdARTCCPA0.Tag = ""

        RESPONSE_BATCH_NO = String.Empty
        ' Mark items that can be applied with a RESPONSE_BATCH_NO
        If dst.Tables("ARTCCPA1").Rows.Count > 0 Then
            RESPONSE_BATCH_NO = ASCMAIN1.Next_Control_No("ARTCCPA1.RESPONSE_BATCH_NO")
            For Each rowARTCCPA1 As DataRow In dst.Tables("ARTCCPA1").Select("")
                Dim CCPA_DATE_SALE As String = rowARTCCPA1.Item("CCPA_DATE_SALE") & String.Empty
                If IsDate(CCPA_DATE_SALE) Then
                    If Val(CDate(CCPA_DATE_SALE).ToString("yyyyMMdd")) < Val(DateTime.Now.ToString("yyyyMMdd")) Then
                        rowARTCCPA1.Item("RESPONSE_BATCH_NO") = RESPONSE_BATCH_NO
                    End If
                End If
            Next
        End If

        Dim tblARTCCPS2 As DataTable = ASCDATA1.SelectDistinct("ARTCCPA1", New String() {"RESPONSE_BATCH_NO", "CUST_CREDIT_CARD_TYPE"})
        For Each row2 As DataRow In tblARTCCPS2.Select($"RESPONSE_BATCH_NO = '{RESPONSE_BATCH_NO}'")
            If row2.Item("RESPONSE_BATCH_NO") & String.Empty = "" Then
                Continue For
            End If

            Dim CUST_CREDIT_CARD_TYPE As String = row2.Item("CUST_CREDIT_CARD_TYPE") & String.Empty
            Dim RESPONSE_PYMT_TYPE_NET_AMT As Double = Val(dst.Tables("ARTCCPA1").Compute("SUM(CCPA_AMT)", $"RESPONSE_BATCH_NO = '{RESPONSE_BATCH_NO}' AND CUST_CREDIT_CARD_TYPE = '{CUST_CREDIT_CARD_TYPE}'") & String.Empty)

            Dim rowARTCCPS2 As DataRow = dst.Tables("ARTCCPS2").NewRow
            rowARTCCPS2.Item("RESPONSE_BATCH_NO") = row2.Item("RESPONSE_BATCH_NO")
            rowARTCCPS2.Item("RESPONSE_PYMT_TYPE") = CUST_CREDIT_CARD_TYPE
            rowARTCCPS2.Item("RESPONSE_PYMT_TYPE_NET_AMT") = RESPONSE_PYMT_TYPE_NET_AMT
            rowARTCCPS2.Item("RESPONSE_PYMT_TYPE_TRANS") = dst.Tables("ARTCCPA1").Select($"RESPONSE_BATCH_NO = '{RESPONSE_BATCH_NO}' AND CUST_CREDIT_CARD_TYPE = '{CUST_CREDIT_CARD_TYPE}'").Length
            dst.Tables("ARTCCPS2").Rows.Add(rowARTCCPS2)
        Next
        grdARTCCPS2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        dst.Tables("ARTCCPA1").AcceptChanges()
        grdARTCCPA1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

    End Sub

    Sub Apply_CC_Payment()

        Dim beginDataTrans As Boolean = False
        Dim rowARTCUST1 As DataRow = Nothing

        Try

            BANK_CODE = MyBase.Absx1.txtFor("BANK_CODE").Text.Trim
            If Not Validate_Code("BANK_CODE") Then
                Exit Sub
            End If

            If grdARTCCPA0.Tag = "" Then
                Call Export_to_Excel(New UltraWinGrid.UltraGrid() {grdARTCCPA0, grdARTCCPS2, grdARTCCPA1})
            End If

            Me.ParentForm.Enabled = False
            Application.UseWaitCursor = True
            Application.DoEvents()

            dst.Tables("ARTPYMT1").Rows.Clear()
            dst.Tables("ARTPYMT2").Rows.Clear()
            dst.Tables("ARTPYMT3").Rows.Clear()
            dst.Tables("ARTPYMT4").Rows.Clear()
            dst.Tables("ARTPYMT5").Rows.Clear()
            dst.Tables("ARTOPEN1").Rows.Clear()

            ' Need this outside the processing to prevent locking other screens
            Dim PYMT_BATCH_NO As String = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
            Dim PYMT_BATCH_DATE As Date = Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy")

            dst.Tables("ARTPYMT1").AcceptChanges()
            dst.Tables("ARTPYMT2").AcceptChanges()

            Call BeginTrans()
            beginDataTrans = True

            Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
            With rowARTPYMT1
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                .Item("PYMT_BATCH_DATE") = PYMT_BATCH_DATE
                .Item("BANK_CODE") = BANK_CODE
                .Item("STATUS") = "1"
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = 1
                .Item("PYMT_SOURCE") = "CC"
                .Item("RESPONSE_BATCH_NO") = RESPONSE_BATCH_NO
            End With
            dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

            Dim PYMT_BATCH_LNO As Integer = 0
            For Each rowARTCCPA1 As DataRow In dst.Tables("ARTCCPA1").Select($"CCPA_STATUS = 'A' AND RESPONSE_BATCH_NO = '{RESPONSE_BATCH_NO}'", "CCPA_NO")
                Dim CCPA_NO As String = rowARTCCPA1.Item("CCPA_NO")
                Dim CCPA_NO_CREDITED As String = rowARTCCPA1.Item("CCPA_NO_CREDITED") & ""
                Dim CUST_CODE As String = rowARTCCPA1.Item("CUST_CODE") & ""
                Dim STMT_NO As String = rowARTCCPA1.Item("STMT_NO") & ""
                Dim ORDR_NO As String = rowARTCCPA1.Item("ORDR_NO") & ""
                Dim INV_NO As String = rowARTCCPA1.Item("INV_NO") & ""
                Dim LENS_BANK_INV_NO As String = rowARTCCPA1.Item("LENS_BANK_INV_NO") & ""
                Dim CCPA_TYPE As String = rowARTCCPA1.Item("CCPA_TYPE") & ""
                Dim CCPA_AMT As Decimal = Val(rowARTCCPA1.Item("CCPA_AMT"))
                Dim CCPA_REASON As String = rowARTCCPA1.Item("CCPA_REASON") & ""

                Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
                rowARTPYMT2.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                PYMT_BATCH_LNO += 1
                rowARTPYMT2.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                rowARTPYMT2.Item("CUST_CODE") = CUST_CODE
                rowARTPYMT2.Item("CUST_NAME") = rowARTCCPA1.Item("CUST_NAME") & ""

                rowARTPYMT2.Item("CUST_PYMT_AMT") = CCPA_AMT
                rowARTPYMT2.Item("CUST_PYMT_AMT_CURR") = CCPA_AMT
                rowARTPYMT2.Item("PYMT_STATUS") = "1"
                rowARTPYMT2.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
                rowARTPYMT2.Item("CUST_CREDIT_CARD_TYPE") = rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE")

                rowARTPYMT2.Item("CURR_CODE") = CURR_CODE ' added by DRC 08/29/2016 - Currency Code changes
                rowARTPYMT2.Item("CURR_EXCH_RATE") = 1

                If Val(rowARTCCPA1.Item("WEB_PYMT_ID") & String.Empty) > 0 Then
                    rowARTPYMT2.Item("CUST_PYMT_WEB_PYMT_ID") = Val(rowARTCCPA1.Item("WEB_PYMT_ID") & String.Empty)
                End If

                rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                dst.Tables("ARTOPEN1").Rows.Clear()

                If rowARTCCPA1.Item("CCPA_TYPE") = "C" AndAlso CCPA_NO_CREDITED.Length > 0 Then ' CC Credit - need to Reverse the Previous Application
                    rowARTPYMT2.Item("PYMT_STATUS") = "2"
                    rowARTPYMT2.Item("LAST_DATE") = DATETIME_STAMP
                    rowARTPYMT2.Item("LAST_OPER") = ASCMAIN1.USER_ID

                    ASCMAIN1.sql = "Select * from ARTPYMT2 where CCPA_NO = '" & CCPA_NO_CREDITED & "'"
                    Dim rowARTPYMT2_to_reverse As DataRow = ASCDATA1.GetDataRow
                    Dim PYMT_BATCH_NO_to_reverse As String = rowARTPYMT2_to_reverse.Item("PYMT_BATCH_NO")
                    Dim PYMT_BATCH_LNO_to_reverse As Integer = rowARTPYMT2_to_reverse.Item("PYMT_BATCH_LNO")

                    For Each TABLE_NAME As String In New String() {"ARTPYMT3", "ARTPYMT4", "ARTPYMT5"}
                        Dim tbl As DataTable = dst.Tables(TABLE_NAME).Clone
                        Fill_Records(TABLE_NAME, New Object() {PYMT_BATCH_NO_to_reverse, PYMT_BATCH_LNO_to_reverse}, , , tbl)
                        For Each row As DataRow In tbl.Rows

                            'MsgBox("ABS Stop Statement - Reversing a Payment when Crediting a CC") 'Stop ' MUST CHECK THIS
                            row.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                            row.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                            Select Case TABLE_NAME
                                Case "ARTPYMT3"
                                    Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").Rows.Find _
                                    (New Object() {CUST_CODE, row.Item("INV_TYPE"), row.Item("INV_NUM")})

                                    If "" <> "" And rowARTOPEN1 Is Nothing Then
                                        rowARTOPEN1 = Fill_Record("ARTOPEN1",
                                            New Object() {CUST_CODE, row.Item("INV_TYPE"), row.Item("INV_NUM")}, , False)
                                    End If

                                    If rowARTOPEN1 Is Nothing Then
                                        ASCMAIN1.sql = "Insert into ARTOPEN1 " _
                                        & " Select * from ARTOPENX " _
                                        & " where CUST_CODE = :PARM1 " _
                                        & "   and INV_TYPE = :PARM2 " _
                                        & "   and INV_NUM = :PARM3"
                                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {CUST_CODE, row.Item("INV_TYPE"), row.Item("INV_NUM")})

                                        ASCMAIN1.sql = "" _
                                        & " Delete from ARTOPENX " _
                                        & " where CUST_CODE = :PARM1 " _
                                        & "   and INV_TYPE = :PARM2 " _
                                        & "   and INV_NUM = :PARM3"
                                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {CUST_CODE, row.Item("INV_TYPE"), row.Item("INV_NUM")})
                                        rowARTOPEN1 = Fill_Record("ARTOPEN1",
                                        New Object() {CUST_CODE, row.Item("INV_TYPE"), row.Item("INV_NUM")}, , False)
                                    End If

                                    TAC.ARCMAIN1.Pay_Open_AR_Item(rowARTOPEN1, rowARTPYMT2,
                                        CURR_CODE, PYMT_BATCH_DATE,
                                        -1 * Val(row.Item("INV_PMT") & ""), -1 * Val(row.Item("INV_DISC_TAKEN") & ""),
                                        -1 * Val(row.Item("INV_WRITE_OFF") & ""), Val(row.Item("PYMT_BATCH_ILNO")) - 1, Me)

                                Case "ARTPYMT4"
                                    row.Item("GL_DIST_AMT") = -1 * Val(row.Item("GL_DIST_AMT") & "")
                                    row.Item("GL_DIST_AMT_CURR") = -1 * Val(row.Item("GL_DIST_AMT_CURR") & "")
                                    row.AcceptChanges()
                                    row.SetAdded()

                                Case "ARTPYMT5"
                                    row.Item("GL_DIST_AMT") = -1 * Val(row.Item("GL_DIST_AMT") & "")
                                    row.Item("GL_DIST_AMT_CURR") = -1 * Val(row.Item("GL_DIST_AMT_CURR") & "")
                                    Load_Open_AR_from_CB(row, CUST_CODE, PYMT_BATCH_DATE)
                                    row.AcceptChanges()
                                    row.SetAdded()
                            End Select
                        Next

                        If TABLE_NAME <> "ARTPYMT3" Then ' BECAUSE THESE ROWS ARE ADDED IN Pay_Open_AR_Item
                            dst.Tables(TABLE_NAME).Merge(tbl, True)
                        End If
                    Next
                Else
                    Select Case rowARTCCPA1.Item("CCPA_REASON")
                        Case "S" ' Statement Payment
                            Dim rowARTSTMT1 As DataRow = LookUp("ARTSTMT1", STMT_NO)
                            If rowARTSTMT1 IsNot Nothing Then
                                If rowARTSTMT1.Item("CUST_CODE") & "" = CUST_CODE Then
                                    Dim TOTAL_DUE As Decimal = Val(rowARTSTMT1.Item("TOTAL_DUE") & "")

                                    Dim sqlARTOPEN1 As String = "Select Sum (ARTOPEN1.INV_BALANCE) from ARTOPEN1" _
                                    & " where ARTOPEN1.CUST_CODE = '" & CUST_CODE & "'" _
                                    & "   and NVL(ARTOPEN1.INV_BALANCE,0) <> 0"
                                    Dim TOTAL_ARTOPEN1 As Decimal = Val(ASCDATA1.GetDataValue(sqlARTOPEN1))

                                    Dim sqlARTSTMT2 As String = "Select Sum (ARTOPEN1.INV_BALANCE) " _
                                    & " from ARTOPEN1,ARTSTMT1,ARTSTMT2" _
                                    & " where ARTSTMT1.STMT_NO = '" & STMT_NO & "'" _
                                    & "   and ARTSTMT2.OPS_YYYYPP = ARTSTMT1.OPS_YYYYPP" _
                                    & "   and ARTSTMT2.CUST_CODE = ARTSTMT1.CUST_CODE" _
                                    & "   and ARTOPEN1.CUST_CODE = ARTSTMT2.CUST_CODE" _
                                    & "   and ARTOPEN1.INV_TYPE = ARTSTMT2.INV_TYPE" _
                                    & "   and ARTOPEN1.INV_NUM = ARTSTMT2.INV_NUM" _
                                    & "   and NVL(ARTOPEN1.INV_BALANCE,0) <> 0"
                                    Dim TOTAL_STILL_DUE As Decimal = Val(ASCDATA1.GetDataValue(sqlARTSTMT2))

                                    Dim sql As String = ""
                                    If CCPA_AMT = TOTAL_STILL_DUE Then
                                        sql = Replace(sqlARTSTMT2, "Sum (ARTOPEN1.INV_BALANCE)", "ARTOPEN1.*")
                                    ElseIf CCPA_AMT = TOTAL_ARTOPEN1 Then
                                        sql = Replace(sqlARTOPEN1, "Sum (ARTOPEN1.INV_BALANCE)", "ARTOPEN1.*")
                                    End If

                                    If sql <> "" Then
                                        Fill_Records("ARTOPEN1", , , sql)
                                        Dim INV_PMT_TOTAL As Decimal = 0
                                        Dim PYMT_BATCH_ILNO As Integer = 0
                                        For Each rowARTOPEN1 As DataRow In dst.Tables("ARTOPEN1").Rows
                                            Dim INV_BALANCE As Double = Val(rowARTOPEN1.Item("INV_BALANCE") & "")
                                            Dim INV_PMT As Double = INV_BALANCE
                                            INV_PMT_TOTAL += INV_PMT
                                            ' WE WILL HAVE TO HANDLE THESE WHEN WE DO A BUYING GROUP WITH ANTIC DISC
                                            Dim INV_DISC_TAKEN As Double = 0
                                            Dim INV_WRITE_OFF As Double = 0

                                            TAC.ARCMAIN1.Pay_Open_AR_Item(rowARTOPEN1, rowARTPYMT2,
                                            CURR_CODE, PYMT_BATCH_DATE,
                                            INV_PMT, INV_DISC_TAKEN, INV_WRITE_OFF, PYMT_BATCH_ILNO, Me)
                                        Next
                                        If System.Math.Round(CCPA_AMT, 2) <> System.Math.Round(INV_PMT_TOTAL, 2) Then
                                            MsgBox("ABS Stop Statment - Amounts not equal " & CStr(CCPA_AMT) & " vs " & CStr(INV_PMT_TOTAL)) ' Stop ' something is very wrong
                                        End If
                                        rowARTPYMT2.Item("PYMT_STATUS") = "2"
                                        rowARTPYMT2.Item("LAST_DATE") = DATETIME_STAMP
                                        rowARTPYMT2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                    End If
                                End If
                            End If

                        Case "C" ' COD CC Payment of a Specific Invoice
                            Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").Rows.Find _
                            (New Object() {CUST_CODE, "I", INV_NO})

                            If rowARTOPEN1 Is Nothing Then
                                rowARTOPEN1 = Fill_Record("ARTOPEN1", New Object() {CUST_CODE, "I", INV_NO}, , False)
                            End If

                            If rowARTOPEN1 IsNot Nothing AndAlso Val(rowARTOPEN1.Item("INV_BALANCE") & "") = CCPA_AMT Then
                                TAC.ARCMAIN1.Pay_Open_AR_Item(rowARTOPEN1, rowARTPYMT2,
                                CURR_CODE, PYMT_BATCH_DATE, CCPA_AMT, 0, 0, 0, Me)
                                rowARTPYMT2.Item("PYMT_STATUS") = "2"
                                rowARTPYMT2.Item("LAST_DATE") = DATETIME_STAMP
                                rowARTPYMT2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            Else
                                ' See if this was a situation where multiple Invoices where Charged at once
                                Fill_Records("ARTCCPAM", String.Empty, True, "SELECT * FROM ARTCCPAM WHERE CCPA_NO = '" & CCPA_NO & "'")
                                Dim TOTAL_CHARGED As Decimal = CCPA_AMT
                                Dim INV_BALANCE As Decimal = 0
                                For Each row As DataRow In dst.Tables("ARTCCPAM").Select("", "INV_BALANCE")
                                    If TOTAL_CHARGED = 0 Then Exit For
                                    INV_BALANCE = (row.Item("INV_BALANCE") & String.Empty)
                                    If TOTAL_CHARGED < INV_BALANCE Then
                                        INV_BALANCE = TOTAL_CHARGED
                                    End If

                                    rowARTOPEN1 = Fill_Record("ARTOPEN1", New Object() {row.Item("CUST_CODE"), row.Item("INV_TYPE"), row.Item("INV_NUM")}, , False)

                                    TAC.ARCMAIN1.Pay_Open_AR_Item(rowARTOPEN1, rowARTPYMT2,
                                        CURR_CODE, PYMT_BATCH_DATE, INV_BALANCE, 0, 0, 0, Me)

                                    TOTAL_CHARGED -= INV_BALANCE
                                Next

                            End If
                    End Select
                End If

                dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)
                rowARTCCPA1.Item("CCPA_STATUS") = "S"

                Update_Record_TDA("ARTOPEN1")

            Next

            ' Moved down here to prevent locking with the warehouse
            INIT_LAST("ARTPYMT1")
            Update_Record_TDA("ARTPYMT1")
            Update_Record_TDA("ARTPYMT2")

            For Each rowARTPYMT2 As DataRow In dst.Tables("ARTPYMT2").Rows
                ASCDATA1.ExecuteSP("ARPCUST6_PYMT", "VN" _
                                   , New Object() {rowARTPYMT2.Item("PYMT_BATCH_NO"), rowARTPYMT2.Item("PYMT_BATCH_LNO")} _
                                   , New String() {"PYMT_BATCH_NO_IN", "PYMT_BATCH_LNO_IN"})
            Next

            Update_Record_TDA("ARTPYMT1")
            Update_Record_TDA("ARTPYMT2")
            Update_Record_TDA("ARTPYMT3")
            Update_Record_TDA("ARTPYMT4")
            Update_Record_TDA("ARTPYMT5")

            If clsTACENCRY.UseEncryption = True Then
                For Each rowARTCCPA1 As DataRow In dst.Tables("ARTCCPA1").Rows
                    For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                        rowARTCCPA1.Item(field & "_E") = clsTACENCRY.EncryptString(rowARTCCPA1.Item(field & "_E") & String.Empty)
                        rowARTCCPA1.Item(field) = DBNull.Value
                    Next
                Next
            End If
            Update_Record_TDA("ARTCCPA1")

            CommitTrans("Settlement Complete")

        Catch ex As Exception
            Me.ParentForm.Enabled = True
            Application.UseWaitCursor = False
            MessageBox.Show("Unable to Settle with Merchant:" & Environment.NewLine & " Error: " & ex.Message, "Settlement", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If beginDataTrans Then
                Rollback(ex.Message)
            End If

            Exit Sub

        Finally
            Me.ParentForm.Enabled = True
            Application.UseWaitCursor = False
        End Try

    End Sub

    Sub Load_Open_AR_from_CB(ByVal rowARTPYMT5 As DataRow, ByVal CUST_CODE As String, ByVal PYMT_BATCH_DATE As Date)

        Dim GL_DIST_AMT As Double = Val(rowARTPYMT5.Item("GL_DIST_AMT") & "")
        Dim INV_TYPE_CB As String = rowARTPYMT5.Item("INV_TYPE_CB") & ""
        If INV_TYPE_CB = "" Then
            If GL_DIST_AMT < 0 Then
                INV_TYPE_CB = "O"
            Else
                INV_TYPE_CB = "B"
            End If
            rowARTPYMT5.Item("INV_TYPE_CB") = INV_TYPE_CB
        End If

        Dim CHARGEBACK_NO As String = ASCMAIN1.Next_Control_No("INV_NUM_" & INV_TYPE_CB)
        rowARTPYMT5.Item("CHARGEBACK_NO") = CHARGEBACK_NO

        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        rowARTOPEN1.Item("CUST_CODE") = CUST_CODE
        rowARTOPEN1.Item("INV_TYPE") = INV_TYPE_CB
        rowARTOPEN1.Item("INV_NUM") = CHARGEBACK_NO
        rowARTOPEN1.Item("INV_DATE") = PYMT_BATCH_DATE
        rowARTOPEN1.Item("INV_DUE_DATE") = rowARTOPEN1.Item("INV_DATE")
        ' rowARTOPEN1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE") & ""
        rowARTOPEN1.Item("INV_CUST_PO") = rowARTPYMT5.Item("CUST_REFERENCE")
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT") = GL_DIST_AMT
        rowARTOPEN1.Item("INV_BALANCE") = GL_DIST_AMT
        rowARTOPEN1.Item("REASON_CODE") = rowARTPYMT5.Item("REASON_CODE")
        rowARTOPEN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTOPEN1.Item("INIT_DATE") = DATETIME_STAMP
        rowARTOPEN1.Item("INV_MISC_CHG") = GL_DIST_AMT
        rowARTOPEN1.Item("CURR_CODE") = CURR_CODE
        rowARTOPEN1.Item("CURR_EXCH_RATE") = 1 ' CURR_EXCH_RATE
        rowARTOPEN1.Item("INV_MISC_CHG_CURR") = GL_DIST_AMT
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = GL_DIST_AMT
        rowARTOPEN1.Item("INV_BALANCE_CURR") = GL_DIST_AMT
        Dim ORDR_TYPE_CODE As String
        Dim rowSOTTYPE1 As DataRow
        If INV_TYPE_CB = "O" Then
            ORDR_TYPE_CODE = "ONA"
        Else
            ORDR_TYPE_CODE = "CHB"
        End If
        rowSOTTYPE1 = dst.Tables("SOTTYPE1").Rows.Find(ORDR_TYPE_CODE)

        Dim POST_CODE As String = rowSOTTYPE1.Item("POST_CODE")
        Dim rowARTPOST1 As DataRow = dst.Tables("ARTPOST1").Rows.Find(POST_CODE)

        rowARTOPEN1.Item("SEG2_CODE") = rowARTPOST1.Item("SEG2_CODE")
        rowARTOPEN1.Item("SEG3_CODE") = rowARTPOST1.Item("SEG3_CODE")
        rowARTOPEN1.Item("SEG4_CODE") = rowARTPOST1.Item("SEG4_CODE")

        rowARTOPEN1.Item("POST_CODE") = rowSOTTYPE1.Item("POST_CODE")
        rowARTOPEN1.Item("TERM_CODE") = rowSOTTYPE1.Item("TERM_CODE")
        If rowARTOPEN1.Item("TERM_CODE") & "" = "" Then
            rowARTOPEN1.Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0")
        End If

        rowARTOPEN1.Item("ORDR_TYPE_CODE") = ORDR_TYPE_CODE
        rowARTOPEN1.Item("INV_REF") = rowARTPYMT5.Item("OUR_REFERENCE")
        rowARTOPEN1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

        rowARTPYMT5.Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
        rowARTPYMT5.Item("SEG2_CODE") = rowARTPOST1.Item("SEG2_CODE")
        rowARTPYMT5.Item("SEG3_CODE") = rowARTPOST1.Item("SEG3_CODE")
        rowARTPYMT5.Item("SEG4_CODE") = rowARTPOST1.Item("SEG4_CODE")

    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdARTCCPA0_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCCPA0.InitializeRow
        If e.Row.Cells("TEXT").Text = "Difference" Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGray
            If Val(e.Row.Cells("TRANS").Value & "") = 0 Then
                e.Row.Cells("TRANS").Appearance.ForeColor = Drawing.Color.Green
            Else
                e.Row.Cells("TRANS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("TEXT").Appearance.ForeColor = Drawing.Color.Red
            End If
            If Val(e.Row.Cells("AMT").Value & "") = 0 Then
                e.Row.Cells("AMT").Appearance.ForeColor = Drawing.Color.Green
            Else
                e.Row.Cells("AMT").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("TEXT").Appearance.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Private Sub grdARTCCPA1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCCPA1.InitializeRow
        Dim CCPA_DATE_AUTH As String = e.Row.Cells("CCPA_DATE_AUTH").Text

        If e.Row.Cells("RESPONSE_BATCH_NO").Text = String.Empty Then
            e.Row.Appearance.BackColor = Drawing.Color.Green
        ElseIf IsDate(CCPA_DATE_AUTH) AndAlso DateDiff(DateInterval.Day, CDate(CCPA_DATE_AUTH), DateTime.Now) > 90 Then
            e.Row.Appearance.BackColor = Drawing.Color.Red
        End If

    End Sub

#End Region

End Class