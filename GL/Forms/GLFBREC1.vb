Public Class GLFBREC1

    Dim rowGLTBANK1 As DataRow
    Dim rowGLTBREC1 As DataRow
    Dim rowGLTPARM2 As DataRow

    Dim BATCH_NO_CLEARED As String
    Dim BANK_CODE As String
    Dim OPS_YYYYPP As String
    Dim BANK_STMT_BALANCE_previous As Decimal

    Dim GLTBREC2 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")

        With dst
            '  Create_TDA(.Tables.Add, "GLTBREC1", "*")

            ASCMAIN1.sql = "Select GLTBREC1.*, GLTBANK1.BANK_DESC, GLTPARM2.LEGEND" _
                & " from GLTBREC1,GLTBANK1,GLTPARM2" _
                & " where GLTBANK1.BANK_CODE = GLTBREC1.BANK_CODE" _
                & "   and GLTPARM2.OPS_YYYYPP = GLTBREC1.OPS_YYYYPP" _
                & "   and GLTBREC1.BANK_CODE = NVL(:PARM1,GLTBREC1.BANK_CODE)"
            Create_TDA(.Tables.Add, "GLTBRECX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select GLTBREC1.*, GLTBANK1.BANK_DESC, GLTPARM2.LEGEND" _
                & " from GLTBREC1,GLTBANK1,GLTPARM2" _
                & " where GLTBANK1.BANK_CODE = GLTBREC1.BANK_CODE" _
                & "   and GLTPARM2.OPS_YYYYPP = GLTBREC1.OPS_YYYYPP" _
                & "   and GLTBREC1.BATCH_NO_CLEARED = :PARM1"
            Create_TDA(.Tables.Add, "GLTBREC1", "**", 0, True, "V")

            GLTBREC2 = ASCMAIN1.Temp_Table("Select GLTBREC2.* from GLTBREC2 where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & GLTBREC2 & " Add Primary Key (BATCH_NO_CLEARED,JOURNAL_TYPE,TRAN_YP,TRAN_KEY,TRAN_KEY_LNO)")
            ASCMAIN1.sql = "Select * from " & GLTBREC2
            Create_TDA(.Tables.Add("GLTBREC2"), GLTBREC2, "**", 0)
            .Tables("GLTBREC2").Columns("TRAN_SEL").DefaultValue = "0"
            With .Tables("GLTBREC2").Columns
                .Add("TRAN_AMT_ARCR", GetType(System.Decimal), "IIF(JOURNAL_TYPE='ARCR',IIF(TRAN_SEL='1',TRAN_AMT,0),NULL)")
                .Add("TRAN_AMT_APCD", GetType(System.Decimal), "IIF(JOURNAL_TYPE='APCD',IIF(TRAN_SEL='1',TRAN_AMT,0),NULL)")
                .Add("TRAN_AMT_GLJE", GetType(System.Decimal), "IIF(JOURNAL_TYPE='ARCR' OR JOURNAL_TYPE='APCD',NULL,IIF(TRAN_SEL='1',TRAN_AMT,0))")
            End With

            With .Tables.Add("GLTBRECT")
                .Columns.Add("T_LNO", GetType(System.Int16))
                .Columns.Add("T_DESC")
                .Columns.Add("T_AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("T_LNO")}
            End With

            ASCMAIN1.sql = "Select * from GLTTYPE1"
            Create_TDA(.Tables.Add, "GLTTYPE1", "**", 0, False)
        End With

        Fill_Records("GLTTYPE1")

        grdGLTBREC2.DataSource = dst.Tables("GLTBREC2")
        grdGLTBRECX.DataSource = dst.Tables("GLTBRECX")
        grdGLTBRECT.DataSource = dst.Tables("GLTBRECT")

        With grdGLTBREC2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "TRAN_SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        Create_Summary(grdGLTBREC2, "TRAN_KEY", "Count")
        Create_Summary(grdGLTBREC2, New String() {"TRAN_AMT", "TRAN_SEL", "TRAN_AMT_ARCR", "TRAN_AMT_APCD", "TRAN_AMT_GLJE"})

        Create_Summary(grdGLTBRECX, "BATCH_NO_CLEARED", "Count")

        With dst.Tables("GLTBRECT").Rows
            .Add(New Object() {1, "Prev Balance", 0})
            .Add(New Object() {2, "Deposits", 0})
            .Add(New Object() {3, "Disbursements", 0})
            .Add(New Object() {4, "Adjustments", 0})
            .Add(New Object() {5, "Roll Forward", 0})
            .Add(New Object() {6, "Stmt Balance", 0})
            .Add(New Object() {7, "Difference", 0})
        End With
        Sort_grdColumns(grdGLTBRECT, "T_LNO", True)

        Bind_Controls(grpStatement, "GLTBREC1")
        'Set_Read_Only_for_ctl(Absx1.numFor("INV_TOTAL_AMOUNT"), True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("BANK_CODE")

                If EMsg = "" Then
                    BANK_CODE = Absx1.txtFor("BANK_CODE").Text
                    If Not ASCMAIN1.Logical_Lock("GLTBREC1", BANK_CODE) Then Exit Sub

                    ASCMAIN1.sql = "Select Max (OPS_YYYYPP) from GLTBREC1 where BANK_CODE = '" & BANK_CODE & "'"
                    OPS_YYYYPP = ASCDATA1.GetDataValue
                    Dim LYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)

                    If OPS_YYYYPP = "" Then
                        BANK_STMT_BALANCE_previous = 0
                        OPS_YYYYPP = LYP
                    Else
                        If OPS_YYYYPP >= LYP Then
                            EMsg &= vbCr & "The last Bank Reconciliation for Bank " & BANK_CODE & " was for Period " & OPS_YYYYPP _
                                & vbCr & "The next Bank Reconciliation will be for the Current Period: " & ASCMAIN1.CYP _
                                & vbCr & "Cannot start a new Bank Reconcilation until the current Period is Closed"
                        Else
                            ASCMAIN1.sql = "Select BANK_STMT_BALANCE from GLTBREC1" _
                                 & " where BANK_CODE = '" & BANK_CODE & "' and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                            BANK_STMT_BALANCE_previous = Val(ASCDATA1.GetDataValue)
                            OPS_YYYYPP = ASCMAIN1.Period_Calc(OPS_YYYYPP, 1)
                        End If
                    End If
                End If

                If EMsg <> "" Then ASCMAIN1.MultiTask_Release()

            Case "View", "Edit"

                If Absx1.txtFor("BATCH_NO_CLEARED").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Batch No"
                Else
                    BATCH_NO_CLEARED = Absx1.txtFor("BATCH_NO_CLEARED").Text
                    ASCMAIN1.sql = "Select * from GLTBREC1 where BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
                    Dim rowGLTBREC1 As DataRow = ASCDATA1.GetDataRow
                    If rowGLTBREC1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Batch No"
                    Else
                        BANK_CODE = rowGLTBREC1.Item("BANK_CODE")

                        If eItemKey = "Edit" Then
                            If Not ASCMAIN1.Logical_Lock("GLTBREC1", BANK_CODE) Then Exit Sub
                            ASCMAIN1.sql = "Select Max (OPS_YYYYPP) from GLTBREC1 where BANK_CODE = '" & BANK_CODE & "'"
                            OPS_YYYYPP = ASCDATA1.GetDataValue
                            If OPS_YYYYPP <> rowGLTBREC1.Item("OPS_YYYYPP") Then
                                EMsg &= vbCr & "Cannot Edit this Bank Reconciliation" _
                                    & vbCr & "- Last Bank Reconciliation for Bank " & BANK_CODE & " was for Period: " & OPS_YYYYPP
                            End If
                        End If
                    End If
                End If

                If EMsg <> "" And eItemKey = "Edit" Then ASCMAIN1.MultiTask_Release()

            Case "Update"

                Dim DT As Date = Absx1.dteFor("BANK_STMT_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Statement Date is Mandatory"
                End If

            Case "Delete"

                If EMsg = "" Then
                    If MsgBox("Do you really want to Delete this Bank Reconciliation", _
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Print"
                Print_Record()
                'Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" And EntryMode <> "E" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = DefaultableBoolean.False
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                        .Items("Edit").Settings.Enabled = iScreenMode
                    End If

                    .Items("Delete").Visible = ScreenMode And EntryMode = "E"
                    .Items("Print").Settings.Enabled = iScreenMode
                End With

                .Groups("Statement").Visible = ScreenMode
                .Groups("Reconciliation").Visible = ScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If EntryMode = "N" Or EntryMode = "E" Then
            grdGLTBREC2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            Set_Read_Only(grpStatement, False)
        Else
            grdGLTBREC2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Set_Read_Only(grpStatement, True)
            Set_Read_Only_for_ctl(chkShowFuture, False)
        End If


        grpDetails.Visible = tf
        grdGLTBREC2.Visible = tf
        grdGLTBRECX.Visible = Not tf

        Set_Read_Only(grpDetails, (EntryMode = "V"))

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GLTBREC1", "GLTBREC2", "GLTBRECX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        grdGLTBREC2.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowGLTBREC1 = dst.Tables("GLTBREC1").NewRow
            BATCH_NO_CLEARED = ASCMAIN1.Next_Control_No("GLTBREC1.BATCH_NO_CLEARED")
            rowGLTBREC1.Item("BATCH_NO_CLEARED") = BATCH_NO_CLEARED
            rowGLTBREC1.Item("BANK_CODE") = BANK_CODE
            rowGLTBREC1.Item("OPS_YYYYPP") = OPS_YYYYPP
            rowGLTBREC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowGLTBREC1.Item("INIT_DATE") = DATETIME_STAMP

            rowGLTPARM2 = LookUp("GLTPARM2", OPS_YYYYPP)
            Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")
            rowGLTBREC1.Item("BANK_STMT_DATE") = PRD_END_DATE

            dst.Tables("GLTBREC1").Rows.Add(rowGLTBREC1)
            Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", rowGLTBREC1.Item("BANK_CODE") & "")
        Else
            rowGLTBREC1 = Fill_Record("GLTBREC1", New String() {Absx1.txtFor("BATCH_NO_CLEARED").Text})
            OPS_YYYYPP = rowGLTBREC1.Item("OPS_YYYYPP")
            rowGLTPARM2 = LookUp("GLTPARM2", OPS_YYYYPP)
        End If

        ASCMAIN1.sql = "Select BANK_STMT_BALANCE from GLTBREC1" _
                              & " where BANK_CODE = '" & BANK_CODE & "' and OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(OPS_YYYYPP, -1) & "'"
        BANK_STMT_BALANCE_previous = Val(ASCDATA1.GetDataValue)

        rowGLTBANK1 = LookUp("GLTBANK1", BANK_CODE)
        UltraExplorerBar1.Groups("Statement").Text = "Statement " & rowGLTPARM2.Item("LEGEND")

        ASCDATA1.ExecuteSQL("Truncate Table " & GLTBREC2)
        ASCMAIN1.sql = "Insert into " & GLTBREC2 _
            & " Select GLTBREC2.* from GLTBREC2" _
            & " where GLTBREC2.BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
        ASCDATA1.ExecuteSQL()

        If EntryMode = "N" Then

            ASCMAIN1.sql = "Select '" & BATCH_NO_CLEARED & "' BATCH_NO_CLEARED, 'APCD' JOURNAL_TYPE" & vbCrLf _
                & ", OPS_YYYYPP TRAN_YP, CHECK_NUM TRAN_KEY, 0 TRAN_KEY_LNO" & vbCrLf _
                & ", CHECK_DATE TRAN_DATE, -1 * CHECK_AMT TRAN_AMT, VEND_NAME TRAN_DESC, '0' TRAN_SEL" _
                & " from APTCHCK1 where BANK_CODE = '" & BANK_CODE & "' and BATCH_NO_CLEARED is Null" & vbCrLf _
                & " and OPS_YYYYPP <= '" & OPS_YYYYPP & "'"
            ASCDATA1.ExecuteSQL("Insert into " & GLTBREC2 & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = "Select '" & BATCH_NO_CLEARED & "' BATCH_NO_CLEARED, 'APCD' JOURNAL_TYPE" & vbCrLf _
                & ", OPS_YYYYPP_F TRAN_YP, CHECK_NUM TRAN_KEY, 1 TRAN_KEY_LNO" & vbCrLf _
                & ", CHECK_DATE TRAN_DATE, CHECK_AMT TRAN_AMT, VEND_NAME TRAN_DESC, '0' TRAN_SEL" _
                & " from APTCHCK1 where BANK_CODE = '" & BANK_CODE & "' and BATCH_NO_CLEARED_F is Null" & vbCrLf _
                & " and CHECK_STATUS = 'V'" & vbCrLf _
                & " and OPS_YYYYPP_F <= '" & OPS_YYYYPP & "'"
            ASCDATA1.ExecuteSQL("Insert into " & GLTBREC2 & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = "Select '" & BATCH_NO_CLEARED & "' BATCH_NO_CLEARED, 'ARCR' JOURNAL_TYPE" & vbCrLf _
                & ", ARTPYMT1.OPS_YYYYPP TRAN_YP, ARTPYMT1.PYMT_BATCH_NO TRAN_KEY, 0 TRAN_KEY_LNO" & vbCrLf _
                & ", ARTPYMT1.PYMT_BATCH_DATE TRAN_DATE, SUM (ARTPYMT2.CUST_PYMT_AMT) TRAN_AMT, 'Payments:' || Count (*) TRAN_DESC, '0' TRAN_SEL" & vbCrLf _
                & " from ARTPYMT1,ARTPYMT2 where ARTPYMT1.BANK_CODE = '" & BANK_CODE & "' and ARTPYMT1.BATCH_NO_CLEARED is Null" & vbCrLf _
                & " and ARTPYMT1.OPS_YYYYPP <= '" & OPS_YYYYPP & "'" & vbCrLf _
                & " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" & vbCrLf _
                & " group by ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_NO, ARTPYMT1.PYMT_BATCH_DATE"
            ASCDATA1.ExecuteSQL("Insert into " & GLTBREC2 & " " & ASCMAIN1.sql)

        End If

        If EntryMode = "N" Or EntryMode = "E" Then

            ASCMAIN1.sql = "Select OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO, ACCT_CODE, DETL_POSTING_AMT, DETL_CTL_DATE" & vbCrLf _
                & " from GLTDETL1 where (OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO) in (" & vbCrLf _
                & "Select GLTDETL1.OPS_YYYYPP, GLTDETL1.JOURNAL_NO, GLTDETL1.JOURNAL_LNO" & vbCrLf _
                & " from GLTDETL1,GLTJRNL1" & vbCrLf _
                & " where GLTDETL1.ACCT_CODE = '" & rowGLTBANK1.Item("ACCT_CODE") & "'" & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO " & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_TYPE = 'GLJE'" & vbCrLf _
                & " minus " & vbCrLf _
                & "Select OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO from GLTBREC0" & vbCrLf _
                & " where ACCT_CODE = '" & rowGLTBANK1.Item("ACCT_CODE") & "'" & vbCrLf _
                & ")"
            ASCDATA1.ExecuteSQL("Insert into GLTBREC0 (OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO, ACCT_CODE, DETL_POSTING_AMT, DETL_CTL_DATE) " & ASCMAIN1.sql)

            ASCMAIN1.sql = "Delete from " & GLTBREC2 & " where JOURNAL_TYPE = 'GLJE' and NVL(TRAN_SEL,'0') <> '1'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select '" & BATCH_NO_CLEARED & "' BATCH_NO_CLEARED, 'GLJE' JOURNAL_TYPE" & vbCrLf _
                & ", GLTBREC0.OPS_YYYYPP TRAN_YP, GLTBREC0.JOURNAL_NO TRAN_KEY, GLTBREC0.JOURNAL_LNO TRAN_KEY_LNO" & vbCrLf _
                & ", GLTBREC0.DETL_CTL_DATE TRAN_DATE, GLTBREC0.DETL_POSTING_AMT TRAN_AMT, GLTJRNL1.JOURNAL_DESC TRAN_DESC, '0' TRAN_SEL" & vbCrLf _
                & " from GLTBREC0,GLTJRNL1 where GLTBREC0.ACCT_CODE = '" & rowGLTBANK1.Item("ACCT_CODE") & "' and GLTBREC0.BATCH_NO_CLEARED is Null" & vbCrLf _
                & " and GLTBREC0.OPS_YYYYPP <= '" & OPS_YYYYPP & "'" & vbCrLf _
                & " and GLTJRNL1.JOURNAL_NO = GLTBREC0.JOURNAL_NO"
            ASCDATA1.ExecuteSQL("Insert into " & GLTBREC2 & " " & ASCMAIN1.sql)

        End If

        Fill_Records("GLTBREC2")
        Set_RowFilter()
        Sort_grdColumns(grdGLTBREC2, "TRAN_DATE,JOURNAL_TYPE,TRAN_KEY,TRAN_KEY_LNO")

        Display_Totals()
        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        ASCDATA1.ExecuteSQL("Truncate Table " & GLTBREC2)

        BeginTrans()

        If EntryMode = "E" Then
            Dependent_Updates(-1)
        End If

        Update_Record_TDA("GLTBREC1")
        Update_Record_TDA("GLTBREC2", "1=1")

        ASCDATA1.ExecuteSQL("Delete from GLTBREC2 where BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'")
        ASCDATA1.ExecuteSQL("Insert into GLTBREC2 Select * from " & GLTBREC2)

        Dependent_Updates(1)

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Dependent_Updates(-1)
        Delete_Records("GLTBREC1")
        Delete_Records("GLTBREC2")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'")
    End Sub

    Sub Print_Record()

        Print_Report_Begin()
        ' CR_params.Add("SUBT", "")
        Dim RPT As String = "GLRBREC1"

        Generate_Report(RPT, Me.Text, "Cleared Items", "{GLTBREC2.TRAN_SEL}='1'", , , False)
        Generate_Report(RPT, Me.Text, "Open Items", "{GLTBREC2.TRAN_SEL}='0' and {GLTBREC2.TRAN_DATE} <= {GLTBREC1.BANK_STMT_DATE}", , , False)
        Print_Report_End()

    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    ' Click_Command("New", e)
                    Load_GLTBRECX()
                End If
            Case "BATCH_NO_CLEARED"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BANK_CODE"
                '   Click_Command("New")
                Load_GLTBRECX()
            Case "BATCH_NO_CLEARED"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "BANK_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BANK_CODE").Text <> "" Then
                        LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                        If cdr IsNot Nothing Then
                            Load_GLTBRECX()
                        End If
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_STMT_DATE"
                set_rowfilter()
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_STMT_BALANCE"
                Display_Totals()
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTBRECX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdGLTBREC2, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Select Selected", "De-Select Selected")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdGLTBREC2"
                tlb_btn = DirectCast(tlb_pop.Tools("Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("TRAN_SEL").Value = IIf(e.Tool.Key = "De-Select Selected", "0", "1")
                    grow.Update()
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            '    If rowICTITEM1 IsNot Nothing Then
            '        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

        End Select
    End Sub

#End Region

    Sub Dependent_Updates(S As Integer)
        If S = -1 Then
            Dim sql0 As String = "Set BATCH_NO_CLEARED = NULL where BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
            ASCDATA1.ExecuteSQL("Update APTCHCK1 " & sql0)
            ASCDATA1.ExecuteSQL("Update ARTPYMT1 " & sql0)
            ASCDATA1.ExecuteSQL("Update GLTBREC0 " & sql0)
            ASCDATA1.ExecuteSQL("Update APTCHCK1 " & Replace(sql0, "BATCH_NO_CLEARED", "BATCH_NO_CLEARED_F"))
        Else
            Dim sql1 As String = "Set BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' where BANK_CODE = '" & BANK_CODE & "'"
            Dim sql2 As String = "in (Select TRAN_KEY from " & GLTBREC2 & " where TRAN_SEL = '1' and JOURNAL_TYPE = "
            ASCDATA1.ExecuteSQL("Update APTCHCK1 " & sql1 & " and CHECK_NUM " & sql2 & "'APCD' and TRAN_KEY_LNO = 0)")
            ASCDATA1.ExecuteSQL("Update APTCHCK1 " & Replace(sql1, "BATCH_NO_CLEARED", "BATCH_NO_CLEARED_F") & " and CHECK_NUM " & sql2 & "'APCD' and TRAN_KEY_LNO = 1)")
            ASCDATA1.ExecuteSQL("Update ARTPYMT1 " & sql1 & " and PYMT_BATCH_NO " & sql2 & "'ARCR')")

            ASCDATA1.ExecuteSQL("Update GLTBREC0 Set BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' where (OPS_YYYYPP,JOURNAL_NO,JOURNAL_LNO) in (Select TRAN_YP,TRAN_KEY,TRAN_KEY_LNO from " & GLTBREC2 & " where TRAN_SEL = '1' and JOURNAL_TYPE = 'GLJE')")
        End If
    End Sub

    Private Sub grdARTOPENA_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTBREC2.AfterCellUpdate
        'If e.Cell.Column.Key = "TRAN_SEL" Then
        '    If e.Cell.Value = "1" Then
        '        e.Cell.Row.Cells("INV_BALANCE_APPLIED").Value = e.Cell.Row.Cells("INV_BALANCE").Value
        '    Else
        '        e.Cell.Row.Cells("INV_BALANCE_APPLIED").Value = 0
        '    End If
        'End If
    End Sub

    Private Sub grdGLTBREC2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGLTBREC2.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdARTOPENA_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTBREC2.BeforeRowUpdate
        'If e.Row.Cells("TRAN_SEL").Value = "1" Then
        '    e.Row.Cells("INV_BALANCE_APPLIED").Value = e.Row.Cells("INV_BALANCE").Value
        'Else
        '    e.Row.Cells("INV_BALANCE_APPLIED").Value = 0
        'End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor

        'If Absx1.txtFor("BANK_CODE").Text <> "" Then
        Load_GLTBRECX()
        'End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdGLTBRECX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdGLTBRECX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("BATCH_NO_CLEARED").Text = e.Row.Cells("BATCH_NO_CLEARED").Text
            Click_Command("View")
        End If
    End Sub

    Sub Load_GLTBRECX()
        Me.Cursor = Cursors.WaitCursor
        Dim BANK_CODE As String = Absx1.txtFor("BANK_CODE").Text
        Fill_Records("GLTBRECX", New String() {BANK_CODE})
        Sort_grdColumns(grdGLTBRECX, "BATCH_NO_CLEARED".ToLower)
        grdGLTBRECX.Text = "Bank Reconciliations" & IIf(BANK_CODE = "", "", " for " & BANK_CODE)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Display_Totals()

        Dim BANK_STMT_BALANCE As Decimal = Val(Absx1.numFor("BANK_STMT_BALANCE").Value & "")
        Dim TRAN_AMT_ARCR As Decimal = Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT_ARCR)", "") & "")
        Dim TRAN_AMT_APCD As Decimal = Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT_APCD)", "") & "")
        Dim TRAN_AMT_GLJE As Decimal = Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT_GLJE)", "") & "")


        With dst.Tables("GLTBRECT").Rows
            .Find(1).Item("T_AMT") = BANK_STMT_BALANCE_previous
            .Find(2).Item("T_AMT") = TRAN_AMT_ARCR
            .Find(3).Item("T_AMT") = TRAN_AMT_APCD
            .Find(4).Item("T_AMT") = TRAN_AMT_GLJE
            .Find(5).Item("T_AMT") = BANK_STMT_BALANCE_previous + TRAN_AMT_ARCR + TRAN_AMT_APCD + TRAN_AMT_GLJE
            .Find(6).Item("T_AMT") = BANK_STMT_BALANCE
            .Find(7).Item("T_AMT") = BANK_STMT_BALANCE - (BANK_STMT_BALANCE_previous + TRAN_AMT_ARCR + TRAN_AMT_APCD + TRAN_AMT_GLJE)
        End With
    End Sub

    Private Sub grdGLTBREC2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGLTBREC2.InitializeRow
        Dim JOURNAL_TYPE As String = e.Row.Cells("JOURNAL_TYPE").Value & ""
        If JOURNAL_TYPE = "ARCR" Then
            e.Row.Cells("JOURNAL_TYPE").Appearance.BackColor = Drawing.Color.LightGreen
        ElseIf JOURNAL_TYPE = "GLJE" Then
            e.Row.Cells("JOURNAL_TYPE").Appearance.BackColor = Drawing.Color.Orange
        ElseIf JOURNAL_TYPE = "APCD" Then
            If Val(e.Row.Cells("TRAN_KEY_LNO").Value & "") = 1 Then ' Void Check
                e.Row.Appearance.ForeColor = Drawing.Color.Red
                e.Row.ToolTipText = "Voided Check"
            End If
        End If
    End Sub

    Private Sub grdGLTBRECT_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGLTBRECT.InitializeRow
        Dim T_LNO As Integer = Val(e.Row.Cells("T_LNO").Value & "")
        If T_LNO = 1 Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGray
        ElseIf T_LNO = 5 Then
            e.Row.Appearance.BackColor = Drawing.Color.LightBlue
        ElseIf T_LNO = 6 Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        ElseIf T_LNO = 7 Then
            Dim T_AMT As Decimal = Val(e.Row.Cells("T_AMT").Value & "")
            If T_AMT <> 0 Then
                e.Row.Cells("T_AMT").Appearance.ForeColor = Drawing.Color.Red
            Else
                e.Row.Cells("T_AMT").Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If
    End Sub

    Private Sub chkShowFuture_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowFuture.CheckedChanged
        set_rowfilter()
    End Sub

    Sub Set_RowFilter()
        Dim dvw As DataView = DirectCast(grdGLTBREC2.DataSource, DataTable).DefaultView
        If chkShowFuture.Checked Then
            dvw.RowFilter = ""
        Else
            Dim BANK_STMT_DATE As Date = Absx1.dteFor("BANK_STMT_DATE").Value
            dvw.RowFilter = "TRAN_SEL = '1' or TRAN_DATE <= '" & Format(BANK_STMT_DATE, "MM/dd/yyyy") & "'"
        End If
    End Sub
End Class