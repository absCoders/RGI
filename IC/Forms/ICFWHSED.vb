Public Class ICFWHSED

    Dim WHSE_CODE As String
    Dim DIVISION_CODE As String
    Dim DIV_WHSE_CODE As String
    Dim TERR_CODE As String
    Dim rowICTWHSE1 As DataRow
    ' REMEMBER TO INITIALIZE WHSE_RATE_SCHEDULE
    ' LOADS OF CODE LOOKING AT WHSE_MIN_STORAGE TO UNDO

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ICTWHSE1", "*", , , , , _
                       "WHSE_WD_CHG,WHSE_XFR_CHG,WHSE_MIN_HANDLING,WHSE_MIN_STORAGE,WHSE_STG_CYCLE,WHSE_WD_CHG_EXP,WHSE_XFR_CHG_EXP,WHSE_NO_STG_ACCRUAL,WHSE_RATE_SCHEDULE,WHSE_CODE_RATE")

            ASCMAIN1.sql = "Select ICTCOSTE.*, ICTCOSTF.WHSE_STORAGE_CLASS_DESC " _
            & " from ICTCOSTE,ICTCOSTF " _
            & " where ICTCOSTE.WHSE_CODE = :PARM1 " _
            & "   and ICTCOSTE.WHSE_STORAGE_CLASS_CODE = ICTCOSTF.WHSE_STORAGE_CLASS_CODE"
            Create_TDA(.Tables.Add, "ICTCOSTE", "**", 0, , "V")

            ASCMAIN1.sql = "Select ICTPROD1.PROD_CODE, ICTPROD1.PROD_DESC" _
            & ", ICTPROD1.WHSE_STORAGE_CLASS_CODE, X.CASES, X.UNITS" _
            & " from ICTPROD1,(Select ICTLOTD2.PROD_CODE, Sum (ICTLOTD2.QTY_ON_HAND) CASES" _
            & ", Sum (ICTLOTD2.QTY_ON_HAND * DECODE(ICTLOTD2.PACK_CODE,'000',ICTLOTD2.CATCH_WEIGHT,ICTPACK1.PACK_FACTOR)) UNITS" _
            & " from ICTLOTD2,ICTPACK1 where ICTPACK1.PACK_CODE = ICTLOTD2.PACK_CODE and ICTLOTD2.WHSE_CODE = :PARM1 group by ICTLOTD2.PROD_CODE) X" _
            & " where X.PROD_CODE = ICTPROD1.PROD_CODE"
            Create_TDA(.Tables.Add, "ICTPROD1", "**", 0, False, "V", 1)

            Create_Relation("ICTCOSTE", "ICTPROD1", "WHSE_STORAGE_CLASS_CODE")

            With .Tables("ICTCOSTE").Columns
                .Add("CASES", GetType(System.Int64), "SUM(CHILD.CASES)")
                .Add("UNITS", GetType(System.Double), "SUM(CHILD.UNITS)")
                .Add("STORAGE", GetType(System.Decimal), "ISNULL(UNITS,0) * ISNULL(WHSE_STORAGE,0) / 100")
            End With

            With .Tables("ICTPROD1").Columns
                .Add("WHSE_STORAGE", GetType(System.Decimal), "PARENT.WHSE_STORAGE")
                .Add("WHSE_HANDLING", GetType(System.Decimal), "PARENT.WHSE_HANDLING")
                .Add("STORAGE", GetType(System.Decimal), "ISNULL(UNITS,0) * ISNULL(WHSE_STORAGE,0) / 100")
            End With

            ASCMAIN1.sql = "Select ICTWHSE1.*,SOTSDIV1.DIV_WHSE_CODE, X.CASES, X.UNITS" _
            & " from ICTWHSE1,SOTTERR1,SOTSDIV1, (Select ICTLOTD2.WHSE_CODE" _
            & ", SUM (ICTLOTD2.QTY_ON_HAND) CASES" _
            & ", SUM (ICTLOTD2.QTY_ON_HAND * DECODE(ICTLOTD2.PACK_CODE,'" & TAC.TACMAIN1.CATCH_PACK & "',ICTLOTD2.CATCH_WEIGHT,ICTPACK1.PACK_FACTOR)) UNITS" _
            & " from ICTLOTD2,ICTPACK1 where ICTPACK1.PACK_CODE = ICTLOTD2.PACK_CODE group by ICTLOTD2.WHSE_CODE) X" _
            & " where SOTTERR1.TERR_CODE (+) = ICTWHSE1.TERR_CODE" _
            & "   and SOTSDIV1.DIVISION_CODE (+) = SOTTERR1.DIVISION_CODE" _
            & "   and X.WHSE_CODE = ICTWHSE1.WHSE_CODE"
            Create_TDA(.Tables.Add, "ICTWHSEL", "**", 0, False, "", 1)
            .Tables("ICTWHSEL").Columns.Add("DIVISIONS")
            .Tables("ICTWHSEL").Columns.Add("USES_WHSE_CODE", GetType(System.String), "IIF(ISNULL(WHSE_RATE_SCHEDULE,'0')='2',WHSE_CODE_RATE,IIF(ISNULL(WHSE_RATE_SCHEDULE,'0')='0',DIV_WHSE_CODE,WHSE_CODE))")

            Create_TDA(.Tables.Add, "ICTCOSTG", "*", 1)
            Create_TDA(.Tables.Add, "ICTCOSTF", "*", 0)
        End With

        Fill_Records("ICTCOSTF")

        'grdICTCOSTE.DataMember = "ICTCOSTE"
        'grdICTCOSTE.DataSource = dst

        grdICTCOSTE.DataSource = dst.Tables("ICTCOSTE")
        grdICTCOSTG.DataSource = dst.Tables("ICTCOSTG")
        grdICTWHSEL.DataSource = dst.Tables("ICTWHSEL")

        Create_Summary(grdICTCOSTE, "CASES")
        Create_Summary(grdICTCOSTE, "UNITS")
        Create_Summary(grdICTCOSTE, "STORAGE")

        Create_Summary(grdICTWHSEL, "WHSE_CODE", "Count")
        Create_Summary(grdICTWHSEL, "CASES")
        Create_Summary(grdICTWHSEL, "UNITS")

        With grdICTCOSTE.DisplayLayout.Bands(0)
            .Columns("WHSE_STORAGE_CLASS_CODE").CellAppearance.BackColor = Color.Beige
            .Columns("WHSE_STORAGE_CLASS_DESC").CellAppearance.BackColor = Color.Beige
            .Columns("CASES").CellAppearance.BackColor = Color.Beige
            .Columns("UNITS").CellAppearance.BackColor = Color.Beige
            .Columns("STORAGE").CellAppearance.BackColor = Color.Beige
        End With
        With grdICTCOSTE.DisplayLayout.Bands(1)
            For Each COLUMN_NAME As String In New String() _
                {"PROD_CODE", "PROD_DESC", "CASES", "UNITS", "WHSE_STORAGE", "WHSE_HANDLING", "STORAGE"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With

        grdICTCOSTE.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay

        cbeYP0.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP0.SelectedItem = cbeYP0.Items(0)
        cbeYP1.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP1.SelectedItem = cbeYP1.Items(0)

        ASCMAIN1.Add_Value_List(grdICTWHSEL, "WHSE_RATE_SCHEDULE", Nothing, New String() {":", "0:Default", "1:Itself", "2:Whse"})

        TABLE_NAME = "ICTWHSE1"
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View", "Edit"
                If Validate_Code("WHSE_CODE") Then
                    If eItemKey = "Edit" Then
                        If Not ASCMAIN1.Logical_Lock("ICTCOSTE", Absx1.txtFor("WHSE_CODE").Text) Then Exit Sub
                    End If

                    WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Warehouse Code (" & WHSE_CODE & ")"
                    Else
                        TERR_CODE = rowICTWHSE1.Item("TERR_CODE") & ""
                        Dim rowSOTTERR1 As DataRow = LookUp("SOTTERR1", TERR_CODE)
                        If rowSOTTERR1 Is Nothing Then
                            EMsg &= vbCr & "Warehouse " & WHSE_CODE & " is not associated with a Valid Territory (" & TERR_CODE & ")"
                        Else
                            DIVISION_CODE = rowSOTTERR1.Item("DIVISION_CODE") & ""
                            Dim rowSOTSDIV1 As DataRow = LookUp("SOTSDIV1", DIVISION_CODE)
                            If rowSOTTERR1 Is Nothing Then
                                EMsg &= vbCr & "Warehouse Territory " & TERR_CODE & " is not associated with a Valid Division (" & DIVISION_CODE & ")"
                            Else
                                DIV_WHSE_CODE = rowSOTSDIV1.Item("DIV_WHSE_CODE") & ""
                                Dim rowICTWHSE1_DEF As DataRow = LookUp("ICTWHSE1", DIV_WHSE_CODE)
                                If rowICTWHSE1_DEF Is Nothing Then
                                    EMsg &= vbCr & "Warehouse Territory Division " & DIVISION_CODE & " is not associated with a Valid Default Warehouse Code (" & DIV_WHSE_CODE & ")"
                                End If
                            End If
                        End If
                    End If
                End If

            Case "Update"
                If optRateSchedule.Value = "0" Then
                    If dst.Tables("ICTCOSTE").Select("ISNULL(WHSE_STORAGE,0) <> 0 OR ISNULL(WHSE_HANDLING,0) <> 0").Length <> 0 _
                    Or dst.Tables("ICTCOSTG").Select.Length <> 0 _
                    Or Val(Absx1.numFor("WHSE_MIN_STORAGE").Value & "") <> 0 _
                    Or Val(Absx1.numFor("WHSE_MIN_HANDLING").Value & "") <> 0 _
                    Or Val(Absx1.numFor("WHSE_WD_CHG_EXP").Value & "") <> 0 _
                    Or Val(Absx1.numFor("WHSE_XFR_CHG_EXP").Value & "") <> 0 Then

                        If MsgBox("Warehouses using a Divisional Default for Storage, Sorting and Handling Rates" _
                                  & vbCrLf & " may not have a non-zero value defined for the following fields:" _
                                  & vbCrLf _
                                  & vbCrLf & "   1) Minimums for Storage & Handling" _
                                  & vbCrLf & "   2) Default Expense Rates for Withdrawal & Transfer in Storage Charges" _
                                  & vbCrLf & "   3) Storage Rates by Storage Class" _
                                  & vbCrLf & "   4) Sorting Charges" _
                                  & vbCrLf & vbCrLf & "By continuing with this Update, the values maintained in these fields" _
                                  & vbCrLf & "  will be cleared (for this warehouse only) so that the Default Values may apply." _
                                  & vbCrLf & vbCrLf & "Continue with Update?", _
                                  MsgBoxStyle.YesNo, _
                                  "Some information may be lost as a result of this Update") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                ElseIf optRateSchedule.Value = "2" Then
                    If Absx1.cmbFor("WHSE_CODE_RATE").Value & "" = "" Then
                        EMsg &= "You Must Specify a Warehouse whose Rate Schedule will be used"
                    Else
                        Dim WHSE_CODE_RATE As String = Absx1.cmbFor("WHSE_CODE_RATE").Value
                        Dim issues As String = ""
                        ASCMAIN1.sql = "Select Count (*) from ICTCOSTE where WHSE_CODE = :PARM1"
                        If Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {WHSE_CODE_RATE})) = 0 Then
                            issues &= vbCrLf & "  No Rate Schedule for Storage"
                        End If
                        ASCMAIN1.sql = "Select Count (*) from ICTCOSTG where WHSE_CODE = :PARM1"
                        If Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {WHSE_CODE_RATE})) = 0 Then
                            issues &= vbCrLf & "  No Rate Schedule for Sorting"
                        End If
                        ASCMAIN1.sql = "Select NVL(WHSE_MIN_STORAGE,0) from ICTWHSE1 where WHSE_CODE = :PARM1"
                        If Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {WHSE_CODE_RATE})) = 0 Then
                            issues &= vbCrLf & "  No Minimum Rate for Storage"
                        End If
                        If issues <> "" Then
                            If MsgBox("Please note the following issues with the Warehouse selected:" _
                                      & vbCrLf & issues _
                                      & vbCrLf & vbCrLf & "Do you still want to Proceed?", _
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If

                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Done"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
            If ScreenMode And EntryMode = "V" Then
                .Groups("Screen Control").Items("Edit").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.True
            Else
                .Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.False
            End If
            If ScreenMode And EntryMode = "V" Then
                .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.False
            Else
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End If

            .Groups("Rate Schedule").Visible = ScreenMode

        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = Not ScreenMode
        splICTWHSED.Visible = ScreenMode

        grdICTWHSED.Visible = False
        grdICTLOTDX_PVT.Visible = False

        Setup_tabMain()

        If ScreenMode Then
            If EntryMode = "E" Then
                grdICTCOSTE.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

                grdICTCOSTG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdICTCOSTG.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdICTCOSTG.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            End If
            Set_Read_Only(grpRateOptions, (EntryMode <> "E"))
            Set_Read_Only(splICTWHSED.Panel1, (EntryMode <> "E"))

        Else
            Clear_Record()
            grdICTCOSTE.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdICTCOSTG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)
        Fill_Records("ICTCOSTE", WHSE_CODE)
        Fill_Records("ICTCOSTG", WHSE_CODE)
        rowICTWHSE1 = Fill_Record("ICTWHSE1", WHSE_CODE)
        Fill_Records("ICTPROD1", WHSE_CODE)

        For Each row As DataRow In ASCDATA1.SelectDistinct("ICTPROD1", "WHSE_STORAGE_CLASS_CODE").Rows
            Dim WHSE_STORAGE_CLASS_CODE As String = row.Item("WHSE_STORAGE_CLASS_CODE") & ""
            If WHSE_STORAGE_CLASS_CODE <> "" Then
                If dst.Tables("ICTCOSTE").Rows.Find(New Object() {WHSE_CODE, WHSE_STORAGE_CLASS_CODE}) Is Nothing Then
                    Dim rowICTCOSTF As DataRow = LookUp("ICTCOSTF", WHSE_STORAGE_CLASS_CODE)
                    Dim rowICTCOSTE As DataRow = dst.Tables("ICTCOSTE").NewRow
                    rowICTCOSTE.Item("WHSE_CODE") = WHSE_CODE
                    rowICTCOSTE.Item("WHSE_STORAGE_CLASS_CODE") = WHSE_STORAGE_CLASS_CODE
                    rowICTCOSTE.Item("WHSE_STORAGE_CLASS_DESC") = rowICTCOSTF.Item("WHSE_STORAGE_CLASS_DESC")
                    dst.Tables("ICTCOSTE").Rows.Add(rowICTCOSTE)
                End If
            End If
        Next
        EnforceConstraints(True)

        For Each rowICTCOSTF As DataRow In dst.Tables("ICTCOSTF").Rows
            Dim WHSE_STORAGE_CLASS_CODE As String = rowICTCOSTF.Item("WHSE_STORAGE_CLASS_CODE")

            Dim rowICTCOSTE As DataRow = dst.Tables("ICTCOSTE").Rows.Find _
                (New String() {WHSE_CODE, WHSE_STORAGE_CLASS_CODE})
            If rowICTCOSTE Is Nothing Then
                rowICTCOSTE = dst.Tables("ICTCOSTE").NewRow
                rowICTCOSTE.Item("WHSE_CODE") = WHSE_CODE
                rowICTCOSTE.Item("WHSE_STORAGE_CLASS_CODE") = WHSE_STORAGE_CLASS_CODE
                rowICTCOSTE.Item("WHSE_STORAGE_CLASS_DESC") = rowICTCOSTF.Item("WHSE_STORAGE_CLASS_DESC")
                dst.Tables("ICTCOSTE").Rows.Add(rowICTCOSTE)
            End If
        Next


        lblDefaultRates.Text = "The Default Rate Schedule for Division " _
        & DIVISION_CODE & " is maintained using Warehouse Code " & DIV_WHSE_CODE & "."

        Absx1.txtFor("DIVISION_CODE").Text = DIVISION_CODE

        Set_ACCRUAL()
        Set_Rate_Option()

        If rowICTWHSE1.Item("WHSE_RATE_SCHEDULE") & "" = "" Then
            rowICTWHSE1.Item("WHSE_RATE_SCHEDULE") = "0"
        End If

        'If Val(rowICTWHSE1.Item("WHSE_MIN_STORAGE") & "") = 0 Then
        '    'If Val(Absx1.numFor("WHSE_MIN_STORAGE").Value & "") = 0 Then
        '    optRateSchedule.Value = "0"
        'Else
        '    optRateSchedule.Value = "1"
        'End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("ICTCOSTE").Rows.Clear()
        dst.Tables("ICTPROD1").Rows.Clear()
        dst.Tables("ICTWHSE1").Rows.Clear()
        dst.Tables("ICTCOSTG").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("DIVISION_CODE").Text = ""

        Load_ICTWHSEL()
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Order")

        If optRateSchedule.Value = "0" Then
            For Each rowICTCOSTE As DataRow In dst.Tables("ICTCOSTE").Select
                rowICTCOSTE.Item("WHSE_STORAGE") = DBNull.Value
                rowICTCOSTE.Item("WHSE_HANDLING") = DBNull.Value
            Next
            dst.Tables("ICTCOSTG").Rows.Clear()
            Absx1.numFor("WHSE_MIN_STORAGE").Value = 0
            Absx1.numFor("WHSE_MIN_HANDLING").Value = 0
            Absx1.numFor("WHSE_WD_CHG_EXP").Value = 0
            Absx1.numFor("WHSE_XFR_CHG_EXP").Value = 0
            Absx1.cmbFor("WHSE_CODE_RATE").Value = DBNull.Value

        ElseIf optRateSchedule.Value = "1" Then
            Absx1.cmbFor("WHSE_CODE_RATE").Value = DBNull.Value
        End If

        BeginTrans()

        Dim sql_delete As String = "WHSE_CODE = '" & WHSE_CODE & "'"
        If optRateSchedule.Value = "0" Then
            ASCDATA1.ExecuteSQL("Delete from ICTCOSTE where " & sql_delete)
            ASCDATA1.ExecuteSQL("Delete from ICTCOSTG where " & sql_delete)
        Else
            Update_Record_TDA("ICTCOSTE", sql_delete)
            Update_Record_TDA("ICTCOSTG", sql_delete)
        End If
        Update_Record_TDA("ICTWHSE1")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTCOSTE, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTWHSEL, "SS", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.Name = "grd" Then
            Exit Sub
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""
                    'If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                    '    e.Cancel = True
                    'End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        If COLUMN_NAME = "WHSE_CODE" Then
            If ctl.Text <> "" Then
                'Call Click_Command("Load Reports")
            End If
        End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "WHSE_CODE" Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then
                Click_Command("View")
            End If
        End If
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)

        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                If txtctl.Text <> "" Then
                    Click_Command("View")
                End If
        End Select
    End Sub

#End Region

#Region "grdICTCOSTE"

    Private Sub grdICTCOSTE_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCOSTE.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "WHSE_STORAGE"

        End Select
    End Sub

    Private Sub grdICTCOSTE_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTCOSTE.AfterRowsDeleted

    End Sub

    Private Sub grdICTCOSTE_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTCOSTE.AfterRowUpdate

    End Sub

    Private Sub grdICTCOSTE_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTCOSTE.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("WHSE_CODE").Value = Absx1.txtFor("WHSE_CODE").Text
        End If
    End Sub

    Private Sub grdICTCOSTE_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCOSTE.InitializeRow

    End Sub
#End Region

#Region "grdICTCOSTG"

    Private Sub grdICTCOSTG_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCOSTG.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "WHSE_STORAGE"

        End Select
    End Sub

    Private Sub grdICTCOSTG_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTCOSTG.AfterRowsDeleted

    End Sub

    Private Sub grdICTCOSTG_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTCOSTG.AfterRowUpdate

    End Sub

    Private Sub grdICTCOSTG_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTCOSTG.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("WHSE_CODE").Value = Absx1.txtFor("WHSE_CODE").Text
        End If
    End Sub

    Private Sub grdICTCOSTG_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCOSTG.InitializeRow

    End Sub
#End Region

    Private Sub chkNoAccrual_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNoAccrual.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_ACCRUAL()
    End Sub

    Sub Set_ACCRUAL()
        Absx1.numFor("WHSE_STG_CYCLE").ReadOnly = (chkNoAccrual.Checked)
    End Sub

    Private Sub optRateSchedule_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRateSchedule.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Rate_Option()
    End Sub

    Sub Set_Rate_Option()

        Absx1.numFor("WHSE_WD_CHG_EXP").ReadOnly = (optRateSchedule.Value = "0")
        Absx1.numFor("WHSE_XFR_CHG_EXP").ReadOnly = (optRateSchedule.Value = "0")
        Absx1.numFor("WHSE_MIN_STORAGE").ReadOnly = (optRateSchedule.Value = "0")
        Absx1.numFor("WHSE_MIN_HANDLING").ReadOnly = (optRateSchedule.Value = "0")

        With grdICTCOSTE.DisplayLayout.Bands(0)
            If optRateSchedule.Value <> "1" Then
                .Columns("WHSE_STORAGE").CellAppearance.BackColor = Color.LightGray
                .Columns("WHSE_HANDLING").CellAppearance.BackColor = Color.LightGray
                .Columns("WHSE_STORAGE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("WHSE_HANDLING").CellActivation = UltraWinGrid.Activation.NoEdit
            Else
                .Columns("WHSE_STORAGE").CellAppearance.BackColor = Color.Yellow
                .Columns("WHSE_HANDLING").CellAppearance.BackColor = Color.Yellow
                .Columns("WHSE_STORAGE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("WHSE_HANDLING").CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End With

        With grdICTCOSTG.DisplayLayout.Bands(0)
            If optRateSchedule.Value <> "1" Then
                .Columns("WHSE_SORTING").CellAppearance.BackColor = Color.LightGray
                .Columns("WHSE_SORTING").CellActivation = UltraWinGrid.Activation.NoEdit
            Else
                .Columns("WHSE_SORTING").CellAppearance.BackColor = Color.Yellow
                .Columns("WHSE_SORTING").CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End With

        cmbWHSE_CODE_RATE.Visible = (optRateSchedule.Value = "2")
    End Sub

    Private Sub cmdFetch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetch.Click
        If Not chkP.Checked And Not chkT.Checked And Not chkC.Checked And Not chkD.Checked Then
            MsgBox("You Must Pick at Least 1 Source (PO Rec, Xfr Rec, Returns, Daily Stg)", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If
        If Not chkSTG.Checked And Not chkSRT.Checked And Not chkHND.Checked Then
            MsgBox("You Must Pick at Least 1 Whse Expense Type (Storage, Sorting, Handling)", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Summary & Pivot")

        Dim DATA As String = ""
        Dim DATA_desc As String = ""
        If chkSTG.Checked Then
            DATA &= "+DECODE(VOUCHER_NO,NULL,NVL(ADDED_STORAGE_TOTAL,0)," & IIf(optAA.Value = "I", "NVL(ACT_STG,0)", "NVL(ADDED_STORAGE_TOTAL,0)") & ")"
            DATA_desc &= ",Storage"
        End If
        If chkSRT.Checked Then
            DATA &= "+DECODE(VOUCHER_NO,NULL,NVL(SORTING,0)," & IIf(optAA.Value = "I", "NVL(ACT_SORT,0)", "NVL(SORTING,0)") & ")"
            DATA_desc &= ",Sorting"
        End If
        If chkHND.Checked Then
            DATA &= "+DECODE(VOUCHER_NO,NULL,NVL(HANDLING,0)," & IIf(optAA.Value = "I", "NVL(ACT_HAND,0)", "NVL(HANDLING,0)") & ")"
            DATA_desc &= ",Handling"
        End If

        Dim TRAN_TYPEs As String = ""
        Dim TRAN_TYPEs_desc As String = ""
        If chkP.Checked Then TRAN_TYPEs &= ",'P'" : TRAN_TYPEs_desc &= ",PO Rec"
        If chkT.Checked Then TRAN_TYPEs &= ",'T'" : TRAN_TYPEs_desc &= ",Xfr Rec"
        If chkC.Checked Then TRAN_TYPEs &= ",'C'" : TRAN_TYPEs_desc &= ",Returns"
        If chkD.Checked Then TRAN_TYPEs &= ",'D'" : TRAN_TYPEs_desc &= ",Daily Stg"

        ASCMAIN1.sql = "Select TRUNC(PROCESS_DATE) PROCESS_DATE, WHSE_CODE" _
        & ", SUM (" & Mid(DATA, 2) & ") AMT" _
        & " from ICTLOTDX" _
        & " where OPS_YYYYPP >= '" & cbeYP0.Value & "' AND OPS_YYYYPP <= '" & cbeYP1.Value & "'" _
        & " and ORIG_TRAN_TYPE IN (" & Mid(TRAN_TYPEs, 2) & ")" _
        & " group by TRUNC(PROCESS_DATE), WHSE_CODE"

        Dim ICTLOTDX_SUM As String = ASCMAIN1.Temp_Table


        Dim sql As String = ""
        Dim WHSE_CODEs As New List(Of String)
        ASCMAIN1.sql = "Select Distinct WHSE_CODE from " & ICTLOTDX_SUM
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "WHSE_CODE")
            Dim WHSE_CODE As String = row.Item("WHSE_CODE")
            WHSE_CODEs.Add(WHSE_CODE)
            sql &= ", Sum (DECODE(WHSE_CODE,'" & WHSE_CODE & "',AMT,0)) W" & WHSE_CODE
        Next

        'ASCMAIN1.sql = ASCMAIN1.Flattened_List("PROCESS_DATE", "AMT", ICTLOTDX_SUM)
        ASCMAIN1.sql = "Select PROCESS_DATE, SUM (AMT) TOTAL" & sql & " from " & ICTLOTDX_SUM & " group by PROCESS_DATE"
        Dim DT As DataTable = ASCDATA1.GetDataTable

        grdICTLOTDX_PVT.DataSource = DT
        Sort_grdColumns(grdICTLOTDX_PVT, "PROCESS_DATE")

        With grdICTLOTDX_PVT.DisplayLayout.Bands(0)
            If .Summaries.Count <> 0 Then
                For i As Integer = .Summaries.Count - 1 To 0 Step -1
                    Dim s As UltraWinGrid.SummarySettings = .Summaries(i)
                    .Summaries.Remove(s)
                Next
            End If
            .Columns("PROCESS_DATE").Header.Fixed = True
            .Columns("TOTAL").Header.Fixed = True
        End With


        For Each WHSE_CODE As String In WHSE_CODEs
            Dim COLUMN_NAME As String = "W" & WHSE_CODE
            Create_Summary(grdICTLOTDX_PVT, COLUMN_NAME)
        Next
        Create_Summary(grdICTLOTDX_PVT, "TOTAL")

        grdICTLOTDX_PVT.Text = "Accrued Warehouse Expense by Date, " & cbeYP0.Text & " thru " & cbeYP1.Text & ", Sources:" & Mid(TRAN_TYPEs_desc, 2) & ", Types:" & Mid(DATA_desc, 2)
        grdICTLOTDX_PVT.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Accrued Whse by Date").Visible = Not ScreenMode AndAlso (tabMain.SelectedTab.Key = "Whse Accrual Summary by Date")
    End Sub

    Private Sub grdICTWHSEL_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEL.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("View")
    End Sub

    Sub Load_ICTWHSEL()
        Fill_Records("ICTWHSEL")
        ASCMAIN1.sql = "Select DIVISION_CODE, DIV_WHSE_CODE from SOTSDIV1 where DIV_WHSE_CODE is not null"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Rows
            Dim WHSE_CODE As String = ROW.Item("DIV_WHSE_CODE")
            Dim rowICTWHSEL As DataRow = dst.Tables("ICTWHSEL").Rows.Find(WHSE_CODE)
            If rowICTWHSEL IsNot Nothing Then
                Dim DIVISIONS As String = rowICTWHSEL.Item("DIVISIONS") & ""
                If DIVISIONS <> "" Then
                    DIVISIONS &= ","
                End If
                DIVISIONS &= ROW.Item("DIVISION_CODE")
                rowICTWHSEL.Item("DIVISIONS") = DIVISIONS
            End If
        Next

        'For Each ROW As DataRow In ASCDATA1.SelectDistinct("ICTWHSEL", "USES_WHSE_CODE").Rows
        '    Dim WHSE_CODE As String = ROW.Item("USES_WHSE_CODE")
        '    Dim rowICTWHSEL As DataRow = dst.Tables("ICTWHSEL").Rows.Find(WHSE_CODE)
        '    If rowICTWHSEL IsNot Nothing Then
        '        rowICTWHSEL.Item("USES_WHSE_CODE") = WHSE_CODE
        '    End If
        'Next

        For Each rowICTWHSEL As DataRow In dst.Tables("ICTWHSEL").Select
            If rowICTWHSEL.Item("WHSE_RATE_SCHEDULE") & "" = "" Then
                rowICTWHSEL.Item("WHSE_RATE_SCHEDULE") = "0"
            End If
        Next
        Sort_grdColumns(grdICTWHSEL, "WHSE_CODE")
    End Sub
End Class