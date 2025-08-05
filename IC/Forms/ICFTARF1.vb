Public Class ICFTARF1

    Dim ORIG_CODE As String

    'AUDITING
    'sqlsort in cmdtest

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select * from ICTORIG1"
            Create_TDA(.Tables.Add, "ICTORIG1", "**", 0, , "")

            ASCMAIN1.sql = "Select ICTCOSTD.*,APTVEND1.VEND_NAME " _
            & " from ICTCOSTD,APTVEND1 " _
            & " where ICTCOSTD.ORIG_CODE = :PARM1 " _
            & "   and APTVEND1.VEND_CODE (+) = ICTCOSTD.VEND_CODE"
            Create_TDA(.Tables.Add, "ICTCOSTD", "**", 0, , "V")

            ASCMAIN1.sql = "Select ICTCOSTC.*,ICTPROD1.PROD_DESC " _
            & " from ICTCOSTC,ICTPROD1 " _
            & " where ICTCOSTC.ORIG_CODE = :PARM1 " _
            & "   and ICTPROD1.PROD_CODE (+) = ICTCOSTC.PROD_CODE"
            Create_TDA(.Tables.Add, "ICTCOSTC", "**", 0, , "V")

            Create_Relation("ICTCOSTD", "ICTCOSTC", "ORIG_CODE,VEND_CODE")

            .Tables.Add("ICTCOSTX")
            With .Tables("ICTCOSTX").Columns
                .Add("ORIG_CODE")
                .Add("ORIG_DESC")
                .Add("VEND_CODES")
            End With

            .Tables.Add("ICTCOSTT")
            With .Tables("ICTCOSTT").Columns
                .Add("FEE")
                .Add("PAYEE")
                .Add("COST")
            End With

            .Tables.Add("ICTCOSTV")
            With .Tables("ICTCOSTV").Columns
                .Add("SOURCE")
                .Add("SUPPLIER")
                .Add("AGENT")
                .Add("INSPECTOR")
                .Add("OTHER")
            End With

        End With

        Fill_Records("ICTORIG1")

        grdICTCOSTX.DataSource = dst.Tables("ICTCOSTX")
        grdICTCOSTC.DataSource = dst.Tables("ICTCOSTC")
        grdICTCOSTD.DataSource = dst.Tables("ICTCOSTD")
        grdICTCOSTT.DataSource = dst.Tables("ICTCOSTT")

        Create_Summary(grdICTCOSTX, "ORIG_CODE", "Count")


        With grdICTCOSTC.DisplayLayout.Bands(0)
            For Each COLUMN_NAME In New String() {"PROD_CODE", "PROD_DESC", "BRAND_CODE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Goldenrod
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            Next
            For Each COLUMN_NAME In New String() {"COMMISSION", "COMM_TYPE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            Next
            For Each COLUMN_NAME In New String() {"INSPECTION", "INSP_TYPE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Green
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            Next
            For Each COLUMN_NAME In New String() {"OTHER", "OTHER_TYPE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            Next
        End With

        With grdICTCOSTD.DisplayLayout.Bands(0)
            For Each COLUMN_NAME In New String() {"VEND_CODE", "VEND_NAME"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Goldenrod
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            Next
            For Each COLUMN_NAME In New String() {"PAYEE_COMMISSION"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            Next
            For Each COLUMN_NAME In New String() {"PAYEE_INSPECTION"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Green
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            Next
            For Each COLUMN_NAME In New String() {"PAYEE_OTHER"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            Next
        End With


        ASCMAIN1.Add_Value_List(grdICTCOSTC, "COMM_TYPE", Nothing, New String() {":", "$:$/LB", "%:%Cost"})
        ASCMAIN1.Add_Value_List(grdICTCOSTC, "INSP_TYPE", Nothing, New String() {":", "$:$/LB", "%:%Cost"})
        ASCMAIN1.Add_Value_List(grdICTCOSTC, "OTHER_TYPE", Nothing, New String() {":", "$:$/LB", "%:%Cost"})


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View", "Edit"
                If Validate_Code("ORIG_CODE") Then
                    If eItemKey = "Edit" Then
                        If Not ASCMAIN1.Logical_Lock("ICTCOSTD", Absx1.txtFor("ORIG_CODE").Text) Then Exit Sub
                    End If

                    ORIG_CODE = Absx1.txtFor("ORIG_CODE").Text
                    Dim rowICTORIG1 As DataRow = Lookup("ICTORIG1", ORIG_CODE)
                    If rowICTORIG1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Origin Code (" & ORIG_CODE & ")"
                    Else

                    End If
                End If

            Case "Update"

                For Each COLUMN_NAME As String In New String() {"VEND_CODE", "PAYEE_COMMISSION", "PAYEE_INSPECTION", "PAYEE_OTHER"}
                    For Each row As DataRow In ASCDATA1.SelectDistinct("ICTCOSTD", COLUMN_NAME).Rows
                        Dim VEND_CODE As String = row.Item(0) & ""
                        If COLUMN_NAME = "VEND_CODE" And VEND_CODE = "Z" Then
                        Else
                            If VEND_CODE <> "" Then
                                If Lookup("APTVEND1", VEND_CODE) Is Nothing Then
                                    EMsg &= vbCr & "Invalide Value Specified for " & grdICTCOSTD.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption & " (" & VEND_CODE & ")"
                                End If
                            End If
                        End If
                        If COLUMN_NAME = "VEND_CODE" Then
                            Dim rows() As DataRow = dst.Tables("ICTCOSTC").Select _
                                ("ORIG_CODE = '" & ORIG_CODE & "' and VEND_CODE = '" & VEND_CODE & "'")
                            If rows.Length = 0 Then
                                EMsg &= vbCr & "No Commission details defined for Supplier " & VEND_CODE & "; Delete or Complete"
                            Else
                                Dim rowICTCOSTD As DataRow = dst.Tables("ICTCOSTD").Rows.Find(New String() {ORIG_CODE, VEND_CODE})
                                For Each rowICTCOSTC As DataRow In rows
                                    Dim BRAND_CODE As String = rowICTCOSTC.Item("BRAND_CODE") & ""
                                    If BRAND_CODE <> "Z" Then
                                        If Lookup("ICTBRAN1", BRAND_CODE) Is Nothing Then
                                            EMsg &= vbCr & "Invalid Value Specified for Brand Code (See Supplier " & VEND_CODE & ")"
                                        End If
                                    End If
                                    Dim PROD_CODE As String = rowICTCOSTC.Item("PROD_CODE") & ""
                                    If PROD_CODE <> "ZZZ" Then
                                        If Lookup("ICTPROD1", PROD_CODE) Is Nothing Then
                                            EMsg &= vbCr & "Invalid Value Specified for Product Code (See Supplier " & VEND_CODE & ")"
                                        End If
                                    End If


                                    For Each COLUMN_NAME2 As String In New String() {"COMMISSION", "INSPECTION", "OTHER"}
                                        Dim COLUMN_NAME3 As String = "PAYEE_" & COLUMN_NAME2

                                        If Val(rowICTCOSTC.Item(COLUMN_NAME2) & "") = 0 Then
                                            If rowICTCOSTD.Item(COLUMN_NAME3) & "" <> "" Then
                                                EMsg &= vbCr & grdICTCOSTD.DisplayLayout.Bands(0).Columns(COLUMN_NAME3).Header.Caption & " Specified, but no value specified for " & grdICTCOSTC.DisplayLayout.Bands(0).Columns(COLUMN_NAME2).Header.Caption & " (See Supplier " & VEND_CODE & ")"
                                            End If
                                        Else
                                            If rowICTCOSTD.Item(COLUMN_NAME3) & "" = "" Then
                                                EMsg &= vbCr & "A Value was specified for " & grdICTCOSTC.DisplayLayout.Bands(0).Columns(COLUMN_NAME2).Header.Caption & ", but no Payee specified for " & grdICTCOSTD.DisplayLayout.Bands(0).Columns(COLUMN_NAME3).Header.Caption & " (See Supplier " & VEND_CODE & ")"
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                        End If
                    Next
                Next

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

            Case "Check Open POs"
                Check_Open_POs()

            Case "Print Report"
                Print_Record()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
            If ScreenMode And EntryMode = "V" Then
                .Groups("Screen Control").Items("Edit").Settings.Enabled = DefaultableBoolean.True
            Else
                .Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            End If
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

            .Groups("Screen Control").Items("Update").Visible = (EntryMode = "E")
            .Groups("Screen Control").Items("Cancel").Visible = (EntryMode = "E")
            .Groups("Screen Control").Items("Done").Visible = (EntryMode = "V")

            .Groups("Test").Visible = ScreenMode

            .Groups("Screen Control").Items("Check Open POs").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Print Report").Settings.Enabled = iScreenMode

        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode And EntryMode = "E" Then
            With grdICTCOSTD.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowDelete = DefaultableBoolean.True
                .AllowUpdate = DefaultableBoolean.True
                .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            End With
            With grdICTCOSTC.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowDelete = DefaultableBoolean.True
                .AllowUpdate = DefaultableBoolean.True
                .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            End With
        Else
            With grdICTCOSTD.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
                .CellClickAction = UltraWinGrid.CellClickAction.Default
            End With
            With grdICTCOSTC.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
                .CellClickAction = UltraWinGrid.CellClickAction.Default
            End With
        End If

        grdICTCOSTX.Visible = Not ScreenMode
        splICTCOSTC.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)
        Fill_Records("ICTCOSTD", ORIG_CODE)
        Sort_grdColumns(grdICTCOSTD, "VEND_CODE")
        Fill_Records("ICTCOSTC", ORIG_CODE)
        Sort_grdColumns(grdICTCOSTC, "BRAND_CODE,PROD_CODE")
        EnforceConstraints(True)

        grdICTCOSTD.Text = "Suppliers with Commission Schedules for Origin " & ORIG_CODE

        Setup_ICTCOSTC()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("ICTCOSTC").Rows.Clear()
        dst.Tables("ICTCOSTD").Rows.Clear()
        dst.Tables("ICTCOSTT").Rows.Clear()
        EnforceConstraints(True)

        Load_ICTCOSTX()
        ' Absx1.txtFor("ORIG_CODE").Text = ""
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Order")

        BeginTrans()

        Dim sql_delete As String = "ORIG_CODE = '" & ORIG_CODE & "'"
        Update_Record_TDA("ICTCOSTC", sql_delete)
        Update_Record_TDA("ICTCOSTD", sql_delete)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdICTCOSTE, "SS", "Show Filter", "Show GroupBox")
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
        If COLUMN_NAME = "ORIG_CODE" Then
            If ctl.Text <> "" Then
                'Call Click_Command("Load Reports")
            End If
        End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "ORIG_CODE" Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then
                Click_Command("View")
            End If
        End If
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)

        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ORIG_CODE"
                If txtctl.Text <> "" Then
                    Click_Command("View")
                End If
        End Select
    End Sub

#End Region

#Region "grdICTCOSTD"

    Private Sub grdICTCOSTD_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCOSTD.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "VEND_CODE"
                If e.Cell.Value & "" <> "Z" Then
                    grdCodeDesc(grdICTCOSTD, "APTVEND1", "VEND_CODE", "VEND_NAME")
                End If
        End Select
    End Sub

    Private Sub grdICTCOSTD_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTCOSTD.AfterRowActivate
        Setup_ICTCOSTC()

        If grdICTCOSTD.ActiveRow.IsAddRow Then
            grdICTCOSTD.DisplayLayout.Bands(0).Columns("VEND_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdICTCOSTD.DisplayLayout.Bands(0).Columns("VEND_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdICTCOSTD_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTCOSTD.AfterRowsDeleted

    End Sub

    Private Sub grdICTCOSTD_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTCOSTD.AfterRowUpdate

    End Sub

    Private Sub grdICTCOSTD_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTCOSTD.BeforeRowUpdate
        Dim EMsg As String = ""

        If e.Row.Cells("VEND_CODE").Value & "" <> "Z" Then
            If Lookup("APTVEND1", e.Row.Cells("VEND_CODE").Value & "") Is Nothing Then
                EMsg &= vbCr & "Invalid Value specifed for Supplier"
                e.Cancel = True
            End If
        End If

        For Each COLUMN_NAME As String In New String() {"PAYEE_COMMISSION", "PAYEE_INSPECTION", "PAYEE_OTHER"}
            If e.Row.Cells(COLUMN_NAME).Value & "" <> "" Then
                If Lookup("APTVEND1", e.Row.Cells(COLUMN_NAME).Value & "") Is Nothing Then
                    EMsg &= vbCr & "Invalid Value specifed for " & grdICTCOSTD.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
                    e.Cancel = True
                End If
            End If
        Next

        If e.Cancel Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Add Row")
        Else
            If e.Row.IsAddRow Then
                e.Row.Cells("ORIG_CODE").Value = ORIG_CODE
            End If
        End If
    End Sub

    Private Sub grdICTCOSTD_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCOSTD.ClickCellButton
        grdClickCellButton(grdICTCOSTD, "", False, e.Cell.Column.Key, "VEND_CODE")
    End Sub

    Private Sub grdICTCOSTD_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCOSTD.InitializeRow
        If e.Row.Cells("VEND_CODE").Value & "" = "Z" Then
            e.Row.Cells("VEND_NAME").Value = "Any Supplier"
            e.Row.Update()
        End If
    End Sub
#End Region

#Region "grdICTCOSTC"

    Private Sub grdICTCOSTC_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCOSTC.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "PROD_CODE"
                If e.Cell.Value & "" <> "ZZZ" Then
                    grdCodeDesc(grdICTCOSTC, "ICTPROD1", "PROD_CODE", "PROD_DESC")
                End If
        End Select
    End Sub

    Private Sub grdICTCOSTC_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTCOSTC.AfterRowActivate
        Setup_ICTCOSTC()

        If grdICTCOSTC.ActiveRow Is Nothing Then

        Else
            If grdICTCOSTC.ActiveRow.IsAddRow Then
                grdICTCOSTC.DisplayLayout.Bands(0).Columns("PROD_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdICTCOSTC.DisplayLayout.Bands(0).Columns("BRAND_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                grdICTCOSTC.DisplayLayout.Bands(0).Columns("PROD_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                grdICTCOSTC.DisplayLayout.Bands(0).Columns("BRAND_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End If
    End Sub

    Private Sub grdICTCOSTC_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTCOSTC.AfterRowsDeleted

    End Sub

    Private Sub grdICTCOSTC_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTCOSTC.AfterRowUpdate

    End Sub

    Private Sub grdICTCOSTC_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTCOSTC.BeforeRowUpdate
        Dim EMsg As String = ""

        If e.Row.Cells("PROD_CODE").Value & "" <> "ZZZ" Then
            If Lookup("ICTPROD1", e.Row.Cells("PROD_CODE").Value & "") Is Nothing Then
                EMsg &= vbCr & "Invalid Value specifed for Product"
                e.Cancel = True
            End If
        End If

        If e.Row.Cells("BRAND_CODE").Value & "" <> "Z" Then
            If Lookup("ICTBRAN1", e.Row.Cells("BRAND_CODE").Value & "") Is Nothing Then
                EMsg &= vbCr & "Invalid Value specifed for Brand"
                e.Cancel = True
            End If
        End If

        If Val(e.Row.Cells("COMMISSION").Value & "") = 0 Then
            e.Row.Cells("COMM_TYPE").Value = ""
        Else
            If Val(e.Row.Cells("COMMISSION").Value & "") < 0 Then
                EMsg &= vbCr & "Invalid Value specifed for Commission"
                e.Cancel = True
            Else
                If e.Row.Cells("COMM_TYPE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Value specifed for Commission Type"
                    e.Cancel = True
                End If
            End If
        End If

        If Val(e.Row.Cells("INSPECTION").Value & "") = 0 Then
            e.Row.Cells("INSP_TYPE").Value = ""
        Else
            If Val(e.Row.Cells("INSPECTION").Value & "") < 0 Then
                EMsg &= vbCr & "Invalid Value specifed for Inspection"
                e.Cancel = True
            Else
                If e.Row.Cells("INSP_TYPE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Value specifed for Inspection Type"
                    e.Cancel = True
                End If
            End If
        End If

        If Val(e.Row.Cells("OTHER").Value & "") = 0 Then
            e.Row.Cells("OTHER_TYPE").Value = ""
        Else
            If Val(e.Row.Cells("OTHER").Value & "") < 0 Then
                EMsg &= vbCr & "Invalid Value specifed for Other"
                e.Cancel = True
            Else
                If e.Row.Cells("OTHER_TYPE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Value specifed for Other Type"
                    e.Cancel = True
                End If
            End If
        End If


        If e.Cancel Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Add Row")
        Else
            If e.Row.IsAddRow Then
                e.Row.Cells("ORIG_CODE").Value = ORIG_CODE
                e.Row.Cells("VEND_CODE").Value = grdICTCOSTD.ActiveRow.Cells("VEND_CODE").Value
            End If
        End If
    End Sub

    Private Sub grdICTCOSTC_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCOSTC.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "PROD_CODE"
                grdClickCellButton(grdICTCOSTC, "", False, e.Cell.Column.Key, "PROD_CODE")
            Case "BRAND_CODE"
                grdClickCellButton(grdICTCOSTC, "", False, e.Cell.Column.Key, "BRAND_CODE")
        End Select
    End Sub

    Private Sub grdICTCOSTC_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCOSTC.InitializeRow
        If e.Row.Cells("PROD_CODE").Value & "" = "ZZZ" Then
            e.Row.Cells("PROD_DESC").Value = "Any Product"
        End If
    End Sub
#End Region

    Sub Setup_ICTCOSTC()
        If grdICTCOSTD.ActiveRow Is Nothing OrElse grdICTCOSTD.ActiveRow.IsAddRow Then
            grdICTCOSTC.Visible = False
        Else
            Dim VEND_CODE As String = grdICTCOSTD.ActiveRow.Cells("VEND_CODE").Value
            Dim dvw As DataView = DirectCast(grdICTCOSTC.DataSource, DataTable).DefaultView
            dvw.RowFilter = "VEND_CODE = '" & VEND_CODE & "'"
            grdICTCOSTC.Text = "Commission Schedules for Origin " & ORIG_CODE & ", Supplier " & VEND_CODE

            grdICTCOSTC.Visible = True
        End If

        grdICTCOSTT.Visible = False

    End Sub

    Sub Load_ICTCOSTX()

        Dim ORIG_CODE As String = ""

        dst.Tables("ICTCOSTX").Rows.Clear()
        Dim rowICTCOSTX As DataRow = Nothing

        ASCMAIN1.sql = ASCMAIN1.Flattened_List("ORIG_CODE", "VEND_CODE", "ICTCOSTD", ",")
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            ORIG_CODE = row.Item("ORIG_CODE")
            Dim rowICTORIG1 As DataRow = Lookup("ICTORIG1", ORIG_CODE)
            rowICTCOSTX = dst.Tables("ICTCOSTX").NewRow
            rowICTCOSTX.Item("ORIG_CODE") = ORIG_CODE
            rowICTCOSTX.Item("ORIG_DESC") = rowICTORIG1.Item("ORIG_DESC") & ""
            rowICTCOSTX.Item("VEND_CODES") = row.Item("VEND_CODES")
            dst.Tables("ICTCOSTX").Rows.Add(rowICTCOSTX)
        Next
        Sort_grdColumns(grdICTCOSTC, "ORIG_CODE")

        'ASCMAIN1.sql = "Select * from ICTCOSTD"
        'For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "ORIG_CODE,VEND_CODE")
        '    If row.Item("ORIG_CODE") & "" <> ORIG_CODE Then
        '        ORIG_CODE = row.Item("ORIG_CODE")
        '        Dim rowICTORIG1 As DataRow = LookUp("ICTORIG1", ORIG_CODE)
        '        rowICTCOSTX = dst.Tables("ICTCOSTX").NewRow
        '        rowICTCOSTX.Item("ORIG_CODE") = ORIG_CODE
        '        rowICTCOSTX.Item("ORIG_DESC") = rowICTORIG1.Item("ORIG_DESC") & ""
        '        dst.Tables("ICTCOSTX").Rows.Add(rowICTCOSTX)
        '    End If
        '    Dim VEND_CODES As String = rowICTCOSTX.Item("VEND_CODES") & ""
        '    If VEND_CODES <> "" Then
        '        VEND_CODES &= ","
        '    End If
        '    VEND_CODES &= row.Item("VEND_CODE")
        '    rowICTCOSTX.Item("VEND_CODES") = VEND_CODES
        'Next

        dst.Tables("ICTCOSTX").AcceptChanges()
    End Sub

    Private Sub grdICTCOSTX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTCOSTX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ORIG_CODE").Text = e.Row.Cells("ORIG_CODE").Value
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTCOSTC_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTCOSTC.InitializeLayout

    End Sub

    Private Sub grdICTCOSTC_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdICTCOSTC.KeyDown
        If e.KeyValue = Keys.Delete Then
            If grdICTCOSTC.ActiveCell IsNot Nothing Then
                Dim COLUMN_NAME As String = grdICTCOSTC.ActiveCell.Column.Key
                If COLUMN_NAME = "COMM_TYPE" Or COLUMN_NAME = "INSP_TYPE" Or COLUMN_NAME = "OTHER_TYPE" Then
                    grdICTCOSTC.ActiveCell.Value = DBNull.Value
                End If
            End If
        End If
    End Sub

    Private Sub grdICTCOSTC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdICTCOSTC.KeyPress
    End Sub

    Private Sub cmdTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest.Click

        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text
        Dim PROD_CODE As String = Absx1.txtFor("PROD_CODE").Text
        Dim VEND_CODE As String = grdICTCOSTD.ActiveRow.Cells("VEND_CODE").Text

        If BRAND_CODE <> "Z" Then
            If Lookup("ICTBRAN1", BRAND_CODE) Is Nothing Then
                MsgBox("Invalid Value specified for Brand Code (" & BRAND_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Perform Test")
                Exit Sub
            End If
        End If
        If PROD_CODE <> "ZZZ" Then
            If Lookup("ICTPROD1", PROD_CODE) Is Nothing Then
                MsgBox("Invalid Value specified for Product Code (" & PROD_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Perform Test")
                Exit Sub
            End If
        End If

        dst.Tables("ICTCOSTT").Rows.Clear()

        Dim sqlwhere As String = "" _
        & "ORIG_CODE = '" & ORIG_CODE & "'" _
        & "   and (BRAND_CODE = '" & BRAND_CODE & "' or BRAND_CODE = 'Z')" _
        & "   and (VEND_CODE = '" & VEND_CODE & "' or VEND_CODE = 'Z')" _
        & "   and (PROD_CODE = '" & PROD_CODE & "' or PROD_CODE = 'ZZZ')"

        'Dim sqlsort As String = "" _
        '& " order by ORIG_CODE" _
        '& ", IIf(BRAND_CODE = 'Z','ZZZZZZZZZZZZ',BRAND_CODE)" _
        '& ", IIf(VEND_CODE = 'Z','ZZZZZZZZZZ',VEND_CODE)" _
        '& ", IIf(PROD_CODE = 'ZZZ','ZZZ',PROD_CODE)"

        Dim sqlsort As String = "" _
        & "ORIG_CODE,BRAND_CODE,VEND_CODE,PROD_CODE"

        Dim rowICTCOSTD As DataRow = dst.Tables("ICTCOSTD").Rows.Find(New String() {ORIG_CODE, VEND_CODE})

        Dim fee(3, 2) As String
        fee(1, 1) = "COMMISSION" : fee(1, 2) = "COMM"
        fee(2, 1) = "INSPECTION" : fee(2, 2) = "INSP"
        fee(3, 1) = "OTHER" : fee(3, 2) = "OTHER"

        For Each rowICTCOSTC As DataRow In dst.Tables("ICTCOSTC").Select(sqlwhere, sqlsort)

            For i As Integer = 1 To 3
                Dim AMT As Decimal = Val(rowICTCOSTC.Item(fee(i, 1)) & "")
                Dim COST As String = ""
                If Val(AMT) <> 0 And rowICTCOSTD.Item("PAYEE_" & fee(i, 1)) & "" <> "" Then
                    If rowICTCOSTC.Item(fee(i, 2) & "_TYPE") & "" = "%" Then
                        COST = Format$(AMT, "##.0000") & "%"
                    ElseIf rowICTCOSTC.Item(fee(1, 2) & "_TYPE") & "" = "$" Then
                        COST = Format$(AMT, "$##.0000")
                    Else
                        COST = "?" & Format$(rowICTCOSTC.Item(fee(i, 1)), "##.0000") & "?"
                    End If
                    dst.Tables("ICTCOSTT").Rows.Add(New Object() {fee(i, 2), rowICTCOSTD.Item("PAYEE_" & fee(i, 1)), COST})
                End If
            Next i

            Exit For
        Next

        grdICTCOSTT.Text = "Test for " & VEND_CODE
        grdICTCOSTT.Visible = True
    End Sub

    Sub Print_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Report")

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report("ICRCOSTC", "Commission Schedule", , , , , False)
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub Check_Open_POs()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Checking POs")

        dst.Tables("ICTCOSTV").Rows.Clear()

        ASCMAIN1.sql = "SELECT PO_ORDER_NO, VEND_CODE, ORIG_CODE" & vbCrLf _
        & ", VEND_CODE_AGENT, VEND_CODE_INSPECTION, VEND_CODE_OTHER" & vbCrLf _
        & " from POTORDR1" & vbCrLf _
        & " WHERE PO_STATUS_CODE = 'O' AND ORIG_CODE = '" & ORIG_CODE & "'" & vbCrLf _
        & " order by VEND_CODE"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim all_match As Boolean = False
            Dim VEND_CODE As String = row.Item("VEND_CODE") & ""
            Dim rowICTCOSTD As DataRow = dst.Tables("ICTCOSTD").Rows.Find(New String() {ORIG_CODE, VEND_CODE})
            If rowICTCOSTD Is Nothing Then
                rowICTCOSTD = dst.Tables("ICTCOSTD").Rows.Find(New String() {ORIG_CODE, "Z"})
            End If
            If rowICTCOSTD IsNot Nothing Then
                If rowICTCOSTD.Item("PAYEE_COMMISSION") & "" = row.Item("VEND_CODE_AGENT") & "" _
                And rowICTCOSTD.Item("PAYEE_INSPECTION") & "" = row.Item("VEND_CODE_INSPECTION") & "" _
                And rowICTCOSTD.Item("PAYEE_OTHER") & "" = row.Item("VEND_CODE_OTHER") & "" Then
                    all_match = True
                End If
            End If

            If Not all_match Then
                If VEND_CODE <> row.Item("VEND_CODE") & "" Then
                    VEND_CODE = row.Item("VEND_CODE") & ""
                    If rowICTCOSTD Is Nothing Then
                        dst.Tables("ICTCOSTV").Rows.Add _
                           (New String() {"Table",
                                          row.Item("VEND_CODE"),
                                          ".",
                                          ".",
                                          "."})
                    Else
                        dst.Tables("ICTCOSTV").Rows.Add _
                            (New String() {"Table",
                                           row.Item("VEND_CODE"),
                                           rowICTCOSTD.Item("PAYEE_COMMISSION"),
                                           rowICTCOSTD.Item("PAYEE_INSPECTION"),
                                           rowICTCOSTD.Item("PAYEE_OTHER")})
                    End If
                End If

                dst.Tables("ICTCOSTV").Rows.Add _
                    (New Object() {row.Item("PO_ORDER_NO"),
                                   row.Item("VEND_CODE"),
                                   row.Item("VEND_CODE_AGENT"),
                                   row.Item("VEND_CODE_INSPECTION"),
                                   row.Item("VEND_CODE_OTHER")})
            End If
        Next

        If dst.Tables("ICTCOSTV").Rows.Count <> 0 Then
            Using F As New ASFMSGBF
                F.Show_grd(dst.Tables("ICTCOSTV"), Me, "Open PO's with Different Payees than the Tables Show")
            End Using

        Else
            MsgBox("There are No Open PO's with different Payees than those reflected in the Tables, as shown." _
                   & vbCrLf & vbCrLf _
                   & "If you have made changes, please make sure that you click Update to save those changes.",
                   MsgBoxStyle.OkOnly, "Verification")
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

End Class