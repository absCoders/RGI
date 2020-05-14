Public Class SOFCSTP1

#Region "Declarations"
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim rowSOTCSTP2 As DataRow
    Dim sqlSOTCSTP2 As String
    Dim rowICTSTYL1 As DataRow
    Dim COLOR_CODEs As New List(Of String)    ' table of COLOR_CODEs associated with a STYLE_CODE
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            sqlSOTCSTP2 = "Select SOTCSTP2.*, ARTCUST1.CUST_NAME" _
                & " from SOTCSTP2,ARTCUST1 where ARTCUST1.CUST_CODE = SOTCSTP2.CUST_CODE"
            ASCMAIN1.sql = sqlSOTCSTP2 _
                & " and SOTCSTP2.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTCSTP2", "**", 0, True, "V")

            ASCMAIN1.sql = "Select SOTCSTP1.*,ICTSTYL1.STYLE_DESC,ICTCOLR1.COLOR_DESC" _
                & " from SOTCSTP1,ICTSTYL1,ICTCOLR1" _
                & " where ICTSTYL1.STYLE_CODE = SOTCSTP1.STYLE_CODE" _
                & "   and ICTCOLR1.COLOR_CODE = SOTCSTP1.COLOR_CODE" _
                & "   and SOTCSTP1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTCSTP1", "**", 0, True, "V")
            .Tables("SOTCSTP1").Columns.Add("STATE")

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
               & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
               & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)
        End With

        grdSOTCSTP2.DataSource = dst.Tables("SOTCSTP2")
        grdSOTCSTP1.DataSource = dst.Tables("SOTCSTP1")

        grdSOTCSTP2.DisplayLayout.UseFixedHeaders = True
        With grdSOTCSTP2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTCSTP1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "STYLE_DESC" Or gcol.Key = "COLOR_DESC" Or gcol.Key = "STATE" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
        End With

        With grdSOTCSTP2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                'If gcol.Key = "PO_QTY" Then
                '    gcol.CellAppearance.BackColor = Drawing.Color.LightYellow
                '    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                '    '  gcol.CellAppearance.BackColor = Drawing.Color.Beige
                '    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                'End If

                'If New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(gcol.Key) Then
                '    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                '    gcol.Header.Appearance.BackColor = Drawing.Color.White
                '    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                'ElseIf New String() {"ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}.Contains(gcol.Key) Then
                '    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                '    gcol.Header.Appearance.BackColor = Drawing.Color.White
                '    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                'ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                '    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                '    gcol.Header.Appearance.BackColor = Drawing.Color.White
                '    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                'ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(gcol.Key) Then
                '    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                '    gcol.Header.Appearance.BackColor = Drawing.Color.White
                '    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                'ElseIf New String() {"ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT"}.Contains(gcol.Key) Then
                '    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                '    gcol.Header.Appearance.BackColor = Drawing.Color.White
                '    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                'ElseIf New String() {"CUST_CODE", "ORDR_GROUP_NO", "SALES_DIVISION_CODE"}.Contains(gcol.Key) Then
                '    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                '    gcol.Header.Appearance.BackColor = Drawing.Color.White
                '    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                'Else
                '    gcol.Header.Appearance.BackColor = Drawing.Color.LightGray
                '    gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                'End If
            Next
        End With

        '  Bind_Controls(grpSOTCSTP2, "SOTCSTP2")

        Create_Summary(grdSOTCSTP2, "CUST_CODE", "Count")
        Create_Summary(grdSOTCSTP1, "STYLE_CODE", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New", "Edit", "View"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTCSTP1", CUST_CODE) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOTCSTP2", CUST_CODE) Then Exit Sub
                End If

            Case "Update"

                'If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                '    EMsg &= vbCr & "Customer PO is required"
                'End If

            Case "Cancel"
                If EMsg = "" Then
                    If MsgBox("Do you really want to Cancel all changes made to this record?", _
                              MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "Verification") <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                If EMsg = "" Then
                    If MsgBox("Do you really want to Delete all Pricing and Cartonization Data for Customer " & CUST_CODE & "?", _
                              MsgBoxStyle.YesNo + MsgBoxStyle.Critical, "WARNING! - Answering 'Yes' will PERMANENTLY DELETE these records") <> MsgBoxResult.Yes Then
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

            Case "Select"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            If (EntryMode = "V" And ScreenMode) Then
                .Groups("Screen Control").Items("Edit").Settings.Enabled = DefaultableBoolean.True
            Else
                .Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            End If
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode

            .Groups("Screen Control").Items("Delete").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            ' .Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode

            .Groups("Screen Control").Items("Done").Visible = (EntryMode = "V" And ScreenMode)
            '.Groups("Screen Control").Items("Print").Visible = False ' ScreenMode
            .Groups("Screen Control").Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
            .Groups("Screen Control").Items("Delete").Visible = (EntryMode = "E")
            .Groups("Screen Control").Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
        End With

        grdSOTCSTP2.Visible = Not ScreenMode
        grdSOTCSTP1.Visible = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpSOTCSTP2.Visible = ScreenMode
        Set_Read_Only(grpSOTCSTP2, Not ScreenMode Or EntryMode = "V")

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTCSTP1}
                If EntryMode = "V" Then
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                Else
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                End If
            Next
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        Absx1.txtFor("CUST_CODE").Text = ""
        CUST_CODE = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTCSTP1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        grdSOTCSTP1.Rows.ColumnFilters.ClearAllFilters()

        Load_SOTCSTP2("")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        '  Load_SOTCSTP2("", CUST_CODE)

        rowSOTCSTP2 = Fill_Record("SOTCSTP2", CUST_CODE)

        Fill_Records("SOTCSTP1", CUST_CODE)
        Sort_grdColumns(grdSOTCSTP1, "STYLE_CODE")

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        For Each TABLE_NAME As String In New String() _
            {"SOTCSTP1", "SOTCSTP2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where CUST_CODE = '" & CUST_CODE & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        'If EntryMode <> "N" Then Delete_Records()
        INIT_LAST("SOTCSTP2", False, , True)
        'Dim sqldelete As String = "RSRV_NO = '" & RSRV_NO & "'"
        Update_Record_TDA("SOTCSTP1") ', sqldelete)
        Update_Record_TDA("SOTCSTP2") ', sqldelete)

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Select"

                Dim CUST_CODE As String = Split(key, ":")(0)
                Dim ORDR_GROUP_NO As String = Split(key & ":", ":")(1)
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Click_Command("Select")
                If ORDR_GROUP_NO <> "" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTCSTP2.Rows
                        If grow.Cells("ORDR_GROUP_NO").Value = ORDR_GROUP_NO Then
                            grdSOTCSTP2.ActiveRow = grow
                        End If

                    Next
                End If
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ARTCUST1"
            E.COLUMN_NAME = "CUST_CODE"
            E.CODE_VALUE = Absx1.txtFor("CUST_CODE").Text
            E.DESC_VALUE = "Customer"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTCSTP2, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTCSTP1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTORDR0"


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select")
                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "CUST_CODE"
            '    If Not ScreenMode Then
            '        Load_SOTORDRX()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("Select")

        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_SOTCSTP2(Optional PARM1 As String = "", Optional CUST_CODE As String = "")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Customers", "")

        ASCMAIN1.sql = sqlSOTCSTP2
        Fill_Records("SOTCSTP2", "", , ASCMAIN1.sql)
        Sort_grdColumns(grdSOTCSTP2, "CUST_CODE")
        grdSOTCSTP2.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub grdSOTCSTP2_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTCSTP2.DoubleClickRow
        If Not ScreenMode Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
            Click_Command("View")
        End If
    End Sub


    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        ' Setup_tabMain()
    End Sub

#Region "grdSOTCSTP1"

    Private Sub grdSOTCSTP1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCSTP1.AfterCellUpdate
        With grdSOTCSTP1.ActiveRow
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value)
                    If STYLE_CODE <> "" Then
                        .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        If COLOR_CODEs.Count = 1 Then
                            .Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                        End If
                    End If

                Case "COLOR_CODE"
                    Dim COLOR_CODE As String = e.Cell.Value & ""
                    If COLOR_CODE <> "" Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        If rowICTCOLR1 IsNot Nothing Then
                            .Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdSOTCSTP1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCSTP1.AfterRowActivate

    End Sub

    Private Sub grdSOTCSTP1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTCSTP1.AfterRowsDeleted

    End Sub

    Private Sub grdSOTCSTP1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCSTP1.AfterRowUpdate

    End Sub

    Private Sub grdSOTCSTP1_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTCSTP1.BeforeCellUpdate

    End Sub

    Private Sub grdSOTCSTP1_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTCSTP1.BeforeExitEditMode
        If grdSOTCSTP1.ActiveCell IsNot Nothing Then
            With grdSOTCSTP1.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value, .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTCSTP1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTCSTP1.BeforeRowsDeleted

    End Sub

    Private Sub grdSOTCSTP1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCSTP1.BeforeRowUpdate

        Validate_Columns(grdSOTCSTP1, "STYLE_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns(grdSOTCSTP1, "COLOR_CODE", e.Cancel)
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = CUST_CODE
            e.Row.Cells("STATE").Value = "Added"
        Else
            e.Row.Cells("STATE").Value = "Edited"
        End If
    End Sub

    Private Sub grdSOTCSTP1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCSTP1.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTCSTP1, sql_where)

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE IN (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE ='" & grdSOTCSTP1.ActiveRow.Cells("STYLE_CODE").Value & "')"

                    grdClickCellButton(grdSOTCSTP1, sql_where)
            End Select
        End With

    End Sub
#End Region

    Sub Validate_Columns(grd As UltraWinGrid.UltraGrid, COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grd.ActiveRow
            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = ""
                    If Trim(.Cells("STYLE_CODE").Value & "") <> "" Then
                        STYLE_CODE = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    End If
                    Cancel = (STYLE_CODE = "")

                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim EMsg As String = ""
        If STYLE_CODE_z = "" Then Return ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

        If rowICTSTYL1 Is Nothing Then
            EMsg = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                EMsg = "Item Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                EMsg = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                EMsg = "Item does not have a valid Division Code" & vbCrLf
            End If
        End If

        If EMsg = "" Then
            COLOR_CODEs.Clear()
            Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In dst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If EMsg <> "" And grdSOTCSTP1.ActiveRow.IsAddRow Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If EMsg = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function
End Class