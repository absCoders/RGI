Public Class ICFQUOT1

    Dim rowICTQUOT1 As DataRow
    Dim sqlICTQUOTX As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")

        With dst
            sqlICTQUOTX = "Select ICTQUOT1.*, X.STYLE_CODE_PLM, X.STYLES" _
            & " from ICTQUOT1, (Select QUOTE_NO, MIN (STYLE_CODE_PLM) STYLE_CODE_PLM, Count (*) STYLES from ICTQUOT2 group by QUOTE_NO) X" _
            & " where X.QUOTE_NO = ICTQUOT1.QUOTE_NO and ICTQUOT1.QUOTE_TYPE = 'P'"
            ASCMAIN1.sql = sqlICTQUOTX
            Create_TDA(.Tables.Add, "ICTQUOTX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "ICTQUOT1", "*")
            With .Tables("ICTQUOT1")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() {.Columns("QUOTE_NO")}
            End With

            'ASCMAIN1.sql = "Select * from ICTQUOT2 where ROWNUM < 1"
            'Dim ICTQUOT2 As String = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTQUOT2 & " Add Primary Key (QUOTE_NO,STYLE_CODE_PLM)")
            'ASCMAIN1.sql = "Select ICTQUOT2.*, ICTPLIN2.SALES_DIVISION_CODE, ICTPLIN2.STYLE_CLASS_CODE" _
            '    & " from " & ICTQUOT2 & " ICTQUOT2, ICTPLIN2 where ICTPLIN2.STYLE_CODE_PLM = ICTQUOT2.STYLE_CODE_PLM"
            'Create_TDA(.Tables.Add("ICTQUOT2"), ICTQUOT2, "**", 0, True, "", 2)

            ASCMAIN1.sql = "Select ICTQUOT2.*, ICTPLIN2.SALES_DIVISION_CODE, ICTPLIN2.STYLE_CLASS_CODE, ICTPLIN2.STYLE_GROUP_CODE, ICTPLIN2.SEASON_CODE" _
                & " from ICTQUOT2, ICTPLIN2 where ICTPLIN2.STYLE_CODE_PLM = ICTQUOT2.STYLE_CODE_PLM" _
                & " and ICTQUOT2.QUOTE_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTQUOT2", "**", 0, True, "V", 2)
            With .Tables("ICTQUOT2")
                .Columns.Add("IMAGE_NAME")
                .Columns.Add("IMAGE", GetType(System.Byte()))
                .Columns("SEQ").DataType = GetType(System.Int32)
                .Columns.Add("WHSE_01")
                .Columns.Add("DATE_01", GetType(System.DateTime))
                .Columns.Add("QTY_01", GetType(System.Int64))
                .Columns.Add("WHSE_02")
                .Columns.Add("DATE_02", GetType(System.DateTime))
                .Columns.Add("QTY_02", GetType(System.Int64))
                .Columns.Add("WHSE_03")
                .Columns.Add("DATE_03", GetType(System.DateTime))
                .Columns.Add("QTY_03", GetType(System.Int64))
                .Columns.Add("WHSE_04")
                .Columns.Add("DATE_04", GetType(System.DateTime))
                .Columns.Add("QTY_04", GetType(System.Int64))
                .Columns.Add("STYLE_SPEC_01")
                .Columns.Add("STYLE_TYPE_DTL_01")
                .Columns.Add("STYLE_PRICE_01", GetType(System.Decimal))
                .Columns.Add("STYLE_SPEC_02")
                .Columns.Add("STYLE_TYPE_DTL_02")
                .Columns.Add("STYLE_PRICE_02", GetType(System.Decimal))
                .Columns.Add("STYLE_SPEC_03")
                .Columns.Add("STYLE_TYPE_DTL_03")
                .Columns.Add("STYLE_PRICE_03", GetType(System.Decimal))
                .Columns.Add("STYLE_SPEC_04")
                .Columns.Add("STYLE_TYPE_DTL_04")
                .Columns.Add("STYLE_PRICE_04", GetType(System.Decimal))
            End With

            Create_TDA(.Tables.Add, "ICTPLIN4", "*", 1, False)

        End With

        grdICTQUOT2.DataSource = dst.Tables("ICTQUOT2")
        grdICTQUOTX.DataSource = dst.Tables("ICTQUOTX")

        Create_Summary(grdICTQUOTX, "QUOTE_NO", "Count")

        Create_Summary(grdICTQUOT2, "SEQ", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("CUST_CODE", , True)

            Case "View", "Edit"
                If Absx1.txtFor("QUOTE_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a valid Quote No"
                Else
                    rowICTQUOT1 = LookUp("ICTQUOT1", Absx1.txtFor("QUOTE_NO").Text)
                    If rowICTQUOT1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Quote No " & Absx1.txtFor("QUOTE_NO").Text
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTQUOT1", Absx1.txtFor("QUOTE_NO").Text) Then Exit Sub
                End If

            Case "Update"
                If dst.Tables("ICTQUOT2").Select("").Length = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowICTQUOT2 As DataRow In dst.Tables("ICTQUOT2").Select("", "", DataViewRowState.CurrentRows)
                    Next
                End If

                If EMsg = "" Then

                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"

                If ASCMAIN1.USER_ID <> rowICTQUOT1.Item("INIT_OPER") & "" Then
                    EMsg &= vbCr & "Only " & rowICTQUOT1.Item("INIT_OPER") & " may Delete this Quote"
                End If

                If EMsg = "" Then
                    If MsgBox("Do you really want to Delete this Quote", _
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Print", "email"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer Code Specified"
                    End If
                End If
                If dst.Tables("ICTQUOT2").Select("").Length = 0 Then
                    EMsg &= vbCr & "No Styles on the Quote Sheet"
                End If

                'Case "Save Quote Sheet"
                '    If txtQUOTE_DESC.Text = "" Then
                '        EMsg &= vbCr & "Please enter a Description for the Quote Sheet"
                '    End If

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

            Case "Print", "email"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Printing")

                '  Update_Record_TDA("ICTQUOT2", "1=1")
                Synch_TABLE_NAME("ICTQUOT1")

                Dim FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG"
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    rowICTQUOT1.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
                End If

                For Each rowICTQUOT2 As DataRow In dst.Tables("ICTQUOT2").Select("")
                    FetchImage(rowICTQUOT2)
                    Load_Pricing(rowICTQUOT2)
                Next

                Print_Report_Begin()
                CR_params.Add("CHKOMITPRICE", IIf(chkOmitPrice.Checked, "1", "0"))
                CR_params.Add("CHKOMITPRICE2", IIf(chkOmitPrice2.Checked, "1", "0"))
                CR_params.Add("CHKOMITAVAIL", "1")

                Dim RPT As String = "ICRQUOT1"
                If chk1perPage.Checked Then RPT = "ICRQUOTN"

                If eItemKey = "email" Then
                    Dim tempFileName As String = rowICTQUOT1.Item("QUOTE_NO")
                    Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)
                    ' Dim FILENAME As String = REPORT_FILENAMES(REPORT_NO)
                    Print_Report_End(, True)
                    email_Quote(tempFileName)
                Else
                    Generate_Report(RPT, "Quote Sheet")
                    Print_Report_End()
                End If

                'Case "Clear Quote Sheet"
                '    dst.Tables("ICTQUOT2").Rows.Clear()
                '    Setup_Style_Quoted()
                '    txtQUOTE_DESC.Text = ""
                '    Absx1.txtFor("CUST_CODE").Text = ""

                'Case "Save Quote Sheet"
                '    Update_Record_TDA("ICTQUOT1")

                'grdICTQUOT2.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
                Setup_Style_Quoted()
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                    .Items("Print").Visible = ScreenMode
                    .Items("email").Visible = ScreenMode
                    .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                    .Items("Delete").Visible = (EntryMode = "E")
                    .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                End With

                .Groups("Style Image").Visible = ScreenMode
                .Groups("Show Quote Sheets").Visible = Not ScreenMode
            End With
        End If

        Setup_Style_Quoted()

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CODE"), False)
        SplitContainer1.Visible = ScreenMode
   
        grdICTQUOTX.Visible = Not ScreenMode

        If ScreenMode Then
            Set_Read_Only(grpHeader, (EntryMode = "V"))
            Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "V" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTQUOT2}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTQUOT2}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Next
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTQUOTX", "ICTQUOT1", "ICTQUOT2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        grdICTQUOT2.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.dteFor("QUOTE_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("QUOTE_NO").Text = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowICTQUOT1 = dst.Tables("ICTQUOT1").NewRow
            rowICTQUOT1.Item("QUOTE_NO") = ASCMAIN1.Next_Control_No("ICTQUOT1.QUOTE_NO")
            rowICTQUOT1.Item("CUST_CODE") = HFs("CUST_CODE")
            rowICTQUOT1.Item("QUOTE_DATE") = HFs("QUOTE_DATE")
            rowICTQUOT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTQUOT1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTQUOT1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTQUOT1.Item("LAST_DATE") = DATETIME_STAMP
            rowICTQUOT1.Item("QUOTE_TYPE") = "P"
            dst.Tables("ICTQUOT1").Rows.Add(rowICTQUOT1)
        Else
            rowICTQUOT1 = Fill_Record("ICTQUOT1", Absx1.txtFor("QUOTE_NO").Text)
            dst.AcceptChanges()
        End If

        Fill_Records("ICTQUOT2", Absx1.txtFor("QUOTE_NO").Text)
        Sort_grdColumns(grdICTQUOT2, "SEQ")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("ICTQUOT1")
        Update_Record_TDA("ICTQUOT2")
        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Delete_Records("ICTQUOT1")
        Delete_Records("ICTQUOT2")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where QUOTE_NO = '" & Absx1.txtFor("QUOTE_NO").Text & "'")
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("QUOTE_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ICTQUOT1"
            E.COLUMN_NAME = "QUOTE_NO"
            E.CODE_VALUE = Absx1.txtFor("QUOTE_NO").Text
            E.DESC_VALUE = "Quote No"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTQUOTX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTQUOT2, "SBBBB", "Show Filter", "Style Status Inquiry", "Product Line Maintenance", "Product Line Inquiry", "Sequence as Shown")

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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdICTQUOT2"
                    tlb_btn = DirectCast(tlb_pop.Tools("Sequence as Shown"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode

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

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)


            Case "Sequence as Shown"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Re-Sequencing by 10's")

                Dim SEQ As Integer = 0
                For Each grow As UltraWinGrid.UltraGridRow In grdICTQUOT2.Rows
                    SEQ += 10
                    grow.Cells("SEQ").Value = SEQ
                    grow.Update()
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Product Line Maintenance"
                Dim STYLE_CODE_PLM As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Text
                Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                If rowICTPLIN2 IsNot Nothing Then
                    Context_Launch("View", STYLE_CODE_PLM, e.Tool.Key, "ICFPLIN1")
                End If

            Case "Product Line Inquiry"
                Dim STYLE_CODE_PLM As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Text
                Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                If rowICTPLIN2 IsNot Nothing Then
                    Context_Launch("View", STYLE_CODE_PLM, e.Tool.Key, "ICFPLINI")
                End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "QUOTE_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View", e)
                End If

            Case "STYLE_CODE_PLM"
                If e.KeyCode = Windows.Forms.Keys.Enter And ScreenMode Then
                    Dim STYLE_CODE_PLM As String = txtQS_STYLE_CODE.Text
                    If LookUp("ICTPLIN2", STYLE_CODE_PLM) IsNot Nothing Then
                        Add_to_Quote(STYLE_CODE_PLM)
                    End If
                End If

        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If Not InquiryMode And Not ScreenMode Then
                    Click_Command("New")
                End If
            Case "QUOTE_NO"
                If Not ScreenMode Then Click_Command("View")
            Case "STYLE_CODE_PLM"
                Dim STYLE_CODE_PLM As String = txtQS_STYLE_CODE.Text
                If LookUp("ICTPLIN2", STYLE_CODE_PLM) IsNot Nothing Then
                    Add_to_Quote(STYLE_CODE_PLM)
                End If
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case ""
        End Select
    End Sub

#End Region

#Region "grdICTQUOT2"

    Private Sub grdICTQUOT2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTQUOT2.AfterCellUpdate
        If Not e.Cell.Row.IsDataRow Then Exit Sub
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_PLM"

                grdCodeDesc(grdICTQUOT2, "ICTPLIN2", "STYLE_CODE_PLM", "STYLE_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE STYLE_DESC
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE_PLM As String = e.Cell.Value
                    e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")
                    e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = cdr.Item("SALES_DIVISION_CODE")
                    e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = cdr.Item("STYLE_CLASS_CODE") & ""
                    '  e.Cell.Row.Cells("STYLE_PRICE").Value = cdr.Item("STYLE_PRICE")
                    e.Cell.Row.Cells("SIZE_SCALE").Value = Get_Colors(STYLE_CODE_PLM)

                    ASCMAIN1.sql = "Select MAX(DECODE(NVL(STYLE_PRICE_OVERRIDE,0),0,STYLE_PRICE,STYLE_PRICE_OVERRIDE)) STYLE_PRICE, COUNT (*) COMPS " _
                        & " from ICTPLIN3 where STYLE_CODE_PLM = '" & STYLE_CODE_PLM & "'"
                    Dim row As DataRow = ASCDATA1.GetDataRow
                    If row IsNot Nothing AndAlso Val(row.Item("COMPS") & "") = 1 Then
                        e.Cell.Row.Cells("STYLE_PRICE").Value = row.Item("STYLE_PRICE")
                    End If

                Else
                    grdICTQUOT2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If
        End Select
    End Sub

    Private Sub grdICTQUOT2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTQUOT2.AfterRowActivate

        If Not grdICTQUOT2.ActiveRow.IsDataRow Then Exit Sub

        Setup_Style_Quoted()

        With grdICTQUOT2.DisplayLayout.Bands(0)
            If grdICTQUOT2.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE_PLM").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdICTQUOT2.ActiveCell = grdICTQUOT2.ActiveRow.Cells("STYLE_CODE_PLM")
                grdICTQUOT2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE_PLM").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

    End Sub

    Private Sub grdICTQUOT2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTQUOT2.AfterRowUpdate

        Dim QUOTE_NO As String = e.Row.Cells("QUOTE_NO").Value
        Dim STYLE_CODE_PLM As String = e.Row.Cells("STYLE_CODE_PLM").Value
        Dim rowICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New Object() {QUOTE_NO, STYLE_CODE_PLM})
        FetchImage(rowICTQUOT2)
    End Sub

    Private Sub grdICTQUOT2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTQUOT2.BeforeExitEditMode
        If grdICTQUOT2.ActiveCell Is Nothing Then Exit Sub
        If Not grdICTQUOT2.ActiveRow.IsDataRow Then Exit Sub
        With grdICTQUOT2.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE_PLM"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTPLIN2", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdICTQUOT2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTQUOT2.BeforeRowUpdate
        With grdICTQUOT2
            If e.Row.Cells("STYLE_CODE_PLM").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTPLIN2", e.Row.Cells("STYLE_CODE_PLM").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE_PLM").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("QUOTE_NO").Text = "" Then
                    .ActiveRow.Cells("QUOTE_NO").Value = Absx1.CtlFor("QUOTE_NO").Text
                    .ActiveRow.Cells("SEQ").Value = Val(dst.Tables("ICTQUOT2").Compute("Max(SEQ)", "") & "") + 10
                End If
            End If
        End With
    End Sub

    Private Sub grdICTQUOT2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTQUOT2.ClickCellButton

        If grdICTQUOT2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_PLM"
        End Select
        grdClickCellButton(grdICTQUOT2, sql_where, False)

    End Sub

#End Region

    Private Sub grdICTQUOTX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTQUOTX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("QUOTE_NO").Text = e.Row.Cells("QUOTE_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Quotes")
        ASCMAIN1.sql = sqlICTQUOTX
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If optShow.Value = "A" And CUST_CODE = "" Then
            grdICTQUOTX.Text = "All Quotes"
        ElseIf optShow.Value = "M" Then
            ASCMAIN1.sql &= " and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
            grdICTQUOTX.Text = "Quotes entered or modified by Me"
        ElseIf optShow.Value = "C" Or CUST_CODE <> "" Then
            ASCMAIN1.sql &= " and CUST_CODE = '" & CUST_CODE & "'"
            grdICTQUOTX.Text = "Quotes associated with " & CUST_CODE
        End If
        Fill_Records("ICTQUOTX")
        Sort_grdColumns(grdICTQUOTX, "QUOTE_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Add_to_Quote(STYLE_CODE_PLM As String) As String
        STYLE_CODE_PLM = STYLE_CODE_PLM.ToUpper
        Dim QUOTE_NO As String = Absx1.txtFor("QUOTE_NO").Text
        Dim rowICTQUOT2 As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New String() {QUOTE_NO, STYLE_CODE_PLM})
        If rowICTQUOT2 Is Nothing Then
            Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
            If rowICTPLIN2 Is Nothing Then Return ""
            rowICTQUOT2 = dst.Tables("ICTQUOT2").NewRow()

            With rowICTQUOT2
                .Item("QUOTE_NO") = QUOTE_NO
                .Item("STYLE_CODE_PLM") = STYLE_CODE_PLM
                .Item("STYLE_CODE_CUST") = STYLE_CODE_PLM
                .Item("STYLE_DESC") = rowICTPLIN2.Item("STYLE_DESC")
                Dim SEQ As Integer = Val(dst.Tables("ICTQUOT2").Compute("MAX(SEQ)", "") & "") + 10
                .Item("SEQ") = SEQ

                ASCMAIN1.sql = "Select MAX(STYLE_PRICE) STYLE_PRICE, COUNT (*) COMPS " _
                      & " from ICTPLIN3 where STYLE_CODE_PLM = '" & STYLE_CODE_PLM & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If row IsNot Nothing AndAlso Val(row.Item("COMPS") & "") = 1 Then
                    .Item("STYLE_PRICE") = row.Item("STYLE_PRICE")
                End If

                '  .Item("STYLE_PRICE") = rowICTPLIN2.Item("STYLE_PRICE")

                .Item("SIZE_SCALE") = Get_Colors(STYLE_CODE_PLM)
                '.Item("STYLE_DESC2") = rowICTPLIN2.Item("STYLE_DESC2")
                .Item("SALES_DIVISION_CODE") = rowICTPLIN2.Item("SALES_DIVISION_CODE")
                .Item("STYLE_CLASS_CODE") = rowICTPLIN2.Item("STYLE_CLASS_CODE")
                .Item("STYLE_GROUP_CODE") = rowICTPLIN2.Item("STYLE_GROUP_CODE")
                .Item("SEASON_CODE") = rowICTPLIN2.Item("SEASON_CODE")
            End With
            dst.Tables("ICTQUOT2").Rows.Add(rowICTQUOT2)
            ' FetchImage(rowICTQUOT2)
            Load_Pricing(rowICTQUOT2)
        End If
        txtQS_STYLE_CODE.Text = ""
        Return STYLE_CODE_PLM
    End Function

    Sub Load_Pricing(rowICTQUOT2 As DataRow)

        Dim STYLE_CODE_PLM As String = rowICTQUOT2.Item("STYLE_CODE_PLM")
        Dim A As Integer = 0

        ASCMAIN1.sql = "Select * from ICTPLIN3 where STYLE_CODE_PLM = :PARM1"
        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {STYLE_CODE_PLM})

        If tbl.Rows.Count = 1 Then Return

        For Each row As DataRow In tbl.Select("", "STYLE_TYPE_SEQ")
            A += 1
            If A <= 4 Then
                rowICTQUOT2.Item("STYLE_SPEC_" & Format(A, "00")) = row.Item("STYLE_SPEC")
                rowICTQUOT2.Item("STYLE_TYPE_DTL_" & Format(A, "00")) = IIf(row.Item("STYLE_TYPE_DTL") & "" <> "", row.Item("STYLE_TYPE_DTL"), row.Item("DUTY_CATGY_CODE"))
                Dim LANDED_COST As Decimal = Val(row.Item("PO_COST") & "") _
                                           + Val(row.Item("OTHER_COST") & "") _
                                           + (Val(row.Item("PO_COST") & "") + Val(row.Item("OTHER_COST") & "")) * Val(row.Item("DUTY_RATE") & "") / 100 _
                                           + Val(row.Item("FREIGHT_COST") & "") _
                                           + Val(row.Item("BRKR_COST") & "") _
                                           + Val(row.Item("MISC_COST") & "") _
                                           + Val(row.Item("INLAND_COST") & "") _
                                           + Val(row.Item("LABOR_COST") & "")
                Dim ROYALTY_PCT As Decimal = 0
                Dim STYLE_PRICE As Decimal = TAC.ICCMAIN1.Calculate_Suggested_SP(LANDED_COST, ROYALTY_PCT)
                'rowICTQUOT2.Item("STYLE_PRICE_" & Format(A, "00")) = row.Item("STYLE_PRICE")
                rowICTQUOT2.Item("STYLE_PRICE_" & Format(A, "00")) = STYLE_PRICE
            End If
        Next
        'Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_PLM)
        'rowICTQUOT2.Item("IMAGE_NAME") = rowICTSTYL1.Item("IMAGE_NAME")
        'FetchImage(rowICTQUOT2)

    End Sub

    Sub Setup_Style_Quoted()
        If grdICTQUOT2.ActiveRow Is Nothing OrElse grdICTQUOT2.ActiveRow.IsAddRow Then
            picStyleImage.Image = Nothing
            UltraExplorerBar1.Groups("Style Image").Text = "Style Image"
        Else
            Dim SALES_DIVISION_CODE As String = grdICTQUOT2.ActiveRow.Cells("SALES_DIVISION_CODE").Value & ""
            Dim STYLE_CODE_PLM As String = grdICTQUOT2.ActiveRow.Cells("STYLE_CODE_PLM").Value & ""
            Dim IMAGE_NAME As String = SALES_DIVISION_CODE & "\" & STYLE_CODE_PLM & ".jpg"
            picStyleImage.Image = Get_Style_Image(IMAGE_NAME)
            UltraExplorerBar1.Groups("Style Image").Text = "Style Image " & STYLE_CODE_PLM
        End If
    End Sub

    Sub email_Quote(tempFileName As String)
        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim CUST_NAME As String = Absx1.txtFor("CUST_NAME").Text
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        ATTACHMENTs.Add(tempFileName & ".pdf", ASCMAIN1.Folders("Temp") & tempFileName & ".pdf")

        Dim SUBJECT As String = "Quote Sheet"
        Dim PFX As String = ""

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        If CUST_CODE <> "" Then
            EMAIL_ADDRESSs.Add(rowARTCUST1.Item("CUST_EMAIL") & "", rowARTCUST1.Item("CUST_CONTACT") & "")
        End If

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                SUBJECT, "ICTQUOT1", False, True, CUST_CODE, CUST_NAME, "Customer")
        If SEND_NO <> "" Then
            TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "QUOEML", "Quote Sheet emailed", SEND_NO)
        End If
    End Sub

    Function FetchImage(rowICTPLIN2 As DataRow) As Byte()
        Dim STYLE_CODE_PLM As String = rowICTPLIN2.Item("STYLE_CODE_PLM") & ""
        Dim SALES_DIVISION_CODE As String = rowICTPLIN2.Item("SALES_DIVISION_CODE") & ""
        Dim IMAGE_NAME As String = SALES_DIVISION_CODE & "\" & STYLE_CODE_PLM & ".jpg"
 
        ' Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & "\" & Absx1.txtFor("SALES_DIVISION_CODE").Text
        If My.Computer.FileSystem.FileExists(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & "\" & SALES_DIVISION_CODE & "\" & STYLE_CODE_PLM & ".png") Then
            IMAGE_NAME = SALES_DIVISION_CODE & "\" & STYLE_CODE_PLM & ".png"
        End If

        Dim imgba() As Byte = Nothing
        If IMAGE_NAME <> "" Then
            Dim bm_source As Bitmap = Get_Style_Image(IMAGE_NAME, imgba)
            'bm.Size.Height = 200

            Dim target As Integer = 90
            Dim size As Integer = bm_source.Width
            If bm_source.Height > bm_source.Width Then
                size = bm_source.Height
            End If
            Dim scale_factor As Single = target / size

            ' Make a bitmap for the result.
            Dim bm_dest As New Bitmap( _
                CInt(bm_source.Width * scale_factor), _
                CInt(bm_source.Height * scale_factor))

            ' Make a Graphics object for the result Bitmap.
            Dim gr_dest As Graphics = Graphics.FromImage(bm_dest)
            gr_dest.Clear(Color.White)
            ' Copy the source image into the destination bitmap.
            gr_dest.DrawImage(bm_source, 0, 0, _
                bm_dest.Width + 1, _
                bm_dest.Height + 1)

            '  Dim backColor As Color = bm_dest.GetPixel(41, 41)
            'bm_dest.MakeTransparent(backColor) ' Color.Black) ' magenta in bitmap will be transparent
            '     imgba = DirectCast(System.ComponentModel.TypeDescriptor.GetConverter(bm_dest).ConvertTo(bm_dest, GetType(Byte())), Byte())

            Dim scale_image As Boolean = False
            If scale_image Then
                Using Bmp As New Bitmap( _
                    bm_source.Width * scale_factor, _
                    bm_source.Height * scale_factor, _
                    Imaging.PixelFormat.Format32bppPArgb)

                    'Set the resolution to 300 DPI
                    Bmp.SetResolution(300, 300)
                    'Create a graphics object from the bitmap
                    Using G = Graphics.FromImage(Bmp)
                        'Paint the canvas white
                        G.Clear(Color.White)
                        'Set various modes to higher quality
                        G.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        G.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        G.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias

                        G.CompositingQuality = Drawing.Drawing2D.CompositingQuality.HighQuality
                        G.PixelOffsetMode = Drawing.Drawing2D.PixelOffsetMode.HighQuality
                        G.CompositingMode = Drawing.Drawing2D.CompositingMode.SourceOver
                        '  G.DrawImage(bm_source, NewX, NewY, NewWidth, NewHeight)

                        G.DrawImage(bm_source, 0, 0, _
                                     Bmp.Width + 1, _
                                     Bmp.Height + 1)
                    End Using

                    imgba = DirectCast(System.ComponentModel.TypeDescriptor.GetConverter(Bmp).ConvertTo(Bmp, GetType(Byte())), Byte())
                End Using
            End If

            picStyleImage.Image = Nothing
            ' imgba = DirectCast(System.ComponentModel.TypeDescriptor.GetConverter(picStyleImage.Image).ConvertTo(picStyleImage.Image, GetType(Byte())), Byte())
            rowICTPLIN2.Item("IMAGE") = imgba

        Else
            rowICTPLIN2.Item("IMAGE") = DBNull.Value
        End If

        Return imgba
    End Function

    Function Get_Style_Image(
        ByVal IMAGE_NAME As String, _
        Optional ByRef imgba() As Byte = Nothing) As System.Drawing.Bitmap

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")
        Return ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)

    End Function

    Private Sub optShow_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShow.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Refresh_Documents()
    End Sub

    Function Get_Colors(STYLE_CODE_PLM As String) As String
        Dim COLORS As String = ""
        Fill_Records("ICTPLIN4", STYLE_CODE_PLM)
        For Each row As DataRow In dst.Tables("ICTPLIN4").Select("", "STYLE_DTL_SEQ")
            Dim STYLE_DTL_QTY As Int32 = Val(row.Item("STYLE_DTL_QTY") & "")
            Dim STYLE_DTL_COLOR As String = row.Item("STYLE_DTL_COLOR") & ""
            If STYLE_DTL_COLOR <> "" Then
                Dim C As String = STYLE_DTL_COLOR
                If STYLE_DTL_QTY <> 0 Then
                    C = CStr(STYLE_DTL_QTY) & " " & STYLE_DTL_COLOR
                End If
                COLORS &= vbCrLf & C
            End If
        Next
        If COLORS <> "" Then
            COLORS = Mid(COLORS, 3)
        End If
        Return COLORS
    End Function

    Private Sub txtQS_STYLE_CODE_EditorButtonClick(sender As Object, e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs) Handles txtQS_STYLE_CODE.EditorButtonClick

    End Sub

    Private Sub txtQS_STYLE_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtQS_STYLE_CODE.ValueChanged

    End Sub

    Private Sub cmdAddMultipleStyles_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddMultipleStyles.Click
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE_PLM")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                For Each STYLE_CODE_PLM As String In ASCMAIN1.CodeSelector.SelectedCodes
                    Add_to_Quote(STYLE_CODE_PLM)
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If
    End Sub
End Class