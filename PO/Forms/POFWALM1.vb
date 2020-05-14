Public Class POFWALM1

    Dim rowPOTWPDM1 As DataRow
    Dim sqlPOTWPDMX As String = ""
    Dim STYLE_GROUP_NO As String
    Dim images_folder As String = "C:\dmp\Images"
    Dim images As New Dictionary(Of String, List(Of System.Drawing.Bitmap))

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")

        With dst
            sqlPOTWPDMX = "Select POTWPDM1.*, X.STYLE_CODE_1, X.STYLES" _
            & " from POTWPDM1, (Select STYLE_GROUP_NO, MIN (STYLE_CODE_PLM) STYLE_CODE_1, Count (*) STYLES from POTWPDM2 group by STYLE_GROUP_NO) X" _
            & " where X.STYLE_GROUP_NO = POTWPDM1.STYLE_GROUP_NO"
            ASCMAIN1.sql = sqlPOTWPDMX
            Create_TDA(.Tables.Add, "POTWPDMX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "POTWPDM1", "*")
            'With .Tables("POTWPDM1")
            '    .Columns.Add("LOGO", GetType(System.Byte()))
            '    '.PrimaryKey = New DataColumn() {.Columns("STYLE_GROUP_NO")}
            'End With

            ASCMAIN1.sql = "Select POTWPDM2.*, ICTPLIN2.STYLE_DESC" _
                & " from POTWPDM2, ICTPLIN2 where ICTPLIN2.STYLE_CODE_PLM = POTWPDM2.STYLE_CODE_PLM" _
                & " and POTWPDM2.STYLE_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTWPDM2", "**", 0, True, "V", 2)
            'With .Tables("POTWPDM2")
            '    .Columns.Add("IMAGE", GetType(System.Byte()))
            '    ' .Columns("SEQ").DataType = GetType(System.Int32)
            'End With

            ASCMAIN1.sql = "Select POTWPDM3.*, ICTCOLR1.COLOR_DESC" _
                & " from POTWPDM3, ICTCOLR1 where ICTCOLR1.COLOR_CODE = POTWPDM3.COLOR_CODE" _
                & " and POTWPDM3.STYLE_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTWPDM3", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select POTWPDM4.*" _
                & " from POTWPDM4" _
                & " where POTWPDM4.STYLE_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTWPDM4", "**", 0, True, "V", 3)

        End With

        grdPOTWPDM2.DataSource = dst.Tables("POTWPDM2")
        grdPOTWPDM3.DataSource = dst.Tables("POTWPDM3")
        grdPOTWPDM4.DataSource = dst.Tables("POTWPDM4")
        grdPOTWPDMX.DataSource = dst.Tables("POTWPDMX")

        Create_Summary(grdPOTWPDMX, "STYLE_GROUP_NO", "Count")

        '   Create_Summary(grdPOTWPDM2, "SEQ", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("SEASON_CODE", , True)

            Case "View", "Edit"
                If Absx1.txtFor("STYLE_GROUP_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a valid Quote No"
                Else
                    STYLE_GROUP_NO = Absx1.txtFor("STYLE_GROUP_NO").Text
                    rowPOTWPDM1 = LookUp("POTWPDM1", STYLE_GROUP_NO)
                    If rowPOTWPDM1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Style Group No " & STYLE_GROUP_NO
                    End If
                End If

            Case "Update"
                If dst.Tables("POTWPDM2").Select("").Length = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowPOTWPDM2 As DataRow In dst.Tables("POTWPDM2").Select("", "", DataViewRowState.CurrentRows)
                    Next
                End If

                If EMsg = "" Then

                End If

            Case "Delete"

                If ASCMAIN1.USER_ID <> rowPOTWPDM1.Item("INIT_OPER") & "" Then
                    EMsg &= vbCr & "Only " & rowPOTWPDM1.Item("INIT_OPER") & " may Delete this Quote"
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
                If dst.Tables("POTWPDM2").Select("").Length = 0 Then
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

                '  Update_Record_TDA("POTWPDM2", "1=1")
                Synch_TABLE_NAME("POTWPDM1")

                Dim FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG"
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    rowPOTWPDM1.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
                End If


                Print_Report_Begin()


                Dim RPT As String = "ICRQUOT1"


                If eItemKey = "email" Then
                    Dim tempFileName As String = rowPOTWPDM1.Item("STYLE_GROUP_NO")
                    Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)
                    ' Dim FILENAME As String = REPORT_FILENAMES(REPORT_NO)
                    Print_Report_End(, True)
                    email_Quote(tempFileName)
                Else
                    Generate_Report(RPT, "Quote Sheet")
                    Print_Report_End()
                End If

                'Case "Clear Quote Sheet"
                '    dst.Tables("POTWPDM2").Rows.Clear()
                '    Setup_Style_Quoted()
                '    txtQUOTE_DESC.Text = ""
                '    Absx1.txtFor("CUST_CODE").Text = ""

                'Case "Save Quote Sheet"
                '    Update_Record_TDA("POTWPDM1")
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

                .Groups("Show Style Groups").Visible = Not ScreenMode
            End With
        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)
        'Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CODE"), False)
        ' SplitContainer1.Visible = ScreenMode

        'grdPOTWPDMX.Visible = Not ScreenMode
        tabPOTWPDMX.Visible = Not ScreenMode

        If ScreenMode Then
            Set_Read_Only(grpHeader, (EntryMode = "V"))

            If EntryMode = "V" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTWPDM2, grdPOTWPDM3, grdPOTWPDM4}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTWPDM2, grdPOTWPDM3, grdPOTWPDM4}
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
        For Each TABLE_NAME As String In New String() {"POTWPDMX", "POTWPDM1", "POTWPDM2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        grdPOTWPDM2.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Absx1.txtFor("STYLE_GROUP_NAME").Text = ""
        Absx1.dteFor("SHIP_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("STYLE_GROUP_NO").Text = ""

        STYLE_GROUP_NO = ""
        images.Clear()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowPOTWPDM1 = dst.Tables("POTWPDM1").NewRow
            STYLE_GROUP_NO = ASCMAIN1.Next_Control_No("POTWPDM1.STYLE_GROUP_NO")
            rowPOTWPDM1.Item("STYLE_GROUP_NO") = STYLE_GROUP_NO
            rowPOTWPDM1.Item("STYLE_GROUP_NAME") = HFs("STYLE_GROUP_NAME")
            'rowPOTWPDM1.Item("SEASON_CODE") = HFs("SEASON_CODE")
            'rowPOTWPDM1.Item("SHIP_DATE") = HFs("SHIP_DATE")
            ' rowPOTWPDM1.Item("TOTAL_QTY") = Val(HFs("TOTAL_QTY"))
            ' rowPOTWPDM1.Item("CUST_CODE") = HFs("CUST_CODE")
            ' rowPOTWPDM1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowPOTWPDM1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTWPDM1.Item("INIT_DATE") = DATETIME_STAMP
            rowPOTWPDM1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTWPDM1.Item("LAST_DATE") = DATETIME_STAMP
            dst.Tables("POTWPDM1").Rows.Add(rowPOTWPDM1)
        Else
            rowPOTWPDM1 = Fill_Record("POTWPDM1", STYLE_GROUP_NO)
            dst.AcceptChanges()
        End If

        images.Clear()

        Fill_Records("POTWPDM4", STYLE_GROUP_NO)

        Fill_Records("POTWPDM2", STYLE_GROUP_NO)
        Sort_grdColumns(grdPOTWPDM2, "STYLE_CODE_PLM")

        Fill_Records("POTWPDM3", STYLE_GROUP_NO)
        Sort_grdColumns(grdPOTWPDM3, "COLOR_CODE")

        Dim b As System.Drawing.Bitmap = ASCMAIN1.Get_Image(images_folder & "\Groups\", "000006.png", False, , , Nothing)
        picCover.Image = b

        UltraGanttView1.CalendarInfo = UltraCalendarInfo1
        UltraCalendarInfo1.Projects.Add("Product Development", Now.AddDays(10))

        UltraCalendarInfo1.Tasks.Add(Now.AddDays(3), TimeSpan.FromDays(3), "First Task")
        UltraCalendarInfo1.Tasks(0).Tasks.Add(Now.AddDays(3), TimeSpan.FromDays(3), "First Task")
        UltraCalendarInfo1.Tasks(0).Tasks.Add(Now.AddDays(3), TimeSpan.FromDays(3), "First Task")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("POTWPDM1")
        Update_Record_TDA("POTWPDM2")
        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Delete_Records("POTWPDM1")
        Delete_Records("POTWPDM2")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where STYLE_GROUP_NO = '" & Absx1.txtFor("STYLE_GROUP_NO").Text & "'")
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

                Absx1.txtFor("STYLE_GROUP_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTWPDM1"
            E.COLUMN_NAME = "STYLE_GROUP_NO"
            E.CODE_VALUE = Absx1.txtFor("STYLE_GROUP_NO").Text
            E.DESC_VALUE = "Style Group No"
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
        Load_Popup_Menu(grdPOTWPDMX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdPOTWPDM2, "SBBBB", "Show Filter", "Get Styles", "Product Line Maintenance", "Product Line Inquiry", "Sequence as Shown")
        Load_Popup_Menu(grdPOTWPDM3, "B", "Get Colors")
        Load_Popup_Menu(grdPOTWPDM4, "B", "Get Specifications")
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

                Case "grdPOTWPDM2"
                    tlb_btn = DirectCast(tlb_pop.Tools("Sequence as Shown"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                    tlb_btn = DirectCast(tlb_pop.Tools("Get Styles"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                Case "grdPOTWPDM3"
                    tlb_btn = DirectCast(tlb_pop.Tools("Get Colors"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                Case "grdPOTWPDM4"
                    tlb_btn = DirectCast(tlb_pop.Tools("Get Specifications"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Get Styles"
                Get_Styles()

            Case "Get Colors"
                Get_Colors()

            Case "Get Specifications"
                Get_Specifications()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)


            Case "Sequence as Shown"
                Dim SEQ As Integer = 0
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTWPDM2.Rows
                    SEQ += 10
                    grow.Cells("SEQ").Value = SEQ
                    grow.Update()
                Next

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
            Case "STYLE_GROUP_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View", e)
                End If

        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If Not InquiryMode And Not ScreenMode Then
                    Click_Command("New")
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

#Region "grdPOTWPDM2"

    Private Sub grdPOTWPDM2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM2.AfterCellUpdate
        If Not e.Cell.Row.IsDataRow Then Exit Sub
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_PLM"

                grdCodeDesc(grdPOTWPDM2, "ICTPLIN2", "STYLE_CODE_PLM", "STYLE_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE STYLE_DESC
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE_PLM As String = e.Cell.Value
                    e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")
                    'e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = cdr.Item("SALES_DIVISION_CODE")
                    'e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = cdr.Item("STYLE_CLASS_CODE") & ""
                    'e.Cell.Row.Cells("STYLE_PRICE").Value = cdr.Item("STYLE_PRICE")

                Else
                    grdPOTWPDM2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If
        End Select
    End Sub

    Private Sub grdPOTWPDM2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTWPDM2.AfterRowActivate

        If Not grdPOTWPDM2.ActiveRow.IsDataRow Then Exit Sub

        With grdPOTWPDM2.DisplayLayout.Bands(0)
            If grdPOTWPDM2.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE_PLM").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdPOTWPDM2.ActiveCell = grdPOTWPDM2.ActiveRow.Cells("STYLE_CODE_PLM")
                grdPOTWPDM2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE_PLM").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        Setup_grdPOTWPDM2()
    End Sub

    Sub Setup_grdPOTWPDM2()

        tplStyle.Tiles.Clear()

        If grdPOTWPDM2.ActiveRow Is Nothing OrElse grdPOTWPDM2.ActiveRow.IsAddRow Or Not grdPOTWPDM2.ActiveRow.IsDataRow Then
            grdPOTWPDM4.Visible = False
        Else
            grdPOTWPDM4.Visible = True
            Dim STYLE_CODE_PLM As String = grdPOTWPDM2.ActiveRow.Cells("STYLE_CODE_PLM").Value & ""
            Dim dvw As DataView = DirectCast(grdPOTWPDM4.DataSource, DataTable).DefaultView
            dvw.RowFilter = "STYLE_GROUP_NO = '" & STYLE_GROUP_NO & "' and STYLE_CODE_PLM = '" & STYLE_CODE_PLM & "'"
            Sort_grdColumns(grdPOTWPDM4, "SEQ")

            If images.ContainsKey(STYLE_CODE_PLM) Then
                For Each I As System.Drawing.Bitmap In images(STYLE_CODE_PLM)
                    Dim t As New Infragistics.Win.Misc.UltraTile
                    Dim P As New UltraWinEditors.UltraPictureBox
                    P.Image = I
                    '   pic.Image = I
                    t.Control = P
                    t.Text = "MY TILE"
                    tplStyle.Tiles.Add(t)
                    P.Visible = True
                    t.Visible = True
                Next
                tplStyle.Visible = True
            End If
        End If
    End Sub
    Private Sub grdPOTWPDM2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTWPDM2.AfterRowUpdate

        Dim STYLE_GROUP_NO As String = e.Row.Cells("STYLE_GROUP_NO").Value
        Dim STYLE_CODE_PLM As String = e.Row.Cells("STYLE_CODE_PLM").Value
        Dim rowPOTWPDM2 As DataRow = dst.Tables("POTWPDM2").Rows.Find(New Object() {STYLE_GROUP_NO, STYLE_CODE_PLM})
        If Not images.ContainsKey(STYLE_CODE_PLM) Then
            Get_Images(STYLE_CODE_PLM)
        End If
    End Sub

    Private Sub grdPOTWPDM2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTWPDM2.BeforeExitEditMode
        If grdPOTWPDM2.ActiveCell Is Nothing Then Exit Sub
        If Not grdPOTWPDM2.ActiveRow.IsDataRow Then Exit Sub
        With grdPOTWPDM2.ActiveCell
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

    Private Sub grdPOTWPDM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDM2.BeforeRowUpdate
        With grdPOTWPDM2
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
                If e.Row.Cells("STYLE_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("STYLE_GROUP_NO").Value = Absx1.CtlFor("STYLE_GROUP_NO").Text
                    .ActiveRow.Cells("SEQ").Value = Val(dst.Tables("POTWPDM2").Compute("Max(SEQ)", "") & "") + 10
                End If
            End If
        End With


    End Sub

    Private Sub grdPOTWPDM2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM2.ClickCellButton

        If grdPOTWPDM2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_PLM"
        End Select
        grdClickCellButton(grdPOTWPDM2, sql_where, sql_where <> "")

    End Sub

#End Region


#Region "grdPOTWPDM3"
    Private Sub grdPOTWPDM3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "COLOR_CODE"
                Dim COLOR_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdPOTWPDM3, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
        End Select
    End Sub

    Private Sub grdPOTWPDM3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDM3.BeforeRowUpdate
        With grdPOTWPDM3
            If Not e.Cancel Then
                If e.Row.Cells("STYLE_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("STYLE_GROUP_NO").Value = Absx1.CtlFor("STYLE_GROUP_NO").Text
                End If
            End If
        End With
    End Sub

    Private Sub grdPOTWPDM3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM3.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdPOTWPDM3, sql_where, sql_where <> "")
    End Sub
#End Region

#Region "grdPOTWPDM4"
    Private Sub grdPOTWPDM4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM4.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "SPEC_CODE"
                Dim SPEC_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdPOTWPDM4, "POTWPDMS", "SPEC_CODE", "SPEC_DESC")
        End Select
    End Sub

    Private Sub grdPOTWPDM4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDM4.BeforeRowUpdate
        With grdPOTWPDM4
            If Not e.Cancel Then
                If e.Row.Cells("STYLE_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("STYLE_GROUP_NO").Value = Absx1.CtlFor("STYLE_GROUP_NO").Text
                    .ActiveRow.Cells("STYLE_CODE_PLM").Value = grdPOTWPDM2.ActiveRow.Cells("STYLE_CODE_PLM").Value
                    Dim sqlx As String = "STYLE_GROUP_NO = '" & .ActiveRow.Cells("STYLE_GROUP_NO").Value & "' and STYLE_CODE_PLM = '" & .ActiveRow.Cells("STYLE_CODE_PLM").Value & "'"
                    .ActiveRow.Cells("SPEC_LNO").Value = Val(dst.Tables("POTWPDM4").Compute("MAX(SPEC_LNO)", sqlx) & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdPOTWPDM4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWPDM4.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdPOTWPDM4, sql_where, sql_where <> "")
    End Sub
#End Region

    Private Sub grdPOTWPDMX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTWPDMX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("STYLE_GROUP_NO").Text = e.Row.Cells("STYLE_GROUP_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        ASCMAIN1.sql = sqlPOTWPDMX
        Dim STYLE_GROUP_NO As String = Absx1.txtFor("STYLE_GROUP_NO").Text
        If optShow.Value = "A" And STYLE_GROUP_NO = "" Then
            grdPOTWPDMX.Text = "All Quotes"
        ElseIf optShow.Value = "M" Then
            ASCMAIN1.sql &= " and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
            grdPOTWPDMX.Text = "Quotes entered or modified by Me"
        ElseIf optShow.Value = "C" Or STYLE_GROUP_NO <> "" Then
            ASCMAIN1.sql &= " and STYLE_GROUP_NO = '" & STYLE_GROUP_NO & "'"
            grdPOTWPDMX.Text = "Quotes associated with " & STYLE_GROUP_NO
        End If
        Fill_Records("POTWPDMX")
        Sort_grdColumns(grdPOTWPDMX, "STYLE_GROUP_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
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
                SUBJECT, "POTWPDM1", False, True, CUST_CODE, CUST_NAME, "Customer")
        If SEND_NO <> "" Then
            TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "QUOEML", "Quote Sheet emailed", SEND_NO)
        End If
    End Sub

    Private Sub optShow_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShow.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Refresh_Documents()
    End Sub


    Sub Get_Styles()
        Add_Codes(grdPOTWPDM2, "ICTPLIN2", "STYLE_CODE_PLM", "Styles")
    End Sub

    Sub Get_Colors()
        Add_Codes(grdPOTWPDM3, "ICTCOLR1", "COLOR_CODE", "Colors")
    End Sub

    Sub Get_Specifications()
        Add_Codes(grdPOTWPDM4, "POTWPDMS", "SPEC_CODE", "Specifications")
    End Sub

    Sub Get_Images(STYLE_CODE_PLM As String)
        Dim I As New List(Of System.Drawing.Bitmap)

        If My.Computer.FileSystem.DirectoryExists(images_folder & "\Styles\" & STYLE_CODE_PLM) Then
            For Each file As String In My.Computer.FileSystem.GetFiles(images_folder & "\Styles\" & STYLE_CODE_PLM)
                Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(file)

                Dim imgba() As Byte = Nothing
                Dim b As System.Drawing.Bitmap = ASCMAIN1.Get_Image(images_folder & "\Styles\" & STYLE_CODE_PLM, fi.Name, True, , , imgba)
                I.Add(b)
                ' we will need imgba when we go to print
            Next
        End If

        If images.ContainsKey(STYLE_CODE_PLM) Then
            images(STYLE_CODE_PLM) = I
        Else
            images.Add(STYLE_CODE_PLM, I)
        End If
    End Sub
End Class