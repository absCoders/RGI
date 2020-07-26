Public Class SOFWBIMG

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFWBIMI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst
            'Dim SQLB As New System.Text.StringBuilder
            'SQLB.Length = 0
            'SQLB.AppendLine("")
            'ASCMAIN1.sql = SQLB.ToString
            'Create_TDA(.Tables.Add, "SOTPRIC1", "**", 0, False)
            'With .Tables("SOTPRIC1").Columns
            '    .Add("PB1", GetType(Double))
            '    .Add("PB2", GetType(Double))
            '    .Add("PB3", GetType(Double))
            '    .Add("PB4", GetType(Double))
            'End With
        End With

        'grdSOFPRIC1.DataSource = dst.Tables("SOTPRIC1")

        'ASCMAIN1.Add_Value_List(grdSOFCSTMX, "REPORT_TYPE", , New String() {":", "I:Initial", "A:Amended", "S:Subsequent", "R:Revised"})

        'Create_Summary(grdSOFPRIC1, "ORDRED_TY", "Sum")
        'Sort_grdColumns(grdSOFPRIC1, "STYLE_CODE", True)

        'With grdSOFPRIC1.DisplayLayout.Bands(0)
        '    For Each COL_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC"}
        '        .Columns(COL_NAME).Header.Fixed = True
        '    Next
        'End With

        TABLE_NAME = "SOFPRIC1"

        EntryMode = "E"
        Call Load_Record()
        Call Mode_Settings(True)
    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Sub Check_Form_Options()
        'With UltraExplorerBar1.Groups("Screen Control")
        '    .Items("New").Visible = (Me.Name = "PMFVIST1")
        'End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
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
            Case "Done"
                Call Mode_Settings(False)
                Me.Close()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'With grdSOFPRIC1.DisplayLayout.Bands(0)
        '    For Each thisCOL As String In New String() {"ORDRED_TY", "SHIPPED_TY", "CANCELLED_TY"}
        '        .Columns.Item(thisCOL).Header.Appearance.BackColor = Drawing.Color.Khaki
        '    Next
        '    For Each thisCOL As String In New String() {"ORDRED_LY", "SHIPPED_LY", "CANCELLED_LY"}
        '        .Columns.Item(thisCOL).Header.Appearance.BackColor = Drawing.Color.Bisque
        '    Next
        'End With


        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        'dst.EnforceConstraints = False
        'dst.Tables("PMTVIST1").Rows.Clear()
        'dst.Tables("PMTVISTH").Rows.Clear()

        'Dim dvw As DataView = DirectCast(grdPMTVIST1.DataSource, DataTable).DefaultView
        'dvw.RowStateFilter = DataViewRowState.CurrentRows

        'Fill_Records("PMTVIST1")
        'Process_SVRs()

        'Sort_grdColumns(grdPMTVIST1, "DATE_VISITED".ToLower)
        'Sort_grdColumns(grdPMTVISTH, "DATE_VISITED".ToLower)
        'dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        'Setup_Summary()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()
        'BeginTrans()
        'INIT_LAST("PMTVIST1", True, "", True)
        'Update_Record_TDA("PMTVIST1")
        'CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        'Call Load_Popup_Menu(grdSOFPRIC1, "SS", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'Dim tlb_btn As UltraWinToolbars.ButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""

                    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'Select Case e.Tool.Key
        '    Case "Project Center"
        '        Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
        '        Context_Launch("Edit", Column_Values("JOB_NO", JOB_NO), e.Tool.Key, "PMFJOBM1")
        '    Case "Show Report"
        '        Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\randfromdrc\RandInvoices\310 West 52nd Street - 30760.pdf"
        '        Show_Document(FILENAME)

        'End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        'Select Case Absx1.GetABSColumnName(sender)
        'Case "EMPLOYEE_CODE"
        '    If e.KeyCode = Windows.Forms.Keys.Enter Then
        '        Setup_Summary()
        '    End If
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        'Select Case COLUMN_NAME
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

#End Region

    Sub Setup_Summary()
        Dim sqlwhere As String = ""
        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        dst.Tables("SOTPRIC1").Rows.Clear()

        Fill_Records("SOTPRIC1")
        'Fill_Pricing()

        dst.EnforceConstraints = False

        ASCMAIN1.Progress("")
        'grdSOFPRIC1.Update()
        'grdSOFPRIC1.Refresh()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs)
        'Setup_Summary()
    End Sub

End Class