Imports System.Drawing

Public Class SOFCGTIN

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFCGTII" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst
            Dim SQLB As New System.Text.StringBuilder
            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("STY.CUST_CODE,")
            SQLB.AppendLine("UPC.STYLE_CODE,")
            SQLB.AppendLine("STY.STYLE_DESC,")
            SQLB.AppendLine("UPC.COLOR_CODE,")
            SQLB.AppendLine("UPC.SIZE_CODE,")
            SQLB.AppendLine("UPC.UPC_CODE,")
            SQLB.AppendLine("GTN.GTIN_PACK_CODE,")
            SQLB.AppendLine("GTN.GTIN_CODE,")
            SQLB.AppendLine("GTN.GTIN_DESC,")
            SQLB.AppendLine("CSY.CUST_STYLE_CODE")
            SQLB.AppendLine("FROM ICTGTINT GTN, ICVLUPC1 UPC, ICTSTYL1 STY, SOTCSTY1 CSY")
            SQLB.AppendLine("WHERE GTN.GTIN_UPC_CODE = UPC.UPC_CODE")
            SQLB.AppendLine("AND UPC.STYLE_CODE = STY.STYLE_CODE")
            SQLB.AppendLine("AND CSY.CUST_CODE = :PARM1")
            SQLB.AppendLine("AND CSY.STYLE_CODE = UPC.STYLE_CODE (+)")
            SQLB.AppendLine("AND CSY.COLOR_CODE = UPC.COLOR_CODE (+)")
            SQLB.AppendLine("AND CSY.SIZE_DESC = UPC.SIZE_CODE (+)")
            SQLB.AppendLine("AND STY.CUST_CODE = :PARM1")
            'SQLB.AppendLine("SELECT")
            'SQLB.AppendLine("STY.CUST_CODE,")
            'SQLB.AppendLine("UPC.STYLE_CODE,")
            'SQLB.AppendLine("STY.STYLE_DESC,")
            'SQLB.AppendLine("UPC.COLOR_CODE,")
            'SQLB.AppendLine("UPC.SIZE_CODE,")
            'SQLB.AppendLine("UPC.UPC_CODE,")
            'SQLB.AppendLine("GTN.GTIN_PACK_CODE,")
            'SQLB.AppendLine("GTN.GTIN_CODE,")
            'SQLB.AppendLine("GTN.GTIN_DESC")
            'SQLB.AppendLine("FROM ICTGTINT GTN, ICVLUPC1 UPC, ICTSTYL1 STY")
            'SQLB.AppendLine("WHERE GTN.GTIN_UPC_CODE = UPC.UPC_CODE")
            'SQLB.AppendLine("AND UPC.STYLE_CODE = STY.STYLE_CODE")
            'SQLB.AppendLine("AND STY.CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOFCGTIN", "**", 0, False, "V")
        End With

        grdSOFCGTIN.DataSource = dst.Tables("SOFCGTIN")

        Sort_grdColumns(grdSOFCGTIN, "STYLE_CODE", False)

        'grdSOTCANG1.DisplayLayout.Bands(0).Columns("ORDR_QTY_SHIP").Format = "###,##0"

        'Create_Summary(grdSOTCANG1, "ORDR_QTY")
        'Create_Summary(grdSOTCANG1, "ORDR_QTY_CANC")

        TABLE_NAME = "SOTCGTIN"

        EntryMode = "E"
        'Call Load_Record()
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
            '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            '.Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

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

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        dst.Tables("SOFCGTIN").Rows.Clear()

        dst.EnforceConstraints = False

        Fill_Records("SOFCGTIN", Absx1.txtFor("CUST_CODE").Text)
        'dst.EnforceConstraints = True

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

    Sub Setup_Summary()
        grdSOFCGTIN.Update()
        grdSOFCGTIN.Refresh()
        Me.Cursor = Cursors.Default

    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
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
        Call Load_Popup_Menu(grdSOFCGTIN, "SS", "Show Filter", "Show GroupBox")
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
    End Sub

#End Region

#Region "ABSColumn Controls"
    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

#End Region

#Region "Custom Methods"

    Private Sub btnFETCH_Click(sender As Object, e As EventArgs) Handles btnFETCH.Click
        If Absx1.txtFor("CUST_CODE").Text = "" Then
            MsgBox("You Must Select A Customer First", vbExclamation, "Customer Selection")
        Else
            Load_Record()
        End If
    End Sub
#End Region

#Region "Form Controls"

#End Region


End Class