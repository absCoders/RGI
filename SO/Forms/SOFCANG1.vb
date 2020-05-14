Imports System.Drawing

Public Class SOFCANG1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFCANGI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst
            Dim SQLB As New System.Text.StringBuilder
            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("SOTORDR1.ORDR_CUST_PO,")
            SQLB.AppendLine("SOTORDR2.EDI_DTL_SEQ AS LNO,")
            'SQLB.AppendLine("SOTORDR1.CUST_STORE_NO,")
            SQLB.AppendLine("SOTORDR2.STYLE_CODE,")
            SQLB.AppendLine("SUM(SOTORDR2.ORDR_QTY) AS ORDR_QTY,")
            SQLB.AppendLine("SUM(SOTORDR2.ORDR_QTY_CANC) AS ORDR_QTY_CANC,")
            SQLB.AppendLine("SOTORDR2.CUST_SKU,")
            SQLB.AppendLine("SOTORDR2.CUST_COLOR_CODE,")
            SQLB.AppendLine("SOTORDR2.CUST_SIZE_CODE")
            SQLB.AppendLine("FROM SOTORDR1,SOTORDR2")
            SQLB.AppendLine("WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO")
            SQLB.AppendLine("AND SOTORDR2.ORDR_QTY_CANC > 0")
            SQLB.AppendLine("AND SOTORDR1.ORDR_STATUS <> 'D'")
            SQLB.AppendLine("AND SOTORDR1.ORDR_GROUP_NO = 'NONE'")
            SQLB.AppendLine("GROUP BY")
            SQLB.AppendLine("SOTORDR1.ORDR_CUST_PO,")
            SQLB.AppendLine("SOTORDR2.EDI_DTL_SEQ,")
            'SQLB.AppendLine("SOTORDR1.CUST_STORE_NO,")
            SQLB.AppendLine("SOTORDR2.STYLE_CODE,")
            SQLB.AppendLine("SOTORDR2.CUST_SKU,")
            SQLB.AppendLine("SOTORDR2.CUST_COLOR_CODE,")
            SQLB.AppendLine("SOTORDR2.CUST_SIZE_CODE")
            SQLB.AppendLine("ORDER BY")
            SQLB.AppendLine("SOTORDR1.ORDR_CUST_PO,")
            SQLB.AppendLine("SOTORDR2.EDI_DTL_SEQ")
            'SQLB.AppendLine("SOTORDR1.CUST_STORE_NO")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTCANG1", "**", 0, False)

            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("SOTORDR1.ORDR_GROUP_NO")
            SQLB.AppendLine("FROM SOTORDR1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTGROUP", "**", 0, False)
        End With

        grdSOTCANG1.DataSource = dst.Tables("SOTCANG1")

        Sort_grdColumns(grdSOTCANG1, "ORDR_CUST_PO, LNO", False)

        'grdSOTCANG1.DisplayLayout.Bands(0).Columns("ORDR_QTY_SHIP").Format = "###,##0"

        'Create_Summary(grdSOTCANG1, "ORDR_QTY")
        'Create_Summary(grdSOTCANG1, "ORDR_QTY_CANC")

        grdGROUPS.DataSource = dst.Tables("SOTGROUP")

        TABLE_NAME = "SOTCANGI"

        EntryMode = "E"
        'Call Load_Record()
        Call Mode_Settings(True)

        SplitContainer2.SplitterDistance = 120

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
        dst.Tables("SOTROYLI").Rows.Clear()

        dst.EnforceConstraints = False

        Fill_Records("SOTROYLI")
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
        ReLoadData()
        grdSOTCANG1.Update()
        grdSOTCANG1.Refresh()
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
        'Call Load_Popup_Menu(grdSOTROYL1, "SSB", "Show Filter", "Show GroupBox")
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
    Private Sub ReLoadData()
        Throw New NotImplementedException
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Absx1.txtFor("CUST_CODE").Text = "" Then
            MsgBox("You Must Select A Customer First")
        Else
            dst.Tables("SOTGROUP").Clear()
            Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
            SQLS.AppendLine(" ORDR_GROUP_NO IN (SELECT ORDR_GROUP_NO FROM SOTORDR1 WHERE ORDR_STATUS <> 'D' AND CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "')")
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ORDR_GROUP_NO", , SQLS.ToString)
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.SelectedCodes.Count > 0 Then
                    For Each SelCode As String In ASCMAIN1.CodeSelector.SelectedCodes
                        Dim newRec As DataRow = dst.Tables("SOTGROUP").NewRow
                        newRec.Item("ORDR_GROUP_NO") = SelCode
                        dst.Tables("SOTGROUP").Rows.Add(newRec)
                    Next
                End If
                F.Dispose()
            End If
        End If
    End Sub

    Private Sub btnFETCH_Click(sender As Object, e As EventArgs) Handles btnFETCH.Click
        Dim lstGroups As String = ""

        Dim lstGroups_min As String = ""
        Dim lstGroups_max As String = ""

        For Each rowSOTGROUP As DataRow In dst.Tables("SOTGROUP").Select()
            lstGroups = lstGroups & "'" & rowSOTGROUP.Item("ORDR_GROUP_NO").ToString() & "',"
            If (rowSOTGROUP.Item("ORDR_GROUP_NO").ToString() < lstGroups_min) Or lstGroups_min = "" Then
                lstGroups_min = rowSOTGROUP.Item("ORDR_GROUP_NO").ToString()
            End If
            If rowSOTGROUP.Item("ORDR_GROUP_NO").ToString() > lstGroups_max Or lstGroups_max = "" Then
                lstGroups_max = rowSOTGROUP.Item("ORDR_GROUP_NO").ToString()
            End If
        Next
        lstGroups = lstGroups.Substring(0, lstGroups.Length - 1)

        Dim SQLB As New System.Text.StringBuilder
        SQLB.Length = 0
        SQLB.AppendLine("SELECT")
        SQLB.AppendLine("SOTORDR1.ORDR_CUST_PO,")
        SQLB.AppendLine("SOTORDR2.EDI_DTL_SEQ AS LNO,")
        'SQLB.AppendLine("SOTORDR1.CUST_STORE_NO,")
        SQLB.AppendLine("SOTORDR2.STYLE_CODE,")
        SQLB.AppendLine("SUM(SOTORDR2.ORDR_QTY) AS ORDR_QTY,")
        SQLB.AppendLine("SUM(SOTORDR2.ORDR_QTY_CANC) AS ORDR_QTY_CANC,")
        SQLB.AppendLine("SOTORDR2.CUST_SKU,")
        SQLB.AppendLine("SOTORDR2.CUST_COLOR_CODE,")
        SQLB.AppendLine("SOTORDR2.CUST_SIZE_CODE")
        SQLB.AppendLine("FROM SOTORDR1,SOTORDR2")
        SQLB.AppendLine("WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO")
        SQLB.AppendLine("AND SOTORDR2.ORDR_QTY_CANC > 0")
        SQLB.AppendLine("AND SOTORDR1.ORDR_STATUS <> 'D'")
        SQLB.AppendLine("AND SOTORDR1.ORDR_GROUP_NO IN (" & lstGroups & ")")
        SQLB.AppendLine("GROUP BY")
        SQLB.AppendLine("SOTORDR1.ORDR_CUST_PO,")
        SQLB.AppendLine("SOTORDR2.EDI_DTL_SEQ,")
        'SQLB.AppendLine("SOTORDR1.CUST_STORE_NO,")
        SQLB.AppendLine("SOTORDR2.STYLE_CODE,")
        SQLB.AppendLine("SOTORDR2.CUST_SKU,")
        SQLB.AppendLine("SOTORDR2.CUST_COLOR_CODE,")
        SQLB.AppendLine("SOTORDR2.CUST_SIZE_CODE")
        SQLB.AppendLine("ORDER BY")
        SQLB.AppendLine("SOTORDR1.ORDR_CUST_PO,")
        SQLB.AppendLine("SOTORDR2.EDI_DTL_SEQ")
        'SQLB.AppendLine("SOTORDR1.CUST_STORE_NO")
        'ASCMAIN1.sql = SQLB.ToString

        Fill_Records("SOTCANG1",,, SQLB.ToString)


        grdSOTCANG1.Text = String.Format("From Group {0} To Group {1}", lstGroups_min, lstGroups_max)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim BegGroup As String = ""
        Dim EndGroup As String = ""
        Dim BegGroupVal As Int64 = 0
        Dim EndGroupVal As Int64 = 0
        If Absx1.txtFor("CUST_CODE").Text = "" Then
            MsgBox("You Must Select A Customer First")
        Else
            dst.Tables("SOTGROUP").Clear()
            Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
            SQLS.AppendLine(" ORDR_GROUP_NO IN (SELECT ORDR_GROUP_NO FROM SOTORDR1 WHERE ORDR_STATUS <> 'D' AND CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "')")
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ORDR_GROUP_NO", , SQLS.ToString)
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                lblGather1.Visible = True
                lblGather2.Visible = True

                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                ASCMAIN1.CodeSelector.Caption = "Select Beginning Group #"
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.SelectedCodes.Count = 1 Then
                    BegGroup = ASCMAIN1.CodeSelector.SelectedCode
                Else
                    MsgBox("You Can Only Select One Group When Using the Begin/End Feature", vbOKOnly, "Begin/End")
                    F.Dispose()
                    Exit Sub
                End If
                F.Dispose()
                lblGather1.Visible = False
                lblGather2.Visible = False
            End If
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                lblGather1.Visible = True
                lblGather2.Visible = True
                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                ASCMAIN1.CodeSelector.Caption = "Select Ending Group #"
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.SelectedCodes.Count = 1 Then
                    EndGroup = ASCMAIN1.CodeSelector.SelectedCode
                Else
                    MsgBox("You Can Only Select One Group When Using the Begin/End Feature", vbOKOnly, "Begin/End")
                    F.Dispose()
                    Exit Sub
                End If
                F.Dispose()
                lblGather1.Visible = False
                lblGather2.Visible = False
            End If
            BegGroupVal = Val(BegGroup)
            EndGroupVal = Val(EndGroup)
            Dim GroupCount As Int64 = 0
            For i As Int64 = BegGroupVal To EndGroupVal
                GroupCount += 1
                Dim newRec As DataRow = dst.Tables("SOTGROUP").NewRow
                newRec.Item("ORDR_GROUP_NO") = Format(i, "0000000000")
                dst.Tables("SOTGROUP").Rows.Add(newRec)
            Next
            MsgBox(GroupCount & " Groups Selected", vbOKOnly, "Selected")
        End If
    End Sub


#End Region

#Region "Form Controls"
    Private Sub btnAddGroup_Click(sender As Object, e As EventArgs) Handles btnAddGroup.Click
        If txtAddGroup.Text.Length = 0 Then
            MsgBox("You Must Provide A Valid Group Number", vbOKOnly, "Invalid Group Specified")
        Else
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT COUNT(*)")
            SQLS.AppendLine("FROM SOTORDR0")
            SQLS.AppendLine(String.Format("WHERE ORDR_GROUP_NO = '{0}'", txtAddGroup.Text & String.Empty))
            SQLS.AppendLine(String.Format("AND CUST_CODE = '{0}'", Absx1.txtFor("CUST_CODE").Text & String.Empty))
            ASCMAIN1.sql = SQLS.ToString()
            Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
            If REC_CNT <> 1 Then
                MsgBox("You Must Provide A Valid Group Number", vbOKOnly, "Invalid Group Specified")
            Else
                Dim Filter As String = String.Format("ORDR_GROUP_NO = '{0}'", txtAddGroup.Text & String.Empty)
                If dst.Tables.Item("SOTGROUP").Select(Filter).Count <> 0 Then
                    MsgBox("That Group Is Already Selected", vbOKOnly, "Invalid Group Specified")
                Else
                    Dim newRec As DataRow = dst.Tables("SOTGROUP").NewRow
                    newRec.Item("ORDR_GROUP_NO") = txtAddGroup.Text
                    dst.Tables("SOTGROUP").Rows.Add(newRec)
                    txtAddGroup.Text = ""
                End If
            End If
        End If
    End Sub
#End Region


End Class