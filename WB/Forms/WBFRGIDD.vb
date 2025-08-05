
Public Class WBFRGIDD
    Dim FormLoading As Boolean = True
    Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}

    Dim RYP0 As String
    Dim RYP1 As String
    'Dim RYW0 As String
    'Dim RYW1 As String
    Dim Periods As Integer

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim defPrd As Int64 = (Val(ASCMAIN1.CYP.Substring(4, 2)) - 1) * -1 '-3
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -48, 0, defPrd)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -48, 0, 0)

        With dst

            'SQLs.Length = 0
            'SQLs.AppendLine("SELECT * ")
            'Create_TDA(.Tables.Add, "WBFFSR01", "**", 0, False)

        End With

        'Fill_Records("ECTECOM1_FILTER")


        'For Each COL As String In valueColsTY
        '    Create_Summary(grdWBFFSR01, COL, "Sum", "", "###,##0")
        '    With grdWBFFSR01.DisplayLayout.Bands(0)
        '        .Columns(COL).Format = "###,##0"
        '        .Columns(COL).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
        '        .Columns(COL).Header.Appearance.BackColor = Drawing.Color.White
        '        .Columns(COL).CellAppearance.BackColor = Drawing.Color.LightBlue
        '        .Columns(COL).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
        '    End With
        'Next

        'For Each COL As String In valueColsLY
        '    Create_Summary(grdWBFFSR01, COL, "Sum", "", "###,##0")
        '    With grdWBFFSR01.DisplayLayout.Bands(0)
        '        .Columns(COL).Format = "###,##0"
        '        .Columns(COL).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
        '        .Columns(COL).Header.Appearance.BackColor = Drawing.Color.White
        '        .Columns(COL).CellAppearance.BackColor = Drawing.Color.LightGreen
        '        .Columns(COL).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
        '    End With
        'Next

        'With grdWBFFSR01.DisplayLayout.Bands(0)
        '    .Columns("CODE").Header.Fixed = True
        '    .Columns("CODE_DESC").Header.Fixed = True
        'End With

        'With grdWBFFSR01.DisplayLayout.Override
        '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '    .AllowDelete = DefaultableBoolean.False
        '    .AllowUpdate = DefaultableBoolean.False
        'End With

        'For i As Integer = 0 To grdWBFFSR01.DisplayLayout.Bands(0).Columns.Count - 1
        '    grdWBFFSR01.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        'Next i

        'Load_Record(False)

        'Sort_grdColumns(grdWBFFSR01, "SALES".ToLower(), False)

        tab.Visible = False

        FormLoading = False
        'grdWBFHORNT.Parent = tab.Parent

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"

            Case "Exit"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                'Load_Record(True)
                'Dim clsWBCRGIDD As New WBCRGIDD(Me)
                'If clsWBCRGIDD.eMsg.Length = 0 Then
                '    clsWBCRGIDD.makeExcel()
                'Else
                '    MsgBox(clsWBCRGIDD.eMsg, vbCritical, "Excel Creation Cancelled")
                'End If
                'Me.Cursor = Cursors.Default
            Case "Exit"
                Call Mode_Settings(False)
                Me.Close()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Visible = False
                .Groups("Screen Control").Items("Exit").Visible = Not ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        dst.Tables("WBFFSR01").Rows.Clear()
    End Sub

    Sub Load_Record(Optional showRefreshing As Boolean = False)
        Me.Cursor = Cursors.WaitCursor
        If showRefreshing Then
            ASCMAIN1.Progress("Refreshing Data", "")
            Me.Cursor = Cursors.WaitCursor
        End If
        Application.DoEvents()
        'Call Save_Header_Fields(UltraGroupBox1)
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        EnforceConstraints(False)

        'Fill_Records("WBTHORNT", , , SQLs.ToString)
        'fillTempTable()
        'loadTempTable()
        'Stop
        'EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Me.Cursor = Cursors.Default
        If showRefreshing Then
            ASCMAIN1.Progress("")
            Me.Cursor = Cursors.Default
        End If
        Application.DoEvents()
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()

        'Call CommitTrans("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'TODO: Remove this stop before going live.
        Stop
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        Print_Report_Begin()
        'Generate_Report("WBRHORNT")
        Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdWBFFSR01, "SSB", "Show Filter", "Show GroupBox")
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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            Case "grdWBFHORNT"

        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Copy To Clipboard"
            '    Dim STYLE_CODE As String = grd.ActiveRow.Cells("RANK_CODE").Text
            '    Clipboard.SetText(STYLE_CODE)
            '    MsgBox($"{STYLE_CODE} Copied To Clipboard.", vbOKOnly, "Clipboard")
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "BANK_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Call Click_Command("New", e)
        '        End If
        '    Case "PYMT_BATCH_NO"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Call Click_Command("Edit", e)
        '        End If
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                'FillStyle()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Call Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub

    Private Sub btnCreateDump_Click(sender As Object, e As EventArgs) Handles btnCreateDump.Click
        Dim clsWBCRGIDD As New WBCRGIDD(Me, Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value, chkOnlyOrderDetails.Checked)
        If clsWBCRGIDD.eMsg.Length = 0 Then
            clsWBCRGIDD.makeExcel()
        Else
            MsgBox(clsWBCRGIDD.eMsg, vbCritical, "Excel Creation Cancelled")
        End If
        Me.Cursor = Cursors.Default
    End Sub
#End Region

#Region "Custom Methods"

#End Region

#Region "Form Controls"

#End Region
End Class