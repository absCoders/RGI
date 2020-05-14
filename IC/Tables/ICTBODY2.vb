Public Class ICTBODY2

    Dim sqlICTBODYS As String = ""
    Dim nodeIndex As Integer
    Private WithEvents UltraTree_DropHightLight_DrawFilter As New UltraTree_DropHightLight_DrawFilter_Class()

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ICTBODYS.*" & vbCrLf _
                & " from ICTBODYS" & vbCrLf _
                & " where ICTBODYS.SUB_BODY_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTBODYS", "**", 0, True, "V", 2)
        End With

        grdICTBODYS.DataSource = dst.Tables("ICTBODYS")

        With grdICTBODYS.DisplayLayout.Bands(0)
            .Columns("SPEC_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
        End With

        If ASCMAIN1.CLIENT = "RGI" Then
            grdICTBODYS.Visible = False
            tvwICTBODYS.Visible = False
        End If

        With tvwICTBODYS
            'If UltraOptionSet1.Value = "T" Then
            .Override.CellClickAction = UltraWinTree.CellClickAction.Default
            .ViewStyle = UltraWinTree.ViewStyle.Standard
            .AllowDrop = True
            .Override.AllowCut = DefaultableBoolean.True
            .Override.AllowCopy = DefaultableBoolean.True
            .Override.AllowPaste = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.True
            'Else
            '.Override.CellClickAction = UltraWinTree.CellClickAction.EditCellSelectText '  .EditCell
            '.ViewStyle = UltraWinTree.ViewStyle.OutlookExpress
            '    .AllowDrop = False
            '    .Override.AllowCut = DefaultableBoolean.False
            '    .Override.AllowCopy = DefaultableBoolean.False
            '    .Override.AllowPaste = DefaultableBoolean.False
            '    .Override.AllowDelete = DefaultableBoolean.False
            'End If
        End With

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTBODYS, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(tvwICTBODYS, "BBB", "Insert Above", "Insert Below", "Insert Within")

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

        Select Case grd.Name

            Case "grdICTBODYS"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New") And (1 <> 1)
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            'Case "Add Codes"
            '    If grd.Name = "grdICTBODYS" Then
            '        Add_Codes(grdICTBODYS, "ICTBODY2", "SUB_BODY_CODE", "Sub-Body Codes")
            '    End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "Overrides"
    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim SUB_BODY_CODE As String = Absx1.txtFor("SUB_BODY_CODE").Text
        Dim sqlDelete = "SUB_BODY_CODE = '" & SUB_BODY_CODE & "'"
        Update_Record_TDA("ICTBODYS", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("ICTBODYS", New String() {Absx1.txtFor("SUB_BODY_CODE").Text})
        Sort_grdColumns(grdICTBODYS, "SUB_BODY_CODE")
        EnforceConstraints(True)

        tvwICTBODYS.Nodes.Clear()
 
        nodeIndex = 0
        Dim SPEC_GROUP As String = ""
        Add_Nodes(SPEC_GROUP)
        If tvwICTBODYS.Nodes.Count > 0 Then
            tvwICTBODYS.ActiveNode = tvwICTBODYS.Nodes(0)
            tvwICTBODYS.ExpandAll()
            tvwICTBODYS.Visible = True

            ''tvwICTBODYS.DataSource = dst.Tables("ICTBODYS")
            'tvwICTBODYS.Refresh()
        End If




        tvw.Nodes.Clear()

        nodeIndex = 0
        SPEC_GROUP = ""
        Add_Nodes2(SPEC_GROUP)
        If tvw.Nodes.Count > 0 Then
            tvw.ActiveNode = tvw.Nodes(0)
            tvw.ExpandAll()
            tvw.Visible = True

            ''tvwICTBODYS.DataSource = dst.Tables("ICTBODYS")
            'tvwICTBODYS.Refresh()
        End If

    End Sub

    Sub Add_Nodes(SPEC_GROUP As String, Optional pNode As Infragistics.Win.UltraWinTree.UltraTreeNode = Nothing)

        Dim aNode As Infragistics.Win.UltraWinTree.UltraTreeNode = Nothing

        Dim sql As String = ""
        Dim rows() As DataRow = dst.Tables("ICTBODYS").Select("ISNULL(SPEC_GROUP,'') = '" & SPEC_GROUP & "'", "SPEC_SEQ")
        If rows.Length > 0 Then
            Dim nodeIndex As Integer = 0
            For Each row As DataRow In rows

                Dim SPEC_CODE As String = row.Item("SPEC_CODE")
                Dim SPEC_DESC As String = row.Item("SPEC_DESC") & ""

                If tvwICTBODYS.Nodes.Count = 0 Then
                    '  aNode = tvwICTBODYS.Nodes.Add(SPEC_CODE, SPEC_DESC)
                    aNode = tvwICTBODYS.Nodes.Add(SPEC_CODE, SPEC_DESC)
                Else
                    aNode = pNode.Nodes.Insert(nodeIndex, SPEC_GROUP & SPEC_CODE, SPEC_DESC)
                End If
                nodeIndex += 1
                aNode.Visible = True
                Add_Nodes(SPEC_CODE, aNode)
                Setup_Node(aNode)
            Next
        End If
    End Sub

    Sub Add_Nodes2(SPEC_GROUP As String, Optional pNode As Infragistics.Win.UltraWinTree.UltraTreeNode = Nothing)

        Dim aNode As Infragistics.Win.UltraWinTree.UltraTreeNode = Nothing

        Dim sql As String = ""
        Dim rows() As DataRow = dst.Tables("ICTBODYS").Select("ISNULL(SPEC_GROUP,'') = '" & SPEC_GROUP & "'", "SPEC_SEQ")
        If rows.Length > 0 Then
            Dim nodeIndex As Integer = 0
            For Each row As DataRow In rows

                Dim SPEC_CODE As String = row.Item("SPEC_CODE")
                Dim SPEC_DESC As String = row.Item("SPEC_DESC") & ""

                If tvw.Nodes.Count = 0 Then
                    '  aNode = tvwICTBODYS.Nodes.Add(SPEC_CODE, SPEC_DESC)
                    aNode = tvw.Nodes.Add(SPEC_CODE, SPEC_DESC)
                Else
                    aNode = pNode.Nodes.Insert(nodeIndex, SPEC_GROUP & SPEC_CODE, SPEC_DESC)
                End If
                nodeIndex += 1
                aNode.Visible = True
                Add_Nodes2(SPEC_CODE, aNode)
                Setup_Node(aNode)
            Next
        End If
    End Sub

    Sub Setup_Node(aNode As Infragistics.Win.UltraWinTree.UltraTreeNode)
        Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "COLUMN_NAME\STMT_LINE_TYPE\"

        If aNode Is Nothing Then Exit Sub
        If aNode.Nodes.Count > 0 Then
            aNode.Override.NodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "H")
            aNode.Override.ExpandedNodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "H" & "_EXP")
            aNode.Override.NodeAppearance.FontData.Bold = DefaultableBoolean.True
        Else
            If aNode IsNot Nothing Then
                aNode.Override.NodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "D")
            End If
        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ICTBODYS"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdICTBODYS.Enabled = tf
        If Not tf Then
            tvwICTBODYS.Nodes.Clear()
        End If
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTBODYS}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub

#End Region

#Region "grdICTBODYS"

    Private Sub grdICTBODYS_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTBODYS.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "SUB_BODY_CODE"
            '    grdCodeDesc(grdICTBODYS, "ICTBODY2", "SUB_BODY_CODE", "SPEC_DESC")
        End Select
    End Sub

    Private Sub grdICTBODYS_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTBODYS.AfterRowActivate

        With grdICTBODYS.DisplayLayout.Bands(0).Columns("SPEC_CODE")
            If grdICTBODYS.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdICTBODYS_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdICTBODYS.AfterRowsDeleted

    End Sub

    Private Sub grdICTBODYS_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTBODYS.AfterRowUpdate

    End Sub

    Private Sub grdICTBODYS_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdICTBODYS.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SUB_BODY_CODE As String = grow.Cells("SUB_BODY_CODE").Value
        '    Dim SUB_BODY_CODE As String = grow.Cells("SUB_BODY_CODE").Value
        '    Dim rowICTBODYS As DataRow = dst.Tables("ICTBODYS").Rows.Find(New String() {SUB_BODY_CODE, SUB_BODY_CODE})
        '    If Not rowICTBODYS.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdICTBODYS_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTBODYS.BeforeRowUpdate

        'Dim row As DataRow = LookUp("ICTBODY2", e.Row.Cells("SUB_BODY_CODE").Text)
        'If row Is Nothing Then
        '    e.Cancel = True
        'End If

        If e.Row.IsAddRow Then
            e.Row.Cells("SUB_BODY_CODE").Value = Absx1.txtFor("SUB_BODY_CODE").Text
            e.Row.Cells("SPEC_CODE").Value = Format(Val(dst.Tables("ICTBODYS").Compute("MAX(SPEC_CODE)", "") & "") + 1, "000000")
        End If

    End Sub

    Private Sub grdICTBODYS_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTBODYS.ClickCellButton
        'Select Case e.Cell.Column.Key
        '    Case "SUB_BODY_CODE"
        '        Dim sql_where As String = Get_List_of_Codes("ICTBODY2.SUB_BODY_CODE not in", "ICTBODYS", "SUB_BODY_CODE")
        '        grdClickCellButton(grdICTBODYS, sql_where, True)
        'End Select
    End Sub

#End Region



    Sub Set_tvw_Columns()
        With tvwICTBODYS

            .Appearances.Add("DropHighLightAppearance")
            With .Appearances("DropHighLightAppearance")
                .BackColor = System.Drawing.Color.Cyan
            End With

            .DrawFilter = UltraTree_DropHightLight_DrawFilter
            .Override.SelectionType = UltraWinTree.SelectType.ExtendedAutoDrag


            Dim rootColumnSet As UltraWinTree.UltraTreeColumnSet = .ColumnSettings.RootColumnSet

            Dim column As UltraWinTree.UltraTreeNodeColumn = rootColumnSet.Columns.Add("MENU_ITEM_DESC")
            column.CanShowExpansionIndicator = DefaultableBoolean.True
            column.LayoutInfo.SpanX = 4
            column.Text = "Description"
            column.LayoutInfo.PreferredCellSize = New System.Drawing.Size(350, 19)
            column.LayoutInfo.PreferredLabelSize = New System.Drawing.Size(350, 0)

            column = rootColumnSet.Columns.Add("MENU_ITEM_OBJECT")
            column.Text = "Object"
            column.LayoutInfo.PreferredCellSize = New System.Drawing.Size(100, 19)
            column.LayoutInfo.PreferredLabelSize = New System.Drawing.Size(100, 0)
            column.MaxLength = dst.Tables("ASTMENU1").Columns("MENU_ITEM_OBJECT").MaxLength

            column = rootColumnSet.Columns.Add("MENU_ITEM_HIDDEN")
            column.Text = "Hidden"
            column.DataType = GetType(Boolean)
            column.CellAppearance.TextHAlign = HAlign.Center
            column.HeaderAppearance.TextHAlign = HAlign.Center
             
            .ColumnSettings.RootColumnSet.AllowCellEdit = UltraWinTree.AllowCellEdit.Full
            .ColumnSettings.RootColumnSet.ShowSortIndicators = DefaultableBoolean.False
            .ColumnSettings.RootColumnSet.AllowSorting = DefaultableBoolean.False

            Setup_Menu_ViewStyle()
            .ColumnSettings.TabNavigation = UltraWinTree.TabNavigation.NextCell

            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_DESC").LayoutInfo.SpanX = 10
            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_OBJECT").LayoutInfo.LabelSpan = 10
            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_SECURITY").LayoutInfo.PreferredCellSize = New System.Drawing.Size(500, 0)
            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_HIDDEN").EditorComponent = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_STANDALONE").EditorComponent = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
          
        End With
    End Sub


    Sub Setup_Menu_ViewStyle()
        With tvw
            If UltraOptionSet1.Value = "T" Then
                .Override.CellClickAction = UltraWinTree.CellClickAction.Default
                .ViewStyle = UltraWinTree.ViewStyle.Standard
                .AllowDrop = True
                .Override.AllowCut = DefaultableBoolean.True
                .Override.AllowCopy = DefaultableBoolean.True
                .Override.AllowPaste = DefaultableBoolean.True
                .Override.AllowDelete = DefaultableBoolean.True
            Else
                .Override.CellClickAction = UltraWinTree.CellClickAction.EditCellSelectText '  .EditCell
                .ViewStyle = UltraWinTree.ViewStyle.OutlookExpress
                .AllowDrop = False
                .Override.AllowCut = DefaultableBoolean.False
                .Override.AllowCopy = DefaultableBoolean.False
                .Override.AllowPaste = DefaultableBoolean.False
                .Override.AllowDelete = DefaultableBoolean.False
            End If
        End With
    End Sub


    Private Sub UltraOptionSet1_ValueChanged(sender As Object, e As EventArgs) Handles UltraOptionSet1.ValueChanged
        Setup_Menu_ViewStyle()
    End Sub
End Class