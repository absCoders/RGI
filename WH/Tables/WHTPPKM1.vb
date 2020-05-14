Public Class WHTPPKM1

    Dim sqlWHTPPKM2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select WHTPPKM2.*, ICTSTYL1.STYLE_DESC" _
                & " from WHTPPKM2,ICTSTYL1" _
                & " where WHTPPKM2.PPK_CODE = :PARM1" _
                & "   and ICTSTYL1.STYLE_CODE = WHTPPKM2.STYLE_CODE"
            Create_TDA(.Tables.Add, "WHTPPKM2", "**", 0, True, "V", 2)
        End With

        grdWHTPPKM2.DataSource = dst.Tables("WHTPPKM2")

        Create_Summary(grdWHTPPKM2, "STYLE_CODE", "Count")
        Create_Summary(grdWHTPPKM2, "PPK_QTY")

        With grdWHTPPKM2.DisplayLayout.Bands(0)
            .Columns("STYLE_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
        End With
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTPPKM2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

            Case "grdWHTPPKM2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
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
            Case "Add Codes"
                If grd.Name = "grdWHTPPKM2" Then
                    Add_Codes(grdWHTPPKM2, "ICTSTYL1", "STYLE_CODE", "Styles")
                End If
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
        Dim PPK_CODE As String = Absx1.txtFor("PPK_CODE").Text
        Dim sqlDelete = "PPK_CODE = '" & PPK_CODE & "'"
        Update_Record_TDA("WHTPPKM2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("WHTPPKM2", New String() {Absx1.txtFor("PPK_CODE").Text})
        Sort_grdColumns(grdWHTPPKM2, "STYLE_CODE,COLOR_CODE")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"WHTPPKM2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdWHTPPKM2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdWHTPPKM2}
            With grd.DisplayLayout.Override
                If (EntryMode = "New" Or EntryMode = "Edit") And "" <> "" Then
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

#Region "grdWHTPPKM2"

    Private Sub grdWHTPPKM2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTPPKM2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                grdCodeDesc(grdWHTPPKM2, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
        End Select
    End Sub

    Private Sub grdWHTPPKM2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTPPKM2.AfterRowActivate

        With grdWHTPPKM2.DisplayLayout.Bands(0).Columns("STYLE_CODE")
            If grdWHTPPKM2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdWHTPPKM2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdWHTPPKM2.AfterRowsDeleted

    End Sub

    Private Sub grdWHTPPKM2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTPPKM2.AfterRowUpdate

    End Sub

    Private Sub grdWHTPPKM2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdWHTPPKM2.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SREP_CODE As String = grow.Cells("SREP_CODE").Value
        '    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
        '    Dim rowWHTPPKM2 As DataRow = dst.Tables("WHTPPKM2").Rows.Find(New String() {SREP_CODE, CUST_CODE})
        '    If Not rowWHTPPKM2.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdWHTPPKM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTPPKM2.BeforeRowUpdate

        Dim row As DataRow = LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)
        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("PPK_CODE").Value = Absx1.txtFor("PPK_CODE").Text
        End If

    End Sub

    Private Sub grdWHTPPKM2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTPPKM2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTSTYL1.STYLE_CODE not in", "WHTPPKM2", "STYLE_CODE")
                grdClickCellButton(grdWHTPPKM2, sql_where, True)
        End Select
    End Sub

#End Region

    Private Sub Panel1_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class