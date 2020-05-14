Public Class POTPGMC1

    Dim sqlPOTPGMCU As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select POTPGMCU.*, ASTUSER1.USER_NAME" & vbCrLf _
                & " from POTPGMCU, ASTUSER1" & vbCrLf _
                & " where POTPGMCU.PROGRAM_CATGY_CODE = :PARM1" & vbCrLf _
                & "   and ASTUSER1.USER_ID = POTPGMCU.USER_ID"
            Create_TDA(.Tables.Add, "POTPGMCU", "**", 0, True, "V", 2)
        End With

        grdPOTPGMCU.DataSource = dst.Tables("POTPGMCU")

        With grdPOTPGMCU.DisplayLayout.Bands(0)
            '.Columns("SPEC_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
        End With
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTPGMCU, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

            Case "grdPOTPGMCU"
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
                If grd.Name = "grdPOTPGMCU" Then
                    Add_Codes(grdPOTPGMCU, "ASTUSER1", "USER_ID", "User Codes")
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
        Dim PROGRAM_CATGY_CODE As String = Absx1.txtFor("PROGRAM_CATGY_CODE").Text
        Dim sqlDelete = "PROGRAM_CATGY_CODE = '" & PROGRAM_CATGY_CODE & "'"
        Update_Record_TDA("POTPGMCU", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("POTPGMCU", New String() {Absx1.txtFor("PROGRAM_CATGY_CODE").Text})
        Sort_grdColumns(grdPOTPGMCU, "USER_ID")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"POTPGMCU"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdPOTPGMCU.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTPGMCU}
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

#Region "grdPOTPGMCU"

    Private Sub grdPOTPGMCU_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTPGMCU.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "USER_ID"
                If e.Cell.Value & "" <> "" Then
                    grdCodeDesc(grdPOTPGMCU, "ASTUSER1", "USER_ID", "USER_NAME")
                End If
        End Select
    End Sub

    Private Sub grdPOTPGMCU_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTPGMCU.AfterRowActivate

        With grdPOTPGMCU.DisplayLayout.Bands(0).Columns("USER_ID")
            If grdPOTPGMCU.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdPOTPGMCU_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTPGMCU.AfterRowsDeleted

    End Sub

    Private Sub grdPOTPGMCU_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTPGMCU.AfterRowUpdate

    End Sub

    Private Sub grdPOTPGMCU_BeforeCellUpdate(sender As Object, e As UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdPOTPGMCU.BeforeCellUpdate
        'If e.Cell.Column.Key = "USER_ID" Then
        '    If e.Cell.Value & "" <> "" Then
        '        e.Cell.Value = e.Cell.Value.ToString.ToLower
        '    End If
        'End If
    End Sub

    Private Sub grdPOTPGMCU_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTPGMCU.BeforeExitEditMode
        'If grdPOTPGMCU.ActiveCell.Column.Key = "USER_ID" Then
        '    If grdPOTPGMCU.ActiveCell.Value & "" <> "" Then
        '        grdPOTPGMCU.ActiveCell.Value = grdPOTPGMCU.ActiveCell.Value.ToString.ToLower
        '    End If
        'End If

        If grdPOTPGMCU.ActiveCell IsNot Nothing Then
            With grdPOTPGMCU.ActiveCell
                Select Case .Column.Key
                    Case "USER_ID"
                        If .EditorResolved IsNot Nothing AndAlso .EditorResolved.IsValid Then
                            If .EditorResolved.Value & "" <> "" Then
                                .EditorResolved.Value = .EditorResolved.Value.ToString.ToLower
                            End If
                        End If

                End Select
            End With
        End If

    End Sub

    Private Sub grdPOTPGMCU_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTPGMCU.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim PROGRAM_CATGY_CODE As String = grow.Cells("PROGRAM_CATGY_CODE").Value
        '    Dim PROGRAM_CATGY_CODE As String = grow.Cells("PROGRAM_CATGY_CODE").Value
        '    Dim rowPOTPGMCU As DataRow = dst.Tables("POTPGMCU").Rows.Find(New String() {PROGRAM_CATGY_CODE, PROGRAM_CATGY_CODE})
        '    If Not rowPOTPGMCU.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdPOTPGMCU_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTPGMCU.BeforeRowUpdate

        Dim row As DataRow = LookUp("ASTUSER1", e.Row.Cells("USER_ID").Text)
        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("PROGRAM_CATGY_CODE").Value = Absx1.txtFor("PROGRAM_CATGY_CODE").Text
        End If

    End Sub

    Private Sub grdPOTPGMCU_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTPGMCU.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "USER_ID"
                Dim sql_where As String = Get_List_of_Codes("ASTUSER1.USER_ID not in", "POTPGMCU", "USER_ID")
                grdClickCellButton(grdPOTPGMCU, sql_where, True)
        End Select
    End Sub

#End Region
 
    Private Sub grdPOTPGMCU_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTPGMCU.InitializeLayout

    End Sub
End Class