Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

Public Class SOFFRGHT

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "SOTFRGHT", "*", 0, True, "", 1)
            .Tables("SOTFRGHT").Columns.Add("STATE_NAME", GetType(String), "STATE_CODE")
            ASCMAIN1.Add_Value_List(grdSOTFRGHT, "STATE_NAME", "SELECT STATE_CODE, STATE_NAME FROM TATSTATE")
        End With

        grdSOTFRGHT.DataSource = dst.Tables("SOTFRGHT")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty

        Select Case eItemKey

            Case "Load"

            Case "Cancel"
                If MessageBox.Show($"Do you want to {eItemKey} changes?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

            Case "Update"
                If MessageBox.Show($"Do you want to {eItemKey} changes?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSOTFRGHT.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("SOTFRGHT").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading data")

        EnforceConstraints(False)
        Fill_Records("SOTFRGHT", String.Empty, True, "SELECT * FROM SOTFRGHT")
        EnforceConstraints(True)

        Sort_grdColumns(grdSOTFRGHT, "STATE_CODE")

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()
            Update_Record_TDA("SOTFRGHT")
            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTFRGHT, "SS", "Show Filter", "Show GroupBox")
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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        '  Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTFRGHT_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdSOTFRGHT.BeforeRowUpdate

        'STATE_CODE   Not Null VARCHAR2(2) 
        'FREIGHT_PERC          Number(5, 2) 

        Dim STATE_CODE As String = e.Row.Cells("STATE_CODE").Value & String.Empty
        STATE_CODE = STATE_CODE.ToUpper.Trim
        e.Row.Cells("STATE_CODE").Value = STATE_CODE
        If STATE_CODE.Length = 0 Then
            MessageBox.Show("State Code is required", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim FREIGHT_PERC As Decimal = Val(e.Row.Cells("FREIGHT_PERC").Value & String.Empty)
        If FREIGHT_PERC < 0 OrElse FREIGHT_PERC > 75 Then
            MessageBox.Show("Freight Percentage must be between 0 and 75 percent", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub grdSOTFRGHT_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdSOTFRGHT.ClickCellButton
        Dim COLUMN_NAME As String = e.Cell.Column.Key

        Select Case COLUMN_NAME
            Case "STATE_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdSOTFRGHT, sql_where)

        End Select
    End Sub

#End Region

End Class