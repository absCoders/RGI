Public Class WHTGUNA1

    Dim sqlWHTPPKM2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from WHTGUNA2 " _
                & " where WHTGUNA2.APP_ID = :PARM1"
            Create_TDA(.Tables.Add, "WHTGUNA2", "**", 0, True, "V", 2)
        End With

        grdWHTGUNA2.DataSource = dst.Tables("WHTGUNA2")

        ASCMAIN1.Add_Value_List(grdWHTGUNA2, "RESPONSE_TYPE", , New String() {":", "S:Scan", "Y:Yes/No"})

        'With grdWHTGUNA2.DisplayLayout.Override
        '    .AllowAddNew = UltraWinGrid.AllowAddNew.Yes
        '    .AllowDelete = DefaultableBoolean.True
        '    .AllowUpdate = DefaultableBoolean.True
        'End With
        ''grdWHTGUNA2.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        'With grdWHTGUNA2.DisplayLayout.Bands("WHTGUNA2")
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        If New String() {"APP_SEQ_NO"}.Contains(gcol.Key) Then
        '            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        '        Else
        '            gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
        '            gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        '            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightSkyBlue
        '            gcol.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
        '        End If
        '    Next
        'End With


    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTGUNA2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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
                    Add_Codes(grdWHTGUNA2, "ICTSTYL1", "STYLE_CODE", "Styles")
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
                ASCMAIN1.sql = "Select Max(APP_ID) as MAX_ID from WHTGUNA1 "
                Dim rowWHTGUNA1 As DataRow = ASCDATA1.GetDataRow
                If rowWHTGUNA1 IsNot Nothing Then
                    Absx1.txtFor("APP_ID").Text = Format(Val(rowWHTGUNA1.Item("MAX_ID")) + 1, "00")
                End If


            Case "Edit"
            Case "Update"

        End Select
    End Sub



    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"WHTGUNA2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("WHTGUNA2", New String() {Absx1.txtFor("APP_ID").Text})
        Sort_grdColumns(grdWHTGUNA2, "APP_SEQ_NO")
        If EntryMode = "New" Then
            Absx1.txtFor("PROCEDURE_NAME").Text = "SCANGUNS."
        End If
        EnforceConstraints(True)
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdWHTGUNA2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdWHTGUNA2}
            With grd.DisplayLayout.Override
                If (EntryMode = "New" Or EntryMode = "Edit") Then
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

    Private Sub grdWHTPPKM2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                grdCodeDesc(grdWHTGUNA2, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
        End Select
    End Sub

    Private Sub grdWHTPPKM2_AfterRowActivate(sender As Object, e As System.EventArgs)

        'With grdWHTGUNA2.DisplayLayout.Bands(0).Columns("STYLE_CODE")
        '    If grdWHTGUNA2.ActiveRow.IsAddRow Then
        '        .CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Else
        '        .CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdWHTPPKM2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs)
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

#End Region

    Private Sub grdWHTGUNA2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTGUNA2.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("APP_ID").Value = Absx1.txtFor("APP_ID").Text
            e.Row.Cells("APP_SEQ_NO").Value = Val(dst.Tables("WHTGUNA2").Compute("COUNT(APP_SEQ_NO)", "") & "") + 1
            e.Row.Cells("VALIDATION_PROCEDURE").Value = IIf(e.Row.Cells("VALIDATION_PROCEDURE").Value = "", "", "SCANGUNS." & e.Row.Cells("VALIDATION_PROCEDURE").Value)
        End If
    End Sub


    Private Sub grdWHTGUNA2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTGUNA2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "VALIDATION_PROCEDURE"
                ASCMAIN1.CodeSelector.Get_SQL("VIEW_NAME")

                ASCMAIN1.CodeSelector.SQL = " Select OBJECT_NAME || '.' || PROCEDURE_NAME as PROCEDURE_NAME from USER_PROCEDURES" _
                & " Where OBJECT_NAME = 'SCANGUNS'"

                ASCMAIN1.CodeSelector.MultipleSelections = False
                Using F As New ASFCODE1
                    F.ShowDialog()
                End Using
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    grdWHTGUNA2.ActiveRow.Cells("VALIDATION_PROCEDURE").Value = ASCMAIN1.CodeSelector.SelectedCode
                End If
        End Select



    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim sqlDelete = ""
        Update_Record_TDA("WHTGUNA2")
    End Sub
End Class