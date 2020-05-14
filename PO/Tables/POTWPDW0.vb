Public Class POTWPDW0

    Dim STEP_LNO_default_task_list As Int32 = 0

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Create_TDA(.Tables.Add, "POTWPDW1", "*", 1, True)
            Create_TDA(.Tables.Add, "POTWPDW2", "*", 1, True)

            Create_Relation("POTWPDW1", "POTWPDW2", "STEP_TEMPLATE,STEP_LNO")
        End With

        grdPOTWPDW1.DataSource = dst.Tables("POTWPDW1")
        grdPOTWPDW2.DataSource = dst.Tables("POTWPDW2")

        Create_Summary(grdPOTWPDW2, "TASK_LNO", "Count")
        Create_Summary(grdPOTWPDW1, "STEP_LNO", "Count")

        ASCMAIN1.Add_Value_List(grdPOTWPDW1, "STEP_STAGE")
        '   ASCMAIN1.Add_Value_List(grdPOTWPDW1, "STEP_ACTION_DATE_NAME", Nothing, New String() {":", "START_DATE:Start Prod", "SHIP_DATE:Ship Date", "ETA_DATE:ETA", "IN_STORE_DATE:In Store"})
        ASCMAIN1.Add_Value_List(grdPOTWPDW2, "TASK_DIR")
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTWPDW1, "BBB", "Re-Number Items using Current Sort", "Re-Number Items using Sequence", "Set as Default Task List")
        Load_Popup_Menu(grdPOTWPDW2, "B", "Set as Default Task List")
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

            Case "grdPOTWPDW1"
                tlb_btn = DirectCast(tlb_pop.Tools("Re-Number Items using Current Sort"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
                tlb_btn = DirectCast(tlb_pop.Tools("Re-Number Items using Sequence"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
                tlb_btn = DirectCast(tlb_pop.Tools("Set as Default Task List"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New") And grdPOTWPDW1.ActiveRow IsNot Nothing AndAlso Not grdPOTWPDW1.ActiveRow.IsAddRow

            Case "grdPOTWPDW2"
                tlb_btn = DirectCast(tlb_pop.Tools("Set as Default Task List"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New") And grdPOTWPDW1.ActiveRow IsNot Nothing AndAlso Not grdPOTWPDW1.ActiveRow.IsAddRow

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

            Case "Re-Number Items using Current Sort"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Re-Numbering Items using Current Sort")

                grdPOTWPDW1.Tag = "X"
                Dim SEQ As Integer = 0
                Dim STEP_LNO_default_task_list_new As Int32 = 0
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTWPDW1.Rows
                    SEQ += 1
                    If grow.Cells("STEP_LNO").Value = STEP_LNO_default_task_list Then
                        STEP_LNO_default_task_list_new = SEQ
                    End If
                    grow.Cells("SEQ").Value = SEQ * 10
                    grow.Cells("STEP_LNO").Value = -1 * SEQ
                    grow.Update()
                Next
                grdPOTWPDW1.Tag = ""
                STEP_LNO_default_task_list = STEP_LNO_default_task_list_new
                Show_Default_Task_List()

                For Each row As DataRow In dst.Tables("POTWPDW1").Select("")
                    row.Item("STEP_LNO") = -1 * Val(row.Item("STEP_LNO"))
                Next
                Sort_grdColumns(grdPOTWPDW1, "STEP_LNO")
                Setup_grdPOTWPDW2()
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("")

            Case "Re-Number Items using Sequence"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Re-Numbering Items using Sequence")

                Dim SEQ As Integer = 0
                Dim STEP_LNO_default_task_list_new As Int32 = 0
                For Each row As DataRow In dst.Tables("POTWPDW1").Select("", "SEQ")
                    SEQ += 1
                    If row.Item("STEP_LNO") = STEP_LNO_default_task_list Then
                        STEP_LNO_default_task_list_new = SEQ
                    End If
                    row.Item("SEQ") = SEQ * 10
                    row.Item("STEP_LNO") = -1 * SEQ
                Next
                STEP_LNO_default_task_list = STEP_LNO_default_task_list_new
                Show_Default_Task_List()
                For Each row As DataRow In dst.Tables("POTWPDW1").Select("")
                    row.Item("STEP_LNO") = -1 * Val(row.Item("STEP_LNO"))
                Next
                Sort_grdColumns(grdPOTWPDW1, "STEP_LNO")
                Setup_grdPOTWPDW2()
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("")

            Case "Set as Default Task List"
                STEP_LNO_default_task_list = grdPOTWPDW1.ActiveRow.Cells("STEP_LNO").Value
                Show_Default_Task_List()

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

    Sub Show_Default_Task_List()
        If STEP_LNO_default_task_list = 0 Then
            lblDefaultTaskList.Visible = False
        Else
            lblDefaultTaskList.Visible = True
            lblDefaultTaskList.Text = "Default Task List: Line " & CStr(STEP_LNO_default_task_list)
        End If

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey

            Case "New"
            Case "Edit"
            Case "Update"

                For Each rowPOTWPDW1 As DataRow In dst.Tables("POTWPDW1").Select("")
                    If rowPOTWPDW1.GetChildRows("POTWPDW1_POTWPDW2").Length = 0 Then
                        EMsg &= vbCr & "No Tasks Defined for Step " & rowPOTWPDW1.Item("STEP_LNO")
                    End If
                Next

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = "STEP_TEMPLATE = '" & Absx1.txtFor("STEP_TEMPLATE").Text & "'"
        Update_Record_TDA("POTWPDW1", sqlDelete)
        Update_Record_TDA("POTWPDW2", sqlDelete)
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        'If EntryMode = "New" Then
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        'Else
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        'End If
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("POTWPDW2", New String() {Absx1.txtFor("STEP_TEMPLATE").Text})
        Fill_Records("POTWPDW1", New String() {Absx1.txtFor("STEP_TEMPLATE").Text})
        Sort_grdColumns(grdPOTWPDW1, "STEP_LNO")
        EnforceConstraints(True)

        grdPOTWPDW2.Text = "Step Tasks - Template " & Absx1.txtFor("STEP_TEMPLATE").Text & ", Step " & ""
        grdPOTWPDW1.Text = "Steps - Template " & Absx1.txtFor("STEP_TEMPLATE").Text

        Setup_grdPOTWPDW2()

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"POTWPDW1", "POTWPDW2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdPOTWPDW2.Enabled = tf
        grdPOTWPDW1.Enabled = tf

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTWPDW2, grdPOTWPDW1}
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

        STEP_LNO_default_task_list = 0
        Show_Default_Task_List()

    End Sub

#End Region

#Region "grdPOTWPDW1"

    Private Sub grdPOTWPDW1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTWPDW1.AfterRowActivate
        Setup_grdPOTWPDW2()
    End Sub

    Private Sub grdPOTWPDW1_AfterRowInsert(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTWPDW1.AfterRowInsert

    End Sub

    Private Sub grdPOTWPDW1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTWPDW1.AfterRowUpdate
        If grdPOTWPDW1.Tag = "X" Then Exit Sub
        Dim rowPOTWPDW1 As DataRow = dst.Tables("POTWPDW1").Rows.Find(New Object() {e.Row.Cells("STEP_TEMPLATE").Value, e.Row.Cells("STEP_LNO").Value})
        If rowPOTWPDW1.GetChildRows("POTWPDW1_POTWPDW2").Length = 0 Then
            If STEP_LNO_default_task_list = 0 Then
                Dim rowPOTWPDW2 As DataRow = dst.Tables("POTWPDW2").NewRow
                rowPOTWPDW2.Item("STEP_TEMPLATE") = Absx1.txtFor("STEP_TEMPLATE").Text
                rowPOTWPDW2.Item("STEP_LNO") = e.Row.Cells("STEP_LNO").Value
                rowPOTWPDW2.Item("TASK_LNO") = 1
                rowPOTWPDW2.Item("TASK_DESC") = ""
                rowPOTWPDW2.Item("TASK_DIR") = DBNull.Value
                dst.Tables("POTWPDW2").Rows.Add(rowPOTWPDW2)
            Else
                For Each rowPOTWPDW2_default As DataRow In dst.Tables("POTWPDW2").Select _
                        ("STEP_LNO = " & CStr(STEP_LNO_default_task_list), "TASK_LNO")
                    Dim rowPOTWPDW2 As DataRow = dst.Tables("POTWPDW2").NewRow
                    rowPOTWPDW2.Item("STEP_TEMPLATE") = rowPOTWPDW2_default.Item("STEP_TEMPLATE")
                    rowPOTWPDW2.Item("STEP_LNO") = e.Row.Cells("STEP_LNO").Value
                    rowPOTWPDW2.Item("TASK_LNO") = rowPOTWPDW2_default.Item("TASK_LNO")
                    rowPOTWPDW2.Item("TASK_DESC") = rowPOTWPDW2_default.Item("TASK_DESC")
                    rowPOTWPDW2.Item("TASK_DIR") = rowPOTWPDW2_default.Item("TASK_DIR")
                    dst.Tables("POTWPDW2").Rows.Add(rowPOTWPDW2)
                Next
            End If
        End If
    End Sub

    Private Sub grdPOTWPDW1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDW1.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("STEP_TEMPLATE").Value = Absx1.txtFor("STEP_TEMPLATE").Text
            e.Row.Cells("STEP_LNO").Value = Val(dst.Tables("POTWPDW1").Compute("MAX(STEP_LNO)", "") & "") + 1
            e.Row.Cells("SEQ").Value = e.Row.Cells("STEP_LNO").Value * 10
        End If
    End Sub
#End Region

    Sub Setup_grdPOTWPDW2()
        If grdPOTWPDW1.ActiveRow Is Nothing OrElse grdPOTWPDW1.ActiveRow.IsAddRow Then
            grdPOTWPDW2.Visible = False
        Else

            Dim STEP_TEMPLATE As String = grdPOTWPDW1.ActiveRow.Cells("STEP_TEMPLATE").Value
            Dim STEP_LNO As Int32 = Val(grdPOTWPDW1.ActiveRow.Cells("STEP_LNO").Value & "")

            grdPOTWPDW2.Visible = True
            Dim dvw As DataView = DirectCast(grdPOTWPDW2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "STEP_TEMPLATE = '" & STEP_TEMPLATE & "' and STEP_LNO = " & CStr(STEP_LNO)
            grdPOTWPDW2.Text = "Step Tasks - Template " & Absx1.txtFor("STEP_TEMPLATE").Text & ", Step " & CStr(STEP_LNO)
        End If

    End Sub

#Region "grdPOTWPDW2"
    Private Sub grdPOTWPDW2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWPDW2.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("STEP_TEMPLATE").Value = Absx1.txtFor("STEP_TEMPLATE").Text
            e.Row.Cells("STEP_LNO").Value = grdPOTWPDW1.ActiveRow.Cells("STEP_LNO").Value
            e.Row.Cells("TASK_LNO").Value = Val(dst.Tables("POTWPDW2").Compute("MAX(TASK_LNO)", "STEP_LNO = " & e.Row.Cells("STEP_LNO").Value) & "") + 1
        End If
    End Sub
#End Region
     
End Class