Public Class SOTSREP1

    Dim sqlSOTSREP2 As String = ""
    Dim sqlSOTSREP3 As String = ""
    Dim sqlSOTSREP4 As String = ""
    Dim sqlSOTSREP5 As String = ""
    Dim sqlSOTSREP6 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            sqlSOTSREP2 = "Select SOTSREP2.*, ARTCUST1.CUST_NAME" _
            & " from ARTCUST1,SOTSREP2" _
            & " where ARTCUST1.CUST_CODE = SOTSREP2.CUST_CODE"
            ASCMAIN1.sql = sqlSOTSREP2 _
            & "  and SOTSREP2.SREP_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "SOTSREP2", "**", 0, True, "V", 3)

            sqlSOTSREP3 = "Select SOTSREP3.*, SOTSDIV1.SALES_DIVISION_NAME" _
            & " from SOTSDIV1,SOTSREP3" _
            & " where SOTSDIV1.SALES_DIVISION_CODE = SOTSREP3.SALES_DIVISION_CODE"
            ASCMAIN1.sql = sqlSOTSREP3 _
            & "  and SOTSREP3.SREP_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "SOTSREP3", "**", 0, True, "V", 2)

            sqlSOTSREP4 = "Select SOTSREP4.*, ICTSGRP1.STYLE_GROUP_DESC" _
            & " from ICTSGRP1,SOTSREP4" _
            & " where ICTSGRP1.STYLE_GROUP_CODE = SOTSREP4.STYLE_GROUP_CODE"
            ASCMAIN1.sql = sqlSOTSREP4 _
            & "  and SOTSREP4.SREP_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "SOTSREP4", "**", 0, True, "V", 2)

            sqlSOTSREP5 = "Select SOTSREP5.*, ARTCUST1.CUST_NAME" _
            & " from ARTCUST1,SOTSREP5" _
            & " where ARTCUST1.CUST_CODE = SOTSREP5.CUST_CODE"
            ASCMAIN1.sql = sqlSOTSREP5 _
            & "  and SOTSREP5.SREP_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "SOTSREP5", "**", 0, True, "V", 2)

            sqlSOTSREP6 = "Select SOTSREP6.*, ICTSGRP1.STYLE_GROUP_DESC" _
            & " from ICTSGRP1,SOTSREP6" _
            & " where ICTSGRP1.STYLE_GROUP_CODE = SOTSREP6.STYLE_GROUP_CODE"
            ASCMAIN1.sql = sqlSOTSREP6 _
            & "  and SOTSREP6.SREP_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "SOTSREP6", "**", 0, True, "V", 3)

            Create_Relation("SOTSREP5", "SOTSREP6", "SREP_CODE,CUST_CODE")
        End With

        grdSOTSREP2.DataSource = dst.Tables("SOTSREP2")
        grdSOTSREP3.DataSource = dst.Tables("SOTSREP3")
        grdSOTSREP4.DataSource = dst.Tables("SOTSREP4")
        grdSOTSREP5.DataSource = dst.Tables("SOTSREP5")

        With grdSOTSREP2.DisplayLayout.Bands(0)
            .Columns("SALES_DIVISION_CODE").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
        End With
        With grdSOTSREP3.DisplayLayout.Bands(0)
            .Columns("SALES_DIVISION_CODE").Header.Fixed = True
            .Columns("SALES_DIVISION_NAME").Header.Fixed = True
        End With
        With grdSOTSREP4.DisplayLayout.Bands(0)
            .Columns("STYLE_GROUP_CODE").Header.Fixed = True
        End With
        With grdSOTSREP5.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
        End With

        Create_Summary(grdSOTSREP2, "SALES_DIVISION_CODE", "Count")
        Create_Summary(grdSOTSREP3, "SALES_DIVISION_CODE", "Count")
        Create_Summary(grdSOTSREP4, "STYLE_GROUP_CODE", "Count")
        Create_Summary(grdSOTSREP5, "CUST_CODE", "Count")

        ReParent_Tabs(tabCommissionDetails)
        splVAN.Visible = (ASCMAIN1.CLIENT = "VAN")
        splRGI.Visible = (ASCMAIN1.CLIENT = "RGI")
        splNYA.Visible = (ASCMAIN1.CLIENT = "NYA")

        chkSREP_GETS_SAMPLES.Visible = (ASCMAIN1.CLIENT = "RGI")
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSREP2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdSOTSREP3, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdSOTSREP4, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdSOTSREP5, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

            Case "grdSOTSREP2", "grdSOTSREP3", "grdSOTSREP4", "grdSOTSREP5"
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
                If grd.Name = "grdSOTSREP3" Then
                    Add_Codes(grdSOTSREP3, "SOTSDIV1", "SALES_DIVISION_CODE", "Divisions")
                ElseIf grd.Name = "grdSOTSREP2" Then
                    Add_Codes(grdSOTSREP2, "ARTCUST1", "CUST_CODE", "Customers")
                ElseIf grd.Name = "grdSOTSREP3" Then
                    Add_Codes(grdSOTSREP2, "ICTSGRP1", "STYLE_GROUP_CODE", "Groups")
                ElseIf grd.Name = "grdSOTSREP4" Then
                    Add_Codes(grdSOTSREP2, "ARTCUST1", "CUST_CODE", "Customers")
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
        Dim sqlDelete = "SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'"
        'INIT_LAST("SOTSREP2", True)
        Update_Record_TDA("SOTSREP2", sqlDelete)
        Update_Record_TDA("SOTSREP3", sqlDelete)
        Update_Record_TDA("SOTSREP4", sqlDelete)
        Update_Record_TDA("SOTSREP5", sqlDelete)
        Update_Record_TDA("SOTSREP6", sqlDelete)
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
        Fill_Records("SOTSREP2", New String() {Absx1.txtFor("SREP_CODE").Text})
        Sort_grdColumns(grdSOTSREP2, "SALES_DIVISION_CODE,CUST_CODE")
        Fill_Records("SOTSREP3", New String() {Absx1.txtFor("SREP_CODE").Text})
        Sort_grdColumns(grdSOTSREP2, "SALES_DIVISION_CODE")
        Fill_Records("SOTSREP4", New String() {Absx1.txtFor("SREP_CODE").Text})
        Sort_grdColumns(grdSOTSREP4, "STYLE_GROUP_CODE")
        Fill_Records("SOTSREP5", New String() {Absx1.txtFor("SREP_CODE").Text})
        Sort_grdColumns(grdSOTSREP5, "CUST_CODE")
        EnforceConstraints(True)

        grdSOTSREP2.Text = "Division / Customer Commission Overrides - SRep " & Absx1.txtFor("SREP_CODE").Text
        grdSOTSREP3.Text = "Division Commission Overrides - SRep " & Absx1.txtFor("SREP_CODE").Text
        grdSOTSREP4.Text = "Commission Overrides by Group - SRep " & Absx1.txtFor("SREP_CODE").Text
        grdSOTSREP5.Text = "Commission Overrides by Account/Group - SRep " & Absx1.txtFor("SREP_CODE").Text
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SOTSREP2", "SOTSREP3", "SOTSREP4", "SOTSREP5", "SOTSREP6"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSOTSREP2.Enabled = tf
        grdSOTSREP3.Enabled = tf
        grdSOTSREP4.Enabled = tf
        grdSOTSREP5.Enabled = tf

        If (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") Then
            splNYA.Panel2Collapsed = Not ScreenMode Or (Absx1.txtFor("SREP_CODE").Text <> "17")
        End If

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTSREP2, grdSOTSREP3, grdSOTSREP4, grdSOTSREP5}
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

#Region "grdSOTSREP3"

    Private Sub grdSOTSREP3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "SALES_DIVISION_CODE"
                grdCodeDesc(grdSOTSREP3, "SOTSDIV1", "SALES_DIVISION_CODE", "SALES_DIVISION_NAME")
        End Select
    End Sub

    Private Sub grdSOTSREP3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSREP3.BeforeRowUpdate
        Dim row As DataRow = LookUp("SOTSDIV1", e.Row.Cells("SALES_DIVISION_CODE").Text)
        If row Is Nothing Then e.Cancel = True

        If e.Row.IsAddRow Then
            e.Row.Cells("SREP_CODE").Value = Absx1.txtFor("SREP_CODE").Text
        End If
    End Sub

    Private Sub grdSOTSREP3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP3.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "SALES_DIVISION_CODE"
                Dim sql_where As String = Get_List_of_Codes("SOTSDIV1.SALES_DIVISION_CODE not in", "SOTSREP3", "SALES_DIVISION_CODE")
                grdClickCellButton(grdSOTSREP3, sql_where, True)
        End Select
    End Sub

    Private Sub grdSOTSREP3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSREP3.InitializeRow
        grd_RowColor(dst.Tables("SOTSREP3"), e.Row)
    End Sub

#End Region

#Region "grdSOTSREP2"

    Private Sub grdSOTSREP2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                grdCodeDesc(grdSOTSREP2, "ARTCUST1", "CUST_CODE", "CUST_NAME")
        End Select
    End Sub

    Private Sub grdSOTSREP2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSREP2.AfterRowActivate
        With grdSOTSREP2.DisplayLayout.Bands(0).Columns("CUST_CODE")
            If grdSOTSREP2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSOTSREP2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTSREP2.AfterRowsDeleted

    End Sub

    Private Sub grdSOTSREP2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTSREP2.AfterRowUpdate

    End Sub

    Private Sub grdSOTSREP2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTSREP2.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        'Next
    End Sub

    Private Sub grdSOTSREP2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSREP2.BeforeRowUpdate
        Dim row As DataRow = LookUp("SOTSDIV1", e.Row.Cells("SALES_DIVISION_CODE").Text)
        If row Is Nothing Then e.Cancel = True
        If e.Row.IsAddRow Then
            e.Row.Cells("SREP_CODE").Value = Absx1.txtFor("SREP_CODE").Text
            If e.Row.Cells("SALES_DIVISION_CODE").Value & "" = "" And grdSOTSREP3.ActiveRow IsNot Nothing Then
                e.Row.Cells("SALES_DIVISION_CODE").Value = grdSOTSREP3.ActiveRow.Cells("SALES_DIVISION_CODE").Value & ""
            End If
        End If
    End Sub

    Private Sub grdSOTSREP2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "SALES_DIVISION_CODE"
                grdClickCellButton(grdSOTSREP2, "", False)
            Case "CUST_CODE"
                Dim sql_where As String = Get_List_of_Codes("ARTCUST1.CUST_CODE not in", "SOTSREP2", "CUST_CODE")
                grdClickCellButton(grdSOTSREP2, sql_where, True)
        End Select
    End Sub

    Private Sub grdSOTSREP2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSREP2.InitializeRow
        grd_RowColor(dst.Tables("SOTSREP2"), e.Row)
    End Sub

#End Region

#Region "grdSOTSREP4"

    Private Sub grdSOTSREP4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP4.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_GROUP_CODE"
                grdCodeDesc(grdSOTSREP4, "ICTSGRP1", "STYLE_GROUP_CODE", "STYLE_GROUP_DESC")
        End Select
    End Sub

    Private Sub grdSOTSREP4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSREP4.BeforeRowUpdate
        Dim row As DataRow = LookUp("ICTSGRP1", e.Row.Cells("STYLE_GROUP_CODE").Text)
        If row Is Nothing Then e.Cancel = True

        If e.Row.IsAddRow Then
            e.Row.Cells("SREP_CODE").Value = Absx1.txtFor("SREP_CODE").Text
        End If
    End Sub

    Private Sub grdSOTSREP4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP4.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "STYLE_GROUP_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTSGRP1.STYLE_GROUP_CODE not in", "SOTSREP4", "STYLE_GROUP_CODE")
                grdClickCellButton(grdSOTSREP4, sql_where, True)
        End Select
    End Sub

    Private Sub grdSOTSREP4_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSREP4.InitializeRow
        grd_RowColor(dst.Tables("SOTSREP4"), e.Row)
    End Sub

#End Region

#Region "grdSOTSREP5"

    Private Sub grdSOTSREP5_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP5.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                grdCodeDesc(grdSOTSREP5, "ARTCUST1", "CUST_CODE", "CUST_NAME")
            Case "STYLE_GROUP_CODE"
                grdCodeDesc(grdSOTSREP5, "ICTSGRP1", "STYLE_GROUP_CODE", "STYLE_GROUP_DESC")
        End Select
    End Sub

    Private Sub grdSOTSREP5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSREP5.BeforeRowUpdate
        Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)
        If row Is Nothing Then e.Cancel = True

        If e.Row.IsAddRow Then
            e.Row.Cells("SREP_CODE").Value = Absx1.txtFor("SREP_CODE").Text
        End If
    End Sub

    Private Sub grdSOTSREP5_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSREP5.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                Dim sql_where As String = Get_List_of_Codes("ARTCUST1.CUST_CODE not in", "SOTSREP5", "CUST_CODE")
                grdClickCellButton(grdSOTSREP5, sql_where, True)
            Case "STYLE_GROUP_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdSOTSREP5, sql_where, True)
        End Select
    End Sub

    Private Sub grdSOTSREP5_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSREP5.InitializeRow
        grd_RowColor(dst.Tables("SOTSREP5"), e.Row)
    End Sub

#End Region

End Class