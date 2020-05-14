Public Class SOTPRIC1

    Dim sqlSOTPRIC2 As String = ""
    Dim isCurrRetailPriceList As Boolean = False

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Get_PARM("SOTPARM1")

        With dst

            '     Create_TDA(.Tables.Add, "TATALRT1", "*")

            ASCMAIN1.sql = "Select SOTPRIC2.*, ICTSTYL1.STYLE_DESC" & vbCrLf _
            & " from SOTPRIC2,ICTSTYL1" & vbCrLf _
            & " where SOTPRIC2.PRICE_LIST_CODE = :PARM1 " & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = SOTPRIC2.STYLE_CODE"
            Create_TDA(.Tables.Add, "SOTPRIC2", "**", 0, True, "V", 2)
        End With

        grdSOTPRIC2.DataSource = dst.Tables("SOTPRIC2")

        grdSOTPRIC2.DisplayLayout.UseFixedHeaders = True
        With grdSOTPRIC2.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTPRIC2.DisplayLayout.Bands(0).Columns
            If New String() {"STYLE_CODE", "STYLE_PRICE", "STYLE_NEW_PRICE", "STYLE_NEW_PRICE_DATE"}.Contains(gcol.Key) Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellAppearance.BackColor = Drawing.Color.LightGray
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        ' HIDE THESE UNTIL WE HAVE DEFINED BUSINESS RULES AS TO WHAT WE NEED TO DO AT MONTH END AND PRICING IN EDI
        grdSOTPRIC2.DisplayLayout.Bands(0).Columns("STYLE_NEW_PRICE").Hidden = True
        grdSOTPRIC2.DisplayLayout.Bands(0).Columns("STYLE_NEW_PRICE_DATE").Hidden = True

        Create_Summary(grdSOTPRIC2, "STYLE_CODE", "Count")

        With grdSOTPRIC2.DisplayLayout.Bands(0)

        End With

        Dim tbl As DataTable = ASCDATA1.GetDataTable("Select * from TATCURR1")
        If tbl.Rows.Count < 2 Or (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT") Then
            lblCURR_CODE.Visible = False
            txtCURR_CODE.Visible = False
        Else
            lblCURR_CODE.Visible = True
            txtCURR_CODE.Visible = True
        End If
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPRIC2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")

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
        'if not new or edit - hide add codes

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

            Case "grdSOTPRIC2"
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
                If grd.Name = "grdSOTPRIC2" Then
                    Add_Codes(grdSOTPRIC2, "ICTSTYL1", "STYLE_CODE", "Items")
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
                If Absx1.txtFor("PRICE_LIST_DESC").Text = "" Then
                    EMsg &= vbCr & "Description is Mandatory"
                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    If Absx1.txtFor("CURR_CODE").Text <> "USD" Then
                        EMsg &= vbCr & "Non-USD prices are not supported"
                    End If
                Else
                    Dim CURR_CODE As String = Absx1.txtFor("CURR_CODE").Text
                    If CURR_CODE = "" Then
                        EMsg &= vbCr & "Currency Code is Mandatory"
                    Else
                        Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", CURR_CODE)
                        If rowTATCURR1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Currency Code (" & CURR_CODE & ")"
                        End If
                    End If

                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    ' LAUREN SAID NOT TO PUT THIS IN. 11/10
                Else
                    For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Select("")
                        Dim STYLE_CODE As String = rowSOTPRIC2.Item("STYLE_CODE") & ""
                        Dim STYLE_NEW_PRICE As Decimal = Val(rowSOTPRIC2.Item("STYLE_NEW_PRICE") & "")
                        If Val(STYLE_NEW_PRICE) < 0 Then
                            EMsg &= vbCr & STYLE_CODE & ":" & "New Item Price must be > 0"
                        End If
                        If rowSOTPRIC2.Item("STYLE_NEW_PRICE_DATE") & "" <> "" Then
                            If Val(STYLE_NEW_PRICE) <= 0 Then
                                EMsg &= vbCr & STYLE_CODE & ":" & "New Item Price Date provided without a New Price"
                            End If
                            Dim DTE As Date = rowSOTPRIC2.Item("STYLE_NEW_PRICE_DATE")
                            If Format(DTE, "dd") <> "01" Or Format(DTE, "yyyyMM") <= ASCMAIN1.CYM Then
                                EMsg &= vbCr & STYLE_CODE & ":" & "New Item Price Date must be the 1st of a Future Month"
                            End If
                        Else
                            If Val(STYLE_NEW_PRICE) > 0 Then
                                EMsg &= vbCr & STYLE_CODE & ":" & "New Item Price Date is Mandatory if specifying a New Price"
                            End If
                        End If
                    Next
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()


        Dim PRICE_LIST_CODE As String = Absx1.txtFor("PRICE_LIST_CODE").Text


        For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Rows
            If rowSOTPRIC2.RowState = DataRowState.Deleted Then
                For Each dcol As DataColumn In dst.Tables("SOTPRIC2").Columns
                    Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                    With rowASTAUDT1
                        .Item("TABLE_NAME") = "SOTPRIC2"
                        .Item("KEY_VALUE") = rowSOTPRIC2.Item("PRICE_LIST_CODE", DataRowVersion.Original) & ":" & rowSOTPRIC2.Item("STYLE_CODE", DataRowVersion.Original)
                        .Item("COLUMN_NAME") = dcol.ColumnName
                        .Item("USER_ID") = ASCMAIN1.USER_ID
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("OLD_VALUE") = rowSOTPRIC2.Item(dcol.ColumnName, DataRowVersion.Original)
                        '.Item("NEW_VALUE") =  
                        .Item("FM_MODE") = "D"
                        '.Item("KEY_VALUE2") =
                        '.Item("KEY_LNO") =
                        .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                        .Item("SELECTION_NO") = SELECTION_NO
                        .Item("XNO") = XNO
                    End With
                    dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                Next
            Else
                ' NOT WRITING ADDED ROWS
                If rowSOTPRIC2.RowState = DataRowState.Added Then
                Else

                    For Each dcol As DataColumn In dst.Tables("SOTPRIC2").Columns

                        If rowSOTPRIC2.RowState = DataRowState.Added OrElse _
                           rowSOTPRIC2.Item(dcol.ColumnName, DataRowVersion.Original) & "" <> _
                           rowSOTPRIC2.Item(dcol.ColumnName, DataRowVersion.Current) & "" Then

                            Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                            With rowASTAUDT1
                                .Item("TABLE_NAME") = "SOTPRIC2"
                                .Item("KEY_VALUE") = rowSOTPRIC2.Item("PRICE_LIST_CODE", DataRowVersion.Original) & ":" & rowSOTPRIC2.Item("STYLE_CODE", DataRowVersion.Original)
                                .Item("COLUMN_NAME") = dcol.ColumnName
                                .Item("USER_ID") = ASCMAIN1.USER_ID
                                .Item("INIT_DATE") = DATETIME_STAMP
                                .Item("OLD_VALUE") = rowSOTPRIC2.Item(dcol.ColumnName, DataRowVersion.Original)
                                .Item("NEW_VALUE") = rowSOTPRIC2.Item(dcol.ColumnName)
                                .Item("FM_MODE") = "E"
                                '.Item("KEY_VALUE2") =
                                '.Item("KEY_LNO") =
                                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                                .Item("SELECTION_NO") = SELECTION_NO
                                .Item("XNO") = XNO
                            End With
                            dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                        End If
                    Next
                End If
            End If
            'Dim STYLE_CODE As String = rowSOTPRIC2.Item("STYLE_CODE")
            'Dim STYLE_PRICE As Decimal = Val(rowSOTPRIC2.Item("STYLE_PRICE") & "")
            'Dim STYLE_NEW_PRICE As Decimal = Val(rowSOTPRIC2.Item("STYLE_NEW_PRICE") & "")
            'Dim STYLE_NEW_PRICE_DATE = rowSOTPRIC2.Item("STYLE_NEW_PRICE_DATE") & ""

        Next
        Update_Record_TDA("ASTAUDT1")

        Dim sqlDelete = "PRICE_LIST_CODE = '" & Absx1.txtFor("PRICE_LIST_CODE").Text & "'"
        Update_Record_TDA("SOTPRIC2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()

        Dim PRICE_LIST_CODE As String = Absx1.txtFor("PRICE_LIST_CODE").Text

        EnforceConstraints(False)
        Fill_Records("SOTPRIC2", New String() {PRICE_LIST_CODE})
        Sort_grdColumns(grdSOTPRIC2, "STYLE_CODE")
        grdSOTPRIC2.Text = "Price List Details for " & PRICE_LIST_CODE

        Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", PRICE_LIST_CODE)
        If rowTATCURR1 IsNot Nothing Then
            isCurrRetailPriceList = True
        Else
            isCurrRetailPriceList = False
        End If

        If EntryMode = "New" Then
            If Not isCurrRetailPriceList Then
                rowASFBASE1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            End If
        End If

        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTPRIC2").Rows.Clear()
            EnforceConstraints(True)
        End If

        grdSOTPRIC2.Text = "Price List Details"
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        '  grdSOTPRIC2.Enabled = tf

        Absx1.txtFor("CURR_CODE").Enabled = (EntryMode = "New")

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTPRIC2}
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

#Region "grdSOTPRIC2"
    Private Sub grdSOTPRIC2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPRIC2.AfterCellUpdate
        If e.Cell Is Nothing OrElse e.Cell.Value Is Nothing Then Exit Sub
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", e.Cell.Value)
                If rowICTSTYL1 IsNot Nothing Then
                    e.Cell.Row.Cells("STYLE_DESC").Value = rowICTSTYL1("STYLE_DESC")
                    e.Cell.Row.Cells("STYLE_PRICE").Value = rowICTSTYL1("STYLE_PRICE")
                End If
        End Select
    End Sub

    Private Sub grdSOTPRIC2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPRIC2.AfterRowActivate

        With grdSOTPRIC2.DisplayLayout.Bands(0).Columns("STYLE_CODE")
            If grdSOTPRIC2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

    End Sub

    Private Sub grdSOTPRIC2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTPRIC2.AfterRowsDeleted

    End Sub

    Private Sub grdSOTPRIC2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPRIC2.AfterRowUpdate

    End Sub

    Private Sub grdSOTPRIC2_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTPRIC2.BeforeExitEditMode
        If grdSOTPRIC2.ActiveCell IsNot Nothing Then
            With grdSOTPRIC2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE"
                        If .EditorResolved.IsValid Then
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTPRIC2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTPRIC2.BeforeRowsDeleted

    End Sub

    Private Sub grdSOTPRIC2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPRIC2.BeforeRowUpdate
        Dim row As DataRow = LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("PRICE_LIST_CODE").Value = Absx1.txtFor("PRICE_LIST_CODE").Text
        End If

    End Sub

    Private Sub grdSOTPRIC2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPRIC2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim sql_where As String = "" ' Get_List_of_Codes("ICTSTYL1.STYLE_CODE not in", "SOTPRIC2", "STYLE_CODE") ORA-01795 because > 1000
                grdClickCellButton(grdSOTPRIC2, sql_where, True)
        End Select
    End Sub

    Private Sub grdSOTPRIC2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPRIC2.InitializeRow
        grd_RowColor(dst.Tables("SOTPRIC2"), e.Row)
    End Sub
#End Region

End Class