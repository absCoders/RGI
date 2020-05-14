Public Class SOTCSTP2

    Dim sqlSOTCSTP1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            sqlSOTCSTP1 = "Select SOTCSTP1.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from ICTCOLR1,ICTSTYL1,SOTCSTP1" _
            & " where ICTSTYL1.STYLE_CODE = SOTCSTP1.STYLE_CODE" _
            & "   and ICTCOLR1.COLOR_CODE = SOTCSTP1.COLOR_CODE"
            ASCMAIN1.sql = sqlSOTCSTP1 _
            & "  and SOTCSTP1.CUST_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "SOTCSTP1", "**", 0, True, "V", 3)
        End With

        grdSOTCSTP1.DataSource = dst.Tables("SOTCSTP1")

        With grdSOTCSTP1.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("COLOR_DESC").Header.Fixed = True
        End With
     
        Create_Summary(grdSOTCSTP1, "STYLE_CODE", "Count")
     

        'For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTCSTP1.DisplayLayout.Bands(0).Columns
        '    If gcol.Key <> "CUST_CODE" Then
        '        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'Next

        ' ASCMAIN1.Add_Value_List(grdSOTCSTP1, "STYLE_COLOR_STATUS", Nothing, New String() {":", "A:Active", "D:Discontinued", "N:Do Not Re-Order"})

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTCSTP1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
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

                'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If

                'Dim rowSOTSREP1 = LookUp("SOTSREP1", Absx1.txtFor("STYLE_CODE").Text)
                'If rowSOTSREP1 Is Nothing Then
                '    EMsg &= vbCr & "Invalid Value entered for Sales Rep Code"
                'End If


        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        'INIT_LAST("SOTCSTP1", True)
        Update_Record_TDA("SOTCSTP1", sqlDelete)
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
        Fill_Records("SOTCSTP1", New String() {Absx1.txtFor("CUST_CODE").Text})
        Sort_grdColumns(grdSOTCSTP1, "STYLE_CODE,COLOR_CODE")
        EnforceConstraints(True)

        grdSOTCSTP1.Text = "Style/Color Parameters for Customer " & Absx1.txtFor("CUST_CODE").Text
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTCSTP1").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSOTCSTP1.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTCSTP1}
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

#Region "grdSOTCSTP1"

    Private Sub grdSOTCSTP1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCSTP1.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                grdCodeDesc(grdSOTCSTP1, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
            Case "COLOR_CODE"
                grdCodeDesc(grdSOTCSTP1, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
        End Select
    End Sub

    Private Sub grdSOTCSTP1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCSTP1.BeforeRowUpdate
        If LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text) Is Nothing Then
            e.Cancel = True
        End If
        If LookUp("ICTCOLR1", e.Row.Cells("COLOR_CODE").Text) Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
        End If

    End Sub

    Private Sub grdSOTCSTP1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCSTP1.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdSOTCSTP1, sql_where, True)
            Case "COLOR_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdSOTCSTP1, sql_where, True)
        End Select

    End Sub

    Private Sub grdSOTCSTP1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCSTP1.InitializeRow
        grd_RowColor(dst.Tables("SOTCSTP1"), e.Row)
    End Sub

#End Region

End Class