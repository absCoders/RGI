Public Class ICFTARF1

    Dim COUNTRY_CODE As String

    'AUDITING
    'sqlsort in cmdtest

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select * from ICTTARF1"
            Create_TDA(.Tables.Add, "ICTTARF1", "**", 0, , "")
            With .Tables("ICTTARF1").Columns
                .Add("EFFECTIVE_TARIFF_PCT", GetType(System.Decimal))
                .Add("DATE_TO_USE", GetType(System.String))
            End With

            ASCMAIN1.sql = "Select * from ICTTARF2"
            Create_TDA(.Tables.Add, "ICTTARF2", "**", 0, , "")

        End With

        Load_ICTTARF1()

        grdICTTARFX.DataSource = dst.Tables("ICTTARF1")
        grdICTTARF2.DataSource = dst.Tables("ICTTARF2")

        Create_Summary(grdICTTARFX, "COUNTRY_CODE", "Count")

        With grdICTTARF2.DisplayLayout.Bands(0)
            '    For Each COLUMN_NAME In New String() {"PROD_CODE", "PROD_DESC", "BRAND_CODE"}
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Goldenrod
            '        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    Next
            '    For Each COLUMN_NAME In New String() {"COMMISSION", "COMM_TYPE"}
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            '        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    Next
            '    For Each COLUMN_NAME In New String() {"INSPECTION", "INSP_TYPE"}
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Green
            '        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    Next
            '    For Each COLUMN_NAME In New String() {"OTHER", "OTHER_TYPE"}
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            '        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    Next
            'End With

            'With grdICTCOSTD.DisplayLayout.Bands(0)
            '    For Each COLUMN_NAME In New String() {"VEND_CODE", "VEND_NAME"}
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Goldenrod
            '        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    Next
            '    For Each COLUMN_NAME In New String() {"PAYEE_COMMISSION"}
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            '        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    Next
            '    For Each COLUMN_NAME In New String() {"PAYEE_INSPECTION"}
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Green
            '        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    Next
            '    For Each COLUMN_NAME In New String() {"PAYEE_OTHER"}
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            '        .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    Next
        End With

        ASCMAIN1.Add_Value_List(grdICTTARFX, "DATE_TO_USE", Nothing, New String() {":", "S:Ship Date", "R:Receive Date", "N:Not Active"})
        ASCMAIN1.Add_Value_List(grdICTTARF2, "TARIFF_DATE_TO_USE", Nothing, New String() {":", "S:Ship Date", "R:Receive Date"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View", "Edit"
                If Validate_Code("COUNTRY_CODE") Then
                    If eItemKey = "Edit" Then
                        If Not ASCMAIN1.Logical_Lock("ICTTARF1", Absx1.txtFor("COUNTRY_CODE").Text) Then Exit Sub
                    End If

                    COUNTRY_CODE = Absx1.txtFor("COUNTRY_CODE").Text
                    Dim rowTATCNTRY As DataRow = LookUp("TATCNTRY", COUNTRY_CODE)
                    If rowTATCNTRY Is Nothing Then
                        EMsg &= vbCr & "Invalid Country Code (" & COUNTRY_CODE & ")"
                    Else

                    End If
                End If

            Case "Update"

                'For Each COLUMN_NAME As String In New String() {"VEND_CODE", "PAYEE_COMMISSION", "PAYEE_INSPECTION", "PAYEE_OTHER"}
                '    For Each row As DataRow In ASCDATA1.SelectDistinct("ICTCOSTD", COLUMN_NAME).Rows
                '        Dim VEND_CODE As String = row.Item(0) & ""
                '        If COLUMN_NAME = "VEND_CODE" And VEND_CODE = "Z" Then
                '        Else
                '            If VEND_CODE <> "" Then
                '                If LookUp("APTVEND1", VEND_CODE) Is Nothing Then
                '                    EMsg &= vbCr & "Invalide Value Specified for " & grdICTCOSTD.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption & " (" & VEND_CODE & ")"
                '                End If
                '            End If
                '        End If
                '        If COLUMN_NAME = "VEND_CODE" Then
                '            Dim rows() As DataRow = dst.Tables("ICTCOSTC").Select _
                '                ("ORIG_CODE = '" & ORIG_CODE & "' and VEND_CODE = '" & VEND_CODE & "'")
                '            If rows.Length = 0 Then
                '                EMsg &= vbCr & "No Commission details defined for Supplier " & VEND_CODE & "; Delete or Complete"
                '            Else
                '                Dim rowICTCOSTD As DataRow = dst.Tables("ICTCOSTD").Rows.Find(New String() {ORIG_CODE, VEND_CODE})
                '                For Each rowICTCOSTC As DataRow In rows
                '                    Dim BRAND_CODE As String = rowICTCOSTC.Item("BRAND_CODE") & ""
                '                    If BRAND_CODE <> "Z" Then
                '                        If LookUp("ICTBRAN1", BRAND_CODE) Is Nothing Then
                '                            EMsg &= vbCr & "Invalid Value Specified for Brand Code (See Supplier " & VEND_CODE & ")"
                '                        End If
                '                    End If
                '                    Dim PROD_CODE As String = rowICTCOSTC.Item("PROD_CODE") & ""
                '                    If PROD_CODE <> "ZZZ" Then
                '                        If LookUp("ICTPROD1", PROD_CODE) Is Nothing Then
                '                            EMsg &= vbCr & "Invalid Value Specified for Product Code (See Supplier " & VEND_CODE & ")"
                '                        End If
                '                    End If


                '                    For Each COLUMN_NAME2 As String In New String() {"COMMISSION", "INSPECTION", "OTHER"}
                '                        Dim COLUMN_NAME3 As String = "PAYEE_" & COLUMN_NAME2

                '                        If Val(rowICTCOSTC.Item(COLUMN_NAME2) & "") = 0 Then
                '                            If rowICTCOSTD.Item(COLUMN_NAME3) & "" <> "" Then
                '                                EMsg &= vbCr & grdICTCOSTD.DisplayLayout.Bands(0).Columns(COLUMN_NAME3).Header.Caption & " Specified, but no value specified for " & grdICTTARF2.DisplayLayout.Bands(0).Columns(COLUMN_NAME2).Header.Caption & " (See Supplier " & VEND_CODE & ")"
                '                            End If
                '                        Else
                '                            If rowICTCOSTD.Item(COLUMN_NAME3) & "" = "" Then
                '                                EMsg &= vbCr & "A Value was specified for " & grdICTTARF2.DisplayLayout.Bands(0).Columns(COLUMN_NAME2).Header.Caption & ", but no Payee specified for " & grdICTCOSTD.DisplayLayout.Bands(0).Columns(COLUMN_NAME3).Header.Caption & " (See Supplier " & VEND_CODE & ")"
                '                            End If
                '                        End If
                '                    Next
                '                Next
                '            End If
                '        End If
                '    Next
                'Next

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
            If ScreenMode And EntryMode = "V" Then
                .Groups("Screen Control").Items("Edit").Settings.Enabled = DefaultableBoolean.True
            Else
                .Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            End If
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

            .Groups("Screen Control").Items("Update").Visible = (EntryMode = "E")
            .Groups("Screen Control").Items("Cancel").Visible = (EntryMode = "E")
            .Groups("Screen Control").Items("Done").Visible = (EntryMode = "V")

            .Groups("Showing").Visible = Not ScreenMode

        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode And EntryMode = "E" Then
            'With grdICTCOSTD.DisplayLayout.Override
            '    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            '    .AllowDelete = DefaultableBoolean.True
            '    .AllowUpdate = DefaultableBoolean.True
            '    .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            'End With
            'With grdICTTARF2.DisplayLayout.Override
            '    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            '    .AllowDelete = DefaultableBoolean.True
            '    .AllowUpdate = DefaultableBoolean.True
            '    .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            'End With
        Else
            'With grdICTCOSTD.DisplayLayout.Override
            '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '    .AllowDelete = DefaultableBoolean.False
            '    .AllowUpdate = DefaultableBoolean.False
            '    .CellClickAction = UltraWinGrid.CellClickAction.Default
            'End With
            'With grdICTTARF2.DisplayLayout.Override
            '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '    .AllowDelete = DefaultableBoolean.False
            '    .AllowUpdate = DefaultableBoolean.False
            '    .CellClickAction = UltraWinGrid.CellClickAction.Default
            'End With
        End If

        grdICTTARFX.Visible = Not ScreenMode
        splICTTARF1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        Dim tblICTTARF2 As DataTable = dst.Tables("ICTTARF2")
        Dim dvw As DataView = tblICTTARF2.DefaultView
        dvw.RowFilter = $"COUNTRY_CODE = '{COUNTRY_CODE}'"
        grdICTTARF2.DataSource = tblICTTARF2

        EnforceConstraints(False)
        'Fill_Records("ICTCOSTD", ORIG_CODE)
        'Sort_grdColumns(grdICTCOSTD, "VEND_CODE")
        'Fill_Records("ICTCOSTC", ORIG_CODE)
        'Sort_grdColumns(grdICTTARF2, "BRAND_CODE,PROD_CODE")
        EnforceConstraints(True)

        grdICTTARF2.Text = $"Tariff Schedule for: {COUNTRY_CODE} "

        'sort by date
        'select current (or most recent)
        'fill form controls with row values
        'Setup_ICTCOSTC()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("ICTTARF1").Rows.Clear()
        dst.Tables("ICTTARF2").Rows.Clear()
        EnforceConstraints(True)

        Load_ICTTARF1()
        Absx1.txtFor("COUNTRY_CODE").Text = ""
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Order")

        BeginTrans()

        Dim sql_delete As String = "COUNTRY_CODE = '" & COUNTRY_CODE & "'"
        Update_Record_TDA("ICTTARF1", "")
        Update_Record_TDA("ICTCOSTD", sql_delete)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdICTCOSTE, "SS", "Show Filter", "Show GroupBox")
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

        If grd.Name = "grd" Then
            Exit Sub
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""
                    'If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                    '    e.Cancel = True
                    'End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        If COLUMN_NAME = "COUNTRY_CODE" Then
            If ctl.Text <> "" Then
                'Call Click_Command("Load Reports")
            End If
        End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "COUNTRY_CODE" Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then
                Click_Command("View")
            End If
        End If
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)

        Select Case Absx1.GetABSColumnName(txtctl)
            Case "COUNTRY_CODE"
                If txtctl.Text <> "" Then
                    Click_Command("View")
                End If
        End Select
    End Sub

#End Region

    Sub Load_ICTTARF1()
        Fill_Records("ICTTARF1")
        Fill_Records("ICTTARF2")
        Calculate_Effective_Tariff_Pct()
    End Sub

    Sub Calculate_Effective_Tariff_Pct()
        Dim today As Date = Date.Today
        Dim tblICTTARF1 As DataTable = dst.Tables("ICTTARF1")
        Dim tblICTTARF2 As DataTable = dst.Tables("ICTTARF2")

        For Each rowICTTARF1 As DataRow In tblICTTARF1.Rows
            Dim TARIFF_ACTIVE As String = rowICTTARF1("TARIFF_ACTIVE") & ""
            If TARIFF_ACTIVE = "1" Then
                Dim COUNTRY_CODE As String = rowICTTARF1("COUNTRY_CODE") & ""
                For Each rowICTTARF2 As DataRow In tblICTTARF2.Select($"COUNTRY_CODE = '{COUNTRY_CODE}'")
                    Dim startDate As Date = rowICTTARF2.Field(Of Date)("TARIFF_START")
                    Dim endDate As Nullable(Of Date) = If(rowICTTARF2.IsNull("TARIFF_END"), CType(Nothing, Nullable(Of Date)), rowICTTARF2.Field(Of Date)("TARIFF_END"))

                    ' Check if today's date is within the range
                    If today >= startDate AndAlso (Not endDate.HasValue OrElse today <= endDate.Value) Then
                        rowICTTARF1("EFFECTIVE_TARIFF_PCT") = Val(rowICTTARF2("TARIFF_PCT"))
                        rowICTTARF1("DATE_TO_USE") = rowICTTARF2("TARIFF_DATE_TO_USE")
                        Exit For ' Only apply the first matching range, that there is only one match is enforced at entry
                    End If
                Next
            Else
                rowICTTARF1("EFFECTIVE_TARIFF_PCT") = 0
                rowICTTARF1("DATE_TO_USE") = "N"
            End If

        Next
    End Sub

End Class