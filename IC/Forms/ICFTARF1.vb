Imports System.ComponentModel
Imports Infragistics.Win.UltraWinGrid

Public Class ICFTARF1

    Dim COUNTRY_CODE As String
    Dim rowICTTARF1 As DataRow
    Dim rowICTTARF2 As DataRow
    Dim rowTATCNTRY As DataRow
    Dim addMode As Boolean = False
    Dim newCountryTariff As Boolean = False
    Dim insertOnUpdate As Boolean = False
    'AUDITING
    'sqlsort in cmdtest

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ICTTARF1.*, TATCNTRY.COUNTRY_NAME 
                from ICTTARF1, TATCNTRY WHERE ICTTARF1.COUNTRY_CODE = TATCNTRY.COUNTRY_CODE"
            Create_TDA(.Tables.Add, "ICTTARF1", "**", 0, True, "",, "TARIFF_ACTIVE")
            With .Tables("ICTTARF1").Columns
                .Add("EFFECTIVE_TARIFF_PCT", GetType(System.Decimal))
                .Add("DATE_TO_USE", GetType(System.String))
                .Add("TARIFF_NOTES", GetType(System.String))
            End With

            ASCMAIN1.sql = "Select * from ICTTARF2"
            Create_TDA(.Tables.Add, "ICTTARF2", "**", 0, True, "")

            ASCMAIN1.sql = "select ICTDUTY4.DUTY_RATE, DUTY_RATE_ADD, DUTY_RATE_BEGIN, DUTY_RATE_END, COUNT (*) RECS
                , MIN (DUTY_RATE_CODE) MINHTSUS
                , MAX (DUTY_RATE_CODE) MAXHTSUS
                from ICTDUTY4 where country_code = :PARM1 AND DUTY_RATE_END IS NULL
                GROUP BY ICTDUTY4.DUTY_RATE, DUTY_RATE_ADD, DUTY_RATE_BEGIN, DUTY_RATE_END"
            Create_TDA(.Tables.Add, "ICTDUTY4", "**", 0, False, "V")

        End With

        grdICTTARFX.DataSource = dst.Tables("ICTTARF1")
        grdICTTARF2.DataSource = dst.Tables("ICTTARF2")
        grdICTDUTY4.DataSource = dst.Tables("ICTDUTY4")

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
        dteTARIFF_START.Nullable = True
        dteTARIFF_END.Nullable = True
        Load_ICTTARF1()
        optTarriffStatus.Value = "A"
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
                addMode = False
                If dst.Tables("ICTTARF2").Rows.Count = 0 Then
                    dteTARIFF_START.Value = Now
                    dteTARIFF_START.Enabled = False
                    numTARIFF_PCT.Value = 0
                    numTARIFF_PCT.Enabled = False
                    dteTARIFF_END.Enabled = False
                    txtTARIFF_NOTES.Value = ""
                    txtTARIFF_NOTES.Enabled = False
                    optTARIFF_DATE_TO_USE.Enabled = False
                End If
            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)
                addMode = False
                If dst.Tables("ICTTARF2").Rows.Count = 0 Then
                    addMode = True
                    dteTARIFF_START.Value = Now
                    dteTARIFF_START.Enabled = True
                    numTARIFF_PCT.Value = 0
                    numTARIFF_PCT.Enabled = True
                    dteTARIFF_END.Enabled = False
                    txtTARIFF_NOTES.Value = ""
                    txtTARIFF_NOTES.Enabled = True
                    optTARIFF_DATE_TO_USE.Enabled = True
                    Set_Add_Mode_Controls()
                End If
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

        grdICTTARFX.Visible = Not ScreenMode
        splICTTARF1.Visible = ScreenMode
        Set_Add_Mode_Controls()

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        rowTATCNTRY = LookUp("TATCNTRY", COUNTRY_CODE)
        newCountryTariff = dst.Tables("ICTTARF1").Select($"COUNTRY_CODE = '{COUNTRY_CODE}'").Length = 0

        If Not newCountryTariff Then
            rowICTTARF1 = dst.Tables("ICTTARF1").Select($"COUNTRY_CODE = '{COUNTRY_CODE}'")(0)
        Else
            rowICTTARF1 = dst.Tables("ICTTARF1").NewRow
            rowICTTARF1("COUNTRY_CODE") = COUNTRY_CODE
            rowICTTARF1("TARIFF_ACTIVE") = "1"
            rowICTTARF1("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTTARF1("INIT_DATE") = Now
            dst.Tables("ICTTARF1").Rows.Add(rowICTTARF1)
            insertOnUpdate = True
        End If

        chkTARIFF_ACTIVE.Checked = rowICTTARF1("TARIFF_ACTIVE") & "" = "1"

        EnforceConstraints(False)

        ASCMAIN1.sql = $"Select * from ICTTARF2 where COUNTRY_CODE = '{COUNTRY_CODE}'"
        Fill_Records("ICTTARF2", "", , ASCMAIN1.sql)
        Fill_Records("ICTDUTY4", COUNTRY_CODE)

        Sort_grdColumns(grdICTDUTY4, "DUTY_RATE_BEGIN".ToLower)
        Sort_grdColumns(grdICTTARF2, "TARIFF_START".ToLower)

        EnforceConstraints(True)

        grdICTTARF2.Text = $"Tariff Schedule for: {rowTATCNTRY("COUNTRY_NAME")} "
        grdICTDUTY4.Text = $"Duty Rate Modifiers for: {rowTATCNTRY("COUNTRY_NAME")} "

        SplitContainer1.Panel2Collapsed = (dst.Tables("ICTDUTY4").Rows.Count = 0)
        chkTARIFF_ACTIVE.Enabled = (EntryMode = "E")

        Set_Add_Mode_Controls()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("ICTTARF1").Rows.Clear()
        dst.Tables("ICTTARF2").Rows.Clear()
        dst.Tables("ICTDUTY4").Rows.Clear()
        EnforceConstraints(True)

        Load_ICTTARF1()
        Absx1.txtFor("COUNTRY_CODE").Text = ""
        insertOnUpdate = False
        newCountryTariff = False
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Tariff Schedule")
        Dim LAST_DATE As DateTime = Now
        Dim TARIFF_ACTIVE As String = If(chkTARIFF_ACTIVE.Checked, "1", "0")
        Dim TARIFF_ACTIVE_curr As String = rowICTTARF1("TARIFF_ACTIVE") & ""
        Dim TARIFF_NOTES As String = txtTARIFF_NOTES.Value & ""
        Dim TARIFF_PCT As Decimal = numTARIFF_PCT.Value

        If rowICTTARF2 IsNot Nothing Then
            Dim INIT_OPER As String = rowICTTARF2("INIT_OPER") & ""
            Dim startDate As Date = dteTARIFF_START.Value
            If INIT_OPER = "" Then
                rowICTTARF2("INIT_OPER") = ASCMAIN1.USER_ID
                rowICTTARF2("INIT_DATE") = LAST_DATE
            End If
            rowICTTARF2("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTTARF2("LAST_DATE") = LAST_DATE
            rowICTTARF2("TARIFF_DATE_TO_USE") = optTARIFF_DATE_TO_USE.Value
            rowICTTARF2("TARIFF_NOTES") = TARIFF_NOTES
            rowICTTARF2("TARIFF_PCT") = TARIFF_PCT
            rowICTTARF2("TARIFF_START") = startDate
        End If

        For Each rowICTTARF2_add As DataRow In dst.Tables("ICTTARF2").Select($"COUNTRY_CODE = '{COUNTRY_CODE}'")
            rowICTTARF2_add.AcceptChanges()
            rowICTTARF2_add.SetAdded()
        Next

        BeginTrans()

        If insertOnUpdate Then
            ASCMAIN1.sql = $"Insert into ICTTARF1 (COUNTRY_CODE, TARIFF_ACTIVE, INIT_DATE, INIT_OPER, LAST_DATE, LAST_OPER)
                        VALUES ('{COUNTRY_CODE}','{TARIFF_ACTIVE}',SYSDATE,'{ASCMAIN1.USER_ID}',SYSDATE,'{ASCMAIN1.USER_ID}')"
            ASCDATA1.ExecuteSQL()
        Else
            If TARIFF_ACTIVE <> TARIFF_ACTIVE_curr Then
                ASCMAIN1.sql = "Update ICTTARF1 Set TARIFF_ACTIVE = :PARM1, LAST_OPER = :PARM2, LAST_DATE = SYSDATE where COUNTRY_CODE = :PARM3"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {TARIFF_ACTIVE, ASCMAIN1.USER_ID, COUNTRY_CODE})
            End If
        End If

        Update_Record_TDA("ICTTARF2", $"COUNTRY_CODE = '{COUNTRY_CODE}'")

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

    Sub Set_Add_Mode_Controls()
        cmdUpdate.Visible = addMode
        cmdCancel.Visible = addMode
        cmdAdd.Visible = (EntryMode = "E" And Not addMode)
        If addMode Then
            dteTARIFF_START.MinDate = Now.Date
        Else
            Dim minTariffStart As Object = dst.Tables("ICTTARF2").Compute("MIN(TARIFF_START)", String.Empty)

            ' Protect against DBNull
            If minTariffStart IsNot DBNull.Value Then
                dteTARIFF_START.MinDate = CDate(minTariffStart)

            End If

        End If
    End Sub
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
            Dim COUNTRY_CODE As String = rowICTTARF1("COUNTRY_CODE") & ""
            If TARIFF_ACTIVE = "1" Then
                For Each rowICTTARF2 As DataRow In tblICTTARF2.Select($"COUNTRY_CODE = '{COUNTRY_CODE}'")
                    Dim startDate As Date = rowICTTARF2.Field(Of Date)("TARIFF_START")
                    Dim endDate As Nullable(Of Date) = If(rowICTTARF2.IsNull("TARIFF_END"), CType(Nothing, Nullable(Of Date)), rowICTTARF2.Field(Of Date)("TARIFF_END"))

                    ' Check if today's date is within the range
                    If today.Date >= startDate.Date AndAlso (Not endDate.HasValue OrElse today <= endDate.Value) Then
                        rowICTTARF1("EFFECTIVE_TARIFF_PCT") = Val(rowICTTARF2("TARIFF_PCT"))
                        rowICTTARF1("DATE_TO_USE") = rowICTTARF2("TARIFF_DATE_TO_USE")
                        rowICTTARF1("TARIFF_NOTES") = rowICTTARF2("TARIFF_NOTES") & ""
                        Exit For ' Only apply the first matching range, that there is only one match is enforced at entry
                    End If
                Next
            Else
                rowICTTARF1("EFFECTIVE_TARIFF_PCT") = 0
                rowICTTARF1("DATE_TO_USE") = "N"
            End If

        Next
        Filter_Tariff_By_Status()
    End Sub

    Private Sub optTarriffStatus_ValueChanged(sender As Object, e As EventArgs) Handles optTarriffStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Filter_Tariff_By_Status()
    End Sub

    Sub Filter_Tariff_By_Status()

        Dim dvwICTTARFX As DataView = DirectCast(grdICTTARFX.DataSource, DataTable).DefaultView
        If optTarriffStatus.Value = "A" Then
            dvwICTTARFX.RowFilter = $"TARIFF_ACTIVE = '1'"
        Else
            dvwICTTARFX.RowFilter = ""
        End If

    End Sub

    Private Sub grdICTTARFX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdICTTARFX.DoubleClickRow
        If grdICTTARFX.ActiveRow IsNot Nothing AndAlso grdICTTARFX.ActiveRow.IsDataRow Then
            Absx1.txtFor("COUNTRY_CODE").Text = grdICTTARFX.ActiveRow.Cells("COUNTRY_CODE").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTTARF2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTTARF2.AfterRowActivate

        If Not addMode Then
            rowICTTARF2 = GetDataRowFromUltraGridRow(grdICTTARF2.ActiveRow)
            Dim minTariffStart As Object = dst.Tables("ICTTARF2").Compute("MIN(TARIFF_START)", String.Empty)
            Dim minTariffEnd As Object = dst.Tables("ICTTARF2").Compute("MIN(TARIFF_END)", String.Empty)

            If minTariffStart IsNot DBNull.Value Then
                dteTARIFF_START.MinDate = CDate(minTariffStart)
            End If
            If minTariffEnd IsNot DBNull.Value Then
                dteTARIFF_END.MinDate = CDate(minTariffEnd)
            End If
            With grdICTTARF2.ActiveRow
                Dim TARIFF_PCT As Decimal = If(IsDBNull(.Cells("TARIFF_PCT").Value), 0D, Convert.ToDecimal(.Cells("TARIFF_PCT").Value))
                Dim TARIFF_START As Date? = If(IsDBNull(.Cells("TARIFF_START").Value), Nothing, CType(.Cells("TARIFF_START").Value, Date))
                Dim tariffHasEndDate As Boolean = Not IsDBNull(.Cells("TARIFF_END").Value)
                Dim TARIFF_END As Date? = If(tariffHasEndDate, .Cells("TARIFF_END").Value, Nothing)
                Dim TARIFF_DATE_TO_USE As String = .Cells("TARIFF_DATE_TO_USE").Value.ToString()
                Dim TARIFF_NOTES As String = .Cells("TARIFF_NOTES").Value.ToString()

                numTARIFF_PCT.Value = TARIFF_PCT
                dteTARIFF_START.Value = If(TARIFF_START.HasValue, TARIFF_START.Value, DBNull.Value)
                dteTARIFF_END.Value = If(TARIFF_END.HasValue, TARIFF_END.Value, DBNull.Value)
                optTARIFF_DATE_TO_USE.Value = TARIFF_DATE_TO_USE
                txtTARIFF_NOTES.Value = TARIFF_NOTES

                Dim allowEdit As Boolean = (Not TARIFF_END.HasValue Or (TARIFF_END.HasValue AndAlso TARIFF_END >= Date.Today))
                Dim hasNotStartedYet As Boolean = dteTARIFF_START.Value >= Now.Date
                numTARIFF_PCT.Enabled = (EntryMode = "E" And allowEdit And hasNotStartedYet)
                dteTARIFF_END.Enabled = (EntryMode = "E" And allowEdit)
                dteTARIFF_START.Enabled = (EntryMode = "E" And allowEdit And hasNotStartedYet)
                optTARIFF_DATE_TO_USE.Enabled = (EntryMode = "E" And allowEdit And hasNotStartedYet)
                txtTARIFF_NOTES.Enabled = (EntryMode = "E" And allowEdit)
                chkTARIFF_ACTIVE.Enabled = (EntryMode = "E")
            End With
        End If


    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        addMode = False
        Dim hasNullTariffEnd As Boolean = dst.Tables("ICTTARF2").Select("TARIFF_END IS NULL").Length > 0
        If hasNullTariffEnd Then
            Dim rowOld As DataRow = dst.Tables("ICTTARF2").Select("TARIFF_END IS NULL")(0)
            Dim dayBefore As Date = dteTARIFF_START.Value.Date.AddDays(-1)
            rowOld("TARIFF_END") = dayBefore
        End If

        Dim rowICTTARF2 As DataRow = dst.Tables("ICTTARF2").NewRow
        rowICTTARF2("COUNTRY_CODE") = COUNTRY_CODE
        rowICTTARF2("TARIFF_PCT") = numTARIFF_PCT.Value
        rowICTTARF2("TARIFF_START") = dteTARIFF_START.Value
        rowICTTARF2("TARIFF_DATE_TO_USE") = optTARIFF_DATE_TO_USE.Value
        rowICTTARF2("TARIFF_NOTES") = txtTARIFF_NOTES.Value & ""
        rowICTTARF2("INIT_DATE") = Now
        rowICTTARF2("INIT_OPER") = ASCMAIN1.USER_ID
        dst.Tables("ICTTARF2").Rows.Add(rowICTTARF2)
        Sort_grdColumns(grdICTTARF2, "TARIFF_START".ToLower)
        Set_Add_Mode_Controls()
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        addMode = False
        Set_Add_Mode_Controls()
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        addMode = True
        If dteTARIFF_START.MinDate < Now.Date Then
            dteTARIFF_START.MinDate = Now.Date
        End If
        dteTARIFF_START.Value = Now
        dteTARIFF_START.Enabled = True
        numTARIFF_PCT.Value = 0
        numTARIFF_PCT.Enabled = True
        dteTARIFF_END.Enabled = False
        txtTARIFF_NOTES.Value = ""
        txtTARIFF_NOTES.Enabled = True
        optTARIFF_DATE_TO_USE.Enabled = True
        Set_Add_Mode_Controls()
    End Sub
    Private Function GetDataRowFromUltraGridRow(row As Infragistics.Win.UltraWinGrid.UltraGridRow) As DataRow
        If row Is Nothing OrElse row.ListObject Is Nothing Then
            Return Nothing
        End If

        Dim drv As DataRowView = TryCast(row.ListObject, DataRowView)
        If drv IsNot Nothing Then
            Return drv.Row
        End If

        Return Nothing
    End Function

    Private Sub dteTARIFF_END_BeforeDropDown(sender As Object, e As CancelEventArgs) Handles dteTARIFF_END.BeforeDropDown
        dteTARIFF_END.MinDate = CDate(dteTARIFF_START.Value).AddDays(1)
        If Not addMode And grdICTTARF2.ActiveRow IsNot Nothing Then
            Dim tariffHasEndDate As Boolean = Not IsDBNull(grdICTTARF2.ActiveRow.Cells("TARIFF_END").Value)
            If Not tariffHasEndDate Then
                dteTARIFF_END.MinDate = Now.Date
            End If
        End If
    End Sub

    Private Sub dteTARIFF_START_BeforeDropDown(sender As Object, e As CancelEventArgs) Handles dteTARIFF_START.BeforeDropDown
        Dim maxTariffStart As Object = dst.Tables("ICTTARF2").Compute("MAX(TARIFF_START)", String.Empty)

        ' Protect against DBNull
        If maxTariffStart IsNot DBNull.Value Then
            Dim maxDate As Date = CDate(maxTariffStart)
            If maxDate > Now.Date Then
                dteTARIFF_START.MinDate = maxDate
            Else
                dteTARIFF_START.MinDate = Now.Date
            End If
        Else
            dteTARIFF_START.MinDate = Now.Date
        End If
        dteTARIFF_START.MinDate = Now.Date
    End Sub
End Class