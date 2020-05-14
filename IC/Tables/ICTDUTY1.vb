Public Class ICTDUTY1
    Dim sqlICTDUTY3 As String = ""
    Dim sqlICTDUTY4 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ' this table is used for the Report
            ASCMAIN1.sql = "Select ICTDUTY1.* From ICTDUTY1"
            Create_TDA(.Tables.Add, "ICTDUTYX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTDUTY3.* " _
            & " from ICTDUTY3" _
            & " where ICTDUTY3.DUTY_RATE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTDUTY3", "**", 0, True, "V", 2)


            sqlICTDUTY4 = "Select ICTDUTY4.*, ICTDUTY1.DUTY_RATE_DESC" _
             & " from ICTDUTY1,ICTDUTY4" _
             & " where ICTDUTY1.DUTY_RATE_CODE = ICTDUTY4.DUTY_RATE_CODE"
            ASCMAIN1.sql = sqlICTDUTY4 _
            & "  and ICTDUTY4.DUTY_RATE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTDUTY4", "**", 0, True, "V")

            ASCMAIN1.sql = sqlICTDUTY4
            Create_TDA(.Tables.Add, "ICTDUTY4X", "**", 0, False, "")
            Create_Relation("ICTDUTYX", "ICTDUTY4X", "DUTY_RATE_CODE")

        End With

        '   grdICTDUTYX.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy
        grdICTDUTYX.DataSource = dst.Tables("ICTDUTYX")
        Show_Filter(grdICTDUTYX, True)
        With grdICTDUTYX.DisplayLayout.Bands(0)
            .Columns("DUTY_RATE_CODE").Header.Fixed = True
            .Columns("DUTY_RATE_DESC").Header.Fixed = True
            .Columns("COUNTRY_CODE").Header.Fixed = True
            .Columns("DUTY_RATE").Header.Fixed = True
        End With

        grdICTDUTY3.DataSource = dst.Tables("ICTDUTY3")
        grdICTDUTY4.DataSource = dst.Tables("ICTDUTY4")
        grdICTDUTY4.Visible = (ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA")
        tabDuty.Tabs("Report").Visible = (ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA")
        Create_Summary(grdICTDUTY3, "OPS_YYYY", "Count")

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTDUTY3, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTDUTY4, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdICTDUTYX, "SS", "Show Filter", "Show GroupBox")
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
            Case "grdICTDUTY4"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select

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

        Select Case e.Tool.Key
            Case "Add Codes"
                If grd.Name = "grdICTDUTY4" Then
                    Add_Codes(grdICTDUTY4, "TATCNTRY", "COUNTRY_CODE", "Country Codes")
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
                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    Dim DUTY_RATE_CODE As String = Absx1.txtFor("DUTY_RATE_CODE").Text
                    If (Len(DUTY_RATE_CODE) <> 12 And Len(DUTY_RATE_CODE) <> 16) Or Mid(DUTY_RATE_CODE, 5, 1) <> "." Or Mid(DUTY_RATE_CODE, 8, 1) <> "." _
                        Or InStr(DUTY_RATE_CODE, " ") <> 0 _
                        Or (Mid(DUTY_RATE_CODE, 13, 1) <> "" And Mid(DUTY_RATE_CODE, 13, 1) <> "-") _
                        Or Not IsNumeric(Mid(DUTY_RATE_CODE, 1, 4)) Or Not IsNumeric(Mid(DUTY_RATE_CODE, 6, 2)) Or Not IsNumeric(Mid(DUTY_RATE_CODE, 9, 4)) Then
                        EMsg &= vbCr & "Duty Rate Code format is 9999.99.9999-CCC (-CCC for Optional Country Code)"
                    ElseIf Len(DUTY_RATE_CODE) = 16 Then
                        Dim COUNTRY_CODE As String = Mid(DUTY_RATE_CODE, 14)
                        If COUNTRY_CODE = "USA" Then
                            EMsg &= vbCr & "No Need for USA-specific Duty Rate Suffix (the default is USA)"
                        End If
                        If LookUp("TATCNTRY", COUNTRY_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Country Code (" & COUNTRY_CODE & ")"
                        End If
                    End If
                End If

            Case "Edit"

            Case "Update"
                If ASCMAIN1.CLIENT = "NYA" Then
                    Dim DUTY_RATE_CODE As String = Absx1.txtFor("DUTY_RATE_CODE").Text

                    If DUTY_RATE_CODE <> Absx1.txtFor("DUTY_HTS_CODE").Text Then
                        EMsg &= vbCr & "Duty Rate Code does NOT Match HTS Code"
                    End If

                    If DUTY_RATE_CODE.Length <> 12 And DUTY_RATE_CODE.Length <> 16 Then
                        EMsg &= vbCr & "Duty Rate Code format is 9999.99.9999-CCC (-CCC for Optional Country Code)"
                    ElseIf DUTY_RATE_CODE.Length = 16 Then
                        If Mid(DUTY_RATE_CODE, 13, 1) <> "-" Then
                            EMsg &= vbCr & "Duty Rate Code Suffix must be a valid Country Code preceded by a dash (-CCC)"
                        ElseIf LookUp("TATCNTRY", Mid(DUTY_RATE_CODE, 14)) Is Nothing Then
                            EMsg &= vbCr & "Duty Rate Code Suffix is not a valid Country Code"
                        ElseIf Mid(DUTY_RATE_CODE, 14) <> Absx1.txtFor("COUNTRY_CODE").Text Then
                            EMsg &= vbCr & "Duty Rate Code Suffix does not match the Country Code"
                        End If
                    End If

                End If

                If ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA" Then
                    Sort_grdColumns(grdICTDUTY4, "COUNTRY_CODE,DUTY_RATE_BEGIN")
                    For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ICTDUTY4"), "COUNTRY_CODE").Select("", "COUNTRY_CODE")
                        Dim COUNTRY_CODE As String = row.Item(0)
                        Dim DUTY_RATE_BEGIN As Date = Nothing
                        Dim DUTY_RATE_END As Date = Nothing
                        Dim first_record As Boolean = True
                        Dim date_end_set As Boolean = False

                        For Each rowICTDUTY4 As DataRow In dst.Tables("ICTDUTY4").Select("COUNTRY_CODE ='" & COUNTRY_CODE & "'", "DUTY_RATE_BEGIN")
                            If rowICTDUTY4.Item("DUTY_RATE_BEGIN") & "" = "" Then
                                EMsg &= vbCr & "Duty Rate Begins Date must not be Blank - " & COUNTRY_CODE
                            Else
                                DUTY_RATE_BEGIN = rowICTDUTY4.Item("DUTY_RATE_BEGIN")
                                If date_end_set And Format(DUTY_RATE_BEGIN, "yyyyMMdd") <= Format(DUTY_RATE_END, "yyyyMMdd") Then
                                    EMsg &= vbCr & "Overlapping Periods Detected for " & COUNTRY_CODE & " " & Format(DUTY_RATE_BEGIN, "MM/dd/yyyy")
                                Else
                                    If rowICTDUTY4.Item("DUTY_RATE_END") & "" = "" Then
                                        DUTY_RATE_END = CDate("12/31/9999")
                                    Else
                                        DUTY_RATE_END = rowICTDUTY4.Item("DUTY_RATE_END")
                                    End If
                                    date_end_set = True
                                    If Format(DUTY_RATE_BEGIN, "yyyyMMdd") > Format(DUTY_RATE_END, "yyyyMMdd") Then
                                        EMsg &= vbCr & "Duty Rate may not End before it Begins - " & COUNTRY_CODE & " " & Format(DUTY_RATE_BEGIN, "MM/dd/yyyy") & "-" & Format(DUTY_RATE_END, "MM/dd/yyyy")
                                    End If
                                End If
                            End If
                            first_record = False
                        Next
                    Next
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = "DUTY_RATE_CODE = '" & Absx1.txtFor("DUTY_RATE_CODE").Text & "'"
        Update_Record_TDA("ICTDUTY3", sqlDelete)
        Update_Record_TDA("ICTDUTY4", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Dim DUTY_RATE_CODE As String = Absx1.txtFor("DUTY_RATE_CODE").Text
        Fill_Records("ICTDUTY3", New String() {DUTY_RATE_CODE})
        Sort_grdColumns(grdICTDUTY3, "OPS_YYYY".ToLower)
        Fill_Records("ICTDUTY4", New String() {DUTY_RATE_CODE})
        Sort_grdColumns(grdICTDUTY4, "COUNTRY_CODE,DUTY_RATE_BEGIN")


        If ASCMAIN1.CLIENT = "NYA" Then
            '  Set_Read_Only_for_ctl(Absx1.txtFor("COUNTRY_CODE"), True)
            If EntryMode = "New" Then
                If DUTY_RATE_CODE.Length = 12 Then
                    Absx1.txtFor("COUNTRY_CODE").Text = ""
                    Absx1.txtFor("DUTY_HTS_CODE").Text = DUTY_RATE_CODE '  Mid(DUTY_RATE_CODE, 1, 12)
                Else
                    Dim COUNTRY_CODE As String = Mid(DUTY_RATE_CODE, 14)
                    Absx1.txtFor("COUNTRY_CODE").Text = COUNTRY_CODE

                    Absx1.txtFor("DUTY_HTS_CODE").Text = DUTY_RATE_CODE '  Mid(DUTY_RATE_CODE, 1, 12)
                End If

            End If
        End If

        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ICTDUTY3").Rows.Clear()
            EnforceConstraints(True)

            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ICTDUTY3", "ICTDUTY4"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdICTDUTY3.Enabled = tf
        grdICTDUTY4.Enabled = tf

        If (ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "NYA") Then
            If tf Then
                tabDuty.Tabs("Report").Visible = False
            Else
                tabDuty.Tabs("Report").Visible = True
            End If
        End If

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTDUTY3, grdICTDUTY4}
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

        If ASCMAIN1.CLIENT = "NYA" Then
            Set_Read_Only_for_ctl(Absx1.txtFor("COUNTRY_CODE"), True)
            Set_Read_Only_for_ctl(Absx1.txtFor("DUTY_HTS_CODE"), True)
        End If


    End Sub

#End Region

#Region "grdICTDUTY3"
    Private Sub grdICTDUTY3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTDUTY3.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("DUTY_RATE_CODE").Value = Absx1.txtFor("DUTY_RATE_CODE").Text
        End If
    End Sub
#End Region


#Region "grdICTDUTY4"

    Private Sub grdICTDUTY4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTDUTY4.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "DUTY_RATE_CODE"
            '    grdCodeDesc(grdICTDUTY4, "ICTDUTY1", "DUTY_RATE_CODE", "DUTY_RATE_DESC")
        End Select
    End Sub

    Private Sub grdICTDUTY4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTDUTY4.BeforeRowUpdate

        Dim row As DataRow = LookUp("TATCNTRY", e.Row.Cells("COUNTRY_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If
 
        If e.Row.IsAddRow Then
            Dim DUTY_RATE_CODE As String = Absx1.txtFor("DUTY_RATE_CODE").Text
            e.Row.Cells("DUTY_RATE_CODE").Value = DUTY_RATE_CODE
        End If
    End Sub

    Private Sub grdICTDUTY4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTDUTY4.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "COUNTRY_CODE"
                Dim sql_where As String = Get_List_of_Codes("TATCNTRY.COUNTRY_CODE not in", "ICTDUTY4", "COUNTRY_CODE")
                grdClickCellButton(grdICTDUTY4, sql_where, True)
        End Select

    End Sub

    Private Sub grdICTDUTY4_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTDUTY4.InitializeRow

    End Sub

#End Region

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        EnforceConstraints(False)
        Fill_Records("ICTDUTYX")
        Fill_Records("ICTDUTY4X", , , sqlICTDUTY4)
        EnforceConstraints(True)

        Sort_grdColumns(grdICTDUTYX, "DUTY_RATE_CODE")
        grdICTDUTYX.Rows.ExpandAll(True)
    End Sub

    Private Sub grdICTDUTYX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTDUTYX.DoubleClickRow
        Try
            Dim DUTY_RATE_CODE As String = e.Row.Cells("DUTY_RATE_CODE").Value & ""
            Absx1.txtFor("DUTY_RATE_CODE").Text = DUTY_RATE_CODE
            Click_Command("Edit")
        Catch ex As Exception

        End Try


    End Sub

End Class