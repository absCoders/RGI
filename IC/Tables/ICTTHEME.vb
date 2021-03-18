
Public Class ICTTHEME
    Dim sqlICTTHEMX As String
    Dim sqlThemeUsage As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            sqlThemeUsage = "SELECT ICTTHEME.THEME_DESC , COUNT(*) USAGE FROM ICTTHEME, ICTSTYC1 " _
                & " WHERE ICTSTYC1.THEME_CODE = ICTTHEME.THEME_CODE " _
                & " AND ICTSTYC1.THEME_CODE IN (SELECT THEME_CODE FROM ICTTHEME WHERE SEASON_CODE = ':SEASON_CODE')" _
                & " GROUP BY ICTTHEME.THEME_DESC"
            sqlICTTHEMX = "Select ICTTHEME.THEME_DESC, X.USAGE" & vbCrLf _
                & ", MIN(SEASON_CODE) SEASON_CODE_MIN, MAX(SEASON_CODE) SEASON_CODE_MAX, MIN(THEME_CODE) THEME_CODE_MIN, MAX(THEME_CODE) THEME_CODE_MAX" & vbCrLf _
                & " From ICTTHEME, (" & sqlThemeUsage & ") X " & vbCrLf _
                & " where X.THEME_DESC (+) = ICTTHEME.THEME_DESC"
            ASCMAIN1.sql = sqlICTTHEMX & " Group by ICTTHEME.THEME_DESC, X.USAGE"
            Create_TDA(.Tables.Add, "ICTTHEMX", "**", 0, False)
            .Tables("ICTTHEMX").PrimaryKey = New DataColumn() {.Tables("ICTTHEMX").Columns("THEME_DESC")}
            .Tables("ICTTHEMX").Columns.Add("THEME_NO", GetType(System.Int64))
            .Tables("ICTTHEMX").Columns.Add("THEME_NO_SEASON_MAX", GetType(System.Int64))

            ASCMAIN1.sql = "Select * from ICTTHEME Where" _
                & " SEASON_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTTHEME_CHECK", "**", 0, False, "V")
        End With

        grdICTTHEMX.DataSource = dst.Tables("ICTTHEMX")
        grdICTTHEMX.DisplayLayout.Bands(0).Columns("USAGE").Format = "#,##0"
        SplitContainer1.Panel2Collapsed = Not (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")

        Fill_Records("ICTTHEMX")
        Sort_grdColumns(grdICTTHEMX, "THEME_DESC")

    End Sub

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SEASON_CODE2"
                'Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE2").Text
                'chkFilterThemes.Enabled = (SEASON_CODE.Length = 5)
                'If SEASON_CODE.Length = 5 Then
                '    Filter_Themes_By_Season()
                'End If
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "SEASON_CODE2"
                Absx1.txtFor("SEASON_CODE2").Text = Absx1.txtFor("SEASON_CODE2").Text.ToUpper
                Filter_Themes_By_Season()
        End Select

    End Sub
    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "SEASON_CODE2"

        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "SEASON_CODE2"

        End Select
    End Sub


#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        If EntryMode = "New" Then
            'ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        ElseIf EntryMode = "Edit" Then
            'ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        End If
    End Sub

    Overrides Sub Show_Record_Special()

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() { _
                "ICTTHEME", "ICTTHEMX", "ICTTHEME_CHECK"}
                dst.Tables(TABLE_NAME).Rows.Clear()
                Absx1.txtFor("SEASON_CODE2").Text = ""
            Next

            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            If tf Then
                chkThemeGenerator.Checked = False
            End If
            chkThemeGenerator.Visible = Not tf
            SplitContainer1.Panel2Collapsed = Not chkThemeGenerator.Checked
            If Not tf Then
                Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE2").Text
                Fill_Records("ICTTHEMX")
                Set_Read_Only(Absx1.txtFor("SEASON_CODE2"), False)
                Set_Read_Only(grdICTTHEMX, False)
                Set_Read_Only(chkThemeGenerator, False)
                Set_Read_Only(chkDeleteSeason, False)
                With grdICTTHEMX.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.False
                End With
                With grdICTTHEMX.DisplayLayout.Bands(0)
                    .Columns("THEME_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("THEME_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("SEASON_CODE_MIN").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("SEASON_CODE_MAX").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("THEME_CODE_MIN").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("THEME_CODE_MAX").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("USAGE").CellActivation = UltraWinGrid.Activation.NoEdit
                End With
            End If
        Else
            chkThemeGenerator.Visible = False
        End If

    End Sub

    Public Overrides Function OK_to_do_View_Lookup(ByVal txtctl As UltraWinEditors.UltraTextEditor) As Boolean

        If txtctl.Name = "txtSeason" Then Return True

        If EntryMode = "" Then
            If htbkey_COLUMN_NAMEs.ContainsKey(Absx1.GetABSColumnName(txtctl)) Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function


#End Region

    Sub Filter_Themes_By_Season()
        Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE2").Text
        If SEASON_CODE.Length = 5 Then
            Dim seasonYear As Integer = Val(SEASON_CODE.Substring(0, 4))
            Dim seasonChar As String = SEASON_CODE.Substring(4, 1)
            Dim lastSeasonYear As Integer = seasonYear - 1
            Dim lySameSeason As String = lastSeasonYear.ToString & seasonChar
            Dim sqlSeason As String = sqlICTTHEMX & " and (SEASON_CODE = '" & lySameSeason & "' or SEASON_CODE = '" & SEASON_CODE & "') Group by ICTTHEME.THEME_DESC, X.USAGE"
            grdICTTHEMX.Text = "Showing Themes used in Seasons " & lySameSeason & " and " & SEASON_CODE
            Fill_Records("ICTTHEMX", , , Replace(sqlSeason, ":SEASON_CODE", SEASON_CODE))
            Show_Existing_Theme_Nos()
            Set_Generator_Mode(True)
        Else
            grdICTTHEMX.Text = "All Theme Codes"
            Fill_Records("ICTTHEMX")
            Set_Generator_Mode(False)
        End If

        For Each rowICTEMEX As DataRow In dst.Tables("ICTTHEMX").Select()
            Dim THEME_DESC As String = rowICTEMEX.Item("THEME_DESC") & ""
            Dim SEASON_CODE_MAX As String = rowICTEMEX.Item("SEASON_CODE_MAX") & ""
            ASCMAIN1.sql = "Select * from ICTTHEME where THEME_DESC = :PARM1 and SEASON_CODE = :PARM2"
            Dim rowICTTHEME As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {THEME_DESC, SEASON_CODE_MAX})
            If rowICTTHEME IsNot Nothing Then
                Dim THEME_NO_SEASON_MAX As Int64 = Val(rowICTTHEME.Item("THEME_NO") & "")
                If THEME_NO_SEASON_MAX > 0 Then
                    rowICTEMEX.Item("THEME_NO_SEASON_MAX") = THEME_NO_SEASON_MAX
                End If
            End If
        Next

        Sort_grdColumns(grdICTTHEMX, "THEME_DESC")
    End Sub

    Sub Show_Existing_Theme_Nos()
        Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE2").Text
        If SEASON_CODE = "" Then
            MsgBox("Please enter a Season to show existing Theme No's.", vbOKOnly, "Cannot Proceed")
            Exit Sub
        Else
            ASCMAIN1.sql = "Select * from ICTTHEME Where NVL(THEME_NO,0) > 0 and SEASON_CODE = :PARM1"
            For Each rowICTTHEME As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New String() {SEASON_CODE}).Select("", "THEME_DESC")
                Dim THEME_DESC As String = rowICTTHEME.Item("THEME_DESC")
                Dim rowICTTHEMXs() As DataRow = dst.Tables("ICTTHEMX").Select(String.Format("THEME_DESC = '{0}'", THEME_DESC.Replace("'", "''")))
                If Not rowICTTHEMXs.Length = 0 Then
                    rowICTTHEMXs(0).Item("THEME_NO") = rowICTTHEME.Item("THEME_NO")
                End If
            Next
        End If
    End Sub

    Private Sub cmdGenerateThemeCodes_Click(sender As Object, e As EventArgs) Handles cmdGenerateThemeCodes.Click
        Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE2").Text
        Fill_Records("ICTTHEME_CHECK", New String() {SEASON_CODE})
        Dim generatedCount As Integer = 0
        Dim updateCount As Integer = 0
        Dim updatedMsg As String = ""
        Dim geeratedMsg As String = ""
        Dim finisheddeMsg As String = ""

        If dst.Tables("ICTTHEME_CHECK").Rows.Count > 0 Then
            If Not chkDeleteSeason.Checked Then
                MsgBox("Theme codes alredy exists for Season " & SEASON_CODE, vbOKOnly, "Cannot Proceed")
                Exit Sub
            End If
        End If

        Dim THEME_NO_prev As Integer = 0
        Dim THEME_NO_curr As Integer = 0
        Dim t As Integer = 0

        For Each rowICTTHEMX As DataRow In dst.Tables("ICTTHEMX").Select("ISNULL(USAGE,0) > 0")
            Dim THEME_DESC_X As String = rowICTTHEMX.Item("THEME_DESC")
            Dim rowICTTHEME_check() As DataRow = dst.Tables("ICTTHEME_CHECK").Select("THEME_DESC = '" & THEME_DESC_X & "'")
            Dim THEME_NO_X As String = rowICTTHEMX.Item("THEME_NO") & ""
            Dim THEME_NO_C As String = rowICTTHEME_check(0).Item("THEME_NO") & ""
            If THEME_NO_X <> THEME_NO_C Then
                Dim msgUsage As String = "Cannot change Theme No for a Theme Code with usage in the selected Season -> " & SEASON_CODE & ":" & THEME_DESC_X
                MsgBox(msgUsage, vbOKOnly, "Cannot Proceed")
                Exit Sub
            End If
        Next

        For Each rowICTTHEMX As DataRow In dst.Tables("ICTTHEMX").Select("ISNULL(THEME_NO,-1) <> -1", "THEME_NO")

            THEME_NO_curr = Val(rowICTTHEMX.Item("THEME_NO"))
            If THEME_NO_curr < 1 Then
                MsgBox("Theme No's must be >= 1.", vbOKOnly, "Cannot Proceed")
                Exit Sub
            End If
            If t > 0 Then
                If Not THEME_NO_prev = (THEME_NO_curr - 1) Then
                    MsgBox("Theme No's must be contiguous.", vbOKOnly, "Cannot Proceed")
                    Exit Sub
                End If
            Else
                If THEME_NO_curr = 0 Then
                    MsgBox("Theme No's must be > 0.", vbOKOnly, "Cannot Proceed")
                    Exit Sub
                End If
                If THEME_NO_curr > 1 Then
                    MsgBox("Theme No's must begin with 1.", vbOKOnly, "Cannot Proceed")
                    Exit Sub
                End If
            End If
            t += 1
            THEME_NO_prev = Val(rowICTTHEMX.Item("THEME_NO"))
        Next

        For Each rowICTTHEMX As DataRow In dst.Tables("ICTTHEMX").Select("ISNULL(THEME_NO,-1) <> -1 AND ISNULL(USAGE,0) = 0", "THEME_NO")
            Dim rowICTTHEME As DataRow = Nothing
            Dim THEME_DESC As String = rowICTTHEMX.Item("THEME_DESC")
            Dim THEME_NO As String = rowICTTHEMX.Item("THEME_NO")
            Dim rowICTTHEME_check() As DataRow = dst.Tables("ICTTHEME_CHECK").Select(String.Format("THEME_DESC = '{0}'", THEME_DESC.Replace("'", "''")))
            If rowICTTHEME_check.Length > 0 Then
                If Not chkDeleteSeason.Checked Then
                    MsgBox("Theme alredy exists for Season " & SEASON_CODE, vbOKOnly, "Cannot Proceed")
                    Exit Sub
                End If
            Else
                Dim THEME_CODE As String = ASCMAIN1.Next_Control_No("ICTTHEME.THEME_CODE")
                rowICTTHEME = dst.Tables("ICTTHEME").NewRow()
                rowICTTHEME.Item("THEME_CODE") = THEME_CODE
                rowICTTHEME.Item("THEME_DESC") = THEME_DESC
                rowICTTHEME.Item("THEME_NO") = THEME_NO
                rowICTTHEME.Item("SEASON_CODE") = SEASON_CODE
                dst.Tables("ICTTHEME").Rows.Add(rowICTTHEME)
            End If
        Next

        Update_Record_TDA("ICTTHEME")

        If generatedCount > 0 Then
            MsgBox("Theme Codes Generated", vbOKOnly, "Done")
        End If

        Reset_Theme_Generator()
    End Sub

    Sub Reset_Theme_Generator()
        Absx1.txtFor("SEASON_CODE2").Text = ""
        chkDeleteSeason.Checked = False
        chkThemeGenerator.Checked = False

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() { _
            "ICTTHEME", "ICTTHEMX", "ICTTHEME_CHECK"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        grdICTTHEMX.Text = "All Theme Codes"
        Fill_Records("ICTTHEMX")

    End Sub
    Sub Set_Generator_Mode(tf As Boolean)
        cmdGenerateThemeCodes.Visible = tf
        chkDeleteSeason.Visible = tf
        If tf Then
            With grdICTTHEMX.DisplayLayout
                .Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .Override.AllowUpdate = DefaultableBoolean.True
                .Bands(0).Columns("THEME_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Bands(0).Columns("THEME_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
            End With
        Else
            With grdICTTHEMX.DisplayLayout
                .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                .Override.AllowUpdate = DefaultableBoolean.False
                .Bands(0).Columns("THEME_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                .Bands(0).Columns("THEME_NO").CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        End If
    End Sub

    Private Sub chkThemeGenerator_CheckedChanged(sender As Object, e As EventArgs) Handles chkThemeGenerator.CheckedChanged
        SplitContainer1.Panel2Collapsed = Not chkThemeGenerator.Checked
        chkDeleteSeason.Checked = False
        Set_Generator_Mode(False)
    End Sub

    Private Sub btnAutoAssignThemeNo_Click(sender As Object, e As EventArgs) Handles btnAutoAssignThemeNo.Click
        For Each rowICTEMEX As DataRow In dst.Tables("ICTTHEMX").Select()
            Dim THEME_NO_SEASON_MAX As Int64 = Val(rowICTEMEX.Item("THEME_NO_SEASON_MAX") & "")
            If THEME_NO_SEASON_MAX > 0 Then
                rowICTEMEX.Item("THEME_NO") = rowICTEMEX.Item("THEME_NO_SEASON_MAX") & ""
            End If
        Next
    End Sub
End Class

