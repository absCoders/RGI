Imports System.Drawing
Imports System.Math

Public Class ICFTHEM1
    Dim rowICTSEAS1 As DataRow
    Dim sqlICTSTYC1 As String = ""
    Dim incActiveUnassigned As Boolean = False


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFTHEMI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")

        With dst
            Dim sqlExclude As String = " AND SUBSTR(ICTSTYL1.STYLE_CODE,1,3) <> 'MTB'" & vbCrLf _
                                       & " AND SUBSTR(ICTSTYL1.STYLE_CODE, -1) NOT IN ('K','M','N','O','Q','R','S','T') " & vbCrLf _
                                       & " AND SUBSTR(ICTSTYL1.STYLE_CODE, -2) NOT IN ('K1','K2','KA','KB')"

            ASCMAIN1.sql = "Select X.SEASON_CODE, ICTSEAS1.SEASON_DESC, X.THEME_CODES, X.STYLE_CODES" _
                & " from ICTSEAS1, (Select ICTTHEME.SEASON_CODE, Count (Distinct ICTTHEME.THEME_CODE) THEME_CODES, Count (*) STYLE_CODES" & vbCrLf _
                & " from ICTSTYC1,ICTTHEME,ICTSTYL1 where ICTTHEME.THEME_CODE (+) = ICTSTYC1.THEME_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" & vbCrLf _
                & sqlExclude & vbCrLf _
                & "   and ICTSTYL1.STYLE_STATUS = 'A' and ICTSTYC1.STYLE_COLOR_STATUS = 'A'" & vbCrLf _
                & " group by ICTTHEME.SEASON_CODE) X where ICTSEAS1.SEASON_CODE = X.SEASON_CODE"
            Create_TDA(.Tables.Add, "ICTSEASX", "**", 0, False)

            Dim ATTR_SQL As String = "SELECT ICTSTYL3.STYLE_CODE, MAX(ATTR_CODE) ATTR_CODE FROM ICTSTYL3 GROUP BY ICTSTYL3.STYLE_CODE"
            ASCMAIN1.sql = "Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE" & vbCrLf _
                & ", ICTSTYC1.STYLE_COLOR_IMAGE_NAME, ICTSTYC1.THEME_CODE, ICTTHEME.THEME_DESC, ICTTHEME.SEASON_CODE, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, ICTSTYL3.ATTR_CODE " & vbCrLf _
                & " from ICTSTYC1,ICTSTYL1,ICTCOLR1,ICTTHEME, (" & ATTR_SQL & ") ICTSTYL3" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = ICTSTYL3.STYLE_CODE (+)" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE" & vbCrLf _
                & "   and ICTTHEME.THEME_CODE (+) = ICTSTYC1.THEME_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_STATUS = 'A' and ICTSTYC1.STYLE_COLOR_STATUS = 'A'" & vbCrLf _
                & sqlExclude & vbCrLf
            sqlICTSTYC1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, True, "V", 2, "THEME_CODE")
            .Tables("ICTSTYC1").Columns.Add("THEME_CODE_NEW")
            .Tables("ICTSTYC1").Columns.Add("THEME_DESC_NEW")
            .Tables("ICTSTYC1").Columns.Add("SEASON_CODE_NEW")
            .Tables("ICTSTYC1").Columns.Add("DISCONTINUE_STYLE")
            .Tables("ICTSTYC1").Columns.Add("DISCONTINUE_COLOR")

            ASCMAIN1.sql = "Select * from ICTTHEME where SEASON_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTTHEME", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select * from ICTSTYT1 where SEASON_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTYT1", "**", 0, True, "V", 3)
            .Tables("ICTSTYT1").Columns.Add("THEME_CODE_NEW")

        End With

        grdICTSTYC1.DataSource = dst.Tables("ICTSTYC1")
        grdICTSEASX.DataSource = dst.Tables("ICTSEASX")
        grdICTTHEME.DataSource = dst.Tables("ICTTHEME")

        Create_Summary(grdICTSEASX, "SEASON_CODE", "Count")
        Create_Summary(grdICTSTYC1, "STYLE_CODE", "Count")

        Show_Filter(grdICTSTYC1, True)
        Sort_grdColumns(grdICTSTYC1, "STYLE_CODE, COLOR_CODE")

        With grdICTSTYC1.DisplayLayout
            .Bands(0).Columns("THEME_DESC_NEW").CellActivation = UltraWinGrid.Activation.NoEdit
        End With

        AUDIT.Add("ICTSTYC1", "*")

        ASCMAIN1.Add_Value_List(cbeSEASON_CODE, "SEASON_CODE")
        cbeSEASON_CODE.Value = cbeSEASON_CODE.Items(0).DataValue
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Edit"
                If Absx1.txtFor("SEASON_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify a Season Code to View"
                Else
                    rowICTSEAS1 = LookUp("ICTSEAS1", Absx1.txtFor("SEASON_CODE").Text)
                    If rowICTSEAS1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Season " & Absx1.txtFor("SEASON_CODE").Text
                    End If
                End If

            Case "Update"
                'If Absx1.txtFor("REASON_CODE").Text = "" Then
                '    EMsg &= vbCr & "You Must Specify a Reason"
                'Else
                '    Dim rowICTREAS1 As DataRow = LookUp("ICTREAS1", Absx1.txtFor("REASON_CODE").Text)
                '    If rowICTREAS1 Is Nothing Then
                '        EMsg &= vbCr & "Invalid Value Specified for Reason"
                '    End If
                'End If

                'If grdICTSTYC1.Rows.Count = 0 Then
                '    EMsg &= vbCr & "No Details Entered"
                'Else
                '    For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("", "", DataViewRowState.CurrentRows)
                '        If rowICTSTYC1.Item("STYLE_CLASS_CODE") & "" = "" Then
                '            EMsg &= vbCr & "Unable to determine Class for " & rowICTSTYC1.Item("STYLE_CODE") & ""
                '        End If
                '        If rowICTSTYC1.Item("SALES_DIVISION_CODE") & "" = "" Then
                '            EMsg &= vbCr & "Unable to determine Division for " & rowICTSTYC1.Item("STYLE_CODE") & ""
                '        End If
                '    Next
                'End If


            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Publish"
                Update_Record(True)
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                End With
                .Groups("Themes in Season").Visible = ScreenMode
                .Groups("Style Image").Visible = ScreenMode
                .Groups("Style Image").Text = "Style Image"
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdICTTHEME.Selected.Rows.Clear()

        grdICTSEASX.Visible = Not ScreenMode
        grdICTSTYC1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTSEASX", "ICTSTYC1", "ICTSTYT1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        incActiveUnassigned = False
        Absx1.txtFor("SEASON_CODE").Text = ""

        Refresh_Documents()
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)
        Dim SEASON_CODE As String = rowICTSEAS1.Item("SEASON_CODE")

        rowICTSEAS1 = LookUp("ICTSEAS1", SEASON_CODE)
        Dim SEASON_CODE_YEAR As String = Mid(SEASON_CODE, 1, 4)
        Dim SEASON_CODE_NEXT_YEAR As String = (CInt(SEASON_CODE_YEAR) + 1).ToString
        Dim rowICTSEAS_check As DataRow = LookUp("ICTSEAS1", SEASON_CODE_NEXT_YEAR & Mid(SEASON_CODE, 5, 1))
        If rowICTSEAS_check IsNot Nothing Then
            cbeSEASON_CODE.Value = rowICTSEAS_check.Item("SEASON_CODE")
        End If

        Load_Season()

        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Season(Optional incActiveUnasigned As Boolean = False)

        Dim sqlSeason As String = ""
        Dim grdCap As String = ""
        If incActiveUnasigned Then
            sqlSeason = "   and (ICTTHEME.SEASON_CODE = :PARM1  or ICTTHEME.SEASON_CODE IS NULL)"
            grdCap = "Style Colors with Themes in Season - Including Active, Unassigned"
        Else
            sqlSeason = "   and ICTTHEME.SEASON_CODE = :PARM1"
            grdCap = "Style Colors with Themes in Season"
        End If

        ASCMAIN1.Progress("Now loading season: " & Absx1.txtFor("SEASON_CODE").Text & "...")

        Dim sqlCombined As String = sqlICTSTYC1 & sqlSeason
        Fill_Records("ICTSTYC1", Absx1.txtFor("SEASON_CODE").Text, , sqlCombined)
        Fill_Records("ICTSTYT1", Absx1.txtFor("SEASON_CODE").Text)
        grdICTSTYC1.Text = grdCap

        For Each rowICTSTYT1 As DataRow In dst.Tables("ICTSTYT1").Select("")
            Dim STYLE_CODE As String = rowICTSTYT1.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTSTYT1.Item("COLOR_CODE")
            Dim THEME_CODE As String = rowICTSTYT1.Item("THEME_CODE")
            Dim rowICTSTYC1s() As DataRow = dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
            If rowICTSTYC1s.Length = 1 Then
                rowICTSTYC1s(0).Item("THEME_CODE_NEW") = rowICTSTYT1.Item("THEME_CODE")
                ASCMAIN1.Progress("Loading saved work for: " & STYLE_CODE & ":" & COLOR_CODE, rowICTSTYC1s(0).Item("THEME_CODE_NEW"))
                Dim rowICTTHEMEs() As DataRow = dst.Tables("ICTTHEME").Select("THEME_CODE = '" & THEME_CODE & "'")
                Dim rowICTTHEME As DataRow = LookUp("ICTTHEME", THEME_CODE)
                If rowICTTHEME IsNot Nothing Then
                    rowICTSTYC1s(0).Item("THEME_DESC_NEW") = rowICTTHEME.Item("THEME_DESC")
                    rowICTSTYC1s(0).Item("SEASON_CODE_NEW") = rowICTTHEME.Item("SEASON_CODE")
                End If
            End If
        Next

        If grdICTSTYC1.Rows.Count > 0 Then
            If Not ASCMAIN1.Running_in_VS Then
                Dim firstRow As UltraWinGrid.UltraGridRow = grdICTSTYC1.Rows(0)
                grdICTSTYC1.ActiveRow = firstRow
            End If
        End If

    End Sub
    Sub Autosave_Work()

        ASCMAIN1.Progress("Autosaving Work....")
        dst.Tables("ICTSTYT1").Rows.Clear()
        Dim SEASON_CODE As String = rowICTSEAS1.Item("SEASON_CODE")

        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("ISNULL(THEME_CODE_NEW,'') <> ''")
            Dim rowICTSTYT1 As DataRow = dst.Tables("ICTSTYT1").NewRow
            Dim STYLE_CODE As String = rowICTSTYC1.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE")
            Dim THEME_CODE As String = rowICTSTYC1.Item("THEME_CODE_NEW")
            rowICTSTYT1.Item("SEASON_CODE") = SEASON_CODE
            rowICTSTYT1.Item("STYLE_CODE") = rowICTSTYC1.Item("STYLE_CODE")
            rowICTSTYT1.Item("COLOR_CODE") = rowICTSTYC1.Item("COLOR_CODE")
            rowICTSTYT1.Item("THEME_CODE") = rowICTSTYC1.Item("THEME_CODE_NEW")
            dst.Tables("ICTSTYT1").Rows.Add(rowICTSTYT1)
        Next

        Update_Record(False, False)

        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record(Optional publish As Boolean = False, Optional showCommitMsg As Boolean = True)

        Dim commitMsg As String = ""
        Dim tbl As String = IIf(publish, "ICTSTYC1", "ICTSTYT1")
        Dim SEASON_CODE As String = rowICTSEAS1.Item("SEASON_CODE")
        Dim sqlDelete As String = IIf(publish, "", "SEASON_CODE = '" & SEASON_CODE & "'")

        BeginTrans()

        For Each ROW As DataRow In dst.Tables(tbl).Select("ISNULL(THEME_CODE_NEW,'') <> ''")
            ROW.Item("THEME_CODE") = ROW.Item("THEME_CODE_NEW")
        Next

        Update_Record_TDA(tbl, sqlDelete)

        If publish Then
            commitMsg = "Publish Complete"
        Else
            If showCommitMsg Then
                commitMsg = "Update Complete"
            End If
        End If

        CommitTrans(commitMsg)

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSEASX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTSTYC1, "BBBBBS", "Style Status Inquiry", "Style Master File", "Assign to Theme", "Select All - Same Theme, Not Reassigned", "Select All - Same Theme", "Show Active - No Theme")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdICTSTYC1"
                    For Each btnKey As String In New String() {"Assign to Theme", "Select All - Same Theme, Not Reassigned", "Select All - Same Theme"}
                        tlb_btn = tlb_pop.Tools(btnKey)

                        If grd.Selected.Rows.Count = 0 Then
                            If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                                grd.ActiveRow.Selected = True
                            Else
                                tlb_btn.SharedProps.Visible = False
                                Exit Sub
                            End If
                        End If

                        If grdICTTHEME.ActiveRow Is Nothing Then
                            tlb_btn.SharedProps.Visible = False
                            Exit Sub
                        End If
                        Select Case btnKey
                            Case "Assign to Theme"
                                tlb_btn.SharedProps.Caption = "Assign to " & grdICTTHEME.ActiveRow.Cells("THEME_DESC").Value
                            Case "Select All - Same Theme, Not Reassigned"
                                tlb_btn.SharedProps.Caption = "Select All " & grdICTSTYC1.ActiveRow.Cells("THEME_DESC").Value & " - Not Reassigned"
                            Case "Select All - Same Theme"
                                tlb_btn.SharedProps.Caption = "Select All " & grdICTSTYC1.ActiveRow.Cells("THEME_DESC").Value
                            Case "Discontinue Style"
                                tlb_btn.SharedProps.Caption = "Discontinue " & grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
                            Case "Discontinue Color"
                                tlb_btn.SharedProps.Caption = "Discontinue " & grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value & ":" & grdICTSTYC1.ActiveRow.Cells("COLOR_DESC").Value
                        End Select
                        tlb_btn.SharedProps.Visible = True
                    Next

                    'tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    'tlb_sbt = DirectCast(tlb_pop.Tools("Show Active - No Theme"), UltraWinToolbars.StateButtonTool)
                    'tlb_sbt.SharedProps.Visible = True ' (Absx1.optFor("ORDR_SOURCE").Value = "K")
                    'tlb_sbt.Tag = "X"
                    'tlb_sbt.Checked = incActiveUnassigned
                    'tlb_sbt.Tag = ""


            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Assign to Theme"
                For Each grow As UltraWinGrid.UltraGridRow In grdICTSTYC1.Selected.Rows
                    grow.Cells("THEME_DESC_NEW").Value = grdICTTHEME.ActiveRow.Cells("THEME_DESC").Value
                    grow.Cells("SEASON_CODE_NEW").Value = grdICTTHEME.ActiveRow.Cells("SEASON_CODE").Value
                    grow.Cells("THEME_CODE_NEW").Value = grdICTTHEME.ActiveRow.Cells("THEME_CODE").Value
                    grow.Update()
                Next
                grdICTSTYC1.Selected.Rows.Clear()
                Autosave_Work()

            Case "Show Active - No Theme"
                incActiveUnassigned = Not incActiveUnassigned

                Load_Season(incActiveUnassigned)
                ASCMAIN1.Progress("")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                If Style_Valid(grd.ActiveRow.Cells("STYLE_CODE").Text) Then
                    Context_Launch("Select", grd.ActiveRow.Cells("STYLE_CODE").Text, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Style Master File"
                If Style_Valid(grd.ActiveRow.Cells("STYLE_CODE").Text) Then
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", grd.ActiveRow.Cells("STYLE_CODE").Text)
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If

            Case "Select All - Same Theme"
                Dim THEME_CODE As String = grd.ActiveRow.Cells("THEME_CODE").Text

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        If grow.Cells("THEME_CODE").Value & "" = THEME_CODE Then
                            If Not grdICTSTYC1.Selected.Rows.Contains(grow) Then
                                grdICTSTYC1.Selected.Rows.Add(grow)
                            End If
                        End If

                        'grow.Cells("SEL").Value = IIf(e.Tool.Key = "De-Select All", "0", "1")
                        'grow.Update()
                    End If
                Next
            Case "Select All - Same Theme, Not Reassigned"
                Dim THEME_CODE As String = grd.ActiveRow.Cells("THEME_CODE").Text

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        If grow.Cells("THEME_CODE").Value & "" = THEME_CODE AndAlso grow.Cells("THEME_CODE_NEW").Value & "" = "" Then
                            If Not grdICTSTYC1.Selected.Rows.Contains(grow) Then
                                grdICTSTYC1.Selected.Rows.Add(grow)
                            End If
                        End If

                        'grow.Cells("SEL").Value = IIf(e.Tool.Key = "De-Select All", "0", "1")
                        'grow.Update()
                    End If
                Next
            Case "Discontinue Style"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                If Style_Valid(STYLE_CODE) Then
                    For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "'")
                        rowICTSTYC1.Item("DISCONTINUE_STYLE") = "1"

                    Next
                End If
            Case "Discontinue Color"
                Dim COLOR_CODE As String = grd.ActiveRow.Cells("COLOR_CODE").Text
                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                If rowICTCOLR1 IsNot Nothing Then
                    For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("COLOR_CODE = '" & COLOR_CODE & "'")
                        rowICTSTYC1.Item("DISCONTINUE_COLOR") = "1"
                    Next
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SEASON_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Edit", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "SEASON_CODE"
                Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "SEASON_CODE"
        End Select
    End Sub

#End Region

    Private Sub grdICTSEASX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSEASX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("SEASON_CODE").Text = e.Row.Cells("SEASON_CODE").Text
            Click_Command("Edit")
        End If
    End Sub
      
    Sub Refresh_Documents()
        Fill_Records("ICTSEASX")
        Sort_grdColumns(grdICTSEASX, "SEASON_CODE".ToLower)
    End Sub

    Private Sub cbeSEASON_CODE_ValueChanged(sender As Object, e As EventArgs) Handles cbeSEASON_CODE.ValueChanged
        Dim SEASON_CODE As String = cbeSEASON_CODE.Value
        Fill_Records("ICTTHEME", SEASON_CODE)
        Sort_grdColumns(GRDICTTHEME, "THEME_NO")

    End Sub
    Private Function Style_Valid(STYLE_CODE As String) As Boolean
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        Return rowICTSTYL1 IsNot Nothing
    End Function

    Private Sub grdICTSTYC1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTSTYC1.AfterRowActivate
        If grdICTSTYC1.ActiveRow.IsDataRow Then
            If Not ASCMAIN1.Running_in_VS Then
                Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
                Dim IMAGE_NAME As String = STYLE_CODE & "-" & COLOR_CODE

                Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
                Dim imgba() As Byte = Nothing
                picStyleImage.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, False, , , imgba)

                UltraExplorerBar1.Groups("Style Image").Text = IMAGE_NAME
            End If

        End If
    End Sub

    Private Sub grdICTSTYC1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYC1.InitializeRow
        Dim DISCONTINUE_STYLE As String = e.Row.Cells("DISCONTINUE_STYLE").Text & ""
        Dim DISCONTINUE_COLOR As String = e.Row.Cells("DISCONTINUE_COLOR").Text & ""
        If DISCONTINUE_STYLE = "1" Then
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Color.Red
            e.Row.Cells("STYLE_DESC").Appearance.ForeColor = Color.Red
        End If
        If DISCONTINUE_COLOR = "1" Then
            e.Row.Cells("COLOR_CODE").Appearance.ForeColor = Color.Red
            e.Row.Cells("COLOR_DESC").Appearance.ForeColor = Color.Red
        End If

    End Sub

    Private Sub picStyleImage_DoubleClick(sender As Object, e As EventArgs) Handles picStyleImage.DoubleClick
        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        Dim IMAGE_NAME As String = UltraExplorerBar1.Groups("Style Image").Text & ".JPG"
        If Not FOLDER_NAME.EndsWith("\") Then
            FOLDER_NAME = FOLDER_NAME & "\"
        End If
        Dim frm As New ICFSIMG1(FOLDER_NAME & IMAGE_NAME)
        frm.ShowDialog()
    End Sub
End Class