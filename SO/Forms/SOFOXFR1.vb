Imports Infragistics.Win.UltraWinGrid

Public Class SOFOXFR1

    Dim SOTOXFRX As String = ""
    Dim SOTORDRX As String = ""
    Dim TABLES_OXFR() As String = {"SOTORDR1", "SOTORDR2", "SOTORDR5", "SOTPICK0", "SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2"}

    Dim sqlSOTORDR0 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFTHEMI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("EDTPARM1")

        Create_WorkTables(True)

        With dst
            ASCMAIN1.sql = $"Select * from {SOTOXFRX}"
            Create_TDA(.Tables.Add, "SOTOXFRX", "**", 0, False)
            With .Tables("SOTOXFRX")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"

                .Columns.Add("UNITS_2_XFR", GetType(System.Int32), "IIF(ISNULL(SEL,'0') = '1', QTY_TO_XFR, 0)")
                .Columns.Add("CASES_2_XFR", GetType(System.Int32), "UNITS_2_XFR / ISNULL(CARTON_PACK_QTY,0)")
                .Columns.Add("CUBE_2_XFR", GetType(System.Decimal), "CASES_2_XFR * ISNULL(CASE_CUBE,0)")
            End With

            ASCMAIN1.sql = $"Select * from {SOTORDRX} where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "VV", 2)
            '.Tables("SOTORDRX").Columns.Add("LINES_PCT_US", GetType(System.Int32), "IIF(ISNULL(LINES_OPEN,0) = 0, 0, 100 * ISNULL(LINES_US,0)/ISNULL(LINES_OPEN,0))")

            For Each TABLE_NAME As String In TABLES_OXFR
                Create_TDA(.Tables.Add, TABLE_NAME, "*")
            Next
            ASCMAIN1.sql = "Select SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR0.ORDR_DATE" & vbCrLf _
                & ", SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR0.WHSE_CODE" & vbCrLf _
                & ", SOTORDR0.ORDR_TYPE_CODE, SOTORDR0.ORDR_SOURCE" & vbCrLf _
                & ", SOTORDR0.ORDR_QTY, SOTORDR0.ORDR_QTY_OPEN, SOTORDR0.ORDR_QTY_PICK, SOTORDR0.ORDR_QTY_SHIP, SOTORDR0.ORDR_QTY_CANC" & vbCrLf _
                & ", SOTORDR1.INIT_DATE, SOTORDR1.INIT_OPER, SOTORDR2_TOTALS.CTNS, SOTORDR2_TOTALS.UNITS, SOTORDR2_TOTALS.CUBE" & vbCrLf _
                & " from SOTORDR0,SOTORDR1, (Select SOTORDR2.ORDR_NO, SUM (SOTORDR2.ORDR_QTY) UNITS, SUM (SOTORDR2.ORDR_QTY / SOTORDR2.CARTON_PACK_QTY) CTNS, SUM ((SOTORDR2.ORDR_QTY / SOTORDR2.CARTON_PACK_QTY) * ICTSTYL1.CASE_CUBE) CUBE from SOTORDR2,ICTSTYL1 where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE group by SOTORDR2.ORDR_NO) SOTORDR2_TOTALS" & vbCrLf _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & "   and SOTORDR2_TOTALS.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_STATUS between 'O' and 'P' and SOTORDR1.ORDR_SOURCE = 'X' and SOTORDR1.ORDR_TYPE_CODE = 'XFR'"
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "", 2)
        End With

        grdSOTOXFRX.DataSource = dst.Tables("SOTOXFRX")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTOXFRX.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = System.Drawing.Color.White
            gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            If gcol.Key.StartsWith("US_") Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
            ElseIf gcol.Key.StartsWith("MS_") Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
            ElseIf gcol.Key = "SEL" Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Goldenrod
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            ElseIf gcol.Key = "ALLO" Or gcol.Key = "QTY_TO_XFR" Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
            ElseIf gcol.Key = "UNITS_2_XFR" Or gcol.Key = "CASES_2_XFR" Or gcol.Key = "CUBE_2_XFR" Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
            End If
        Next

        Create_Summary(grdSOTOXFRX, "STYLE_CODE", "Count")
        Create_Summary(grdSOTOXFRX, New String() {"SEL", "UNITS_2_XFR", "CASES_2_XFR", "CUBE_2_XFR"})

        Show_Filter(grdSOTOXFRX, True)

        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORDRX.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = System.Drawing.Color.White
            gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            If New String() {"ORDR", "OPEN", "SHIP", "PICK", "CANC", "ALLO"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
            ElseIf gcol.Key.StartsWith("ORDR_AMT") Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                'ElseIf gcol.Key = "LINES_OPEN" Or gcol.Key = "LINES_US" Or gcol.Key = "LINES_PCT_US" Then
                '    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
            End If
        Next

        Create_Summary(grdSOTORDRX, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDRX, New String() {"ORDR", "OPEN", "SHIP", "PICK", "CANC", "ALLO"})
        Create_Summary(grdSOTORDRX, New String() {"ORDR_AMT", "ORDR_AMT_OPEN"}) ', "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP"

        Show_Filter(grdSOTORDRX, True)
        Sort_grdColumns(grdSOTORDRX, "ORDR_SHIP_DATE")


        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORDR0.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = System.Drawing.Color.White
            gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            If New String() {"ORDR", "OPEN", "SHIP", "PICK", "CANC", "ALLO"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
            ElseIf gcol.Key.StartsWith("ORDR_AMT") Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
            End If
        Next

        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDR0, New String() {"CTNS", "UNITS", "CUBE"})

        Show_Filter(grdSOTORDR0, True)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load"

            Case "Update"

                If dst.Tables("SOTOXFRX").Select("SEL='1'").Length = 0 Then
                    EMsg &= vbCr & "Nothing Selected to Transfer"
                End If

                Dim ava_check() As DataRow = dst.Tables("SOTOXFRX").Select("SEL='1' AND ISNULL(UNITS_2_XFR,0) > ISNULL(US_AVA,0)")

                If ava_check.Length > 0 Then
                    Dim SCs As New List(Of String)
                    For Each row As DataRow In ava_check
                        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                        SCs.Add(STYLE_CODE & "-" & COLOR_CODE)
                    Next
                    EMsg &= vbCr & "SCs Queued for Transfer where Units2Xfr is greater than Ava US" & vbCr & Join(SCs.ToArray, ",")
                End If

                If ASCMAIN1.Running_in_VS Then Stop

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
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

            Case "Refresh"
                Refresh_Documents()

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
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
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode
        SplitContainer2.Visible = Not ScreenMode

        If ScreenMode Then
            grdSOTOXFRX.Parent = SplitContainer1.Panel1
        Else
            Clear_Record()
            grdSOTOXFRX.Parent = SplitContainer2.Panel1
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"SOTOXFRX", "SOTORDRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        For Each TABLE_NAME As String In TABLES_OXFR
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)
        'Absx1.txtFor("WHSE_CODE").Text = ""

        Refresh_Documents()
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Create_WorkTables(False)
        Fill_Records("SOTOXFRX")
        Sort_grdColumns(grdSOTOXFRX, "STYLE_CODE, COLOR_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record(Optional publish As Boolean = False, Optional showCommitMsg As Boolean = True)

        BeginTrans()

        ' Reduce Qty OPEN by the Qty ALLO for all selected SCs to Transfer

        Dim WHSE_CODE As String = "US"

        TAC.SOCMAIN1.Create_Transfer_Order(Me, dst.Tables("SOTOXFRX").Select("SEL = '1'"), "UNITS_2_XFR") ' Create a Single XFR Order for the selected SCs to Transfer
        TAC.SOCMAIN1.Release_Transfer_Order(Me) ' Release that XFR Order

        For Each TABLE_NAME As String In TABLES_OXFR
            Update_Record_TDA(TABLE_NAME)
        Next

        Dim SHIP_BOL_NO As String = dst.Tables("SOTSHIP1").Rows(0).Item("SHIP_BOL_NO")
        Dim ORDR_GROUP_NO As String = dst.Tables("SOTSHIP1").Rows(0).Item("ORDR_GROUP_NO")
        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

        For Each rowSOTOXFRX As DataRow In dst.Tables("SOTOXFRX").Select("SEL = '1'")
            Dim STYLE_CODE As String = rowSOTOXFRX.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTOXFRX.Item("COLOR_CODE")
            Dim QTY As Int32 = -1 * Val(rowSOTOXFRX.Item("ALLO"))
            TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY)

            ' Update Status of Transfer Queue Records
            ASCMAIN1.sql = $"Update SOTOXFR1 SET OXFR_STATUS = '1', SHIP_BOL_NO = '{SHIP_BOL_NO}'" & vbCrLf _
            & " where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2 and OXFR_STATUS = '0'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {STYLE_CODE, COLOR_CODE})
        Next

        ' Incease Qty In PICK by the PICK_QTY for all selected SCs to Transfer

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim QTY As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK"))
            TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_PICK", QTY)
        Next

        ExportPickTckts(SHIP_BOL_NO)

        CommitTrans("Update Successful")

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
        Load_Popup_Menu(grdSOTOXFRX, "SSBBBBB", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdSOTORDRX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTORDR0, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")

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

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Sales Order Inquiry", "Sales Order Entry"

                Dim ORDR_NO As String = ""

                If grd.Name = "grdSOTORDR0" Then
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Else
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                End If

                If e.Tool.Key = "Sales Order Entry" Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDR1")
                Else
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                Next
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "SEASON_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Click_Command("Edit", e)
            '    End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "SEASON_CODE"
            '    Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "SEASON_CODE"
        End Select
    End Sub

#End Region

    Sub Refresh_Documents()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building Order from Records Selected in Transfer Queue", "")

        Fill_Records("SOTORDR0")
        Sort_grdColumns(grdSOTORDR0, "ORDR_GROUP_NO".ToLower)

        Create_WorkTables(False)
        Fill_Records("SOTOXFRX")
        Sort_grdColumns(grdSOTOXFRX, "STYLE_CODE, COLOR_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdSOTOXFRX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTOXFRX.AfterRowActivate

        If grdSOTOXFRX.ActiveRow Is Nothing OrElse Not grdSOTOXFRX.ActiveRow.IsDataRow Then
            grdSOTORDRX.Visible = False
        Else

            Dim STYLE_CODE As String = grdSOTOXFRX.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdSOTOXFRX.ActiveRow.Cells("COLOR_CODE").Value
            grdSOTORDRX.Text = $"Released Sales Orders with Style / Color {STYLE_CODE} / {COLOR_CODE}"

            Fill_Records("SOTORDRX", New String() {STYLE_CODE, COLOR_CODE})

            grdSOTORDRX.Visible = True
        End If

    End Sub

    Sub Create_WorkTables(initialize As Boolean)

        If initialize Then
            SOTOXFRX = ASCMAIN1.Temp_Table(Get_SQL("SOTOXFRX") & " and ROWNUM < 1")
            ASCDATA1.ExecuteSQL($"Alter Table {SOTOXFRX} add Primary Key (STYLE_CODE, COLOR_CODE)")

            SOTORDRX = ASCMAIN1.Temp_Table(Get_SQL("SOTORDRX") & " and ROWNUM < 1")
            ASCDATA1.ExecuteSQL($"Alter Table {SOTORDRX} add Primary Key (ORDR_TYPE, ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE)")
            ASCDATA1.ExecuteSQL($"Create Index I_{SOTORDRX}_1 on {SOTORDRX} (STYLE_CODE, COLOR_CODE)")

        Else
            ASCDATA1.ExecuteSQL($"Truncate Table {SOTOXFRX}")
            ASCDATA1.ExecuteSQL($"Insert into {SOTOXFRX} {Get_SQL("SOTOXFRX")}")

            ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDRX}")
            ASCDATA1.ExecuteSQL($"Insert into {SOTORDRX} {Get_SQL("SOTORDRX")}")
            'ASCDATA1.ExecuteSQL($"Update {SOTORDRX} SOTORDRX Set LINES_OPEN = (Select Count (*) from SOTORDR2 where ORDR_NO = SOTORDRX.ORDR_NO and ORDR_QTY_OPEN > 0)")
            'ASCDATA1.ExecuteSQL($"Update {SOTORDRX} SOTORDRX Set LINES_US = (Select Count (*) from SOTORDR2 where ORDR_NO = SOTORDRX.ORDR_NO and ORDR_QTY_OPEN > 0 and (STYLE_CODE, COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE from {SOTOXFRX}))")

        End If
    End Sub

    Function Get_SQL(TABLE_NAME As String) As String

        Dim SQL As String = ""

        Select Case TABLE_NAME

            Case "SOTOXFRX"
                SQL = "Select X.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                    & ", US_TRAN, US_ONHD, US_PICK, US_OPEN, NVL(US_ONHD,0) - NVL(US_PICK,0) - NVL(US_OPEN,0) US_AVA" & vbCrLf _
                    & ", MS_ONHD, MS_PICK, NVL(MS_ONHD,0) - NVL(MS_PICK,0) MS_AVA" & vbCrLf _
                    & ", CASE WHEN MOD(NVL(ALLO,0), NVL(ICTSTYL1.CARTON_PACK_QTY,0)) = 0 THEN NVL(ALLO,0)" & vbCrLf _
                    & "       ELSE NVL(ALLO,0) +  NVL(ICTSTYL1.CARTON_PACK_QTY,0) - MOD(NVL(ALLO,0), NVL(ICTSTYL1.CARTON_PACK_QTY,0)) END QTY_TO_XFR" & vbCrLf _
                    & "from ICTSTYL1,ICTCOLR1, (" & vbCrLf _
                    & "Select SOTOXFR1.STYLE_CODE, SOTOXFR1.COLOR_CODE" & vbCrLf _
                    & ", Sum (ALLO) ALLO" & vbCrLf _
                    & "from SOTOXFR1" & vbCrLf _
                    & " where SOTOXFR1.OXFR_STATUS = '0' and SOTOXFR1.ALLO <> 0" & vbCrLf _
                    & "group by SOTOXFR1.STYLE_CODE, SOTOXFR1.COLOR_CODE) X" & vbCrLf _
                    & ", (Select STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND US_ONHD, WHSE_QTY_OPEN US_OPEN, WHSE_QTY_PICK US_PICK, WHSE_QTY_TRAN US_TRAN from ICTSTAT2 where WHSE_CODE = 'US') US" & vbCrLf _
                    & ", (Select STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND MS_ONHD, WHSE_QTY_OPEN MS_OPEN, WHSE_QTY_PICK MS_PICK, WHSE_QTY_TRAN MS_TRAN from ICTSTAT2 where WHSE_CODE = 'MS') MS" & vbCrLf _
                    & "where ICTSTYL1.STYLE_CODE = X.STYLE_CODE and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                    & "and US.STYLE_CODE (+) = X.STYLE_CODE and US.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                    & "and MS.STYLE_CODE (+) = X.STYLE_CODE and MS.COLOR_CODE (+) = X.COLOR_CODE"

            Case "SOTORDRX"
                SQL = "SELECT 'O' ORDR_TYPE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                    & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                    & ", SOTORDR1.CUST_CODE, SOTORDR1.ORDR_NO" & vbCrLf _
                    & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                    & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY ORDR" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_OPEN OPEN" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_PICK PICK" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_ALLO ALLO" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_SHIP SHIP" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_CANC CANC" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_OPEN" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_PICK" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
                    & ", SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_CANC" & vbCrLf _
                    & ", SOTORDR1.CUST_NAME" & vbCrLf _
                    & ", SOTORDR1.ORDR_DATE_RECD, SOTORDR1.INIT_DATE" & vbCrLf _
                    & " From SOTORDR2, SOTORDR1" & vbCrLf _
                    & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                    & $"   And (SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE from {SOTOXFRX})" & vbCrLf _
                    & "   And SOTORDR1.WHSE_CODE = 'MS' and SOTORDR2.ORDR_QTY_PICK > 0"


        End Select

        Return SQL

    End Function
    Private Sub ExportPickTckts(SHIP_BOL_NO As String)

        ASCMAIN1.sql = $"select TO_CHAR(SYSDATE, 'YYYYMMDD') as ""Date"", SOTPICK1.PICK_NO as ""P.O.#"", SOTORDR5.CUST_NAME Customer
                            , SOTORDR5.CUST_ADDR1 ""Ship To Addr 1"", SOTORDR5.CUST_ADDR2 ""Ship To Addr 2"", SOTORDR5.CUST_CITY
                            , SOTORDR5.CUST_STATE, SOTORDR5.CUST_ZIP_CODE, TATCNTRY.COUNTRY_CODE2  ""Country""
                            , SOTORDR2.STYLE_CODE || '-' || SOTORDR2.COLOR_CODE PRODUCT, '' LOT#, SOTPICK2.PICK_QTY QTY
                            , SOTORDR1.SHIP_VIA_CODE ""SHP Via"", ''  ACCT#, SOTORDR1.ORDR_SHIP_INSTR ""Ship Inst 1"",
                            case when RESIDENTIAL_ORDR = '1' then 'Residential Order ' end || 
                            case when INSIDE_REQ = '1' then 'Inside Delivery ' end ||
                            case when GATE_LIFT_REQ = '1' then 'Lift Gate Req ' end ||
                            case when LIMITED_ACCESS = '1' then 'Limited Access- ' || LIMITED_ACCESS_NOTE || ' ' end ||
                            case when IRREGULAR_HOURS = '1' then 'Hours- ' || IRREGULAR_HOURS_NOTE || ' ' end ||
                            case when APPOINTMENT_REQUIRED = '1' then 'Appointment Req- ' || APPOINTMENT_REQUIRED_NOTE || ' ' end ||
                            case when BROKER = '1' then 'Broker- ' || BROKER_NOTE || ' ' end
                            as ""Ship Inst 2"", '' ""Ship Inst 3"", '' ""Ship Inst 4""
                            from SOTPICK1, SOTORDR1, SOTORDR5, SOTPICK2, SOTORDR2, ARTCUSTQ, TATCNTRY
                            where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                            and SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO
                            and SOTORDR5.CUST_ADDR_TYPE = 'ST'
                            and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO
                            and SOTPICK2.PICK_QTY > 0
                            and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO
                            and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO
                            and ARTCUSTQ.CUST_CODE(+) = SOTORDR1.CUST_CODE
                            and ARTCUSTQ.CUST_ADDR_CODE(+) = SOTORDR1.CUST_STORE_NO
                            AND TATCNTRY.COUNTRY_CODE3(+) = SOTORDR5.CUST_COUNTRY
                            And SOTPICK1.SHIP_BOL_NO = '{SHIP_BOL_NO}'"


        Dim tblEXPORT As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        'this is USL specific so add USL to dir plus outbound dir and find a filename
        Dim csvFileName = $"Order_Standard{SHIP_BOL_NO}_" & DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".txt"
        Dim WorkDir = ASCMAIN1.Folders("Work")

        Dim ED_PARM_3PL_FTP_DIR As String = ROWs("EDTPARM1")("ED_PARM_3PL_FTP_DIR") & "USL\Order\"
        If ASCMAIN1.Running_in_VS Then
            ED_PARM_3PL_FTP_DIR = ASCMAIN1.Folders("Work")
        End If
        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        worksheet.Cells(0, 1).EntireColumn.NumberFormat = "@" ' PO#/PICK_NO - preserve leading zeros
        worksheet.Cells(0, 7).EntireColumn.NumberFormat = "@" ' ZipCode - preserve leading zeros
        worksheet.Cells(0, 13).EntireColumn.NumberFormat = "@" ' Acct# - preserve leading zeros
        Dim range As SpreadsheetGear.IRange = worksheet.Cells("A1")
        range.CopyFromDataTable(tblEXPORT, SpreadsheetGear.Data.SetDataFlags.None)
        workbook.SaveAs(WorkDir & csvFileName, SpreadsheetGear.FileFormat.UnicodeText)
        range = Nothing
        worksheet = Nothing
        workbook = Nothing

        If ASCMAIN1.Running_in_VS Then
            Show_Document(WorkDir & csvFileName)
        Else
            'Copy to sftp EDI machine for transmitting
            My.Computer.FileSystem.CopyFile(WorkDir & csvFileName, ED_PARM_3PL_FTP_DIR & csvFileName, True)
        End If


    End Sub

    Private Sub grdSOTOXFRX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTOXFRX.InitializeRow

        If e.Row.IsDataRow Then
            Dim UNITS_2_XFR As Int32 = Val(e.Row.Cells("UNITS_2_XFR").Value & "")
            Dim US_AVA As Int32 = Val(e.Row.Cells("US_AVA").Value & "")

            If UNITS_2_XFR > US_AVA Then
                e.Row.Cells("UNITS_2_XFR").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("UNITS_2_XFR").ToolTipText = "Units Ava in US is less than Units 2 Xfr"
            Else
                e.Row.Cells("UNITS_2_XFR").Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If


    End Sub
End Class