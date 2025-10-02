Imports Infragistics.Win.UltraWinGrid

Public Class SOFOXFR1

    Dim SOTOXFRX As String = ""
    Dim SOTORDRX As String = ""
    Dim TABLES_OXFR() As String = {"SOTORDR1", "SOTORDR2", "SOTORDR5", "SOTPICK0", "SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2"}

    Dim sqlSOTORDR0 As String
    Dim one_and_done As Boolean = False
    Dim sqlICTSTATQ As String = ""
    Dim bulk_transfer As Boolean = False


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
            Create_TDA(.Tables.Add, "SOTOXFRX", "**", 0, False, , 2)
            With .Tables("SOTOXFRX")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"

                .Columns.Add("UNITS_2_XFR", GetType(System.Int32), "IIF(ISNULL(SEL,'0') = '1', QTY_TO_XFR, 0)")
                .Columns.Add("CASES_2_XFR", GetType(System.Int32), "UNITS_2_XFR / ISNULL(CARTON_PACK_QTY,0)")
                .Columns.Add("CUBE_2_XFR", GetType(System.Decimal), "CASES_2_XFR * ISNULL(CASE_CUBE,0)")
                '.Columns.Add("NET_SHORT", GetType(System.Decimal), "IIF(SHORT + ISNULL(US_PICK,0) >= 0, NULL, SHORT + ISNULL(US_PICK,0))")
                .Columns.Add("NET_SHORT", GetType(System.Decimal), "IIF(NEEDED <= 0, NULL, NEEDED)")
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

            ASCMAIN1.sql = "Select SOTOXFR1.* from SOTOXFR1,SOTPICK1" & vbCrLf _
                & " where (SOTOXFR1.OXFR_STATUS = '0' or SOTOXFR1.OXFR_STATUS = '1')" & vbCrLf _
                & " and SOTPICK1.SHIP_BOL_NO (+) = SOTOXFR1.SHIP_BOL_NO" & vbCrLf _
                & " and NVL(SOTPICK1.PICK_STATUS,'?') <> 'F'" & vbCrLf _
                & " and SOTOXFR1.STYLE_CODE = :PARM1 and SOTOXFR1.COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add, "SOTOXFR1", "**", 0, True, "VV", 3, "OXFR_STATUS,LAST_DATE,LAST_OPER")
            With .Tables("SOTOXFR1")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
                '.Columns.Add("NET_SHORT", GetType(System.Decimal), "IIF(SHORT + ISNULL(US_PICK,0) >= 0, NULL, SHORT + ISNULL(US_PICK,0))")
                .Columns.Add("NET_SHORT", GetType(System.Decimal), "IIF(NEEDED <= 0, NULL, NEEDED)")
            End With

            ASCMAIN1.sql = "Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY" & vbCrLf _
                & ", MS.ONHD MS_ONHD, MS.PICK MS_PICK, NVL(MS.ONHD,0) - NVL(MS.PICK,0) MS_OTS" & vbCrLf _
                & ", US.ONHD US_ONHD, US.PICK US_PICK, NVL(US.ONHD,0) - NVL(US.PICK,0) US_OTS" & vbCrLf _
                & "" & vbCrLf _
                & " from ICTSTYC1, ICTSTYL1" & vbCrLf _
                & ",(SELECT STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND ONHD, WHSE_QTY_PICK PICK FROM ICTSTAT2 WHERE WHSE_CODE = 'MS') MS" & vbCrLf _
                & ",(SELECT STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND ONHD, WHSE_QTY_PICK PICK FROM ICTSTAT2 WHERE WHSE_CODE = 'US') US" & vbCrLf _
                & "where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" & vbCrLf _
                & "  and MS.STYLE_CODE (+) = ICTSTYC1.STYLE_CODE AND MS.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE" & vbCrLf _
                & "  and US.STYLE_CODE (+) = ICTSTYC1.STYLE_CODE AND US.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE" & vbCrLf
            ASCMAIN1.sql = $"Select * from ({ASCMAIN1.sql }) X where X.MS_OTS < 0"
            Create_TDA(.Tables.Add, "ICTSTATS", "**", 0, False, "", 2)
            With .Tables("ICTSTATS")
                .Columns.Add("NET_SHORT", GetType(System.Decimal), "IIF(ISNULL(MS_OTS,0) + ISNULL(US_PICK,0) >= 0, NULL, ISNULL(MS_OTS,0) + ISNULL(US_PICK,0))")
            End With

            ASCMAIN1.sql = "Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY" & vbCrLf _
                & ", MS.ONHD MS_ONHD, MS.PICK MS_PICK, NVL(MS.ONHD,0) - NVL(MS.PICK,0) MS_OTS" & vbCrLf _
                & ", US.ONHD US_ONHD, US.PICK US_PICK, NVL(US.ONHD,0) - NVL(US.PICK,0) US_OTS, (NVL(MS.ONHD,0) - NVL(MS.PICK,0)) + US.PICK MS_OTS_US_PICK" & vbCrLf _
                & "" & vbCrLf _
                & " from ICTSTYC1, ICTSTYL1" & vbCrLf _
                & ",(SELECT STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND ONHD, WHSE_QTY_PICK PICK FROM ICTSTAT2 WHERE WHSE_CODE = 'MS') MS" & vbCrLf _
                & ",(SELECT STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND ONHD, WHSE_QTY_PICK PICK FROM ICTSTAT2 WHERE WHSE_CODE = 'US') US" & vbCrLf _
                & "where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" & vbCrLf _
                & "  and MS.STYLE_CODE (+) = ICTSTYC1.STYLE_CODE AND MS.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE" & vbCrLf _
                & "  and US.STYLE_CODE (+) = ICTSTYC1.STYLE_CODE AND US.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE" & vbCrLf
            ASCMAIN1.sql = $"Select * from ({ASCMAIN1.sql }) X where X.US_OTS > 0"
            Create_TDA(.Tables.Add, "ICTXFRBL", "**", 0, False, "", 2)
            With .Tables("ICTXFRBL")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
                '.Columns.Add("NET_SHORT", GetType(System.Decimal), "ISNULL(US_ONHD,0) - ISNULL(US_PICK,0)")
                .Columns.Add("NET_SHORT", GetType(System.Decimal))
            End With

            ASCMAIN1.sql = "SELECT * FROM (" & vbCrLf _
                & "SELECT STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", SUM (WHSE_QTY_OPEN) WHSE_QTY_OPEN" & vbCrLf _
                & ", SUM (ALLO) ALLO, SUM (SHORT) SHORT" & vbCrLf _
                & ", SUM (COMING) COMING, SUM (NEEDED) NEEDED" & vbCrLf _
                & ", SUM (WHSE_QTY_OPEN) - SUM (NEEDED) DIFF" & vbCrLf _
                & "FROM (" & vbCrLf _
                & "SELECT WHSE_CODE, STYLE_CODE, COLOR_CODE, WHSE_QTY_OPEN, 0 ALLO, 0 SHORT, 0 COMING, 0 NEEDED" & vbCrLf _
                & "FROM ICTSTAT2 WHERE WHSE_CODE = 'US' AND WHSE_QTY_OPEN <> 0" & vbCrLf _
                & "UNION" & vbCrLf _
                & "SELECT 'US' WHSE_CODE, STYLE_CODE, COLOR_CODE, 0 WHSE_QTY_OPEN, ALLO, -1 * SHORT SHORT, COMING, NEEDED" & vbCrLf _
                & "FROM SOTOXFR1 WHERE OXFR_STATUS =  '0' AND NEEDED > 0" & vbCrLf _
                & ") GROUP BY STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ")"
            sqlICTSTATQ = ASCMAIN1.sql
            ASCMAIN1.sql = $"Select X.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY from ICTSTYL1, ({ASCMAIN1.sql}) X where ICTSTYL1.STYLE_CODE = X.STYLE_CODE"
            Create_TDA(.Tables.Add, "ICTSTATQ", "**", 0, False, "", 2)

        End With

        grdSOTOXFRX.DataSource = dst.Tables("SOTOXFRX")
        grdSOTOXFR1.DataSource = dst.Tables("SOTOXFR1")
        grdICTSTATS.DataSource = dst.Tables("ICTSTATS")
        grdICTSTATQ.DataSource = dst.Tables("ICTSTATQ")
        grdICTXFRBL.DataSource = dst.Tables("ICTXFRBL")

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
            ElseIf gcol.Key = "ALLO" Or gcol.Key = "COMING" Or gcol.Key = "NEEDED" Or gcol.Key = "QTY_TO_XFR" Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
            ElseIf gcol.Key = "NET_SHORT" Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.PaleVioletRed
            ElseIf gcol.Key = "UNITS_2_XFR" Or gcol.Key = "CASES_2_XFR" Or gcol.Key = "CUBE_2_XFR" Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
            End If
        Next

        For Each grd As UltraGrid In New UltraGrid() {grdICTSTATS, grdICTXFRBL}
            For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key.StartsWith("US_") Then
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                ElseIf gcol.Key.StartsWith("MS_") Then
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf gcol.Key = "NET_SHORT" Then
                    If grd.Name = "grdICTSTATS" Then
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.PaleVioletRed
                    Else
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Lavender
                        gcol.Header.Caption = "XFR Qty"
                    End If
                End If
            Next
            Show_Filter(grd, True)
            Create_Summary(grd, "STYLE_CODE", "Count")
        Next

        With grdICTXFRBL.DisplayLayout
            .Override.AllowUpdate = DefaultableBoolean.True
            With .Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns
                    If c.Key = "SEL" Or c.Key = "NET_SHORT" Then
                        .Columns(c.Key).CellActivation = Activation.AllowEdit
                    Else
                        .Columns(c.Key).CellActivation = Activation.NoEdit
                    End If
                Next
            End With
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdICTSTATQ.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = System.Drawing.Color.White
            gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            If gcol.Key = ("STYLE_CODE") Or gcol.Key = ("STYLE_DESC") Or gcol.Key = ("CARTON_PACK_QTY") Or gcol.Key = ("COLOR_CODE") Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
            ElseIf gcol.Key = ("DIFF") Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.PaleVioletRed
            End If
        Next

        Show_Filter(grdICTSTATQ, True)
        Create_Summary(grdICTSTATQ, "STYLE_CODE", "Count")
        Create_Summary(grdICTSTATQ, New String() {"WHSE_QTY_OPEN", "ALLO", "SHORT", "DIFF"})

        With grdSOTOXFR1.DisplayLayout.Override
            .AllowAddNew = AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTOXFR1.DisplayLayout.Bands(0).Columns
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
            ElseIf gcol.Key = "NET_SHORT" Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.PaleVioletRed
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

        ASCMAIN1.Add_Value_List(grdSOTOXFR1, "OXFR_STATUS", Nothing, New String() {":", "0:Pending Xfr", "1:Sent to USL"})

        tab1.Tabs("US Qty Open Diagnostic").Visible = ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz"

        MakeTransparent(chkShowOnlyNetShort)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load"
                Select Case optAction.Value
                    Case "B"
                        'If dst.Tables("ICTXFRBL").Select("SEL = '1'").Length = 0 Then
                        '    EMsg &= vbCr & "You must select 1 or more Style/Colors"
                        'End If
                    Case "D"
                    Case Else
                        If one_and_done Then
                            EMsg &= vbCr & "Please exit and re-enter this screen to start another Transfer"
                        End If
                End Select



            Case "Update"

                If optAction.Value = "D" Then

                    Dim selRows() As DataRow = dst.Tables("SOTOXFR1").Select("SEL='1'")

                    If selRows.Length = 0 Then
                        EMsg &= vbCr & "Nothing Selected to Delete"
                    End If

                    If EMsg = "" Then
                        If MsgBox($"OK to Delete {selRows.Length} Records with False Demand?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                Else

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

                '.Groups("Action").Enabled = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpAction, ScreenMode)

        SplitContainer1.Visible = ScreenMode
        SplitContainer2.Visible = Not ScreenMode


        If ScreenMode Then
            grdSOTOXFRX.Parent = SplitContainer1.Panel1
            'grdSOTOXFRX.DisplayLayout.Bands(0).Columns("SEL").Hidden = False

            grdSOTOXFRX.DisplayLayout.Bands(0).Columns("SEL").Hidden = (optAction.Value = "D")
            grdSOTOXFR1.DisplayLayout.Bands(0).Columns("SEL").Hidden = (optAction.Value = "X")

            If (optAction.Value = "D") Then
                grdSOTOXFRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTOXFR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            Else
                grdSOTOXFRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTOXFR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            End If
        Else
            Clear_Record()
            grdSOTOXFRX.Parent = SplitContainer2.Panel1
            grdSOTOXFRX.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
            bulk_transfer = False
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"SOTOXFRX", "SOTORDRX", "SOTOXFR1", "ICTSTATS", "ICTXFRBL"}
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

        bulk_transfer = (optAction.Value = "B")

        'Dim sqlOTS As String = "ISNULL(MS_AVA,0) >= ISNULL(SHORT,0)"
        If Not bulk_transfer Then
            Dim sqlOTS As String = "ISNULL(MS_AVA,0) >= ISNULL(ALLO,0)"

            Dim rowsWithNewOTS() As DataRow = dst.Tables("SOTOXFRX").Select(sqlOTS)
            If rowsWithNewOTS.Length > 0 Then
                If MsgBox($"There are {CStr(rowsWithNewOTS.Length)} Style/Color(s) with new OTS positions in MS that satisfy the Shortage" & vbCrLf & vbCrLf & "OK to clear those shortages?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    BeginTrans()
                    For Each rowSOTOXFRX As DataRow In rowsWithNewOTS
                        Dim US_OPEN As Int32 = Val(rowSOTOXFRX.Item("US_OPEN") & "")
                        Dim STYLE_CODE As String = rowSOTOXFRX.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTOXFRX.Item("COLOR_CODE")

                        TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, "US", "WHSE_QTY_OPEN", -1 * US_OPEN)

                        ' Update Status of Transfer Queue Records
                        ASCMAIN1.sql = $"Update SOTOXFR1 SET OXFR_STATUS = 'M', LAST_DATE = SYSDATE, LAST_OPER = '{ASCMAIN1.USER_ID}'" & vbCrLf _
                        & " where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2 and OXFR_STATUS = '0'"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {STYLE_CODE, COLOR_CODE})

                        rowSOTOXFRX.Delete()
                    Next
                    dst.Tables("SOTOXFRX").AcceptChanges()

                    CommitTrans()
                End If
            End If
        End If
        dst.Tables("SOTOXFR1").Rows.Clear()
        For Each rowSOTOXFRX As DataRow In dst.Tables("SOTOXFRX").Select("")
            Dim STYLE_CODE As String = rowSOTOXFRX.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTOXFRX.Item("COLOR_CODE")
            Fill_Records("SOTOXFR1", New String() {STYLE_CODE, COLOR_CODE}, False)
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record(Optional publish As Boolean = False, Optional showCommitMsg As Boolean = True)

        BeginTrans()

        ' Reduce Qty OPEN by the Qty ALLO for all selected SCs to Transfer

        Dim WHSE_CODE As String = "US"

        If optAction.Value = "D" Then ' Delete False Demand

            For Each rowSOTOXFR1 As DataRow In dst.Tables("SOTOXFR1").Select("SEL = '1'")
                Dim STYLE_CODE As String = rowSOTOXFR1.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTOXFR1.Item("COLOR_CODE")
                'Dim QTY As Int32 = Val(rowSOTOXFRX.Item("SHORT"))
                Dim QTY As Int32 = -1 * Val(rowSOTOXFR1.Item("NEEDED"))
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY)

                rowSOTOXFR1.Item("OXFR_STATUS") = "D"
                rowSOTOXFR1.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTOXFR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            Next

            Update_Record_TDA("SOTOXFR1")

        Else ' Create Transfer Order

            TAC.SOCMAIN1.Create_Transfer_Order(Me, dst.Tables("SOTOXFRX").Select("SEL = '1'"), "UNITS_2_XFR") ' Create a Single XFR Order for the selected SCs to Transfer
            TAC.SOCMAIN1.Release_Transfer_Order(Me) ' Release that XFR Order

            one_and_done = True

            For Each TABLE_NAME As String In TABLES_OXFR
                Update_Record_TDA(TABLE_NAME)
            Next

            Dim SHIP_BOL_NO As String = dst.Tables("SOTSHIP1").Rows(0).Item("SHIP_BOL_NO")
            Dim ORDR_GROUP_NO As String = dst.Tables("SOTSHIP1").Rows(0).Item("ORDR_GROUP_NO")
            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

            For Each rowSOTOXFRX As DataRow In dst.Tables("SOTOXFRX").Select("SEL = '1'")
                Dim STYLE_CODE As String = rowSOTOXFRX.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTOXFRX.Item("COLOR_CODE")
                'Dim QTY As Int32 = Val(rowSOTOXFRX.Item("SHORT"))
                Dim QTY As Int32 = -1 * Val(rowSOTOXFRX.Item("NEEDED"))
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY)

                ' Update Status of Transfer Queue Records
                ASCMAIN1.sql = $"Update SOTOXFR1 SET OXFR_STATUS = '1', SHIP_BOL_NO = '{SHIP_BOL_NO}', LAST_DATE = SYSDATE, LAST_OPER = '{ASCMAIN1.USER_ID}'" & vbCrLf _
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
        End If

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
        Load_Popup_Menu(grdICTSTATS, "SSBB", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Add to Transfer Queue")
        Load_Popup_Menu(grdICTXFRBL, "SSBB", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Bulk Transfer Selected")

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
                Case "grdSOTOXFRX"
                    For Each tt As String In New String() {"Select All", "De-Select All", "Select Selected", "De-Select Selected"}
                        tlb_pop.Tools(tt).SharedProps.Visible = ScreenMode And (optAction.Value = "X")
                    Next

                Case "grdICTSTATS"
                    tlb_pop.Tools("Add to Transfer Queue").SharedProps.Visible = chkShowOnlyNetShort.Checked
                Case "grdICTXFRBL"
                    tlb_pop.Tools("Bulk Transfer Selected").SharedProps.Visible = (dst.Tables("ICTXFRBL").Select("SEL = '1'").Length > 0)
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
            Case "Bulk Transfer Selected"
                bulk_transfer = True
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    Dim SEL As String = grow.Cells("SEL").Value & ""
                    grow.Selected = (SEL = "1")
                Next

                'If grd.Selected.Rows.Count = 0 Then
                '    If grd.ActiveRow IsNot Nothing Then
                '        grd.ActiveRow.Selected = True
                '    End If
                'End If

                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("No Rows Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                If MsgBox($"OK to Queue up {grd.Selected.Rows.Count} Style / Colors selected?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                dst.Tables("SOTOXFR1").Rows.Clear()

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
                    Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value
                    Dim US_ONHD As Int32 = Val(grow.Cells("US_ONHD").Value & "")
                    Dim NET_SHORT As Int32 = Val(grow.Cells("NET_SHORT").Value & "")

                    If dst.Tables("SOTOXFRX").Rows.Find(New String() {STYLE_CODE, COLOR_CODE}) Is Nothing Then
                    Else
                        MsgBox($"Style / Color {STYLE_CODE} / {COLOR_CODE} is already in the Transfer Queue", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Exit Sub
                    End If

                    Dim rowSOTOXFR1 As DataRow = dst.Tables("SOTOXFR1").NewRow
                    With rowSOTOXFR1
                        Dim PICK_BATCH_NO As String = ""
                        .Item("PICK_BATCH_NO") = "Q" & Mid(ASCMAIN1.Next_Control_No("SOTOXFR1.PICK_BATCH_NO_Q"), 6, 5)
                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE

                        .Item("US_ONHD") = US_ONHD
                        .Item("US_PICK") = grow.Cells("US_PICK").Value
                        .Item("US_AVA") = grow.Cells("US_OTS").Value

                        .Item("MS_ONHD") = grow.Cells("MS_ONHD").Value
                        .Item("MS_PICK") = grow.Cells("MS_PICK").Value
                        .Item("MS_AVA") = grow.Cells("MS_OTS").Value

                        .Item("ALLO") = NET_SHORT
                        .Item("SHORT") = 0
                        .Item("OXFR_STATUS") = "0"
                        .Item("COMING") = Val(grow.Cells("US_PICK").Value & "") + Val(grow.Cells("US_PICK").Value & "")
                        .Item("NEEDED") = NET_SHORT

                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    End With
                    dst.Tables("SOTOXFR1").Rows.Add(rowSOTOXFR1)
                Next

                Update_Record_TDA("SOTOXFR1")

                MsgBox("Transfer Queue Record(s) Added - Transfer Queue will be Refreshed")

                Refresh_Documents()
            Case "Add to Transfer Queue"

                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("No Rows Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                If MsgBox($"OK to Queue up {grd.Selected.Rows.Count} Style / Colors selected?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                dst.Tables("SOTOXFR1").Rows.Clear()

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
                    Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value
                    Dim US_ONHD As Int32 = Val(grow.Cells("US_ONHD").Value & "")
                    Dim NET_SHORT As Int32 = Val(grow.Cells("NET_SHORT").Value & "")

                    If dst.Tables("SOTOXFRX").Rows.Find(New String() {STYLE_CODE, COLOR_CODE}) Is Nothing Then
                    Else
                        MsgBox($"Style / Color {STYLE_CODE} / {COLOR_CODE} is already in the Transfer Queue", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Exit Sub
                    End If

                    Dim rowSOTOXFR1 As DataRow = dst.Tables("SOTOXFR1").NewRow
                    With rowSOTOXFR1
                        Dim PICK_BATCH_NO As String = ""
                        .Item("PICK_BATCH_NO") = "Q" & Mid(ASCMAIN1.Next_Control_No("SOTOXFR1.PICK_BATCH_NO_Q"), 6, 5)
                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE

                        .Item("US_ONHD") = grow.Cells("US_ONHD").Value
                        .Item("US_PICK") = grow.Cells("US_PICK").Value
                        .Item("US_AVA") = grow.Cells("US_OTS").Value

                        .Item("MS_ONHD") = grow.Cells("MS_ONHD").Value
                        .Item("MS_PICK") = grow.Cells("MS_PICK").Value
                        .Item("MS_AVA") = grow.Cells("MS_OTS").Value

                        .Item("ALLO") = -1 * NET_SHORT
                        .Item("SHORT") = NET_SHORT
                        .Item("OXFR_STATUS") = "0"
                        .Item("COMING") = Val(grow.Cells("US_PICK").Value & "") + Val(grow.Cells("US_PICK").Value & "")
                        .Item("NEEDED") = -1 * NET_SHORT

                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    End With
                    dst.Tables("SOTOXFR1").Rows.Add(rowSOTOXFR1)
                Next

                Update_Record_TDA("SOTOXFR1")

                MsgBox("Transfer Queue Record(s) Added - Transfer Queue will be Refreshed")

                Refresh_Documents()

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

        Fill_Records("ICTSTATS")
        Sort_grdColumns(grdICTSTATS, "STYLE_CODE, COLOR_CODE")
        ShowOnlyNetShort()

        Fill_Records("ICTSTATQ")
        Sort_grdColumns(grdICTSTATQ, "STYLE_CODE,COLOR_CODE")

        Fill_Records("ICTXFRBL")
        Sort_grdColumns(grdICTXFRBL, "STYLE_CODE, COLOR_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdSOTOXFRX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTOXFRX.AfterRowActivate

        If grdSOTOXFRX.ActiveRow Is Nothing OrElse Not grdSOTOXFRX.ActiveRow.IsDataRow Then
            grdSOTORDRX.Visible = False
            grdSOTOXFR1.Visible = False
        Else

            Dim STYLE_CODE As String = grdSOTOXFRX.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdSOTOXFRX.ActiveRow.Cells("COLOR_CODE").Value

            grdSOTORDRX.Text = $"Released Sales Orders with Style / Color {STYLE_CODE} / {COLOR_CODE}"
            Fill_Records("SOTORDRX", New String() {STYLE_CODE, COLOR_CODE})
            grdSOTORDRX.Visible = True

            grdSOTOXFR1.Text = $"Release Batches calling for Style / Color {STYLE_CODE} / {COLOR_CODE}"
            Dim dvw As DataView = dst.Tables("SOTOXFR1").DefaultView
            dvw.RowFilter = $"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'"
            grdSOTOXFR1.Visible = True

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
            'ASCDATA1.ExecuteSQL($"Truncate Table {SOTOXFRX}")
            ASCDATA1.ExecuteSQL($"Delete from {SOTOXFRX}")
            ASCDATA1.ExecuteSQL($"Insert into {SOTOXFRX} {Get_SQL("SOTOXFRX")}")

            'ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDRX}")
            ASCDATA1.ExecuteSQL($"Delete from {SOTORDRX}")
            ASCDATA1.ExecuteSQL($"Insert into {SOTORDRX} {Get_SQL("SOTORDRX")}")
            'ASCDATA1.ExecuteSQL($"Update {SOTORDRX} SOTORDRX Set LINES_OPEN = (Select Count (*) from SOTORDR2 where ORDR_NO = SOTORDRX.ORDR_NO and ORDR_QTY_OPEN > 0)")
            'ASCDATA1.ExecuteSQL($"Update {SOTORDRX} SOTORDRX Set LINES_US = (Select Count (*) from SOTORDR2 where ORDR_NO = SOTORDRX.ORDR_NO and ORDR_QTY_OPEN > 0 and (STYLE_CODE, COLOR_CODE) in (Select STYLE_CODE, COLOR_CODE from {SOTOXFRX}))")

        End If
    End Sub

    Function Get_SQL(TABLE_NAME As String) As String

        Dim SQL As String = ""

        Select Case TABLE_NAME

            Case "SOTOXFRX"
                Dim QTY_TO_XFR_sql As String = "CASE WHEN MOD(NVL(NEEDED,0), NVL(ICTSTYL1.CARTON_PACK_QTY,0)) = 0 THEN NVL(NEEDED,0)" & vbCrLf _
                    & "       ELSE NVL(NEEDED,0) +  NVL(ICTSTYL1.CARTON_PACK_QTY,0) - MOD(NVL(NEEDED,0), NVL(ICTSTYL1.CARTON_PACK_QTY,0)) END"
                If bulk_transfer Then
                    QTY_TO_XFR_sql = " NVL(US_ONHD,0) "
                End If
                SQL = "Select X.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                    & ", US_TRAN, US_ONHD, US_PICK, US_OPEN, NVL(US_ONHD,0) - NVL(US_PICK,0) US_AVA" & vbCrLf _
                    & ", MS_ONHD, MS_PICK, NVL(MS_ONHD,0) - NVL(MS_PICK,0) MS_AVA" & vbCrLf _
                    & $", {QTY_TO_XFR_sql} QTY_TO_XFR" & vbCrLf _
                    & "from ICTSTYL1,ICTCOLR1, (" & vbCrLf _
                    & "Select SOTOXFR1.STYLE_CODE, SOTOXFR1.COLOR_CODE" & vbCrLf _
                    & ", Sum (ALLO) ALLO" & vbCrLf _
                    & ", Sum (SHORT) SHORT" & vbCrLf _
                    & ", Sum (COMING) COMING" & vbCrLf _
                    & ", Sum (NEEDED) NEEDED" & vbCrLf _
                    & "from SOTOXFR1" & vbCrLf _
                    & " where SOTOXFR1.OXFR_STATUS = '0' and SOTOXFR1.NEEDED <> 0" & vbCrLf _
                    & "group by SOTOXFR1.STYLE_CODE, SOTOXFR1.COLOR_CODE) X" & vbCrLf _
                    & ", (Select STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND US_ONHD, WHSE_QTY_OPEN US_OPEN, WHSE_QTY_PICK US_PICK, WHSE_QTY_TRAN US_TRAN from ICTSTAT2 where WHSE_CODE = 'US') US" & vbCrLf _
                    & ", (Select STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND MS_ONHD, WHSE_QTY_OPEN MS_OPEN, WHSE_QTY_PICK MS_PICK, WHSE_QTY_TRAN MS_TRAN from ICTSTAT2 where WHSE_CODE = 'MS') MS" & vbCrLf _
                    & "where ICTSTYL1.STYLE_CODE = X.STYLE_CODE and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                    & "and US.STYLE_CODE (+) = X.STYLE_CODE and US.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                    & "and MS.STYLE_CODE (+) = X.STYLE_CODE and MS.COLOR_CODE (+) = X.COLOR_CODE"

                'SQL = "Select X.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                '    & ", US_TRAN, US_ONHD, US_PICK, US_OPEN, NVL(US_ONHD,0) - NVL(US_PICK,0) US_AVA" & vbCrLf _
                '    & ", MS_ONHD, MS_PICK, NVL(MS_ONHD,0) - NVL(MS_PICK,0) MS_AVA" & vbCrLf _
                '    & ", CASE WHEN MOD(NVL(ALLO,0), NVL(ICTSTYL1.CARTON_PACK_QTY,0)) = 0 THEN NVL(ALLO,0)" & vbCrLf _
                '    & "       ELSE NVL(ALLO,0) +  NVL(ICTSTYL1.CARTON_PACK_QTY,0) - MOD(NVL(ALLO,0), NVL(ICTSTYL1.CARTON_PACK_QTY,0)) END QTY_TO_XFR" & vbCrLf _
                '    & "from ICTSTYL1,ICTCOLR1, (" & vbCrLf _
                '    & "Select SOTOXFR1.STYLE_CODE, SOTOXFR1.COLOR_CODE" & vbCrLf _
                '    & ", Sum (ALLO) ALLO" & vbCrLf _
                '    & ", Sum (SHORT) SHORT" & vbCrLf _
                '    & "from SOTOXFR1" & vbCrLf _
                '    & " where SOTOXFR1.OXFR_STATUS = '0' and SOTOXFR1.ALLO <> 0" & vbCrLf _
                '    & "group by SOTOXFR1.STYLE_CODE, SOTOXFR1.COLOR_CODE) X" & vbCrLf _
                '    & ", (Select STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND US_ONHD, WHSE_QTY_OPEN US_OPEN, WHSE_QTY_PICK US_PICK, WHSE_QTY_TRAN US_TRAN from ICTSTAT2 where WHSE_CODE = 'US') US" & vbCrLf _
                '    & ", (Select STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND MS_ONHD, WHSE_QTY_OPEN MS_OPEN, WHSE_QTY_PICK MS_PICK, WHSE_QTY_TRAN MS_TRAN from ICTSTAT2 where WHSE_CODE = 'MS') MS" & vbCrLf _
                '    & "where ICTSTYL1.STYLE_CODE = X.STYLE_CODE and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                '    & "and US.STYLE_CODE (+) = X.STYLE_CODE and US.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                '    & "and MS.STYLE_CODE (+) = X.STYLE_CODE and MS.COLOR_CODE (+) = X.COLOR_CODE"

                'SQL = "Select X.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                '    & ", US_TRAN, US_ONHD, US_PICK, US_OPEN, NVL(US_ONHD,0) - NVL(US_PICK,0) US_AVA" & vbCrLf _
                '    & ", MS_ONHD, MS_PICK, NVL(MS_ONHD,0) - NVL(MS_PICK,0) MS_AVA" & vbCrLf _
                '    & ", CASE WHEN MOD(NVL(SHORT,0), NVL(ICTSTYL1.CARTON_PACK_QTY,0)) = 0 THEN NVL(SHORT,0)" & vbCrLf _
                '    & "       ELSE NVL(SHORT,0) +  NVL(ICTSTYL1.CARTON_PACK_QTY,0) - MOD(NVL(SHORT,0), NVL(ICTSTYL1.CARTON_PACK_QTY,0)) END QTY_TO_XFR" & vbCrLf _
                '    & "from ICTSTYL1,ICTCOLR1, (" & vbCrLf _
                '    & "Select SOTOXFR1.STYLE_CODE, SOTOXFR1.COLOR_CODE" & vbCrLf _
                '    & ", Sum (-1 * SHORT) SHORT" & vbCrLf _
                '    & ", Sum (ALLO) ALLO" & vbCrLf _
                '    & "from SOTOXFR1" & vbCrLf _
                '    & " where SOTOXFR1.OXFR_STATUS = '0' and -1 * SOTOXFR1.SHORT <> 0 and SOTOXFR1.ALLO <> 0" & vbCrLf _
                '    & "group by SOTOXFR1.STYLE_CODE, SOTOXFR1.COLOR_CODE) X" & vbCrLf _
                '    & ", (Select STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND US_ONHD, WHSE_QTY_OPEN US_OPEN, WHSE_QTY_PICK US_PICK, WHSE_QTY_TRAN US_TRAN from ICTSTAT2 where WHSE_CODE = 'US') US" & vbCrLf _
                '    & ", (Select STYLE_CODE, COLOR_CODE, WHSE_QTY_ON_HAND MS_ONHD, WHSE_QTY_OPEN MS_OPEN, WHSE_QTY_PICK MS_PICK, WHSE_QTY_TRAN MS_TRAN from ICTSTAT2 where WHSE_CODE = 'MS') MS" & vbCrLf _
                '    & "where ICTSTYL1.STYLE_CODE = X.STYLE_CODE and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                '    & "and US.STYLE_CODE (+) = X.STYLE_CODE and US.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                '    & "and MS.STYLE_CODE (+) = X.STYLE_CODE and MS.COLOR_CODE (+) = X.COLOR_CODE"

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
                    & " and SOTORDR1.ORDR_NO <> '0000865352'" & vbCrLf _
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

    Private Sub grdSOTOXFR1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTOXFR1.InitializeRow
        Dim OXFR_STATUS As String = e.Row.Cells("OXFR_STATUS").Value
        If OXFR_STATUS = "1" Then
            e.Row.Appearance.ForeColor = System.Drawing.Color.Blue
            e.Row.Cells("SEL").Appearance.BackColor = Drawing.Color.Blue
        Else
            e.Row.Appearance.ForeColor = System.Drawing.Color.Empty
            e.Row.Cells("SEL").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdSOTOXFR1_BeforeRowActivate(sender As Object, e As RowEventArgs) Handles grdSOTOXFR1.BeforeRowActivate
        Dim OXFR_STATUS As String = e.Row.Cells("OXFR_STATUS").Value
        With grdSOTOXFR1.DisplayLayout.Bands(0).Columns("SEL")
            If OXFR_STATUS = "1" Then
                .CellActivation = Activation.NoEdit
            Else
                .CellActivation = Activation.AllowEdit
            End If
        End With

    End Sub

    Private Sub grdICTSTATS_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTSTATS.InitializeRow
        If e.Row.IsDataRow Then
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value
            If dst.Tables("SOTOXFRX").Rows.Find(New String() {STYLE_CODE, COLOR_CODE}) Is Nothing Then
                e.Row.Appearance.ForeColor = System.Drawing.Color.Red
            Else
                e.Row.Appearance.ForeColor = System.Drawing.Color.Empty
            End If
        End If
    End Sub

    Private Sub chkShowOnlyNetShort_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowOnlyNetShort.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        ShowOnlyNetShort()
    End Sub

    Sub ShowOnlyNetShort()
        Dim dvw As DataView = dst.Tables("ICTSTATS").DefaultView
        If chkShowOnlyNetShort.Checked Then
            dvw.RowFilter = "US_ONHD > 0 AND ISNULL(NET_SHORT,0) < 0"
        Else
            dvw.RowFilter = ""
        End If
    End Sub

    Private Sub cmdFixOpen_Click(sender As Object, e As EventArgs) Handles cmdFixOpen.Click

        If MsgBox("OK to Reset Qty Open?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        ASCMAIN1.sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS" & vbCrLf _
            & sqlICTSTATQ & vbCrLf _
            & " WHERE DIFF <> 0;" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "UPDATE ICTSTAT2 SET WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) - R1.DIFF" & vbCrLf _
            & "WHERE WHSE_CODE = 'US' AND STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        Fill_Records("ICTSTATQ")
        Sort_grdColumns(grdICTSTATQ, "STYLE_CODE,COLOR_CODE")
    End Sub

    Private Sub tab1_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab1.SelectedTabChanged

    End Sub

    Private Sub grdICTXFRBL_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdICTXFRBL.ClickCellButton

    End Sub

    Private Sub grdICTXFRBL_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTXFRBL.BeforeRowUpdate
        Dim NET_SHORT As Decimal = Val(e.Row.Cells("NET_SHORT").Value & "")
        Dim US_ONHD As Decimal = Val(e.Row.Cells("US_ONHD").Value & "")
        Dim US_PICK As Decimal = Val(e.Row.Cells("US_PICK").Value & "")
        Dim XFR_AVAILABLE As Decimal = US_ONHD - US_PICK

        If NET_SHORT > XFR_AVAILABLE Or NET_SHORT < 0 Then
            e.Cancel = True
        End If
    End Sub
    Sub Calculate_Selected_Cube()

    End Sub

    Private Sub grdICTXFRBL_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdICTXFRBL.ClickCell
        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim SEL As String = e.Cell.Row.Cells("SEL").Text

        Select Case COLUMN_NAME
            Case "SEL"
                Dim US_ONHD As Decimal = Val(e.Cell.Row.Cells("US_ONHD").Value & "")
                Dim US_PICK As Decimal = Val(e.Cell.Row.Cells("US_PICK").Value & "")
                e.Cell.Row.Cells("NET_SHORT").Value = If(SEL = "1", US_ONHD - US_PICK, 0)
                Calculate_Selected_Cube()
        End Select
    End Sub
End Class