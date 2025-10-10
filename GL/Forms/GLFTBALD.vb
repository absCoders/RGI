Public Class GLFTBALD
    Dim SQL As New Text.StringBuilder With {.Length = 0}
    'Dim ACCT_CODE As String = "1000"
    'Dim SEG2_CODE As String = "000"
    'Dim SEG3_CODE As String = "00"
    Dim CURR_PERIOD As String = ""
    Dim CURR_YEAR As String = ""
    Dim CURR_MO As String = ""
    Dim SEL_PERIOD As String = ""
    Dim SEL_PERIOD_BEG As Date = Now()
    Dim SEL_PERIOD_END As Date = Now()
    Dim SEL_YEAR As String = ""
    Dim SEL_MO As String = ""
    Dim DAY_COLS As New Dictionary(Of String, Date)
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")

        With dst
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("G1.ACCT_CODE,")
            SQL.AppendLine("G3.SEG2_CODE,")
            SQL.AppendLine("G3.SEG3_CODE,")
            SQL.AppendLine("G1.ACCT_TYPE,")
            SQL.AppendLine("G1.ACCT_DESC,")
            SQL.AppendLine("0.00 ACCT_BEG_BAL")
            SQL.AppendLine("FROM GLTACCT1 G1, GLTACCT3 G3")
            SQL.AppendLine("WHERE G1.ACCT_CODE = G3.ACCT_CODE")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "GLTTBALX", "**", 0, False)
            With .Tables("GLTTBALX").Columns
                For i As Int64 = 1 To 31
                    Dim DAY_COL As String = $"DAY{Format(i, "00")}"
                    .Add(DAY_COL, GetType(System.Double))
                Next
            End With

            SQL.Length = 0
            SQL.AppendLine("SELECT *")
            SQL.AppendLine(" from GLTTBALD")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "GLTTBALD", "**", 0, True)

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("OPS_YYYYPP,")
            SQL.AppendLine("ACCT_CODE,")
            SQL.AppendLine("SEG2_CODE,")
            SQL.AppendLine("SEG3_CODE,")
            SQL.AppendLine("DETL_CTL_DATE,")
            SQL.AppendLine("DETL_POSTING_AMT")
            SQL.AppendLine(" from GLTDETL1")
            SQL.AppendLine("WHERE OPS_YYYYPP = :PARM1")
            SQL.AppendLine("AND ACCT_CODE = :PARM2")
            SQL.AppendLine("AND SEG2_CODE = :PARM3")
            SQL.AppendLine("AND SEG3_CODE = :PARM4")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "GLTDETL1", "**", 0, False, "VVVV")

        End With

        grdGLTTBALX.DataSource = dst.Tables("GLTTBALX")
        grdGLFTBALD.DataSource = dst.Tables("GLTTBALD")

        Fill_Records("GLTTBALD")

        'Create_Summary(grdGLTTBALX, New String() {"ACCT_BEG_BAL", "ACCT_DR", "ACCT_CR", "ACCT_END_BAL", "ACCT_TRANS", "P01", "P02", "P03", "P04", "P05", "P06", "P07", "P08", "P09", "P10", "P11", "P12"})

        'With grdGLTTBALX.DisplayLayout
        'For Each COLUMN_NAME As String In New String() _
        '    {"ACCT_BEG_BAL", "ACCT_DR", "ACCT_CR", "ACCT_END_BAL", "ACCT_TRANS"}
        '    With .Columns(COLUMN_NAME)
        '        .Width = IIf(COLUMN_NAME = "ACCT_TRANS", 60, 110)
        '        .Header.Appearance.BackColor = Drawing.Color.White
        '        .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
        '        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '    End With
        'Next
        'For P As Integer = 1 To 12
        '    With .Columns("P" & Format(P, "00"))
        '        .Width = 110
        '        .Header.Appearance.BackColor = Drawing.Color.White
        '        .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
        '        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '    End With
        'Next
        'End With

    End Sub
    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub
    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Refresh"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Update_Record()
                Call Mode_Settings(False)
                Me.Close()
        End Select

    End Sub
    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub
    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GLTTBALX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        'Absx1.txtFor("OPS_YYYYPP").Text = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")

        Dim CYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        ASCMAIN1.Period_Calc(CYP, 1)
        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.Period_Calc(CYP, 1)
    End Sub
    Sub Load_Record()

        CalcPeriods()

        ASCMAIN1.Progress("Now Loading Summary Data")

        setDAY_COLS()
        'grdGLTTBALX.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Dim ACCT_YEAR As String = Mid$(HFs("OPS_YYYYPP"), 1, 4)
        Dim P As Integer = Val(Mid$(HFs("OPS_YYYYPP"), 5, 2))

        dst.Tables.Item("GLTTBALX").Clear()
        For Each rowGLTTBALD As DataRow In dst.Tables("GLTTBALD").Select()
            Dim ACCT_CODE As String = rowGLTTBALD.Item("ACCT_CODE").ToString & String.Empty
            Dim SEG2_CODE As String = rowGLTTBALD.Item("SEG2_CODE").ToString & String.Empty
            Dim SEG3_CODE As String = rowGLTTBALD.Item("SEG3_CODE").ToString & String.Empty
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("G1.ACCT_CODE,")
            SQL.AppendLine("G3.SEG2_CODE,")
            SQL.AppendLine("G3.SEG3_CODE,")
            SQL.AppendLine("G1.ACCT_TYPE,")
            SQL.AppendLine("G1.ACCT_DESC,")
            SQL.AppendLine("0.00 ACCT_BEG_BAL,")
            For i As Int64 = 1 To 31
                Dim DAY_COL As String = $"DAY{Format(i, "00")}"
                If i = 31 Then
                    SQL.AppendLine($"0.00 {DAY_COL}")
                Else
                    SQL.AppendLine($"0.00 {DAY_COL},")
                End If
            Next
            SQL.AppendLine("FROM GLTACCT1 G1, GLTACCT3 G3")
            SQL.AppendLine("WHERE G1.ACCT_CODE = G3.ACCT_CODE")
            SQL.AppendLine($"AND G1.ACCT_CODE = '{ACCT_CODE}'")
            SQL.AppendLine($"AND G3.SEG2_CODE = '{SEG2_CODE}'")
            SQL.AppendLine($"AND G3.SEG3_CODE = '{SEG3_CODE}'")
            SQL.AppendLine($"AND G3.ACCT_YEAR = '{ACCT_YEAR}'")
            Fill_Records("GLTTBALX",, False, SQL.ToString)
        Next

        For Each rowGLTTBALX As DataRow In dst.Tables("GLTTBALX").Select()
            calcBegBal(rowGLTTBALX)
            calcDays(rowGLTTBALX)
            hideDays()
        Next

        EnforceConstraints(True)

        With grdGLTTBALX.DisplayLayout.Bands(0)
            For i As Integer = 1 To 31
                Dim DAY_COL As String = $"DAY{Format(i, "00")}"
                .Columns(DAY_COL).Format = "###,###,###,##0"
                Create_Summary(grdGLTTBALX, DAY_COL,,, "###,###,###,##0")
            Next
            .Columns("ACCT_BEG_BAL").Format = "###,###,###,##0"
            Create_Summary(grdGLTTBALX, "ACCT_BEG_BAL",,, "###,###,###,##0")
        End With

        'Sort_grdColumns(grdGLTTBALX, "ACCT_TYPE_SEQ", False, 0)
        'Sort_grdColumns(grdGLFTBALD, "JOURNAL_TYPE", False, 0)

        'For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdGLTTBALX, grdGLFTBALD}
        '    With grd.DisplayLayout
        '        For B As Integer = 0 To 1
        '            With .Bands(B)
        '                For PX As Integer = 1 To 12
        '                    With .Columns("P" & Format(PX, "00"))
        '                        Dim LEGEND As String = ASCMAIN1.Get_Legend(ACCT_YEAR & Format(PX, "00"))
        '                        .Header.Caption = Mid(LEGEND, 10, 6)
        '                        .Hidden = (PX > P)
        '                    End With
        '                Next
        '            End With
        '        Next

        '        Dim X As Integer = 1
        '        If chkSEG2_CODE.Checked Then X += 1
        '        If chkSEG3_CODE.Checked Then X += 1

        '        Dim COLUMN_NAME_cs As String = "ACCT_TYPE"
        '        If grd.Name = "grdGLTSUMJ1" Then
        '            COLUMN_NAME_cs = "JOURNAL_TYPE"
        '        End If
        '        .Bands(0).Columns(COLUMN_NAME_cs).ColSpan = X
        '        .Bands(0).Columns(COLUMN_NAME_cs).Width = X * 60 + 10

        '        With .Bands(1)
        '            For Each COLUMN_NAME As String In New String() _
        '                        {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"}
        '                With .Columns(COLUMN_NAME)
        '                    .Width = 60
        '                End With
        '            Next
        '        End With
        '    End With
        'Next
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Call Update_Record_TDA("GLTTBALD")
        CommitTrans("Update Complete")
    End Sub
#End Region
#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTTBALX, "SSB", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdGLFTBALD, "SSBBB", "Show Filter", "Show GroupBox", "Add Account", "Remove Account")
    End Sub
    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        If (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow) And e.Tool.Key <> "grdGLFTBALD" Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "grdGLFTBALD"
                    If (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow) Then
                        If tlb_pop.Tools.Exists("Show Filter") Then
                            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
                            tlb_sbt.SharedProps.Visible = False
                        End If
                        If tlb_pop.Tools.Exists("Show GroupBox") Then
                            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
                            tlb_sbt.SharedProps.Visible = False
                        End If
                        If tlb_pop.Tools.Exists("Remove Account") Then
                            tlb_btn = DirectCast(tlb_pop.Tools("Remove Account"), UltraWinToolbars.ButtonTool)
                            tlb_btn.SharedProps.Visible = False
                        End If
                        If tlb_pop.Tools.Exists("Add Account") Then
                            tlb_btn = DirectCast(tlb_pop.Tools("Add Account"), UltraWinToolbars.ButtonTool)
                            tlb_btn.SharedProps.Visible = True
                        End If

                    End If

            End Select

        End If
    End Sub
    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case ""
        End Select

        If (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow) And e.Tool.Key <> "Add Account" Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Account Inquiry"
                If grd.ActiveRow.Band.Index = 1 Then
                    Dim ACCT_CODE As String = grd.ActiveRow.Cells("ACCT_CODE").Value
                    Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                    If rowGLTACCT1 IsNot Nothing Then
                        Context_Launch("Load", ACCT_CODE, e.Tool.Key, "GLFACTI1")
                    End If
                End If
            Case "Add Account"
                Dim S As New Text.StringBuilder With {.Length = 0}
                S.AppendLine("SELECT")
                S.AppendLine("G1.ACCT_CODE,")
                S.AppendLine("G1.ACCT_DESC,")
                S.AppendLine("G3.SEG2_CODE,")
                S.AppendLine("G3.SEG3_CODE")
                S.AppendLine("FROM GLTACCT1 G1, GLTACCT3 G3")
                S.AppendLine("WHERE G1.ACCT_CODE = G3.ACCT_CODE")
                S.AppendLine("AND G1.ACCT_STATUS = 'A'")
                S.AppendLine("AND G1.ACCT_TYPE IN ('A','L')")
                S.AppendLine("AND G3.ACCT_YEAR = TO_CHAR(SYSDATE, 'YYYY')")
                S.AppendLine("GROUP BY")
                S.AppendLine("G1.ACCT_CODE,")
                S.AppendLine("G1.ACCT_DESC,")
                S.AppendLine("G3.SEG2_CODE,")
                S.AppendLine("G3.SEG3_CODE")
                S.AppendLine("ORDER BY")
                S.AppendLine("G1.ACCT_CODE,")
                S.AppendLine("G1.ACCT_DESC,")
                S.AppendLine("G3.SEG2_CODE,")
                S.AppendLine("G3.SEG3_CODE")
                With ASCMAIN1.CodeSelector
                    .SQL = S.ToString
                    .MultipleSelections = False
                    .PreviouslySelectedCodes0 = ""
                    .Caption = "SELECT ACCOUNT"
                    .TABLE_NAME = ""
                    .VIEW_NAME = ""
                    .VIEW_DESC = ""
                    .COLUMN_NAME = ""
                    .COLUMN_PREKEYs = New Dictionary(Of String, String)
                    '.Custom_sql_where = ""
                    .tblASTVIEW1 = New DataTable
                End With
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    Dim ACCT_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("ACCT_CODE") & ""
                    Dim ACCT_DESC As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("ACCT_DESC") & ""
                    Dim SEG2_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("SEG2_CODE") & ""
                    Dim SEG3_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("SEG3_CODE") & ""
                    Dim fltr As String = $"ACCT_CODE = '{ACCT_CODE}' AND SEG2_CODE = '{SEG2_CODE}' AND SEG3_CODE = '{SEG3_CODE}'"
                    If dst.Tables.Item("GLTTBALD").Select(fltr).Count > 0 Then
                        MsgBox("Account Already Selected", vbCritical, "Huh?")
                    Else
                        Dim rowGLTTBALD As DataRow = dst.Tables.Item("GLTTBALD").NewRow
                        rowGLTTBALD.Item("ACCT_CODE") = ACCT_CODE
                        rowGLTTBALD.Item("ACCT_DESC") = ACCT_DESC
                        rowGLTTBALD.Item("SEG2_CODE") = SEG2_CODE
                        rowGLTTBALD.Item("SEG3_CODE") = SEG3_CODE
                        rowGLTTBALD.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowGLTTBALD.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowGLTTBALD.Item("INIT_DATE") = DATETIME_STAMP
                        rowGLTTBALD.Item("LAST_DATE") = DATETIME_STAMP
                        dst.Tables.Item("GLTTBALD").Rows.Add(rowGLTTBALD)
                        Call Update_Record_TDA("GLTTBALD")
                    End If
                End If
            Case "Remove Account"
                Dim ACCT_CODE As String = grd.ActiveRow.Cells("ACCT_CODE").Value
                Dim SEG2_CODE As String = grd.ActiveRow.Cells("SEG2_CODE").Value
                Dim SEG3_CODE As String = grd.ActiveRow.Cells("SEG3_CODE").Value
                Dim fltr As String = $"ACCT_CODE = '{ACCT_CODE}' AND SEG2_CODE = '{SEG2_CODE}' AND SEG3_CODE = '{SEG3_CODE}'"
                Dim rowGLTTBALD As DataRow = dst.Tables.Item("GLTTBALD").Select(fltr).FirstOrDefault
                If Not IsNothing(rowGLTTBALD) Then
                    rowGLTTBALD.Delete()
                    Call Update_Record_TDA("GLTTBALD")
                End If
        End Select
    End Sub
#End Region
#Region "Form Controls"

#End Region
#Region "Custom Methods"
    Private Sub CalcPeriods()
        CURR_PERIOD = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")
        CURR_YEAR = CURR_PERIOD.Substring(0, 4)
        CURR_MO = CURR_PERIOD.Substring(4, 2)
        SEL_PERIOD = Absx1.txtFor("OPS_YYYYPP").Text
        SEL_YEAR = SEL_PERIOD.Substring(0, 4)
        SEL_MO = SEL_PERIOD.Substring(4, 2)
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("PRD_END_DATE")
        SQLS.AppendLine("FROM GLTPARM2")
        SQLS.AppendLine($"WHERE OPS_YYYYPP = '{SEL_PERIOD}'")
        ASCMAIN1.sql = SQLS.ToString()
        SEL_PERIOD_END = CDate(ASCDATA1.GetDataValue)
        SEL_PERIOD_BEG = DateSerial(SEL_PERIOD_END.Year, SEL_PERIOD_END.Month, 1)
    End Sub
    Private Sub calcDays(ByRef rowGLTTBALX As DataRow)
        Dim ACCT_CODE As String = rowGLTTBALX.Item("ACCT_CODE").ToString & String.Empty
        Dim SEG2_CODE As String = rowGLTTBALX.Item("SEG2_CODE").ToString & String.Empty
        Dim SEG3_CODE As String = rowGLTTBALX.Item("SEG3_CODE").ToString & String.Empty
        Fill_Records("GLTDETL1", {SEL_PERIOD, ACCT_CODE, SEG2_CODE, SEG3_CODE})
        For Each DAYP As KeyValuePair(Of String, Date) In DAY_COLS
            Dim THIS_COL As String = DAYP.Key
            Dim THIS_DAY As Date = DAYP.Value
            Dim BEG_BAL As Double = Val(rowGLTTBALX.Item("ACCT_BEG_BAL").ToString & String.Empty)
            Dim TOT_DAY As Double = BEG_BAL
            Dim fltr As String = $"DETL_CTL_DATE <= '{Format(CDate(THIS_DAY), "MM/dd/yyy")}'"
            For Each rowGLTDETL1 As DataRow In dst.Tables("GLTDETL1").Select(fltr, "DETL_CTL_DATE")
                TOT_DAY = TOT_DAY + Val(rowGLTDETL1.Item("DETL_POSTING_AMT").ToString & String.Empty)
            Next
            rowGLTTBALX.Item(THIS_COL) = TOT_DAY
        Next
    End Sub
    Private Sub calcBegBal(ByRef rowGLTTBALX As DataRow)
        Dim ACCT_CODE As String = rowGLTTBALX.Item("ACCT_CODE").ToString & String.Empty
        Dim SEG2_CODE As String = rowGLTTBALX.Item("SEG2_CODE").ToString & String.Empty
        Dim SEG3_CODE As String = rowGLTTBALX.Item("SEG3_CODE").ToString & String.Empty
        Dim SQL As New System.Text.StringBuilder With {.Length = 0}
        SQL.AppendLine("SELECT ACCT_BEG_BAL")
        SQL.AppendLine("FROM GLTACCT3")
        SQL.AppendLine($"WHERE ACCT_YEAR = '{CURR_YEAR}'")
        SQL.AppendLine($"AND ACCT_CODE = '{ACCT_CODE}'")
        SQL.AppendLine($"AND SEG2_CODE = '{SEG2_CODE}'")
        SQL.AppendLine($"AND SEG3_CODE = '{SEG3_CODE}'")
        ASCMAIN1.sql = SQL.ToString()
        Dim ACCT_BEG_BAL As Double = Val(ASCDATA1.GetDataValue)
        'rowGLTTBALX.Item("ACCT_BEG_BAL") = ACCT_BEG_BAL
        Dim ACCT_ACTIVITY As Double = 0
        For YR As Int64 = Val(CURR_YEAR) To Val(SEL_YEAR)
            If YR <> Val(SEL_YEAR) Then
                SQL.Length = 0
                SQL.AppendLine("SELECT")
                SQL.AppendLine("SUM(")
                For i As Int64 = 1 To 13
                    If i <> 13 Then
                        SQL.AppendLine($"ACCT_ACT_P{Format(i, "00")} + ")
                    Else
                        SQL.AppendLine($"ACCT_ACT_P{Format(i, "00")}) As YR_ACT")
                    End If
                Next
                SQL.AppendLine(" FROM GLTACCT3")
                SQL.AppendLine($"WHERE ACCT_YEAR = '{YR.ToString}'")
                SQL.AppendLine($"AND ACCT_CODE = '{ACCT_CODE}'")
                SQL.AppendLine($"AND SEG2_CODE = '{SEG2_CODE}'")
                SQL.AppendLine($"AND SEG3_CODE = '{SEG3_CODE}'")
                ASCMAIN1.sql = SQL.ToString()
                Dim ACCT_YR_ACT As Double = Val(ASCDATA1.GetDataValue)
                ACCT_ACTIVITY = ACCT_ACTIVITY + ACCT_YR_ACT
            Else
                SQL.Length = 0
                SQL.AppendLine("SELECT")
                SQL.AppendLine("SUM(")
                For i As Int64 = 1 To Val(SEL_MO) - 1
                    If i <> Val(SEL_MO) - 1 Then
                        SQL.AppendLine($"ACCT_ACT_P{Format(i, "00")} + ")
                    Else
                        SQL.AppendLine($"ACCT_ACT_P{Format(i, "00")}) As YR_ACT")
                    End If
                Next
                SQL.AppendLine(" FROM GLTACCT3")
                SQL.AppendLine($"WHERE ACCT_YEAR = '{YR.ToString}'")
                SQL.AppendLine($"AND ACCT_CODE = '{ACCT_CODE}'")
                SQL.AppendLine($"AND SEG2_CODE = '{SEG2_CODE}'")
                SQL.AppendLine($"AND SEG3_CODE = '{SEG3_CODE}'")
                ASCMAIN1.sql = SQL.ToString()
                Dim ACCT_YR_ACT As Double = Val(ASCDATA1.GetDataValue)
                ACCT_ACTIVITY = ACCT_ACTIVITY + ACCT_YR_ACT
            End If
        Next
        rowGLTTBALX.Item("ACCT_BEG_BAL") = ACCT_BEG_BAL + ACCT_ACTIVITY
    End Sub
    Private Sub hideDays()
        For Each DAYP As KeyValuePair(Of String, Date) In DAY_COLS
            If DAYP.Value > SEL_PERIOD_END Then
                grdGLTTBALX.DisplayLayout.Bands(0).Columns.Item(DAYP.Key).Hidden = True
            End If
        Next
    End Sub
    Private Sub setDAY_COLS()
        For i As Int64 = 1 To 31
            Dim DY As Date = DateSerial(SEL_PERIOD_BEG.Year, SEL_PERIOD_BEG.Month, i)
            DAY_COLS.Add($"DAY{Format(i, "00")}", DY)
        Next
        For Each grdCol As UltraWinGrid.UltraGridColumn In grdGLTTBALX.DisplayLayout.Bands(0).Columns
            If DAY_COLS.ContainsKey(grdCol.Key) Then
                grdCol.Header.Caption = $"{Format(DAY_COLS(grdCol.Key), "MM/dd")}"
            End If
        Next
    End Sub
#End Region
End Class