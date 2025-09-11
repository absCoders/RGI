
Public Class WBFFSR01
    Dim FormLoading As Boolean = True
    Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
    Dim tmpWBFFSR01 As String = ""
    Dim TY_DAYS As New List(Of Date)
    Dim LY_DAYS As New List(Of Date)
    Dim DATES As New Dictionary(Of String, Date)
    Dim QTY_DOLLAR As String = "I2.ORDR_QTY_SHIP"

    Dim SEL_CODE As String = "CUST_CODE"
    Dim SEL_DESC As String = "CUST_NAME"
    Dim SEL_TABLE As String = "ARTCUST1"

    Dim valueColsTY As String() = {"TY_WK1", "TY_WK2", "TY_WK3", "TY_WK4", "TY_WK5", "TY_WK6", "TY_WK7", "TOT_TY_WK", "TY_MTD", "TY_YTD"}
    Dim valueColsLY As String() = {"LY_WK1", "LY_WK2", "LY_WK3", "LY_WK4", "LY_WK5", "LY_WK6", "LY_WK7", "TOT_LY_WK", "LY_MTD", "LY_YTD", "LY_FULL_MO", "LY_FULL_YR", "PCT_TY_LY", "PCT_TY_FY"}
    Dim valueColsYOY As String() = {"WTD_FULL_WK_YOY", "MTD_YOY", "YTD_YOY", "MTD_FULL_MO_YOY", "YTD_FULL_YR_YOY"}
    Dim valueColsPCT As String() = {"WTD_FULL_WK_YOY_PCT", "MTD_YOY_PCT", "YTD_YOY_PCT", "MTD_FULL_MO_YOY_PCT", "YTD_FULL_YR_YOY_PCT"}

    Dim EOM_MODE As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        grdWBFFSR01.Visible = False

        DATES.Add("EOW_TY", Now())
        DATES.Add("EOW_LY", Now())
        DATES.Add("BOY_TY", Now())
        DATES.Add("BOY_LY", Now())
        DATES.Add("BOM_TY", Now())
        DATES.Add("BOM_LY", Now())
        DATES.Add("EOM_TY", Now())
        DATES.Add("EOM_LY", Now())

        dteSaturday.Value = calcSaturdayEOW(Now())

        With dst

            SQLs.Length = 0
            SQLs.AppendLine("SELECT ")
            SQLs.AppendLine($"I1.{SEL_CODE} AS CODE,")
            SQLs.AppendLine($"W1.{SEL_DESC} AS CODE_DESC,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK1,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK2,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK3,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK4,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK5,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK6,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK7,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_MTD,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_YTD,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK1,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK2,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK3,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK4,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK5,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK6,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK7,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_MTD,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_YTD,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_FULL_MO,")
            SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_FULL_YR")
            SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1")
            SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
            SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
            SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
            SQLs.AppendLine("AND I1.INV_DATE = '01-JAN-1900'")
            SQLs.AppendLine("GROUP BY ")
            SQLs.AppendLine($"I1.{SEL_CODE}, ")
            SQLs.AppendLine($"W1.{SEL_DESC}")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBFFSR01", "**", 0, False)
            With .Tables("WBFFSR01").Columns
                .Add("TOT_TY_WK", GetType(System.Decimal), "TY_WK1 + TY_WK2 + TY_WK3 + TY_WK4 + TY_WK5 + TY_WK6 + TY_WK7")
                .Add("TOT_LY_WK", GetType(System.Decimal), "LY_WK1 + LY_WK2 + LY_WK3 + LY_WK4 + LY_WK5 + LY_WK6 + LY_WK7")
                .Add("PCT_TY_LY", GetType(System.Decimal), "IIF(ISNULL(TOT_LY_WK,0)=0,0,(TOT_TY_WK/TOT_LY_WK) * 100)")
                .Add("PCT_TY_FY", GetType(System.Decimal), "IIF(ISNULL(LY_FULL_YR,0)=0,0,(TY_YTD/LY_FULL_YR) * 100)")
                '"IIF(ISNULL(STYLE_PRICE,0)=0,0,100*DISC_AMT/ISNULL(STYLE_PRICE,0))"

                .Add("WTD_FULL_WK_YOY", GetType(System.Decimal), "(ISNULL(TOT_TY_WK,0) - ISNULL(TOT_LY_WK,0))")
                .Add("MTD_YOY", GetType(System.Decimal), "(ISNULL(TY_MTD,0) - ISNULL(LY_MTD,0))")
                .Add("YTD_YOY", GetType(System.Decimal), "(ISNULL(TY_YTD,0) - ISNULL(LY_YTD,0))")
                .Add("MTD_FULL_MO_YOY", GetType(System.Decimal), "(ISNULL(TY_MTD,0) - ISNULL(LY_FULL_MO,0))")
                .Add("YTD_FULL_YR_YOY", GetType(System.Decimal), "(ISNULL(TY_YTD,0) - ISNULL(LY_FULL_YR,0))")

                .Add("WTD_FULL_WK_YOY_PCT", GetType(System.Decimal), "IIF(ISNULL(TOT_LY_WK,0)=0,0,(ISNULL(WTD_FULL_WK_YOY,0) / ISNULL(TOT_LY_WK,0)) * 100)")
                .Add("MTD_YOY_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_MTD,0)=0,0,(ISNULL(MTD_YOY,0) / ISNULL(LY_MTD,0)) * 100)")
                .Add("YTD_YOY_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_YTD,0)=0,0,(ISNULL(YTD_YOY,0) / ISNULL(LY_YTD,0)) * 100)")
                .Add("MTD_FULL_MO_YOY_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_FULL_MO,0)=0,0,(ISNULL(MTD_FULL_MO_YOY,0) / ISNULL(LY_FULL_MO,0)) * 100)")
                .Add("YTD_FULL_YR_YOY_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_FULL_YR,0)=0,0,(ISNULL(YTD_FULL_YR_YOY,0) / ISNULL(LY_FULL_YR,0)) * 100)")
            End With

            tmpWBFFSR01 = ASCMAIN1.Temp_Table

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM GLTPARM3")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False)
            Fill_Records("GLTPARM3")

        End With

        'Fill_Records("ECTECOM1_FILTER")

        grdWBFFSR01.DataSource = dst.Tables("WBFFSR01")

        setGridGrouping()

        For Each COL As String In valueColsTY
            Create_Summary(grdWBFFSR01, COL, "Sum", "", "###,###,##0")


            'For Each grp As UltraWinGrid.UltraGridGroup In grdWBFFSR01.DisplayLayout.Bands(0).Groups

            '    'For Each C1 As UltraWinGrid.UltraGridColumn In grp.Columns
            '    '    grdWBFFSR01.DisplayLayout.Bands(0).Columns(C1.Key).Group = grp
            '    'Next
            'Next


        Next

        For Each COL As String In valueColsLY
            If COL = "PCT_TY_LY" Or COL = "PCT_TY_FY" Then
                'Create_Summary(grdWBFFSR01, COL, "Avg", "", "###,###,##0.00")
                'grdWBFFSR01.DisplayLayout.Bands(0).Columns(COL).Format = "###,###,##0.00"
            Else
                Create_Summary(grdWBFFSR01, COL, "Sum", "", "###,###,##0")
                'grdWBFFSR01.DisplayLayout.Bands(0).Columns(COL).Format = "###,###,##0"
            End If
            'With grdWBFFSR01.DisplayLayout.Bands(0)
            '    .Columns(COL).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            '    .Columns(COL).Header.Appearance.BackColor = Drawing.Color.White
            '    .Columns(COL).CellAppearance.BackColor = Drawing.Color.LightGreen
            '    .Columns(COL).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            'End With
            If COL.Substring(0, 5) = "LY_WK" Then
                grdWBFFSR01.DisplayLayout.Bands(0).Columns(COL).Hidden = True
            End If
        Next

        For Each COL As String In valueColsYOY
            Create_Summary(grdWBFFSR01, COL, "Sum", "", "###,###,##0")
            With grdWBFFSR01.DisplayLayout.Bands(0)
                '.Columns(COL).Format = "###,###,##0"
                '.Columns(COL).Header.Appearance.BackColor2 = Drawing.Color.LightPink
                '.Columns(COL).Header.Appearance.BackColor = Drawing.Color.White
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightGreen
                '.Columns(COL).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
        Next

        For Each COL As String In valueColsPCT
            'Create_Summary(grdWBFFSR01, COL, "Avg", "", "###,###,##0.00")
            With grdWBFFSR01.DisplayLayout.Bands(0)
                '.Columns(COL).Format = "###,##0.00"
                '.Columns(COL).Header.Appearance.BackColor2 = Drawing.Color.LightYellow
                '.Columns(COL).Header.Appearance.BackColor = Drawing.Color.White
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightGreen
                '.Columns(COL).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
        Next

        'With grdWBFFSR01.DisplayLayout.Bands(0)
        '    .Columns("CODE").Header.Fixed = True
        '    .Columns("CODE_DESC").Header.Fixed = True
        'End With

        With grdWBFFSR01.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        For i As Integer = 0 To grdWBFFSR01.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBFFSR01.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        'Load_Record(False)

        'Sort_grdColumns(grdWBFFSR01, "SALES".ToLower(), False)

        tab.Visible = False

        FormLoading = False
        'grdWBFHORNT.Parent = tab.Parent

    End Sub



    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"
                Dim inputDate As DateTime = dteSaturday.DateTime
                Dim lastDay As Integer = DateTime.DaysInMonth(inputDate.Year, inputDate.Month)
                Dim lastDateOfMonth As DateTime = New DateTime(inputDate.Year, inputDate.Month, lastDay)
                If inputDate = lastDateOfMonth Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "EOM Version"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("You Select The Last Day Of The Month.")
                    iMSG.AppendLine("Do You Want To Run The End-Of-Month")
                    iMSG.AppendLine("Version Of The Report?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult = MsgBoxResult.Yes Then
                        EOM_MODE = True
                        If setDates() = False Then
                            EMsg &= vbCr & $"Problem Setting Dates For {Format(dteSaturday.DateTime, "MM/dd/yy")}."
                        End If
                    Else
                        EOM_MODE = False
                    End If
                Else
                    EOM_MODE = False
                End If

                If EOM_MODE = False Then
                    If dteSaturday.DateTime.DayOfWeek <> DayOfWeek.Saturday Then
                        EMsg &= vbCr & $"{Format(dteSaturday.DateTime, "MM/dd/yy")} In Not A Saturday."
                    Else
                        If setDates() = False Then
                            EMsg &= vbCr & $"Problem Setting Dates For {Format(dteSaturday.DateTime, "MM/dd/yy")}."
                        End If
                    End If
                End If

            Case "Exit"

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
                If rdoUnits.Checked Then
                    QTY_DOLLAR = "I2.ORDR_QTY_SHIP"
                Else
                    If rdoDollars.Checked Then
                        QTY_DOLLAR = "I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE"
                    End If
                End If

                If rdoWSHE_CODE.Checked Then
                    SEL_CODE = "WHSE_CODE"
                    SEL_DESC = "WHSE_DESC"
                    SEL_TABLE = "ICTWHSE1"
                    grdWBFFSR01.DisplayLayout.Bands(0).Columns("CODE").Header.Caption = "Whse"
                End If
                If rdoCUST_CODE.Checked Then
                    SEL_CODE = "CUST_CODE"
                    SEL_DESC = "CUST_NAME"
                    SEL_TABLE = "ARTCUST1"
                    grdWBFFSR01.DisplayLayout.Bands(0).Columns("CODE").Header.Caption = "Cust"
                End If
                For Each COL As String In valueColsLY
                    If COL.Substring(0, 5) = "LY_WK" Then
                        If chkHideLYDays.Checked Then
                            grdWBFFSR01.DisplayLayout.Bands(0).Columns(COL).Hidden = True
                        Else
                            grdWBFFSR01.DisplayLayout.Bands(0).Columns(COL).Hidden = False
                        End If
                    End If
                Next
                Clear_Record()
                Load_Record(True)
                setGridTitle()
            Case "Exit"
                Call Mode_Settings(False)
                Me.Close()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Exit").Visible = Not ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        dst.Tables("WBFFSR01").Rows.Clear()
    End Sub

    Sub Load_Record(Optional showRefreshing As Boolean = False)
        Me.Cursor = Cursors.WaitCursor
        If showRefreshing Then
            ASCMAIN1.Progress("Refreshing Data", "")
            Me.Cursor = Cursors.WaitCursor
        End If
        Application.DoEvents()
        'Call Save_Header_Fields(UltraGroupBox1)
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        EnforceConstraints(False)

        SQLs.Length = 0
        SQLs.AppendLine($"DELETE FROM {tmpWBFFSR01}")
        ASCMAIN1.sql = SQLs.ToString
        ASCDATA1.ExecuteSQL()

        'Fill_Records("WBTHORNT", , , SQLs.ToString)
        If chkSeperateWebEDI.Checked = False Then
            fillTempTable("A")
        Else
            fillTempTable("X")
            fillTempTable("E")
            fillTempTable("W")
        End If

        loadTempTable()
        'Stop
        'EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        setEOM_MODES()

        Me.Cursor = Cursors.Default
        If showRefreshing Then
            ASCMAIN1.Progress("")
            Me.Cursor = Cursors.Default
        End If
        Application.DoEvents()
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()

        'Call CommitTrans("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'TODO: Remove this stop before going live.
        Stop
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        Print_Report_Begin()
        'Generate_Report("WBRHORNT")
        Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWBFFSR01, "SSB", "Show Filter", "Show GroupBox")
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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            Case "grdWBFHORNT"

        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Copy To Clipboard"
            '    Dim STYLE_CODE As String = grd.ActiveRow.Cells("RANK_CODE").Text
            '    Clipboard.SetText(STYLE_CODE)
            '    MsgBox($"{STYLE_CODE} Copied To Clipboard.", vbOKOnly, "Clipboard")
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "BANK_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Call Click_Command("New", e)
        '        End If
        '    Case "PYMT_BATCH_NO"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Call Click_Command("Edit", e)
        '        End If
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                'FillStyle()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Call Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub fillTempTable(ByVal OrdrType As String)
        ' --- OrdrType ---
        ' A = All Order Types
        ' X = All Order Types Except EDI & Web
        ' E = EDI
        ' W = Web
        ' Anything Else Don't Fill.

        Dim SCODES_SEL As String = ""
        Dim SCODES_GRP As String = ""
        Dim SCODES_WHR As String = ""
        Select Case OrdrType
            Case "A"
                SCODES_SEL = $"I1.{SEL_CODE} AS CODE, W1.{SEL_DESC} AS CODE_DESC,"
                SCODES_GRP = $"I1.{SEL_CODE}, W1.{SEL_DESC}"
                SCODES_WHR = ""
            Case "X"
                SCODES_SEL = $"I1.{SEL_CODE} AS CODE, W1.{SEL_DESC} AS CODE_DESC,"
                SCODES_GRP = $"I1.{SEL_CODE}, W1.{SEL_DESC}"
                SCODES_WHR = "AND O1.ORDR_SOURCE NOT IN ('E','W')"
            Case "E"
                SCODES_SEL = $"'EDI' AS CODE, 'EDI Orders' AS CODE_DESC,"
                SCODES_GRP = $"'EDI', 'EDI Orders'"
                SCODES_WHR = "AND O1.ORDR_SOURCE = 'E'"
            Case "W"
                SCODES_SEL = $"'WEB' AS CODE, 'Web Orders' AS CODE_DESC,"
                SCODES_GRP = $"'WEB', 'Web Orders'"
                SCODES_WHR = "AND O1.ORDR_SOURCE = 'W'"
            Case Else
                Exit Sub
        End Select

        Dim MOs As Int64 = 0
        If EOM_MODE Then
            MOs = 6
        End If

        For i As Int64 = MOs To 6
            SQLs.Length = 0
            SQLs.AppendLine($"INSERT INTO {tmpWBFFSR01}")
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine(SCODES_SEL)
            If i = 0 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK1,")
            Else
                SQLs.AppendLine("0 TY_WK1,")
            End If
            If i = 1 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK2,")
            Else
                SQLs.AppendLine("0 TY_WK2,")
            End If
            If i = 2 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK3,")
            Else
                SQLs.AppendLine("0 TY_WK3,")
            End If
            If i = 3 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK4,")
            Else
                SQLs.AppendLine("0 TY_WK4,")
            End If
            If i = 4 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK5,")
            Else
                SQLs.AppendLine("0 TY_WK5,")
            End If
            If i = 5 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK6,")
            Else
                SQLs.AppendLine("0 TY_WK6,")
            End If
            If i = 6 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_WK7,")
            Else
                SQLs.AppendLine("0 TY_WK7,")
            End If
            SQLs.AppendLine("0 TY_MTD,")
            SQLs.AppendLine("0 TY_YTD,")
            SQLs.AppendLine("0 LY_WK1,")
            SQLs.AppendLine("0 LY_WK2,")
            SQLs.AppendLine("0 LY_WK3,")
            SQLs.AppendLine("0 LY_WK4,")
            SQLs.AppendLine("0 LY_WK5,")
            SQLs.AppendLine("0 LY_WK6,")
            SQLs.AppendLine("0 LY_WK7,")
            SQLs.AppendLine("0 LY_MTD,")
            SQLs.AppendLine("0 LY_YTD,")
            SQLs.AppendLine("0 LY_FULL_MO,")
            SQLs.AppendLine("0 LY_FULL_YR")
            SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1, SOTORDR1 O1")
            SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
            SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
            SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
            SQLs.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
            If EOM_MODE Then
                SQLs.AppendLine($"AND (I1.INV_DATE >= '{Format(DATES("BOM_TY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOW_TY"), "dd-MMM-yyyy")}')")
            Else
                SQLs.AppendLine($"AND I1.INV_DATE = '{Format(TY_DAYS(i), "dd-MMM-yyyy")}'")
            End If
            If chkHideCredits.Checked Then
                SQLs.AppendLine("AND I1.INV_TYPE = 'I'")
            End If
            SQLs.AppendLine(SCODES_WHR)
            SQLs.AppendLine("GROUP BY ")
            SQLs.AppendLine(SCODES_GRP)
            ASCMAIN1.sql = SQLs.ToString
            ASCDATA1.ExecuteSQL()

            SQLs.Length = 0
            SQLs.AppendLine($"INSERT INTO {tmpWBFFSR01}")
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine(SCODES_SEL)
            SQLs.AppendLine("0 TY_WK1,")
            SQLs.AppendLine("0 TY_WK2,")
            SQLs.AppendLine("0 TY_WK3,")
            SQLs.AppendLine("0 TY_WK4,")
            SQLs.AppendLine("0 TY_WK5,")
            SQLs.AppendLine("0 TY_WK6,")
            SQLs.AppendLine("0 TY_WK7,")
            SQLs.AppendLine("0 TY_MTD,")
            SQLs.AppendLine("0 TY_YTD,")
            If i = 0 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK1,")
            Else
                SQLs.AppendLine("0 LY_WK1,")
            End If
            If i = 1 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK2,")
            Else
                SQLs.AppendLine("0 LY_WK2,")
            End If
            If i = 2 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK3,")
            Else
                SQLs.AppendLine("0 LY_WK3,")
            End If
            If i = 3 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK4,")
            Else
                SQLs.AppendLine("0 LY_WK4,")
            End If
            If i = 4 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK5,")
            Else
                SQLs.AppendLine("0 LY_WK5,")
            End If
            If i = 5 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK6,")
            Else
                SQLs.AppendLine("0 LY_WK6,")
            End If
            If i = 6 Then
                SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_WK7,")
            Else
                SQLs.AppendLine("0 LY_WK7,")
            End If
            SQLs.AppendLine("0 LY_MTD,")
            SQLs.AppendLine("0 LY_YTD,")
            SQLs.AppendLine("0 LY_FULL_MO,")
            SQLs.AppendLine("0 LY_FULL_YR")
            SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1, SOTORDR1 O1")
            SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
            SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
            SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
            SQLs.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
            If EOM_MODE Then
                SQLs.AppendLine($"AND (I1.INV_DATE >= '{Format(DATES("BOM_LY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOM_LY"), "dd-MMM-yyyy")}')")
            Else
                SQLs.AppendLine($"AND I1.INV_DATE = '{Format(LY_DAYS(i), "dd-MMM-yyyy")}'")
            End If
            If chkHideCredits.Checked Then
                SQLs.AppendLine("AND I1.INV_TYPE = 'I'")
            End If
            SQLs.AppendLine(SCODES_WHR)
            SQLs.AppendLine("GROUP BY ")
            SQLs.AppendLine(SCODES_GRP)
            ASCMAIN1.sql = SQLs.ToString
            ASCDATA1.ExecuteSQL()
        Next

        'Fill TY_MTD
        SQLs.Length = 0
        SQLs.AppendLine($"INSERT INTO {tmpWBFFSR01}")
        SQLs.AppendLine("SELECT ")
        SQLs.AppendLine(SCODES_SEL)
        SQLs.AppendLine("0 TY_WK1,")
        SQLs.AppendLine("0 TY_WK2,")
        SQLs.AppendLine("0 TY_WK3,")
        SQLs.AppendLine("0 TY_WK4,")
        SQLs.AppendLine("0 TY_WK5,")
        SQLs.AppendLine("0 TY_WK6,")
        SQLs.AppendLine("0 TY_WK7,")
        SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_MTD,")
        SQLs.AppendLine("0 TY_YTD,")
        SQLs.AppendLine("0 LY_WK1,")
        SQLs.AppendLine("0 LY_WK2,")
        SQLs.AppendLine("0 LY_WK3,")
        SQLs.AppendLine("0 LY_WK4,")
        SQLs.AppendLine("0 LY_WK5,")
        SQLs.AppendLine("0 LY_WK6,")
        SQLs.AppendLine("0 LY_WK7,")
        SQLs.AppendLine("0 LY_MTD,")
        SQLs.AppendLine("0 LY_YTD,")
        SQLs.AppendLine("0 LY_FULL_MO,")
        SQLs.AppendLine("0 LY_FULL_YR")
        SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1, SOTORDR1 O1")
        SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
        SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
        SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
        SQLs.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
        SQLs.AppendLine($"AND I1.INV_DATE >= '{Format(DATES("BOM_TY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOW_TY"), "dd-MMM-yyyy")}'")
        If chkHideCredits.Checked Then
            SQLs.AppendLine("AND I1.INV_TYPE = 'I'")
        End If
        SQLs.AppendLine(SCODES_WHR)
        SQLs.AppendLine("GROUP BY ")
        SQLs.AppendLine(SCODES_GRP)
        ASCMAIN1.sql = SQLs.ToString()
        ASCDATA1.ExecuteSQL()

        'Fill TY_YTD
        SQLs.Length = 0
        SQLs.AppendLine($"INSERT INTO {tmpWBFFSR01}")
        SQLs.AppendLine("SELECT ")
        SQLs.AppendLine(SCODES_SEL)
        SQLs.AppendLine("0 TY_WK1,")
        SQLs.AppendLine("0 TY_WK2,")
        SQLs.AppendLine("0 TY_WK3,")
        SQLs.AppendLine("0 TY_WK4,")
        SQLs.AppendLine("0 TY_WK5,")
        SQLs.AppendLine("0 TY_WK6,")
        SQLs.AppendLine("0 TY_WK7,")
        SQLs.AppendLine("0 TY_MTD,")
        SQLs.AppendLine($"SUM({QTY_DOLLAR}) TY_YTD,")
        SQLs.AppendLine("0 LY_WK1,")
        SQLs.AppendLine("0 LY_WK2,")
        SQLs.AppendLine("0 LY_WK3,")
        SQLs.AppendLine("0 LY_WK4,")
        SQLs.AppendLine("0 LY_WK5,")
        SQLs.AppendLine("0 LY_WK6,")
        SQLs.AppendLine("0 LY_WK7,")
        SQLs.AppendLine("0 LY_MTD,")
        SQLs.AppendLine("0 LY_YTD,")
        SQLs.AppendLine("0 LY_FULL_MO,")
        SQLs.AppendLine("0 LY_FULL_YR")
        SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1, SOTORDR1 O1")
        SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
        SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
        SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
        SQLs.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
        SQLs.AppendLine($"AND I1.INV_DATE >= '{Format(DATES("BOY_TY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOW_TY"), "dd-MMM-yyyy")}'")
        If chkHideCredits.Checked Then
            SQLs.AppendLine("AND I1.INV_TYPE = 'I'")
        End If
        SQLs.AppendLine(SCODES_WHR)
        SQLs.AppendLine("GROUP BY ")
        SQLs.AppendLine(SCODES_GRP)
        ASCMAIN1.sql = SQLs.ToString()
        ASCDATA1.ExecuteSQL()

        'Fill LY_MTD
        SQLs.Length = 0
        SQLs.AppendLine($"INSERT INTO {tmpWBFFSR01}")
        SQLs.AppendLine("SELECT ")
        SQLs.AppendLine(SCODES_SEL)
        SQLs.AppendLine("0 TY_WK1,")
        SQLs.AppendLine("0 TY_WK2,")
        SQLs.AppendLine("0 TY_WK3,")
        SQLs.AppendLine("0 TY_WK4,")
        SQLs.AppendLine("0 TY_WK5,")
        SQLs.AppendLine("0 TY_WK6,")
        SQLs.AppendLine("0 TY_WK7,")
        SQLs.AppendLine("0 TY_MTD,")
        SQLs.AppendLine("0 TY_YTD,")
        SQLs.AppendLine("0 LY_WK1,")
        SQLs.AppendLine("0 LY_WK2,")
        SQLs.AppendLine("0 LY_WK3,")
        SQLs.AppendLine("0 LY_WK4,")
        SQLs.AppendLine("0 LY_WK5,")
        SQLs.AppendLine("0 LY_WK6,")
        SQLs.AppendLine("0 LY_WK7,")
        SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_MTD,")
        SQLs.AppendLine("0 LY_YTD,")
        SQLs.AppendLine("0 LY_FULL_MO,")
        SQLs.AppendLine("0 LY_FULL_YR")
        SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1, SOTORDR1 O1")
        SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
        SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
        SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
        SQLs.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
        If EOM_MODE Then
            SQLs.AppendLine($"AND I1.INV_DATE >= '{Format(DATES("BOM_LY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOM_LY"), "dd-MMM-yyyy")}'")
        Else
            SQLs.AppendLine($"AND I1.INV_DATE >= '{Format(DATES("BOM_LY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOW_LY"), "dd-MMM-yyyy")}'")
        End If
        If chkHideCredits.Checked Then
            SQLs.AppendLine("AND I1.INV_TYPE = 'I'")
        End If
        SQLs.AppendLine(SCODES_WHR)
        SQLs.AppendLine("GROUP BY ")
        SQLs.AppendLine(SCODES_GRP)
        ASCMAIN1.sql = SQLs.ToString()
        ASCDATA1.ExecuteSQL()

        'Fill LY_YTD
        SQLs.Length = 0
        SQLs.AppendLine($"INSERT INTO {tmpWBFFSR01}")
        SQLs.AppendLine("SELECT ")
        SQLs.AppendLine(SCODES_SEL)
        SQLs.AppendLine("0 TY_WK1,")
        SQLs.AppendLine("0 TY_WK2,")
        SQLs.AppendLine("0 TY_WK3,")
        SQLs.AppendLine("0 TY_WK4,")
        SQLs.AppendLine("0 TY_WK5,")
        SQLs.AppendLine("0 TY_WK6,")
        SQLs.AppendLine("0 TY_WK7,")
        SQLs.AppendLine("0 TY_MTD,")
        SQLs.AppendLine("0 TY_YTD,")
        SQLs.AppendLine("0 LY_WK1,")
        SQLs.AppendLine("0 LY_WK2,")
        SQLs.AppendLine("0 LY_WK3,")
        SQLs.AppendLine("0 LY_WK4,")
        SQLs.AppendLine("0 LY_WK5,")
        SQLs.AppendLine("0 LY_WK6,")
        SQLs.AppendLine("0 LY_WK7,")
        SQLs.AppendLine("0 LY_MTD,")
        SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_YTD,")
        SQLs.AppendLine("0 LY_FULL_MO,")
        SQLs.AppendLine("0 LY_FULL_YR")
        SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1, SOTORDR1 O1")
        SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
        SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
        SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
        SQLs.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
        If EOM_MODE Then
            SQLs.AppendLine($"AND I1.INV_DATE >= '{Format(DATES("BOY_LY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOM_LY"), "dd-MMM-yyyy")}'")
        Else
            SQLs.AppendLine($"AND I1.INV_DATE >= '{Format(DATES("BOY_LY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOW_LY"), "dd-MMM-yyyy")}'")
        End If
        If chkHideCredits.Checked Then
            SQLs.AppendLine("AND I1.INV_TYPE = 'I'")
        End If
        SQLs.AppendLine(SCODES_WHR)
        SQLs.AppendLine("GROUP BY ")
        SQLs.AppendLine(SCODES_GRP)
        ASCMAIN1.sql = SQLs.ToString()
        ASCDATA1.ExecuteSQL()

        'Fill LY_FULL_MO
        'Dim LY_EOM As Date = DateSerial(DATES("BOM_LY").Year, DATES("BOM_LY").Month, DATES("BOM_LY").AddMonths(1).AddDays(-1).Day)
        SQLs.Length = 0
        SQLs.AppendLine($"INSERT INTO {tmpWBFFSR01}")
        SQLs.AppendLine("SELECT ")
        SQLs.AppendLine(SCODES_SEL)
        SQLs.AppendLine("0 TY_WK1,")
        SQLs.AppendLine("0 TY_WK2,")
        SQLs.AppendLine("0 TY_WK3,")
        SQLs.AppendLine("0 TY_WK4,")
        SQLs.AppendLine("0 TY_WK5,")
        SQLs.AppendLine("0 TY_WK6,")
        SQLs.AppendLine("0 TY_WK7,")
        SQLs.AppendLine("0 TY_MTD,")
        SQLs.AppendLine("0 TY_YTD,")
        SQLs.AppendLine("0 LY_WK1,")
        SQLs.AppendLine("0 LY_WK2,")
        SQLs.AppendLine("0 LY_WK3,")
        SQLs.AppendLine("0 LY_WK4,")
        SQLs.AppendLine("0 LY_WK5,")
        SQLs.AppendLine("0 LY_WK6,")
        SQLs.AppendLine("0 LY_WK7,")
        SQLs.AppendLine("0 LY_MTD,")
        SQLs.AppendLine("0 LY_YTD,")
        SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_FULL_MO,")
        SQLs.AppendLine("0 LY_FULL_YR")
        SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1, SOTORDR1 O1")
        SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
        SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
        SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
        SQLs.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
        SQLs.AppendLine($"AND I1.INV_DATE >= '{Format(DATES("BOM_LY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(DATES("EOM_LY"), "dd-MMM-yyyy")}'")
        If chkHideCredits.Checked Then
            SQLs.AppendLine("AND I1.INV_TYPE = 'I'")
        End If
        SQLs.AppendLine(SCODES_WHR)
        SQLs.AppendLine("GROUP BY ")
        SQLs.AppendLine(SCODES_GRP)
        ASCMAIN1.sql = SQLs.ToString()
        ASCDATA1.ExecuteSQL()

        'Fill LY_FULL_YR
        Dim LY_EOY As Date = CDate(DATES("BOY_TY").AddDays(-1))
        SQLs.Length = 0
        SQLs.AppendLine($"INSERT INTO {tmpWBFFSR01}")
        SQLs.AppendLine("SELECT ")
        SQLs.AppendLine(SCODES_SEL)
        SQLs.AppendLine("0 TY_WK1,")
        SQLs.AppendLine("0 TY_WK2,")
        SQLs.AppendLine("0 TY_WK3,")
        SQLs.AppendLine("0 TY_WK4,")
        SQLs.AppendLine("0 TY_WK5,")
        SQLs.AppendLine("0 TY_WK6,")
        SQLs.AppendLine("0 TY_WK7,")
        SQLs.AppendLine("0 TY_MTD,")
        SQLs.AppendLine("0 TY_YTD,")
        SQLs.AppendLine("0 LY_WK1,")
        SQLs.AppendLine("0 LY_WK2,")
        SQLs.AppendLine("0 LY_WK3,")
        SQLs.AppendLine("0 LY_WK4,")
        SQLs.AppendLine("0 LY_WK5,")
        SQLs.AppendLine("0 LY_WK6,")
        SQLs.AppendLine("0 LY_WK7,")
        SQLs.AppendLine("0 LY_MTD,")
        SQLs.AppendLine("0 LY_YTD,")
        SQLs.AppendLine("0 LY_FULL_MO,")
        SQLs.AppendLine($"SUM({QTY_DOLLAR}) LY_FULL_YR")
        SQLs.AppendLine($"FROM SOTINVH1 I1, SOTINVH2 I2, {SEL_TABLE} W1, SOTORDR1 O1")
        SQLs.AppendLine($"WHERE I1.{SEL_CODE} = W1.{SEL_CODE}")
        SQLs.AppendLine("AND I1.INV_TYPE = I2.INV_TYPE")
        SQLs.AppendLine("AND I1.INV_NO = I2.INV_NO")
        SQLs.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
        SQLs.AppendLine($"AND I1.INV_DATE >= '{Format(DATES("BOY_LY"), "dd-MMM-yyyy")}' AND I1.INV_DATE <= '{Format(LY_EOY, "dd-MMM-yyyy")}'")
        If chkHideCredits.Checked Then
            SQLs.AppendLine("AND I1.INV_TYPE = 'I'")
        End If
        SQLs.AppendLine(SCODES_WHR)
        SQLs.AppendLine("GROUP BY ")
        SQLs.AppendLine(SCODES_GRP)
        ASCMAIN1.sql = SQLs.ToString()
        ASCDATA1.ExecuteSQL()

    End Sub

    Private Sub loadTempTable()
        SQLs.Length = 0
        SQLs.AppendLine("SELECT ")
        SQLs.AppendLine("CODE,")
        SQLs.AppendLine("CODE_DESC,")
        SQLs.AppendLine("SUM(TY_WK1) TY_WK1,")
        SQLs.AppendLine("SUM(TY_WK2) TY_WK2,")
        SQLs.AppendLine("SUM(TY_WK3) TY_WK3,")
        SQLs.AppendLine("SUM(TY_WK4) TY_WK4,")
        SQLs.AppendLine("SUM(TY_WK5) TY_WK5,")
        SQLs.AppendLine("SUM(TY_WK6) TY_WK6,")
        SQLs.AppendLine("SUM(TY_WK7) TY_WK7,")
        SQLs.AppendLine("SUM(TY_MTD) TY_MTD,")
        SQLs.AppendLine("SUM(TY_YTD) TY_YTD,")
        SQLs.AppendLine("SUM(LY_WK1) LY_WK1,")
        SQLs.AppendLine("SUM(LY_WK2) LY_WK2,")
        SQLs.AppendLine("SUM(LY_WK3) LY_WK3,")
        SQLs.AppendLine("SUM(LY_WK4) LY_WK4,")
        SQLs.AppendLine("SUM(LY_WK5) LY_WK5,")
        SQLs.AppendLine("SUM(LY_WK6) LY_WK6,")
        SQLs.AppendLine("SUM(LY_WK7) LY_WK7,")
        SQLs.AppendLine("SUM(LY_MTD) LY_MTD,")
        SQLs.AppendLine("SUM(LY_YTD) LY_YTD,")
        SQLs.AppendLine("SUM(LY_FULL_MO) LY_FULL_MO,")
        SQLs.AppendLine("SUM(LY_FULL_YR) LY_FULL_YR")
        SQLs.AppendLine($"FROM {tmpWBFFSR01}")
        SQLs.AppendLine("GROUP BY ")
        SQLs.AppendLine("CODE, ")
        SQLs.AppendLine("CODE_DESC")
        ASCMAIN1.sql = SQLs.ToString()
        Fill_Records("WBFFSR01",,, SQLs.ToString())
    End Sub

    Private Function setDates() As Boolean
        Dim Retval As Boolean = False
        Try
            TY_DAYS.Clear()
            LY_DAYS.Clear()
            Dim fltr As String = $"WEEK_END_DATE = '{Format(dteSaturday.DateTime, "dd-MMM-yyyy")}'"
            If EOM_MODE Then
                Dim daysToSubtract As Integer = (dteSaturday.DateTime.DayOfWeek - DayOfWeek.Saturday + 7) Mod 7
                Dim previousSaturday As DateTime = dteSaturday.DateTime.AddDays(-daysToSubtract)
                fltr = $"WEEK_END_DATE = '{Format(previousSaturday, "dd-MMM-yyyy")}'"
            End If
            Dim rowGLTPARM3TY As DataRow = dst.Tables.Item("GLTPARM3").Select(fltr).FirstOrDefault
            Dim rowGLTPARM3LY As DataRow = Nothing
            If Not IsNothing(rowGLTPARM3TY) Then
                Dim YYYYMM As String = rowGLTPARM3TY.Item("YYYYMM").ToString
                YYYYMM = Val(YYYYMM.Substring(0, 4)) - 1 & YYYYMM.Substring(4, 2)
                Dim REL_WEEK As Int64 = Val(rowGLTPARM3TY.Item("REL_WEEK").ToString)
                fltr = $"YYYYMM = '{YYYYMM}' AND REL_WEEK = {REL_WEEK}"
                rowGLTPARM3LY = dst.Tables.Item("GLTPARM3").Select(fltr).FirstOrDefault
                If Not IsNothing(rowGLTPARM3TY) Then
                    DATES("EOW_TY") = dteSaturday.DateTime
                    DATES("EOW_LY") = CDate(rowGLTPARM3LY.Item("WEEK_END_DATE").ToString)
                    DATES("BOY_TY") = DateSerial(DATES("EOW_TY").Year, 4, 1)
                    DATES("BOM_TY") = DateSerial(DATES("EOW_TY").Year, DATES("EOW_TY").Month, 1)
                    DATES("BOY_LY") = DateSerial(DATES("EOW_LY").Year, 4, 1)
                    DATES("BOM_LY") = DateSerial(DATES("EOW_LY").Year, DATES("EOW_TY").Month, 1)
                    DATES("EOM_TY") = DateSerial(DATES("BOM_TY").Year, DATES("BOM_TY").Month, DATES("BOM_TY").AddMonths(1).AddDays(-1).Day)
                    DATES("EOM_LY") = DateSerial(DATES("BOM_LY").Year, DATES("BOM_TY").Month, DATES("BOM_LY").AddMonths(1).AddDays(-1).Day)

                    Dim C As Int64 = 0
                    For i As Int64 = 6 To 0 Step -1
                        C += 1
                        TY_DAYS.Add(DATES("EOW_TY").AddDays(-1 * i))
                        LY_DAYS.Add(DATES("EOW_LY").AddDays(-1 * i))
                        grdWBFFSR01.DisplayLayout.Bands(0).Columns($"TY_WK" & C).Header.Caption = Format(DATES("EOW_TY").AddDays(-1 * i), "MM/dd/yy")
                        grdWBFFSR01.DisplayLayout.Bands(0).Columns($"LY_WK" & C).Header.Caption = Format(DATES("EOW_LY").AddDays(-1 * i), "MM/dd/yy")
                    Next

                    Retval = True
                    grdWBFFSR01.Visible = True
                End If
            End If
        Catch ex As Exception
            'Swallow Error.
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                Stop
            End If
        End Try
        grdWBFFSR01.Visible = Retval
        Return Retval
    End Function

    Private Function calcSaturdayEOW(ByVal inDate As Date) As Date
        Dim retval As Date = Now()
        Dim daysUntilSaturday As Integer = DayOfWeek.Saturday - inDate.DayOfWeek
        If daysUntilSaturday < 0 Then
            daysUntilSaturday += 7
        End If
        Dim endOfWeekSaturday As Date = inDate.AddDays(daysUntilSaturday)
        Return CDate(endOfWeekSaturday.ToString("MM/dd/yyy"))
    End Function

    Private Sub setGridGrouping()
        grdWBFFSR01.DisplayLayout.Bands(0).Groups.Clear()
        'grdWBFFSR01.DisplayLayout.Override.CellMultiLine = DefaultableBoolean.True
        'With grdWBFFSR01.DisplayLayout.Override.GroupByRowAppearance
        '    .TextTrimming = TextTrimming.None
        '    .TextVAlign = VAlign.Top
        '    .TextHAlign = HAlign.Center
        'End With

        With grdWBFFSR01.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup

            'Codes
            Dim COLS As String() = {"CODE", "CODE_DESC"}
            G = .Groups.Add("Codes", "")
            G.Header.Appearance.TextHAlign = HAlign.Center
            'G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                '.Columns(COL).Format = "###,##0"
            Next
            G.Header.Fixed = True

            'This Week
            COLS = {"TY_WK1", "TY_WK2", "TY_WK3", "TY_WK4", "TY_WK5", "TY_WK6", "TY_WK7"}
            G = .Groups.Add("This Week", "This Week")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                .Columns(COL).Format = "###,##0"
            Next

            'Space00
            COLS = {"Space00"}
            G = .Groups.Add("Space00", "")
            G.Header.Appearance.TextHAlign = HAlign.Center
            'G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                '.Columns(COL).Format = "###,##0"
            Next

            ''Last Year Full Wk
            'COLS = {"TOT_TY_WK", "TOT_LY_WK"}
            'G = .Groups.Add("TOT_LY_WK", "")
            'G.Header.Appearance.TextHAlign = HAlign.Center
            ''G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            'G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
            'For Each COL As String In COLS
            '    .Columns(COL).Group = G
            '    '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
            '    '.Columns(COL).Width = 80
            '    '.Columns(COL).Format = "###,##0"
            'Next

            'Week To Date TYLY
            COLS = {"TOT_TY_WK", "TOT_LY_WK", "MTD_YOY", "MTD_YOY_PCT"}
            G = .Groups.Add("WTDTYLY", "Current Week-to-Date vs. Last Year FULL Week")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                If COL.EndsWith("_PCT") Then
                    .Columns(COL).Format = "###,##0.0"
                Else
                    .Columns(COL).Format = "###,##0"
                End If
            Next

            'Space01
            COLS = {"Space01"}
            G = .Groups.Add("Space01", "")
            G.Header.Appearance.TextHAlign = HAlign.Center
            'G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                '.Columns(COL).Format = "###,##0"
            Next

            'MTD
            COLS = {"TY_MTD", "LY_MTD", "WTD_FULL_WK_YOY", "WTD_FULL_WK_YOY_PCT"}
            G = .Groups.Add("MTD", "Month-to-Date")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.LightPink
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                If COL.EndsWith("_PCT") Then
                    .Columns(COL).Format = "###,##0.0"
                Else
                    .Columns(COL).Format = "###,##0"
                End If
            Next

            'Space02
            COLS = {"Space02"}
            G = .Groups.Add("Space02", "")
            G.Header.Appearance.TextHAlign = HAlign.Center
            'G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                '.Columns(COL).Format = "###,##0"
            Next

            'YTD
            COLS = {"TY_YTD", "LY_YTD", "YTD_YOY", "YTD_YOY_PCT"}
            G = .Groups.Add("YTD", "Year-to-Date")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.LightSalmon
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                If COL.EndsWith("_PCT") Then
                    .Columns(COL).Format = "###,##0.0"
                Else
                    .Columns(COL).Format = "###,##0"
                End If
            Next

            'Space03
            COLS = {"Space03"}
            G = .Groups.Add("Space03", "")
            G.Header.Appearance.TextHAlign = HAlign.Center
            'G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                '.Columns(COL).Format = "###,##0"
            Next

            'Curr Month TYLY
            COLS = {"LY_FULL_MO", "MTD_FULL_MO_YOY", "MTD_FULL_MO_YOY_PCT"}
            G = .Groups.Add("CMTYLY", "Current Month-to-Date vs. Last Year FULL Month")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                If COL.EndsWith("_PCT") Then
                    .Columns(COL).Format = "###,##0.0"
                Else
                    .Columns(COL).Format = "###,##0"
                End If
            Next

            'Space04
            COLS = {"Space04"}
            G = .Groups.Add("Space04", "")
            G.Header.Appearance.TextHAlign = HAlign.Center
            'G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                '.Columns(COL).Format = "###,##0"
            Next

            'Curr Year TYLY
            COLS = {"LY_FULL_YR", "YTD_FULL_YR_YOY", "YTD_FULL_YR_YOY_PCT"}
            G = .Groups.Add("CYTYLY", "Current Year-to-Date vs. Last Year FULL Year")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Appearance.TextTrimming = TextTrimming.None
            G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            G.Header.Appearance.BackColor2 = Drawing.Color.Yellow
            For Each COL As String In COLS
                .Columns(COL).Group = G
                '.Columns(COL).CellAppearance.BackColor = Drawing.Color.LightYellow
                '.Columns(COL).Width = 80
                If COL.EndsWith("_PCT") Then
                    .Columns(COL).Format = "###,##0.0"
                Else
                    .Columns(COL).Format = "###,##0"
                End If
            Next


        End With
    End Sub

    Private Sub setGridTitle()
        Dim titleRoot As String = "Flash Sales Report"
        Dim titleSelection As String = ""
        Dim titlePeriod As String = ""
        Dim titleValues As String = ""
        If rdoWSHE_CODE.Checked Then
            titleSelection = "Warehouse"
        End If
        If rdoCUST_CODE.Checked Then
            titleSelection = "Customer"
        End If
        titlePeriod = "For The Period Ending " & Format(dteSaturday.DateTime, "MM/dd/yy")
        If EOM_MODE Then
            titlePeriod = titlePeriod & " (EOM Mode)"
        End If
        If rdoUnits.Checked Then
            titleValues = "** SHOWING UNITS ***"
        End If
        If rdoDollars.Checked Then
            titleValues = "** SHOWING DOLLARS **"
        End If
        grdWBFFSR01.Text = $"{titleSelection} {titleRoot} {titlePeriod} {titleValues}."
    End Sub

    Private Sub rdoCUST_CODE_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCUST_CODE.CheckedChanged
        setWebEDIOption()
    End Sub

    Private Sub setWebEDIOption()
        If rdoCUST_CODE.Checked Then
            chkSeperateWebEDI.Checked = False
            chkSeperateWebEDI.Visible = False
        Else
            chkSeperateWebEDI.Checked = True
            chkSeperateWebEDI.Visible = True
        End If
    End Sub

    Private Sub rdoWSHE_CODE_CheckedChanged(sender As Object, e As EventArgs) Handles rdoWSHE_CODE.CheckedChanged
        setWebEDIOption()
    End Sub

    Private Sub setEOM_MODES()
        Dim COLS As String() = {"TY_WK1", "TY_WK2", "TY_WK3", "TY_WK4", "TY_WK5", "TY_WK6", "TY_WK7", "TOT_TY_WK", "TOT_LY_WK", "LY_MTD", "WTD_FULL_WK_YOY", "WTD_FULL_WK_YOY_PCT"}
        For Each COL As String In COLS
            grdWBFFSR01.DisplayLayout.Bands(0).Columns(COL).Hidden = EOM_MODE
        Next
    End Sub

#End Region

#Region "Form Controls"

#End Region
End Class