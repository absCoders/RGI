Imports System.Drawing
Imports System.Math
Imports System.IO

Imports System.Collections
Imports System.Xml.Serialization

Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Resources.Appearance
Imports Infragistics.UltraChart.Core
Imports Infragistics.UltraChart.Core.ColorModel
Imports Infragistics.UltraChart.Data
Imports Infragistics.UltraChart.Core.Layers
Imports Infragistics.UltraChart.Core.Primitives
Imports Infragistics.UltraChart.Shared.Styles

Public Class TAFDASH1

    ' facts
    ' cmoplete rest of dash
    ' best worst
    ' bug where click off trend and no column headings show up

    ' sotsrep1 on the backbone
    ' bad sales rep codes for COS people - s/b F000
    ' svia codes for henry


    Dim RYP As String
    Dim TATDASHX As String

    Dim DASH_CODE As String
    Dim DASH_VIEW As Int64

    Dim SCOPE() As String

    Dim YWP(,) As String
    Dim YWPD() As Date
    Dim YWF(,) As String
    Dim YWFD() As Date
    Dim YWN(,) As String
    Dim YWND() As Date

    Dim COL_M() As String
    Dim COL_W() As String
    Dim YP() As String
    Dim YW() As String

    Dim sqlSUM As String
    Dim rowTATDASH2 As DataRow

    Dim US_STATES() As String
    Dim USMap As MapLayer

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select DASH_CODE, DASH_VIEW from TATDASH2 where ROWNUM < 1"
            TATDASHX = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("ALTER TABLE " & TATDASHX & " ADD PRIMARY KEY (DASH_CODE, DASH_VIEW)")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & TATDASHX & " ADD CODE1 VARCHAR2(30)")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & TATDASHX & " ADD CODE2 VARCHAR2(30)")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & TATDASHX & " ADD DESC1 VARCHAR2(60)")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & TATDASHX & " ADD DESC2 VARCHAR2(60)")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & TATDASHX & " ADD STATE_CODE VARCHAR2(2)")
            ASCDATA1.ExecuteSQL("CREATE INDEX I_" & TATDASHX & "_1 ON " & TATDASHX & " (DASH_CODE,DASH_VIEW,CODE1,CODE2,STATE_CODE)")
            sqlSUM = ""
            For Each MW As String In New String() {"M", "W"}
                For iVAL As Integer = 1 To 6
                    For i As Integer = 5 To 0 Step -1
                        Dim COLUMN_NAME As String = "VAL" & CStr(iVAL) & MW & Format(i, "0")
                        sqlSUM &= ", SUM (" & COLUMN_NAME & ") " & COLUMN_NAME
                        ASCDATA1.ExecuteSQL("ALTER TABLE " & TATDASHX & " ADD " & COLUMN_NAME & " NUMBER (13,2)")
                    Next
                Next
            Next


            With .Tables.Add("TATDASHF")
                .Columns.Add("FACT_NO")
                .Columns.Add("FACT_DESC")
                .Columns.Add("FACT_VALUE")
            End With

            ASCMAIN1.sql = "Select * from TATDASH2"
            Create_TDA(.Tables.Add, "TATDASH2", "**", 0, False)
            .Tables("TATDASH2").Columns.Add("DESCRIPTION", GetType(System.String), "DASH_CODE1_CAPTION + '/' + DASH_CODE2_CAPTION")
            .Tables("TATDASH2").Columns.Add("SQL_DETAILS")

            ASCMAIN1.sql = "Select * from TATDASH1"
            Create_TDA(.Tables.Add, "TATDASH1", "**", 0, False)
            .Tables("TATDASH1").Columns.Add("DASH_SEL")

            ASCMAIN1.sql = "Select TATDASHX.DASH_CODE, TATDASHX.DASH_VIEW" _
            & ", TATDASHX.CODE1, TATDASHX.CODE2" _
            & ", TATDASHX.DESC1, TATDASHX.DESC2" _
            & sqlSUM & " from " & TATDASHX & " TATDASHX" _
            & " where TATDASHX.DASH_CODE = :PARM1 and TATDASHX.DASH_VIEW = :PARM2 " _
            & " group by TATDASHX.DASH_CODE, TATDASHX.DASH_VIEW" _
            & ", TATDASHX.CODE1, TATDASHX.CODE2" _
            & ", TATDASHX.DESC1, TATDASHX.DESC2"
            Create_TDA(.Tables.Add, "TATDASHX", "**", 0, False, "VN")
            For Each MW As String In New String() {"M", "W"}
                For i As Integer = 5 To 0 Step -1
                    Dim COLUMN_NAME As String = "VAL0" & MW & Format(i, "0")
                    .Tables("TATDASHX").Columns.Add(COLUMN_NAME, GetType(System.Decimal))
                Next
            Next

            ASCMAIN1.sql = "Select TATDASHX.DASH_CODE, TATDASHX.DASH_VIEW" _
            & ", TATDASHX.CODE1 CODE_VALUE, TATDASHX.DESC1 DESC_VALUE" _
            & sqlSUM & " from " & TATDASHX & " TATDASHX" _
            & " group by TATDASHX.DASH_CODE, TATDASHX.DASH_VIEW" _
            & ", TATDASHX.CODE1, TATDASHX.DESC1"
            Create_TDA(.Tables.Add, "TATDASHY", "**", 0, False)
            For Each MW As String In New String() {"M", "W"}
                For i As Integer = 5 To 0 Step -1
                    Dim COLUMN_NAME As String = "VAL0" & MW & Format(i, "0")
                    .Tables("TATDASHY").Columns.Add(COLUMN_NAME, GetType(System.Decimal))
                Next
            Next

            ASCMAIN1.sql = "Select TATDASHX.STATE_CODE" _
            & sqlSUM & " from " & TATDASHX & " TATDASHX" _
            & " where TATDASHX.DASH_CODE = :PARM1 and TATDASHX.DASH_VIEW = :PARM2 " _
            & " and (TATDASHX.CODE1 = :PARM3 or :PARM4 = '0') and (TATDASHX.CODE2 = :PARM5 or :PARM6 <> '2')" _
            & " group by TATDASHX.STATE_CODE"
            Create_TDA(.Tables.Add, "TATDASHS", "**", 0, False, "VNVVVV")
            For Each MW As String In New String() {"M", "W"}
                For i As Integer = 5 To 0 Step -1
                    Dim COLUMN_NAME As String = "VAL0" & MW & Format(i, "0")
                    .Tables("TATDASHS").Columns.Add(COLUMN_NAME, GetType(System.Decimal))
                Next
            Next


            Create_Relation("TATDASH2", "TATDASHX", "DASH_CODE,DASH_VIEW")

            For Each MW As String In New String() {"M", "W"}
                For i As Integer = 5 To 0 Step -1
                    Dim COLUMN_NAME As String = "VAL0" & MW & Format(i, "0")
                    .Tables("TATDASH1").Columns.Add(COLUMN_NAME, GetType(System.Decimal))
                    .Tables("TATDASH2").Columns.Add(COLUMN_NAME, GetType(System.Decimal), "SUM(CHILD." & COLUMN_NAME & ")")
                Next
            Next

            Create_TDA(.Tables.Add, "TATSTATE", "*", 0, False)
            With .Tables("TATSTATE")
                .Columns.Add("AMT", GetType(System.Int32))
                .Columns.Add("MAP_INDEX", GetType(System.Int32))
            End With

        End With

        Fill_Records("TATSTATE")
        Dim rowTATSTATE As DataRow = dst.Tables("TATSTATE").NewRow
        rowTATSTATE.Item("STATE_CODE") = "??"
        rowTATSTATE.Item("STATE_NAME") = "Unknown"
        dst.Tables("TATSTATE").Rows.Add(rowTATSTATE)

        grdTATDASH1.DataSource = dst.Tables("TATDASH1")
        grdTATDASHF.DataSource = dst.Tables("TATDASHF")
        grdTATDASHF.DisplayLayout.Bands(0).ColHeadersVisible = False

        grdTATDASHX.DataSource = dst.Tables("TATDASHX")
        grdTATDASHY.DataSource = dst.Tables("TATDASHY")

        Fill_Records("TATDASH2")

        Fill_Records("TATDASH1")
        For Each rowTATDASH1 As DataRow In dst.Tables("TATDASH1").Rows
            rowTATDASH1.Item("DASH_SEL") = "1"
        Next
        Sort_grdColumns(grdTATDASH1, "DASH_CODE")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdTATDASHX, grdTATDASHY}
            With grd.DisplayLayout.Bands(0)
                For Each MW As String In New String() {"M", "W"}
                    For iVAL As Integer = 0 To 6
                        For i As Integer = 5 To 0 Step -1
                            Dim COLUMN_NAME As String = "VAL" & CStr(iVAL) & MW & Format(i, "0")
                            With .Columns(COLUMN_NAME)
                                If iVAL = 0 Then
                                    If MW = "M" Then
                                        .Header.Appearance.BackColor2 = Color.Green
                                        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                                        If i = 0 Then
                                            .CellAppearance.BackColor = Color.LightGreen
                                        End If
                                    Else
                                        .Header.Appearance.BackColor2 = Color.Violet
                                        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                                        If i = 0 Then
                                            .CellAppearance.BackColor = Color.Pink
                                        End If
                                    End If
                                    .Width = 100
                                    .Format = "###,##0"
                                    Create_Summary(grd, COLUMN_NAME)
                                Else
                                    .Hidden = True
                                End If
                            End With
                        Next
                    Next
                Next

                If grd.Name = "grdTATDASHX" Then
                    .Columns("CODE1").CellAppearance.BackColor = Color.Beige
                    .Columns("CODE1").CellAppearance.ForeColor = Color.Green
                    .Columns("CODE2").CellAppearance.BackColor = Color.Beige
                    .Columns("CODE2").CellAppearance.ForeColor = Color.Purple
                    .Columns("DESC1").CellAppearance.BackColor = Color.Beige
                    .Columns("DESC1").CellAppearance.ForeColor = Color.Green
                    .Columns("DESC2").CellAppearance.BackColor = Color.Beige
                    .Columns("DESC2").CellAppearance.ForeColor = Color.Purple
                    .Columns("CODE1").Header.Fixed = True
                    .Columns("CODE2").Header.Fixed = True
                    .Columns("DESC1").Header.Fixed = True
                    .Columns("DESC2").Header.Fixed = True
                    .Columns("CODE1").Width = 80
                    .Columns("CODE2").Width = 80
                    .Columns("DESC1").Width = 140
                    .Columns("DESC2").Width = 140
                Else
                    .Columns("CODE_VALUE").Header.Caption = "Code"
                    .Columns("DESC_VALUE").Header.Caption = "Description"

                    .Columns("CODE_VALUE").CellAppearance.BackColor = Color.Beige
                    .Columns("CODE_VALUE").CellAppearance.ForeColor = Color.Blue
                    .Columns("DESC_VALUE").CellAppearance.BackColor = Color.Beige
                    .Columns("DESC_VALUE").CellAppearance.ForeColor = Color.Blue
                    .Columns("CODE_VALUE").Header.Fixed = True
                    .Columns("DESC_VALUE").Header.Fixed = True
                    .Columns("CODE_VALUE").Width = 80
                    .Columns("DESC_VALUE").Width = 140 + 80 + 140 + 4
                End If

                .Columns("DASH_VIEW").Hidden = True
                .Columns("DASH_CODE").Hidden = True

            End With
        Next


        grdTATDASH1.DisplayLayout.Bands(0).Columns("VAL0M0").Header.Caption = "Value"

        Create_Summary(grdTATDASH1, "DASH_CODE", "Count")


        With grdTATDASH1.DisplayLayout.Bands("TATDASH1")
            .Columns("DASH_CODE").Header.Fixed = True
        End With

        cbeOPS_YYYYPP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' and OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' order by OPS_YYYYPP DESC")
        cbeOPS_YYYYPP.SelectedIndex = 0

        Dim dvw As DataView = dst.Tables("TATSTATE").DefaultView
        dvw.RowFilter = "AMT <> 0"
        grdTATSTATE.DataSource = dvw

        Create_Summary(grdTATSTATE, "STATE_CODE", "Count")
        Create_Summary(grdTATSTATE, "AMT")

        With chtTotals
            .Axis.X.ScrollScale.Visible = True
            .Axis.Y.ScrollScale.Visible = True

            .Axis.X.ScrollScale.Scale = 1 ' 0.25
            .Axis.Y.ScrollScale.Scale = 1 ' 0.25
            .EnableCrossHair = True
            '.ColorModel.ModelStyle = ColorModels.CustomLinear '  CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        End With

        Setup_Map()


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                If cbeOPS_YYYYPP.Value = "" Then
                    EMsg &= vbCrLf & "You must specify a Period to View"
                End If

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
                Call Load_Record()
                Call Mode_Settings(True)


            Case "Cancel", "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

                .Groups("Data").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpDashboardControls, False)
        Set_Read_Only(cbeTATDASH2, Not ScreenMode)

        'Set_Read_Only(cbeOPS_YYYYPP, ScreenMode)

        splDashboard.Visible = ScreenMode
        grpDashboardControls.Visible = ScreenMode
        cbeTATDASH2.Visible = ScreenMode

        Setup_tabMain()

        If ScreenMode Then
        Else
            Clear_Record()
        End If


        With grdTATDASH1
            .DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ExtendLastColumn
            If ScreenMode Then
                .DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect
                'grdTATDASH1.Parent = splGeneralInfo.Panel1
                'grdTATDASH1.DisplayLayout.CaptionVisible = DefaultableBoolean.True
            Else
                .DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
                'grdTATDASH1.Parent = grpTATDASH1
                'grdTATDASH1.DisplayLayout.CaptionVisible = DefaultableBoolean.False
            End If

            With .DisplayLayout.Bands(0)

                .Columns("DASH_SEL").Hidden = ScreenMode

                .Columns("VAL0M0").Hidden = Not ScreenMode
                .Columns("VAL0M0").Format = "###,##0"
            End With

        End With
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        ' dst.Tables("TATDASH2").Rows.Clear()
        EnforceConstraints(True)

        DASH_CODE = ""
        DASH_VIEW = 0
        'cbeOPS_YYYYPP.Value = ""

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)
        RYP = cbeOPS_YYYYPP.Value

        Load_TATDASH2()
        Load_TATDASHF()
        Setup_Dashboard()
        Setup_MW()

        Sort_grdColumns(grdTATDASH1, "DASH_CODE")
        If grdTATDASH1.Rows.Count > 0 Then
            'grdTATDASH1.ActiveRow = grdTATDASH1.Rows(0)
            grdTATDASH1.Rows(0).Activate()
            Setup_grdTATDASH1()
            Setup_grdTATDASHX()
        End If

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()
        Stop
        Call CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdTATSTATE, "CC", "Best", "Worst")
        Load_Popup_Menu(grdTATDASHX, "SSB", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If


        'If tlb_pop.Tools.Exists("Show On Hand") Then
        '    tlb_sbt = DirectCast(tlb_pop.Tools("Show On Hand"), UltraWinToolbars.StateButtonTool)
        '    tlb_sbt.Tag = "H"
        '    tlb_sbt.Checked = InStr(DATA_TYPES_visible, tlb_sbt.Tag)
        'End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

                'Case "Show On Hand", "Show On PO", "Show Sales", "Show Weeks of Supply"
                '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                '    If tlb_sbt.Checked Then
                '        If InStr(DATA_TYPES_visible, tlb_sbt.Tag) = 0 Then
                '            DATA_TYPES_visible &= tlb_sbt.Tag
                '        End If
                '    Else
                '        DATA_TYPES_visible = Replace(DATA_TYPES_visible, tlb_sbt.Tag, "")
                '    End If
                '    Show_Visible_DATA_TYPES()

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Lot Inquiry"

            '    Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Text
            '    Dim LOT_NO As String = grd.ActiveRow.Cells("LOT_NO").Text
            '    Dim LOT_SEQ_NO As Int64 = grd.ActiveRow.Cells("LOT_SEQ_NO").Text
            '    Context_Launch("Load", WHSE_CODE & vbTab & LOT_NO & vbTab & CStr(LOT_SEQ_NO), e.Tool.Key, "ICFLOTH1")

        End Select
    End Sub

    Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)
        Select Case e.Tool.Key
            Case "Best"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorEnd = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                Application.DoEvents()
                grdTATSTATE.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

            Case "Worst"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                Application.DoEvents()
                grdTATSTATE.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        End Select

    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "OPS_YYYYPP"
                Call Click_Command("View")
        End Select
    End Sub
#End Region

    Sub Load_TATDASHF()
        dst.Tables("TATDASHF").Rows.Clear()

        ASCMAIN1.sql = "Select Count (*) from SOTORDR1 where SO_STATUS_CODE = 'O'"
        dst.Tables("TATDASHF").Rows.Add(New Object() {1, "Open Orders", ASCDATA1.GetDataValue})

    End Sub

    Sub Load_TATDASH2()

        Dim RYP = cbeOPS_YYYYPP.Value
        Dim RYW As String = ""

        If RYP = ASCMAIN1.CYP Then
            RYW = ASCMAIN1.CYW
        Else
            ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & RYP & "'"
            RYW = Val(ASCDATA1.GetDataValue)
        End If
        ASCMAIN1.Get_Week_Range(-60, YWPD, YWP)
        ASCMAIN1.Get_Week_Range(60, YWFD, YWF)
        ASCMAIN1.Get_Week_Range(60, YWND, YWN, ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -52))

        ReDim COL_M(5)
        ReDim COL_W(5)
        ReDim YP(5)
        ReDim YW(5)
        For i As Integer = 5 To 0 Step -1
            YP(i) = ASCMAIN1.Period_Calc(RYP, -1 * i)
            COL_M(i) = ASCMAIN1.Get_Legend(YP(i), False, True)
        Next
        For i As Integer = 5 To 0 Step -1
            YW(i) = ASCMAIN1.Week_Calc(RYW, -1 * i)
            If i = 0 Then
                COL_W(i) = "TW"
            ElseIf i = 1 Then
                COL_W(i) = "LW"
            Else
                COL_W(i) = "W" & Format(i, "0") & "ago"
            End If
        Next

        ASCDATA1.ExecuteSQL("Truncate Table " & TATDASHX)

        dst.Tables("TATDASHX").Rows.Clear()

        For Each rowTATDASH1 As DataRow In dst.Tables("TATDASH1").Select("DASH_SEL = '1'")
            Dim DASH_CODE As String = rowTATDASH1.Item("DASH_CODE") & ""

            For Each rowTATDASH2 As DataRow In dst.Tables("TATDASH2").Select("DASH_CODE = '" & DASH_CODE & "'")

                Dim DASH_VIEW As String = Val(rowTATDASH2.Item("DASH_VIEW") & "")

                Dim DASH_CODE1_COLUMN_NAME As String = rowTATDASH2.Item("DASH_CODE1_COLUMN_NAME") & ""
                Dim DASH_CODE2_COLUMN_NAME As String = rowTATDASH2.Item("DASH_CODE2_COLUMN_NAME") & ""
                Dim DASH_DESC1_COLUMN_NAME As String = rowTATDASH2.Item("DASH_DESC1_COLUMN_NAME") & ""
                Dim DASH_DESC2_COLUMN_NAME As String = rowTATDASH2.Item("DASH_DESC2_COLUMN_NAME") & ""
                Dim DASH_TABLES As String = rowTATDASH2.Item("DASH_TABLES") & ""
                Dim DASH_JOIN As String = rowTATDASH2.Item("DASH_JOIN") & ""
                Dim DASH_STATE_COLUMN_NAME As String = rowTATDASH2.Item("DASH_STATE_COLUMN_NAME") & ""
                Dim DASH_DETAIL_CODE_COLUMN_NAME As String = rowTATDASH2.Item("DASH_DETAIL_CODE_COLUMN_NAME") & ""
                Dim DASH_DETAIL_DESC_COLUMN_NAME As String = rowTATDASH2.Item("DASH_DETAIL_DESC_COLUMN_NAME") & ""
                Dim DASH_DETAIL_TABLE_NAME As String = rowTATDASH2.Item("DASH_DETAIL_TABLE_NAME") & ""
                Dim DASH_DETAIL_TABLE_JOIN As String = rowTATDASH2.Item("DASH_DETAIL_TABLE_JOIN") & ""

                Dim sqlSelect As String = "Select '" & DASH_CODE & "' DASH_CODE" _
                        & ", " & CStr(DASH_VIEW) & " DASH_VIEW " _
                        & ", " & DASH_CODE1_COLUMN_NAME & ", " & DASH_CODE2_COLUMN_NAME _
                        & ", " & DASH_DESC1_COLUMN_NAME & ", " & DASH_DESC2_COLUMN_NAME _
                        & ", " & IIf(DASH_STATE_COLUMN_NAME = "", "NULL", DASH_STATE_COLUMN_NAME)

                Dim sqlSelectDetails As String = "Select '" & DASH_CODE & "' DASH_CODE" _
                        & ", " & CStr(DASH_VIEW) & " DASH_VIEW " _
                        & ", " & DASH_DETAIL_CODE_COLUMN_NAME & " CODE_VALUE" _
                        & ", " & DASH_DETAIL_DESC_COLUMN_NAME & " DESC_VALUE"

                Dim sqlGroupBy As String = " group by " _
                    & DASH_CODE1_COLUMN_NAME & ", " & DASH_CODE2_COLUMN_NAME _
                    & ", " & DASH_DESC1_COLUMN_NAME & ", " & DASH_DESC2_COLUMN_NAME _
                    & IIf(DASH_STATE_COLUMN_NAME = "", "", ", " & DASH_STATE_COLUMN_NAME)

                Dim sqlGroupByDetails As String = " group by " _
                    & DASH_DETAIL_CODE_COLUMN_NAME & ", " & DASH_DETAIL_DESC_COLUMN_NAME


                Dim sql As String = ""
                Dim COLS As String = "DASH_CODE, DASH_VIEW, CODE1, CODE2, DESC1, DESC2, STATE_CODE"
                Select Case DASH_CODE
                    Case "A"

                    Case "B"

                        sql = sqlSelect _
                        & ", SUM (ARTOPEN1.INV_BALANCE) VAL1M0" _
                        & " from ARTOPEN1" & DASH_TABLES _
                        & " where ARTOPEN1.INV_BALANCE <> 0 " _
                        & DASH_JOIN _
                        & sqlGroupBy

                        ASCMAIN1.sql = "Insert into " & TATDASHX _
                        & " (" & COLS & ", VAL1M0)" _
                        & " Select * from (" & sql & ")"

                        ASCDATA1.ExecuteSQL()


                    Case "C"

                        sql = sqlSelect _
                        & ", SUM (APTINVH1.INV_BALANCE) VAL1M0" _
                        & " from APTINVH1,APTVEND1" & DASH_TABLES _
                        & " where APTINVH1.INV_STATUS IN ('O','H')" _
                        & "   and APTVEND1.VEND_CODE = APTINVH1.VEND_CODE " _
                        & DASH_JOIN _
                        & sqlGroupBy

                        ASCMAIN1.sql = "Insert into " & TATDASHX _
                        & " (" & COLS & ", VAL1M0)" _
                        & " Select * from (" & sql & ")"

                        ASCDATA1.ExecuteSQL()


                    Case "D"

                        sql = sqlSelect _
                        & ", SUM (ICTLOTD2.QTY_ON_HAND * DECODE(ICTLOTD2.PACK_CODE,'000',ICTLOTD2.CATCH_WEIGHT,ICTPACK1.PACK_FACTOR)) VAL1M0" _
                        & ", SUM (ICTLOTD2.QTY_ON_HAND * DECODE(ICTLOTD2.PACK_CODE,'000',ICTLOTD2.CATCH_WEIGHT,ICTPACK1.PACK_FACTOR) * ICTLOTD2.STANDARD_COST) VAL2M0" _
                        & " from ICTLOTD2,ICTPACK1" & DASH_TABLES _
                        & " where ICTLOTD2.QTY_ON_HAND <> 0 " _
                        & "   and ICTPACK1.PACK_CODE = ICTLOTD2.PACK_CODE " _
                        & DASH_JOIN _
                        & sqlGroupBy

                        ASCMAIN1.sql = "Insert into " & TATDASHX _
                        & " (" & COLS & ", VAL1M0, VAL2M0)" _
                        & " Select * from (" & sql & ")"

                        ASCDATA1.ExecuteSQL()

                    Case "E"

                        sql = sqlSelect & vbCrLf

                        Dim COLUMN_ALIASes As String = ""
                        For i As Integer = 0 To 0 Step -1
                            Dim iCOL As Integer = 0
                            Dim COLUMN_ALIAS As String = ""
                            For Each COLUMN_NAME As String In New String() { _
                            "SOTORDR3.SO_LOT_CASES", _
                            "SOTORDR3.SO_LOT_UNITS", _
                            "SOTORDR3.SO_LOT_UNITS * SOTORDR2.ORDR_PRICE_GRS", _
                            "SOTORDR3.SO_LOT_UNITS * SOTORDR2.ORDR_PRICE_NET", _
                            "SOTORDR3.SO_LOT_UNITS * SOTORDR2.REBATE", _
                            "SOTORDR3.SO_LOT_UNITS * SOTORDR2.ORDR_PRICE_NET - SOTORDR3.STD_COST_EXT"}
                                iCOL += 1
                                COLUMN_ALIAS = "VAL" & CStr(iCOL) & "M" & CStr(i)
                                COLUMN_ALIASes &= "," & COLUMN_ALIAS
                                sql &= ", SUM (" & COLUMN_NAME & ") " & COLUMN_ALIAS & vbCrLf
                                COLUMN_ALIAS = "VAL" & CStr(iCOL) & "W" & CStr(i)
                                COLUMN_ALIASes &= "," & COLUMN_ALIAS
                                sql &= ", SUM (" & COLUMN_NAME & ") " & COLUMN_ALIAS & vbCrLf
                            Next
                        Next

                        sql &= "" _
                        & " from SOTORDR1,SOTORDR2,SOTORDR3" & DASH_TABLES _
                        & " where SOTORDR3.SO_LOT_UNITS <> 0 " _
                        & "   and SOTORDR1.SO_ORDER_NO = SOTORDR2.SO_ORDER_NO " _
                        & "   and SOTORDR2.SO_ORDER_NO = SOTORDR3.SO_ORDER_NO " _
                        & "   and SOTORDR2.SO_ORDER_LNO = SOTORDR3.SO_ORDER_LNO " _
                        & DASH_JOIN _
                        & sqlGroupBy

                        ASCMAIN1.sql = "Insert into " & TATDASHX _
                        & " (" & COLS & COLUMN_ALIASes & ")" _
                        & " Select * from (" & sql & ")"

                        ASCDATA1.ExecuteSQL()

                    Case "F"

                        sql = sqlSelect & vbCrLf

                        Dim COLUMN_ALIASes As String = ""
                        For i As Integer = 5 To 0 Step -1
                            Dim iCOL As Integer = 0
                            Dim COLUMN_ALIAS As String = ""
                            For Each COLUMN_NAME As String In New String() { _
                            "SOTINVH0.QTY_CASES", _
                            "SOTINVH0.QTY_UNITS", _
                            "SOTINVH0.QTY_UNITS * SOTINVH0.ORDR_PRICE_GRS", _
                            "SOTINVH0.QTY_UNITS * SOTINVH0.ORDR_PRICE_NET", _
                            "SOTINVH0.REBATE", _
                            "SOTINVH0.QTY_UNITS * SOTINVH0.ORDR_PRICE_NET - SOTINVH0.STD_COST_EXT"}
                                iCOL += 1
                                COLUMN_ALIAS = "VAL" & CStr(iCOL) & "M" & CStr(i)
                                COLUMN_ALIASes &= "," & COLUMN_ALIAS
                                sql &= ", SUM (DECODE(SOTINVH0.OPS_YYYYPP,'" & YP(i) & "'," & COLUMN_NAME & ",0)) " & COLUMN_ALIAS & vbCrLf
                                COLUMN_ALIAS = "VAL" & CStr(iCOL) & "W" & CStr(i)
                                COLUMN_ALIASes &= "," & COLUMN_ALIAS
                                sql &= ", SUM (DECODE(SOTINVH0.OPS_YYYYWW,'" & YW(i) & "'," & COLUMN_NAME & ",0)) " & COLUMN_ALIAS & vbCrLf
                            Next
                        Next

                        sql &= "" _
                        & " from SOTINVH0,SOTINVH1" & DASH_TABLES _
                        & " where SOTINVH0.QTY_UNITS <> 0 " _
                        & "   and SOTINVH1.SO_ORDER_NO = SOTINVH0.SO_ORDER_NO " _
                        & "   and SOTINVH0.OPS_YYYYPP >= '" & YP(5) & "' and SOTINVH0.OPS_YYYYPP <= '" & YP(0) & "' " _
                        & DASH_JOIN _
                        & sqlGroupBy

                        ASCMAIN1.sql = "Insert into " & TATDASHX _
                        & " (" & COLS & COLUMN_ALIASes & ")" _
                        & " Select * from (" & sql & ")"

                        ASCDATA1.ExecuteSQL()


                    Case "G"
                        sql = sqlSelect & vbCrLf

                        Dim COLUMN_ALIASes As String = ""
                        For i As Integer = 0 To 0 Step -1
                            Dim iCOL As Integer = 0
                            Dim COLUMN_ALIAS As String = ""
                            For Each COLUMN_NAME As String In New String() { _
                            "POTORDR2.PO_CASES", _
                            "POTORDR2.PO_UNITS", _
                            "POTORDR2.PO_UNITS * POTORDR2.PURCHASE_COST", _
                            "POTORDR2.PO_CASES_PRESOLD", _
                            "POTORDR2.PO_UNITS_PRESOLD"}
                                iCOL += 1
                                COLUMN_ALIAS = "VAL" & CStr(iCOL) & "M" & CStr(i)
                                COLUMN_ALIASes &= "," & COLUMN_ALIAS
                                sql &= ", SUM (" & COLUMN_NAME & ") " & COLUMN_ALIAS & vbCrLf
                                COLUMN_ALIAS = "VAL" & CStr(iCOL) & "W" & CStr(i)
                                COLUMN_ALIASes &= "," & COLUMN_ALIAS
                                sql &= ", SUM (" & COLUMN_NAME & ") " & COLUMN_ALIAS & vbCrLf
                            Next
                        Next

                        sql &= "" _
                        & " from POTORDR1,POTORDR2" & DASH_TABLES _
                        & " where POTORDR1.PO_STATUS_CODE = 'O'" _
                        & "   and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO " _
                        & DASH_JOIN _
                        & sqlGroupBy

                        ASCMAIN1.sql = "Insert into " & TATDASHX _
                        & " (" & COLS & COLUMN_ALIASes & ")" _
                        & " Select * from (" & sql & ")"

                        ASCDATA1.ExecuteSQL()

                End Select

                If DASH_DETAIL_CODE_COLUMN_NAME <> "" Then
                    sql = Replace(sql, sqlSelect, sqlSelectDetails)
                    sql = Replace(sql, sqlGroupBy, sqlGroupByDetails)
                    If DASH_DETAIL_TABLE_NAME <> "" Then
                        sql = Replace(sql, DASH_TABLES, DASH_TABLES & "," & DASH_DETAIL_TABLE_NAME)
                        sql = Replace(sql, DASH_JOIN, DASH_JOIN & " and " & DASH_DETAIL_TABLE_JOIN)
                    End If
                    rowTATDASH2.Item("SQL_DETAILS") = sql
                End If

                If ASCMAIN1.sql <> "" Then
                    Fill_Records("TATDASHX", New String() {DASH_CODE, CStr(DASH_VIEW)}, False)
                End If

            Next
        Next

        Setup_DataType()

        grdTATDASHX.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdTATDASHX.DisplayLayout.Bands(0).SortedColumns.Add("CODE1", False, True)

        For Each rowTATDASH1 As DataRow In dst.Tables("TATDASH1").Rows
            Dim DASH_CODE As String = rowTATDASH1.Item("DASH_CODE")
            For Each MW As String In New String() {"M", "W"}
                For i As Integer = 5 To 0 Step -1
                    Dim COLUMN_NAME As String = "VAL0" & MW & Format(i, "0")
                    Dim COLUMN_NAME_VALUE As Decimal = Val(dst.Tables("TATDASH2").Compute("SUM(" & COLUMN_NAME & ")", "DASH_CODE = '" & DASH_CODE & "' AND DASH_VIEW = 1") & "")
                    rowTATDASH1.Item(COLUMN_NAME) = COLUMN_NAME_VALUE
                Next
            Next
        Next

    End Sub

    Sub Set_DataField(ByVal TABLE_NAME As String)

        Dim iVAL As Integer = Val(optDataType.Value & "")
        If iVAL = 0 Then iVAL = 1
        For Each MW As String In New String() {"M", "W"}
            For i As Integer = 5 To 0 Step -1
                Dim COLUMN_NAME As String
                If chkShowTrend.Checked Then
                    COLUMN_NAME = "VAL" & CStr(iVAL) & MW & Format(i, "0")
                Else
                    COLUMN_NAME = "VAL" & CStr(6 - i) & MW & Format(0, "0")
                End If
                Dim COLUMN_NAME_0 As String = "VAL" & CStr(0) & MW & Format(i, "0")
                dst.Tables(TABLE_NAME).Columns(COLUMN_NAME_0).Expression = COLUMN_NAME
            Next
        Next

    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Charts").Visible = ScreenMode And chkShowDetails.Checked AndAlso tabMain.SelectedTab.Key = "Charts"
        optLevelCharts.Visible = (ScreenMode And chkShowDetails.Checked AndAlso tabMain.SelectedTab.Key = "Charts")
        optLevelMap.Visible = (ScreenMode And chkShowDetails.Checked AndAlso tabMain.SelectedTab.Key = "Map")
    End Sub

#Region "Charts"

    Sub CreateMap()

        Dim CODE1 As String = ""
        Dim CODE2 As String = ""

        If grdTATDASHX.ActiveRow IsNot Nothing Then
            If grdTATDASHX.ActiveRow.IsDataRow Then
                CODE1 = grdTATDASHX.ActiveRow.Cells("CODE1").Text
                CODE2 = grdTATDASHX.ActiveRow.Cells("CODE2").Text
            ElseIf grdTATDASHX.ActiveRow.IsGroupByRow Then
                Dim grow As UltraWinGrid.UltraGridRow = grdTATDASHX.ActiveRow
                Do While grow IsNot Nothing
                    Dim gby As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow, UltraWinGrid.UltraGridGroupByRow)
                    If gby.Column.Key = "CODE1" Then CODE1 = gby.Value & ""
                    If gby.Column.Key = "CODE2" Then CODE2 = gby.Value & ""
                    If gby.ParentRow IsNot Nothing Then
                        grow = gby.ParentRow
                    Else
                        grow = Nothing
                    End If
                Loop
            End If
        End If

        Fill_Records("TATDASHS", New String() {DASH_CODE, DASH_VIEW, CODE1, optLevelMap.Value, CODE2, optLevelMap.Value})

        For Each ROW As DataRow In dst.Tables("TATSTATE").Rows
            ROW.Item("AMT") = 0
        Next

        Dim DATA_TYPE As String
        If chkShowTrend.Checked Then
            DATA_TYPE = "VAL0" & optMW.Value & CStr(0)
        Else
            Dim iVAL As Integer = Val(optDataType.Value & "")
            If iVAL = 0 Then iVAL = 1
            DATA_TYPE = "VAL0" & optMW.Value & CStr(6 - iVAL)
        End If

        For Each rowTATDASHS As DataRow In dst.Tables("TATDASHS").Rows
            Dim rowTATSTATE As DataRow = dst.Tables("TATSTATE").Rows.Find(rowTATDASHS.Item("STATE_CODE"))
            If rowTATSTATE IsNot Nothing Then
                rowTATSTATE.Item("AMT") = Val(rowTATSTATE.Item("AMT") & "") + Val(rowTATDASHS.Item(DATA_TYPE) & "")
            Else
                rowTATSTATE = dst.Tables("TATSTATE").Rows.Find("??")
                rowTATSTATE.Item("AMT") = Val(rowTATSTATE.Item("AMT") & "") + Val(rowTATDASHS.Item(DATA_TYPE) & "")
            End If
        Next

        Me.UltraChart1.Data.DataSource = StatesData()
        Me.UltraChart1.Data.DataBind()
        tabMain.Tabs("Map").Visible = True

        grdTATSTATE.DisplayLayout.Bands(0).Columns("AMT").Header.Caption = grdTATDASH1.ActiveRow.Cells("DASH_DESC").Text
        Sort_grdColumns(grdTATSTATE, "AMT".ToLower)


        Dim CAPTION As String = optDataType.Text
        If optLevelMap.Value = "0" Then
            'CAPTION = ""
        ElseIf optLevelMap.Value = "1" Then
            CAPTION &= " for " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE1").Header.Caption & ":" & CODE1
        ElseIf optLevelMap.Value = "2" Then
            CAPTION &= " for " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE1").Header.Caption & ":" & CODE1 _
            & ", " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE2").Header.Caption & ":" & CODE2
        End If
        CAPTION &= " - " & IIf(optMW.Value = "M", COL_M(0), COL_W(0))
        'If chkShowTrend.Checked Then
        '    CAPTION &= " - " & COL_M(0) ' grdTATDASHX.DisplayLayout.Bands(0).Columns(DATA_TYPE).Header.Caption
        'Else
        '    CAPTION &= " - " & COL_M(0)
        'End If
        grdTATSTATE.Text = CAPTION

    End Sub

    Sub CreateGraph_Totals()

        Dim chtIsVisible As Boolean = chtTotals.Visible
        chtTotals.Visible = False

        chtTotals.DataSource = Nothing

        Dim DATA_TYPE As String
        If chkShowTrend.Checked Then
            DATA_TYPE = "VAL0" & optMW.Value & CStr(0)
        Else
            Dim iVAL As Integer = Val(optDataType.Value & "")
            If iVAL = 0 Then iVAL = 1
            DATA_TYPE = "VAL0" & optMW.Value & CStr(6 - iVAL)
        End If

        'Me.SuspendLayout()

        Dim RL() As String

        chtTotals.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTotals.LabelHash = labelHash

        chtTotals.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTotals.Tooltips.FormatString = "<HIGHLOW>"

        Dim RLi As Integer = 0

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Decimal))
        End With

        Dim CODE1 As String = ""
        Dim CHARTED_CODE As String = "CODE1"
        If optLevelCharts.Value = "2" Then
            If grdTATDASHX.ActiveRow Is Nothing Then
                Exit Sub
            End If
            If grdTATDASHX.ActiveRow.IsGroupByRow Then
                Dim grow As UltraWinGrid.UltraGridRow = grdTATDASHX.ActiveRow
                Do While grow IsNot Nothing
                    Dim gby As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow, UltraWinGrid.UltraGridGroupByRow)
                    If gby.Column.Key = "CODE1" Then CODE1 = gby.Value & ""
                    If gby.ParentRow IsNot Nothing Then
                        grow = gby.ParentRow
                    Else
                        grow = Nothing
                    End If
                Loop

            ElseIf grdTATDASHX.ActiveRow.IsDataRow Then
                CODE1 = grdTATDASHX.ActiveRow.Cells("CODE1").Value & ""
            Else
                Stop ' what now batman
            End If
            CHARTED_CODE = "CODE2"
        End If

        Dim SQL1 As String = "DASH_CODE = '" & DASH_CODE & "' AND DASH_VIEW = " & CStr(DASH_VIEW)
        Dim SQL2 As String = SQL1 & IIf(optLevelCharts.Value = "1", "", " and ISNULL(CODE1,'') = '" & CODE1 & "'")

        Dim DTX As DataTable = ASCDATA1.SelectDistinct(dst.Tables("TATDASHX").Select(SQL2), CHARTED_CODE)
        DTX.Columns.Add(DATA_TYPE, GetType(System.Decimal))
        For Each rowDTX As DataRow In DTX.Rows

            Dim SQL3 As String = SQL2 & " and ISNULL(" & CHARTED_CODE & ",'') = '" & rowDTX.Item(0) & "'"
            Dim VALUE As Decimal = Val(dst.Tables("TATDASHX").Compute("SUM(" & DATA_TYPE & ")", SQL3) & "")
            rowDTX.Item(DATA_TYPE) = VALUE
        Next

        Dim PCT_at_TOP_N As Decimal = 0
        Dim VALUE_TOTAL As Decimal = Val(DTX.Compute("SUM(" & DATA_TYPE & ")", "") & "")
        Dim VALUE_CHARTED As Decimal = 0

        ReDim RL(DTX.Rows.Count - 1)
        For Each row As DataRow In DTX.Select("", DATA_TYPE & " DESC")
            RL(RLi) = row.Item(CHARTED_CODE) & "" ' & ":" & row("DESC_VALUE")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item(CHARTED_CODE), row.Item(DATA_TYPE)})

            If optChartTrend.Value = "N" And RLi <= Val(numChartTrend.Value & "") Then
                PCT_at_TOP_N = 100 * Val(row.Item(DATA_TYPE & "00")) / VALUE_TOTAL
            End If
        Next

        Dim CAPTION As String = optDataType.Text ' grdTATDASHX.DisplayLayout.Bands(0).Columns(DATA_TYPE).Header.Caption
        If optLevelCharts.Value = "1" Then
            CAPTION &= " by " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE1").Header.Caption
        Else
            CAPTION &= " for " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE1").Header.Caption & ":" & CODE1 _
            & " by " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE2").Header.Caption
        End If
        chtTotals.TitleTop.Text = CAPTION & "-" & IIf(optMW.Value = "M", COL_M(0), COL_W(0))
        chtTotals.Data.SetRowLabels(RL)
        'chtTotals.Data.SetColumnLabels(CL)

        chtTotals.DataSource = DTY
        chtTotals.PieChart.ColumnIndex = -1

        chtTotals.PieChart.OthersCategoryPercent = 2
        If optChartTrend.Value = "C" Then
            chtTotals.PieChart.OthersCategoryPercent = Val(numChartTrend.Value & "")
        Else
            chtTotals.PieChart.OthersCategoryPercent = PCT_at_TOP_N
        End If
        chtTotals.DataBind()

        chtTotals.Visible = True

        'Me.ResumeLayout()
        'Application.DoEvents()
    End Sub

    Sub CreateGraph_Trend()

        Dim chtIsVisible As Boolean = chtTrend.Visible
        chtTrend.Visible = False

        Dim periods As Integer = 6

        Dim iVAL As Integer = Val(optDataType.Value & "") + 1
        Dim DATA_TYPE As String = "VAL0" & optMW.Value & "0" ' CStr(iVAL)
        Dim S As Integer = 1
        'If DATA_TYPE = "R" Then
        '    S = -1
        'End If

        chtTrend.DataSource = Nothing

        'Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String
        ReDim CL(periods)

        For i As Integer = 1 To periods
            If optMW.Value = "M" Then
                CL(i - 1) = COL_M(6 - i)
            Else
                CL(i - 1) = COL_W(6 - i)
            End If
        Next

        chtTrend.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtTrend.LabelHash = labelHash


        Dim CODE1 As String = ""
        Dim CHARTED_CODE As String = "CODE1"
        If optLevelCharts.Value = "2" Then
            If grdTATDASHX.ActiveRow Is Nothing Then
                Exit Sub
            End If
            If grdTATDASHX.ActiveRow.IsGroupByRow Then
                Dim grow As UltraWinGrid.UltraGridRow = grdTATDASHX.ActiveRow
                Do While grow IsNot Nothing
                    Dim gby As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow, UltraWinGrid.UltraGridGroupByRow)
                    If gby.Column.Key = "CODE1" Then CODE1 = gby.Value & ""
                    If gby.ParentRow IsNot Nothing Then
                        grow = gby.ParentRow
                    Else
                        grow = Nothing
                    End If
                Loop

            ElseIf grdTATDASHX.ActiveRow.IsDataRow Then
                CODE1 = grdTATDASHX.ActiveRow.Cells("CODE1").Value & ""
            Else
                Stop ' what now batman
            End If
            CHARTED_CODE = "CODE2"
        End If


        Dim CAPTION As String = optDataType.Text ' grdTATDASHX.DisplayLayout.Bands(0).Columns(DATA_TYPE).Header.Caption
        If optLevelCharts.Value = "1" Then
            CAPTION &= " by " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE1").Header.Caption
        Else
            CAPTION &= " for " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE1").Header.Caption & ":" & CODE1 _
            & " by " & grdTATDASHX.DisplayLayout.Bands(0).Columns("CODE2").Header.Caption
        End If
        chtTrend.TitleTop.Text = CAPTION & " (" & COL_M(5) & " thru " & COL_M(0) & ")"

        chtTrend.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtTrend.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To periods
            DT.Columns.Add("P" & Format(P, "0"), GetType(System.Decimal))
        Next



        Dim SQL1 As String = "DASH_CODE = '" & DASH_CODE & "' AND DASH_VIEW = " & CStr(DASH_VIEW)
        Dim SQL2 As String = SQL1 & IIf(optLevelCharts.Value = "1", "", " and ISNULL(CODE1,'') = '" & CODE1 & "'")

        Dim DTX As DataTable = ASCDATA1.SelectDistinct(dst.Tables("TATDASHX").Select(SQL2), CHARTED_CODE, Replace(CHARTED_CODE, "CODE", "DESC"))
        For P As Integer = 1 To periods
            Dim COLUMN_NAME_period As String = "VAL0" & optMW.Value & Format(periods - P, "0")
            DTX.Columns.Add(COLUMN_NAME_period, GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0

        Dim VALUE_TOTAL As Decimal = S * Val(dst.Tables("TATDASHX").Compute("SUM(" & DATA_TYPE & ")", SQL2) & "")
        Dim VALUE_CHARTED As Decimal = 0

        Dim chart_all_others As Boolean = False

        ReDim RL(DTX.Rows.Count - 1)
        ''chtTrend.TitleTop.Text = "Trend " & optTD.Text & " " & optTrend.Text & ", by " & optRSTSLSA1.Text

        Dim rowDT As DataRow = Nothing

        For Each rowDTX As DataRow In DTX.Rows
            Dim SQL3 As String = SQL2 & " and ISNULL(" & CHARTED_CODE & ",'') = '" & rowDTX.Item(0) & "'"

            'Dim VALUE As Decimal = Val(dst.Tables("TATDASHX").Compute("SUM(" & DATA_TYPE & ")", SQL3) & "")
            'rowDTX.Item(DATA_TYPE) = VALUE

            Dim this_record_is_others As Boolean = False

            Dim U00 As Decimal = S * Val(dst.Tables("TATDASHX").Compute("SUM(" & DATA_TYPE & ")", SQL3) & "") ' S * Val(rowDTX.Item(DATA_TYPE) & "")
            Dim CODE_VALUE As String = rowDTX.Item("CODE1") & ""
            Dim DESC_VALUE As String = rowDTX.Item("DESC1") & ""

            If (optChartTrend.Value = "C" And VALUE_TOTAL > 0 AndAlso 100 * U00 / VALUE_TOTAL > Val(numChartTrend.Value & "")) _
            Or (optChartTrend.Value = "N" And RLi < Val(numChartTrend.Value & "")) Then
            Else
                this_record_is_others = True
                CODE_VALUE = "Z"
                DESC_VALUE = "All Others"
            End If

            If Not this_record_is_others Or chart_all_others Then
                If RLi <> 0 AndAlso RL(RLi - 1) = CODE_VALUE & ":" & DESC_VALUE Then
                Else
                    RL(RLi) = CODE_VALUE & ":" & DESC_VALUE
                    RLi += 1
                    rowDT = DT.NewRow
                    rowDT.Item("CODE_VALUE") = CODE_VALUE
                    rowDT.Item("DESC_VALUE") = DESC_VALUE
                    DT.Rows.Add(rowDT)
                End If

                VALUE_CHARTED += +Val(rowDTX.Item(DATA_TYPE) & "")

                For P As Integer = 1 To periods
                    Dim COLUMN_NAME_period As String = "VAL0" & optMW.Value & Format(periods - P, "0")
                    Dim UP As Decimal = S * Val(dst.Tables("TATDASHX").Compute("SUM(" & COLUMN_NAME_period & ")", SQL3) & "") ' S * Val(rowDTX.Item(DATA_TYPE) & "")

                    rowDT.Item("P" & Format(P, "0")) = Val(rowDT.Item("P" & Format(P, "0")) & "") _
                                                      + UP
                Next
            End If

        Next


        chtTrend.Data.SetRowLabels(RL)
        chtTrend.Data.SetColumnLabels(CL)

        Dim CHART_CAPTION As String = ""
        Dim VALUE_PCT As Decimal = 0
        If VALUE_TOTAL <> 0 Then
            VALUE_PCT = VALUE_CHARTED / VALUE_TOTAL
        End If
        If optChartTrend.Value = "C" Then
            CHART_CAPTION = "Cut-off " & numChartTrend.Value & "%, Charting " & CStr(DT.Rows.Count) & " of " & CStr(DTX.Rows.Count) & ", " & Format(VALUE_PCT, "##.0%")
        Else
            CHART_CAPTION = "Top " & numChartTrend.Value & " of " & CStr(DTX.Rows.Count) & ", " & Format(VALUE_PCT, "##.0%")
        End If
        chtTrend.TitleBottom.Text = CHART_CAPTION

        chtTrend.DataSource = DT
        chtTrend.DataBind()
        chtTrend.Visible = chkShowTrend.Checked

        'Me.ResumeLayout()
        'Application.DoEvents()
    End Sub

    Private Sub chtTrend_ChartDataClicked(ByVal sender As System.Object, ByVal e As Infragistics.UltraChart.Shared.Events.ChartDataEventArgs) Handles chtTrend.ChartDataClicked
        Select_CODE_VALUE_from_TATDASHX(Split(e.RowLabel & ":", ":")(0))
    End Sub

    Private Sub cmdChartRedraw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChartRedraw.Click
        CreateGraph_Totals()
        CreateGraph_Trend()
    End Sub

    Private Sub tbkChartTrend_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbkChartTrend.Scroll
        chtTrend.Axis.Y.ScrollScale.Scale = (100 - Me.tbkChartTrend.Value) / 100.0
    End Sub

    Private Sub chtTotals_ChartDataClicked(ByVal sender As System.Object, ByVal e As Infragistics.UltraChart.Shared.Events.ChartDataEventArgs) Handles chtTotals.ChartDataClicked
        'Select_CODE_VALUE_from_TATDASHX(Split(e.RowLabel & ":", ":")(0))
    End Sub

    Sub Select_CODE_VALUE_from_TATDASHX(ByVal CODE_VALUE As String)
        Exit Sub
        'For Each grow As UltraWinGrid.UltraGridRow In grdTATDASHX.Rows
        '    If grow.Cells("CODE1").Value & "" = CODE_VALUE Then
        '        grdTATDASHX.ActiveRow = grow
        '        grdTATDASHX.Selected.Rows.Clear()
        '        grow.Selected = True
        '        Exit Sub
        '    End If
        'Next
    End Sub

#End Region


    Private Sub grdTATSTATE_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdTATSTATE.InitializeRow
        If USMap.COLORS.ContainsKey(e.Row.Cells("STATE_NAME").Text) Then
            e.Row.Cells("AMT").Appearance.ForeColor = USMap.COLORS(e.Row.Cells("STATE_NAME").Text)
        End If
    End Sub

    Sub Setup_Map()
        '' create the layer
        Dim points As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), ASCMAIN1.Folders("Images") & "ABS\UsMap\US_STATES.xml")
        USMap = New MapLayer(points)

        US_STATES = USMap.STATES
        For i As Integer = 0 To USMap.STATES.Length - 1
            Dim rowTATSTATE() As DataRow = dst.Tables("TATSTATE").Select("STATE_NAME = '" & USMap.STATES(i) & "'")
            If rowTATSTATE.Length = 1 Then
                rowTATSTATE(0).Item("MAP_INDEX") = i
                rowTATSTATE(0).Item("STATE_NAME") = USMap.STATES(i)
            End If
            ' Add(New Object() {"", USMap.STATES(i), 0})
        Next

        '' set the layer
        Me.UltraChart1.ChartType = ChartType.Composite
        Me.UltraChart1.CompositeChart.ChartAreas.Add(New ChartArea())
        Me.UltraChart1.UserLayerIndex = New String() {"USMap"}
        Me.UltraChart1.Layer.Add("USMap", USMap)

        '' set the tooltip.
        Dim labelRenderers As New Hashtable()
        labelRenderers.Add("USMap", New USMapLabelRenderer(dst.Tables("TATSTATE")))
        Me.UltraChart1.LabelHash = labelRenderers
        Me.UltraChart1.Tooltips.FormatString = "<USMap>"

        ''set border
        Me.UltraChart1.Border.CornerRadius = 20
        Me.UltraChart1.Border.Thickness = 0
        Me.UltraChart1.BackColor = Color.White

        '' set color model
        'Me.UltraChart1.ColorModel.ColorBegin = Color.AliceBlue
        Me.UltraChart1.ColorModel.ColorBegin = Color.Red
        Me.UltraChart1.ColorModel.ColorEnd = Color.Blue '  Color.Yellow ' Color.FromArgb(24, 89, 165)
        Me.UltraChart1.ColorModel.AlphaLevel = 255
        Me.UltraChart1.ColorModel.ModelStyle = ColorModels.DataValueLinearRange

        '' legend
        Me.UltraChart1.Legend.Visible = True
        Me.UltraChart1.Axis.X.Extent = 10
        Me.UltraChart1.Legend.SpanPercentage = 10
        Me.UltraChart1.Legend.Location = LegendLocation.Right

        '' set the data
        Me.UltraChart1.Data.DataSource = StatesData()
        Me.UltraChart1.Data.DataBind()
    End Sub

#Region "Create StateDataView Data"

    Private Function StatesData() As StateDataInfo()
        Dim StatesDataFromDataSource() As StateDataInfo
        ReDim StatesDataFromDataSource(49)
        If SELECTION_NO <> 0 Then
            For I As Integer = 0 To US_STATES.Length - 1
                Debug.Print(US_STATES(I))
                'Dim rows() As DataRow = dst.Tables("TATSTATE").Select("STATE_NAME = '" & US_STATES(I) & "'")
                Dim rows() As DataRow = dst.Tables("TATSTATE").Select("MAP_INDEX = " & CStr(I))
                Dim SALES As Int32 = 0
                If rows.Length = 1 Then
                    SALES = Val(rows(0).Item("AMT") & "")
                End If
                StatesDataFromDataSource(I) = New StateDataInfo(US_STATES(I), SALES, "")
            Next
        End If
        'StatesDataFromDataSource(0) = New StateExpenseViewInfo("Alabama", 1915560.96, "")
        Return StatesDataFromDataSource
    End Function
#End Region

    Private Sub chkShowDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        Setup_Dashboard()
    End Sub

    Sub Setup_Dashboard()
        splDashboard.Panel2Collapsed = Not chkShowDetails.Checked
        Setup_tabMain()
    End Sub

    Private Sub grdTATDASHX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdTATDASHX.AfterRowActivate

        If optLevelCharts.Value = "2" Then
            CreateGraph_Totals()
            CreateGraph_Trend()
        End If

        If optLevelMap.Value <> "0" Then
            CreateMap()
        End If

        Setup_grdTATDASHY()
    End Sub

    Private Sub grdTATDASHX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdTATDASHX.InitializeLayout

    End Sub

    Private Sub grdTATDASH1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdTATDASH1.AfterRowActivate
        If ScreenMode Then
            Setup_grdTATDASH1()
        End If
    End Sub

    Sub Setup_grdTATDASH1()
        DASH_CODE = grdTATDASH1.ActiveRow.Cells("DASH_CODE").Text

        Dim dvwTATDASH2 As DataView = dst.Tables("TATDASH2").DefaultView
        dvwTATDASH2.RowFilter = "DASH_CODE = '" & DASH_CODE & "'"
        dvwTATDASH2.Sort = "DASH_VIEW"
        cbeTATDASH2.DataSource = dvwTATDASH2
        cbeTATDASH2.Value = 1

        Dim rowTATDASH1 As DataRow = dst.Tables("TATDASH1").Rows.Find(DASH_CODE)

        'optDataType.ValueList = Nothing

        Dim VL As New ValueList
        For iVAL As Integer = 1 To 6
            If rowTATDASH1.Item("DASH_VALUE" & CStr(iVAL)) & "" = "" Then
                Exit For
            Else
                VL.ValueListItems.Add(CStr(iVAL), rowTATDASH1.Item("DASH_VALUE" & CStr(iVAL)) & "")
            End If
        Next
        optDataType.ValueList = VL
        If VL.ValueListItems.Count = 0 Then
            optDataType.Text = ""
        Else
            optDataType.Value = "1"
            optDataType.Text = VL.ValueListItems(0).DisplayText
            Setup_DataType()
        End If
    End Sub

    Sub Setup_grdTATDASHX()

        DASH_VIEW = 0
        Dim DASH_CODE1_CAPTION As String = ""
        Dim DASH_CODE2_CAPTION As String = ""
        Dim DASH_DESC1_CAPTION As String = ""
        Dim DASH_DESC2_CAPTION As String = ""
        Dim DESCRIPTION As String = ""
        If cbeTATDASH2.Value & "" <> "" Then
            DASH_VIEW = Val(cbeTATDASH2.Value & "")
            rowTATDASH2 = dst.Tables("TATDASH2").Rows.Find(New Object() {DASH_CODE, DASH_VIEW})
            DASH_CODE1_CAPTION = rowTATDASH2.Item("DASH_CODE1_CAPTION") & ""
            DASH_CODE2_CAPTION = rowTATDASH2.Item("DASH_CODE2_CAPTION") & ""
            DASH_DESC1_CAPTION = rowTATDASH2.Item("DASH_DESC1_CAPTION") & ""
            DASH_DESC2_CAPTION = rowTATDASH2.Item("DASH_DESC2_CAPTION") & ""
            DESCRIPTION = cbeTATDASH2.Text
        End If

        If grdTATDASH1.ActiveRow.Cells("DASH_NO_TREND").Value & "" = "1" Then
            chkShowTrend.Checked = False
            chkShowTrend.Visible = False
        Else
            chkShowTrend.Visible = True
        End If

        optLevelCharts.ValueList.ValueListItems(0).DisplayText = DASH_CODE1_CAPTION
        optLevelCharts.ValueList.ValueListItems(1).DisplayText = DASH_CODE2_CAPTION & " within " & DASH_CODE1_CAPTION
        optLevelMap.ValueList.ValueListItems(1).DisplayText = "Selected " & DASH_CODE1_CAPTION
        optLevelMap.ValueList.ValueListItems(2).DisplayText = "Selected " & DASH_CODE2_CAPTION

        With grdTATDASHX.DisplayLayout.Bands(0)
            .Columns("CODE1").Header.Caption = DASH_CODE1_CAPTION
            .Columns("CODE2").Header.Caption = DASH_CODE2_CAPTION
            .Columns("DESC1").Header.Caption = DASH_DESC1_CAPTION
            .Columns("DESC2").Header.Caption = DASH_DESC2_CAPTION
        End With

        grdTATDASHX.Text = grdTATDASH1.ActiveRow.Cells("DASH_DESC").Text & " by " & DESCRIPTION
        Dim dvw As DataView = dst.Tables("TATDASHX").DefaultView
        dvw.RowFilter = "DASH_CODE = '" & DASH_CODE & "' and DASH_VIEW = " & CStr(DASH_VIEW)
        'grdTATDASHX.Rows.ExpandAll(True)
        grdTATDASHX.Rows.CollapseAll(True)

        CreateMap()
        CreateGraph_Totals()
        CreateGraph_Trend()
        Setup_grdTATDASHY()
    End Sub
    Private Sub grdTATDASH1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdTATDASH1.InitializeLayout

    End Sub

    Private Sub grdTATDASHX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdTATDASHX.InitializeRow

    End Sub

    Private Sub optMW_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMW.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_MW()

        CreateMap()
        CreateGraph_Totals()
        CreateGraph_Trend()
    End Sub

    Sub Setup_MW()
        If SELECTION_NO = 0 Then Exit Sub

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdTATDASHX, grdTATDASHY}
            With grd.DisplayLayout.Bands(0)
                For Each MW As String In New String() {"M", "W"}
                    For i As Integer = 5 To 0 Step -1
                        Dim COLUMN_NAME As String = "VAL0" & MW & Format(i, "0")
                        If chkShowTrend.Checked Then
                            .Columns(COLUMN_NAME).Hidden = Not (optMW.Value = MW)
                            If MW = "M" Then
                                .Columns(COLUMN_NAME).Header.Caption = COL_M(i)
                                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Green
                            Else
                                .Columns(COLUMN_NAME).Header.Caption = COL_W(i)
                            End If
                        Else
                            If MW = "M" Then
                                If optDataType.ValueList.ValueListItems.Count >= 6 - i Then
                                    Dim CAPTION As String = optDataType.ValueList.ValueListItems(6 - i - 1).DisplayText
                                    .Columns(COLUMN_NAME).Header.Caption = CAPTION
                                    .Columns(COLUMN_NAME).Hidden = False
                                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.Orange
                                Else
                                    .Columns(COLUMN_NAME).Hidden = True
                                End If
                            Else
                                .Columns(COLUMN_NAME).Hidden = True
                            End If
                        End If
                    Next
                Next
            End With
        Next
    End Sub

    Private Sub cbeTATDASH2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeTATDASH2.ValueChanged
        If ScreenMode And DASH_CODE <> "" Then
            Setup_grdTATDASHX()
        End If
    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
    End Sub

    Private Sub optLevelCharts_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLevelCharts.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        CreateGraph_Totals()
        CreateGraph_Trend()
    End Sub

    Private Sub optLevelMap_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optLevelMap.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        CreateMap()
    End Sub

    Private Sub optDataType_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optDataType.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_DataType()
    End Sub

    Sub Setup_DataType()

        Set_DataField("TATDASHX")
        Set_DataField("TATDASHY")
        Set_DataField("TATDASHS")

        grdTATDASHX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

        CreateMap()
        CreateGraph_Totals()
        CreateGraph_Trend()

    End Sub

    Private Sub chkShowTrend_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowTrend.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_DataType()
        If Not chkShowTrend.Checked Then
            optMW.Value = "M"
        End If
        optMW.Visible = chkShowTrend.Checked
        Setup_MW()

        chtTrend.Visible = chkShowTrend.Checked
    End Sub

    Sub Setup_grdTATDASHY()
        Dim SQL_DETAILS As String = rowTATDASH2.Item("SQL_DETAILS") & ""
        If SQL_DETAILS <> "" Then

            Dim CODE1 As String = ""
            Dim CODE2 As String = ""

            Dim CODE2_ESTABLISHED As Boolean = False

            If grdTATDASHX.ActiveRow IsNot Nothing Then
                If grdTATDASHX.ActiveRow.IsDataRow Then
                    CODE1 = grdTATDASHX.ActiveRow.Cells("CODE1").Text
                    CODE2 = grdTATDASHX.ActiveRow.Cells("CODE2").Text
                    CODE2_ESTABLISHED = True
                ElseIf grdTATDASHX.ActiveRow.IsGroupByRow Then
                    Dim grow As UltraWinGrid.UltraGridRow = grdTATDASHX.ActiveRow
                    Do While grow IsNot Nothing
                        Dim gby As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow, UltraWinGrid.UltraGridGroupByRow)
                        If gby.Column.Key = "CODE1" Then CODE1 = gby.Value & ""
                        If gby.Column.Key = "CODE2" Then
                            CODE2 = gby.Value & ""
                            CODE2_ESTABLISHED = True
                        End If
                        If gby.ParentRow IsNot Nothing Then
                            grow = gby.ParentRow
                        Else
                            grow = Nothing
                        End If
                    Loop
                End If
            End If

            Dim DASH_CODE1_COLUMN_NAME As String = rowTATDASH2.Item("DASH_CODE1_COLUMN_NAME")
            Dim DASH_CODE2_COLUMN_NAME As String = rowTATDASH2.Item("DASH_CODE2_COLUMN_NAME")
            Dim DASH_CODE1_CAPTION As String = rowTATDASH2.Item("DASH_CODE1_CAPTION")
            Dim DASH_CODE2_CAPTION As String = rowTATDASH2.Item("DASH_CODE2_CAPTION")

            'If CODE1 = "" And CODE2 = "" Then
            '    grdTATDASHY.Visible = False
            '    Exit Sub
            'End If


            SQL_DETAILS = Replace(SQL_DETAILS, " where ", " where " & DASH_CODE1_COLUMN_NAME & IIf(CODE1 = "", " IS NULL ", " = '" & CODE1 & "'") & " and ")
            If CODE2_ESTABLISHED Then
                SQL_DETAILS = Replace(SQL_DETAILS, " where ", " where " & DASH_CODE2_COLUMN_NAME & IIf(CODE2 = "", " IS NULL ", " = '" & CODE2 & "'") & " and ")
            End If

            Fill_Records("TATDASHY", "", , SQL_DETAILS)
            Sort_grdColumns(grdTATDASHY, "CODE_VALUE")

            grdTATDASHY.Text = "Details for " & DASH_CODE1_CAPTION & ":" & CODE1 & IIf(CODE2_ESTABLISHED, ", " & DASH_CODE2_CAPTION & ":" & CODE2, "")
            grdTATDASHY.Visible = True
        Else
            grdTATDASHY.Visible = False
        End If
    End Sub
End Class


Public Class MyCustomTooltip
    Implements IRenderLabel

    Public Sub New()

    End Sub 'New

    Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
        'Return Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        'Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        Return Context("SERIES_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))

    End Function 'ToString 
End Class 'MyCustomTooltip
#Region "USMap"

Public Class MapLayer
    Implements ILayer
    Private shapeFile As shapeFile = Nothing


    Public Sub New(ByVal filename As String)
        'Load the shape file which contains each states shape.
        shapeFile = shapeFile.Load(filename)
    End Sub 'New

    'Public Shared STATES As String() = {"Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming"}
    Public STATES As String() = {"Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming"}
    Public COLORS As New Dictionary(Of String, Color)

    '/ <summary>
    '/ Method which loops through each state, locates the appropriate polygon
    '/ shape and then determines how it sohuld be added to the SceneGraph
    '/ </summary>
    '/ <param name="scene"></param>
    Public Sub FillSceneGraph(ByVal scene As SceneGraph) Implements Infragistics.UltraChart.Core.Layers.ILayer.FillSceneGraph
        'Create a background Box for the layer and color it white
        '            Box bkgnd = new Box(this._OuterBound);
        '            bkgnd.PE.Fill = Color.White;
        '            bkgnd.PE.FillOpacity = 255;
        '            scene.Add(bkgnd);
        COLORS.Clear()
        Dim i As Integer
        For i = 0 To STATES.Length - 1
            Dim state As String = STATES(i)
            Dim color As Color = Drawing.Color.Empty
            If state.StartsWith("Michigan") Then
                'Since Michigan requires two polygons (for the LP and UP) we have to treat it different
                color = AddPolygons(i, New PolygonShape() {shapeFile("Michigan0"), shapeFile("Michigan1")}, scene)
            ElseIf state.StartsWith("Hawaii") Then
                'Since Hawaii is several polygons, we have to treat it different
                color = AddPolygons(i, New PolygonShape() {shapeFile("Hawaii0"), shapeFile("Hawaii1"), shapeFile("Hawaii2"), shapeFile("Hawaii3"), shapeFile("Hawaii4")}, scene)
            Else
                color = AddPolygons(i, New PolygonShape() {shapeFile(state)}, scene)
            End If
            COLORS.Add(STATES(i), color)
        Next i
    End Sub 'FillSceneGraph


    '/ <summary>
    '/ Method which creates each new polygon and sets its properties 
    '/ and actually adds the polygon to the SceneGraph
    '/ </summary>
    '/ <param name="index"></param>
    '/ <param name="polygonshapes"></param>
    '/ <param name="scene"></param>
    Private Function AddPolygons(ByVal index As Integer, ByVal polygonshapes() As PolygonShape, ByVal scene As SceneGraph) As Color
        Dim i As Integer
        Dim shape_color As Color = Drawing.Color.Empty
        Dim objectValue As Double = CDbl(Me.ChartData.GetObjectValue(index, 0))
        'Console.WriteLine(objectValue.ToString())
        shape_color = Me._ChartColorModel.getFillColor(index, 0, objectValue)

        For i = 0 To polygonshapes.Length - 1
            Dim polygon As New Polygon(Infragistics.UltraChart.Core.Util.Transform.viewingTransform(shapeFile.Bounds, Me.OuterBound, polygonshapes(i).Points.ToArray(), True))

            polygon.PE.Fill = shape_color ' Me._ChartColorModel.getFillColor(index, 0, objectValue)
            polygon.PE.Stroke = Me._ChartColorModel.getOutlineColor(index, 0, objectValue)
            polygon.Caps = PCaps.HitTest Or PCaps.Tooltip Or PCaps.Skin

            polygon.Row = index
            polygon.Column = 0
            polygon.Value = polygonshapes(i).Name
            polygon.Layer = Me

            scene.Add(polygon)
        Next i
        Return shape_color
    End Function 'AddPolygons

#Region "ILayer Members"

    Private innerBounds As Rectangle

    Public Function GetInnerBounds() As Rectangle Implements Infragistics.UltraChart.Core.Layers.ILayer.GetInnerBounds
        Return Me.innerBounds
    End Function 'GetInnerBounds


    Public Function GetDataInvalidMessage() As String Implements Infragistics.UltraChart.Core.Layers.ILayer.GetDataInvalidMessage
        Return "United States"
    End Function 'GetDataInvalidMessage

    Private _Grid As New Hashtable()

    Public Property Grid() As Hashtable Implements Infragistics.UltraChart.Core.Layers.ILayer.Grid
        Get
            Return _Grid
        End Get
        Set(ByVal Value As Hashtable)
            _Grid = Value
        End Set
    End Property

    Private _LayerID As String

    Public Property LayerID() As String Implements Infragistics.UltraChart.Core.Layers.ILayer.LayerID
        Get
            Return _LayerID
        End Get
        Set(ByVal Value As String)
            _LayerID = Value
        End Set
    End Property

    Private _ChartCore As ChartCore

    Public Property ChartCore() As ChartCore Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartCore
        Get
            Return _ChartCore
        End Get
        Set(ByVal Value As ChartCore)
            _ChartCore = Value
        End Set
    End Property

    Private _ChartData As IChartData

    Public Property ChartData() As IChartData Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartData
        Get
            Return _ChartData
        End Get
        Set(ByVal Value As IChartData)
            _ChartData = Value
        End Set
    End Property

    Private _ChartColorModel As IColorModel

    Public Property ChartColorModel() As IColorModel Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartColorModel
        Get
            Return _ChartColorModel
        End Get
        Set(ByVal Value As IColorModel)
            _ChartColorModel = Value
        End Set
    End Property

    Private _Visible As Boolean

    Public Property Visible() As Boolean Implements Infragistics.UltraChart.Core.Layers.ILayer.Visible
        Get
            Return _Visible
        End Get
        Set(ByVal Value As Boolean)
            _Visible = Value
        End Set
    End Property

    Private _ChartComponent As IChartComponent

    Public Property ChartComponent() As IChartComponent Implements Infragistics.UltraChart.Core.Layers.ILayer.ChartComponent
        Get
            Return _ChartComponent
        End Get
        Set(ByVal Value As IChartComponent)
            _ChartComponent = Value
        End Set
    End Property

    Private _OuterBound As New Rectangle(0, 0, 0, 0)

    Public Property OuterBound() As Rectangle Implements Infragistics.UltraChart.Core.Layers.ILayer.OuterBound
        Get
            Return _OuterBound
        End Get
        Set(ByVal Value As Rectangle)
            _OuterBound = Value
            CalculateInnerBounds()
        End Set
    End Property


    Protected Sub CalculateInnerBounds()
        Me.innerBounds = New Rectangle(Me._OuterBound.X, Me._OuterBound.Y, Me._OuterBound.Width, Me._OuterBound.Height)
    End Sub 'CalculateInnerBounds

#End Region
End Class 'MapLayer

Public Class USMapLabelRenderer
    Implements IRenderLabel


    Public Sub New(ByVal info As DataTable)
        Me._InformationPerState = info
    End Sub 'New ''New
    Private _InformationPerState As DataTable

#Region "IRenderLabel Members"

    '/ <summary>
    '/ Locate the proper data value for the current state, 
    '/ construct and return the proper tooltip string
    '/ </summary>
    '/ <param name="Context"></param>
    '/ <returns></returns>
    Overloads Function ToString(ByVal Context As Hashtable) As String Implements Infragistics.UltraChart.Resources.IRenderLabel.ToString
        Dim row As Integer
        If Not (Context("DATA_ROW") Is Nothing) Then
            row = CInt(Context("DATA_ROW"))
        Else
            row = CInt(Context("ITEM_NUMBER"))
        End If

        Dim tip As String = ""

        Dim rowState() As DataRow = _InformationPerState.Select("MAP_INDEX = " & CStr(row))
        If rowState.Length <> 1 Then
            tip = ""
        Else
            Dim SALES As Decimal = Val(rowState(0).Item("AMT") & "")
            If SALES = 0 Then
                tip = rowState(0).Item("STATE_NAME") & ""
            Else
                tip = rowState(0).Item("STATE_NAME") & ": " & Format(SALES, "###,##0")
            End If
        End If

        'Try
        '    If Val(_InformationPerState.Rows(row)(2) & "") <> 0 Then
        '        tip = _InformationPerState.Rows(row)(1) & ": " & System.Convert.ToDouble(_InformationPerState.Rows(row)(2)).ToString("#,##0")
        '    Else
        '        tip = _InformationPerState.Rows(row)(1)
        '    End If
        'Catch ex As Exception
        '    tip = ""
        'End Try
        Return tip
    End Function 'IRenderLabel.ToString
#End Region
End Class 'USMapLabelRenderer ''USMapLabelRenderer

Public Class ShapeFile
    Private _Shapes As New PolygonShapeCollection()


    Public ReadOnly Property Shapes() As PolygonShapeCollection
        Get
            Return _Shapes
        End Get
    End Property


    '/ <summary>
    '/ Loads the shapes from an external file
    '/ </summary>
    '/ <param name="filename"></param>
    '/ <returns></returns>
    Public Overloads Shared Function Load(ByVal filename As String) As ShapeFile
        Dim serializer As New XmlSerializer(GetType(ShapeFile))
        Dim result As ShapeFile = Nothing
        Dim reader As New StreamReader(filename)
        result = Load(reader)
        reader.Close()
        Return result
    End Function 'Load
    ''Load
    '/ <summary>
    '/ Loads the shapes from a TextReader
    '/ </summary>
    '/ <param name="reader"></param>
    '/ <returns></returns>
    Public Overloads Shared Function Load(ByVal reader As TextReader) As ShapeFile
        Dim serializer As New XmlSerializer(GetType(ShapeFile))
        Dim result As ShapeFile = Nothing
        result = CType(serializer.Deserialize(reader), ShapeFile)
        Return result
    End Function 'Load
    ''Load
    '/ <summary>
    '/ Save the existing shapes to an XML file
    '/ </summary>
    '/ <param name="filename"></param>
    Public Sub Save(ByVal filename As String)
        Dim writer As New StreamWriter(filename)
        Dim serializer As New XmlSerializer(GetType(ShapeFile))
        serializer.Serialize(writer, Me)
        writer.Close()
    End Sub 'Save ''Save
    Private BoundsUptoDate As Boolean = False
    Private _Bounds As Rectangle


    Public ReadOnly Property Bounds() As Rectangle
        Get
            If Not Me.BoundsUptoDate Then
                Dim minX As Integer = Int32.MaxValue
                Dim minY As Integer = Int32.MaxValue
                Dim maxX As Integer = Int32.MinValue
                Dim maxY As Integer = Int32.MinValue

                Dim ps As PolygonShape
                For Each ps In Me.Shapes
                    If ps.Bounds.X < minX Then
                        minX = ps.Bounds.X
                    End If
                    If ps.Bounds.Right > maxX Then
                        maxX = ps.Bounds.Right
                    End If
                    If ps.Bounds.Y < minY Then
                        minY = ps.Bounds.Y
                    End If
                    If ps.Bounds.Bottom > maxY Then
                        maxY = ps.Bounds.Bottom
                    End If
                Next ps

                Me._Bounds = New Rectangle(minX, minY, maxX - minX, maxY - minY)
                BoundsUptoDate = True
            End If
            Return Me._Bounds
        End Get
    End Property


    Default Public Property Item(ByVal id As String) As PolygonShape
        Get
            Return Me._Shapes(id)
        End Get
        Set(ByVal Value As PolygonShape)
            Me._Shapes(id) = Value
        End Set
    End Property
End Class 'ShapeFile ''ShapeFile

Public Class PointCollection
    Inherits CollectionBase

    Public Overridable Function Add(ByVal point As Point) As Integer
        Return Me.List.Add(point)
    End Function 'Add


    Default Public Overridable Property Item(ByVal index As Integer) As Point
        Get
            Return CType(Me.List(index), Point)
        End Get
        Set(ByVal Value As Point)
            Me(index) = Value
        End Set
    End Property


    Public Overridable Function ToArray() As Point()
        Dim points(Me.Count - 1) As Point
        Dim current As Integer
        For current = 0 To (Me.Count) - 1
            points(current) = Me(current)
        Next current
        Return points
    End Function 'ToArray
End Class 'PointCollection

Public Class PolygonShape
    Private _Name As String


    <XmlAttributeAttribute()> _
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal Value As String)
            _Name = Value
        End Set
    End Property

    Private _Points As New PointCollection()

    Public ReadOnly Property Points() As PointCollection
        Get
            Return _Points
        End Get
    End Property

    Private BoundsUptoDate As Boolean = False
    Private _Bounds As Rectangle

    Public ReadOnly Property Bounds() As Rectangle
        Get
            If Not Me.BoundsUptoDate Then
                Dim minX As Integer = Int32.MaxValue
                Dim minY As Integer = Int32.MaxValue
                Dim maxX As Integer = Int32.MinValue
                Dim maxY As Integer = Int32.MinValue


                Dim p As Point
                For Each p In Me._Points
                    If p.X < minX Then
                        minX = p.X
                    End If
                    If p.X > maxX Then
                        maxX = p.X
                    End If
                    If p.Y < minY Then
                        minY = p.Y
                    End If
                    If p.Y > maxY Then
                        maxY = p.Y
                    End If
                Next p
                Me._Bounds = New Rectangle(minX, minY, maxX - minX, maxY - minY)
                BoundsUptoDate = True
            End If
            Return Me._Bounds
        End Get
    End Property
End Class 'PolygonShape ''PolygonShape

Public Class PolygonShapeCollection
    Inherits CollectionBase


    Default Public Property Item(ByVal id As String) As PolygonShape
        Get
            Return SearchForId(id)
        End Get
        Set(ByVal Value As PolygonShape)
            Dim e As PolygonShape = SearchForId(id)
            If e Is Nothing Then
                Me.Add(Value)
            Else
                Me(Me.IndexOf(e)) = Value
            End If
        End Set
    End Property


    Private Function SearchForId(ByVal id As String) As PolygonShape
        Dim result As PolygonShape = Nothing

        Dim ef As PolygonShape
        For Each ef In Me
            If ef.Name.Equals(id) Then
                Return ef
            End If
        Next ef

        Return result
    End Function 'SearchForId 
    ''SearchForId


    Default Public Property Item(ByVal index As Integer) As PolygonShape
        Get
            Return CType(Me(index), PolygonShape)
        End Get
        Set(ByVal Value As PolygonShape)
            Me(index) = Value
        End Set
    End Property


    Public Function Add(ByVal value As PolygonShape) As Integer
        Return List.Add(value)
    End Function 'Add
    ''Add
    Public Function IndexOf(ByVal value As PolygonShape) As Integer
        Return Me.IndexOf(value)
    End Function 'IndexOf
    ''IndexOf
    Public Sub Insert(ByVal index As Integer, ByVal value As PolygonShape)
        Me.Insert(index, value)
    End Sub 'Insert
    ''Insert
    Public Sub Remove(ByVal value As PolygonShape)
        Me.Remove(value)
    End Sub 'Remove
    ''Remove
    Public Function Contains(ByVal value As PolygonShape) As Boolean
        '' If value is not of type PolygonShape, this will return false.
        Return Me.Contains(value)
    End Function 'Contains ''Contains
End Class 'PolygonShapeCollection

Public Class StateDataInfo

#Region "Private Member Variables"
    Private _State As String = ""
    Private _Amount As Double = 0.0
#End Region

#Region "Constructors"

    Public Sub New(ByVal state As String, ByVal amount As Double, ByVal category As String)
        _State = state
        _Amount = amount
    End Sub 'New

#End Region

#Region "Public Properties"

    Public Property State() As String
        Get
            Return _State
        End Get

        Set(ByVal Value As String)
            _State = Value
        End Set
    End Property


    Public Property Amount() As Double
        Get
            Return _Amount
        End Get

        Set(ByVal Value As Double)
            _Amount = Value
        End Set
    End Property

#End Region
End Class 'StateData

#End Region