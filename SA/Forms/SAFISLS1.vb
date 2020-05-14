Imports ABSolution
Imports Infragistics.Win
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO

Imports System.Collections
Imports System.Xml.Serialization

Imports Infragistics.UltraChart.Shared.Styles
Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Resources.Appearance
Imports Infragistics.UltraChart.Core
Imports Infragistics.UltraChart.Core.ColorModel
Imports Infragistics.UltraChart.Data
Imports Infragistics.UltraChart.Core.Layers
Imports Infragistics.UltraChart.Core.Primitives

Public Class SAFISLS1
    Dim SATISLS1 As String
    Dim SATISLS2 As String
    Dim sqlSATISLS1 As String
    Dim sqlSATISLS2 As String
    Dim sqlSOTINVHX As String
    Dim sqlSATISLS1_STORES As String
    Dim STORES As List(Of String)
    Dim STORES_XX As String
    Dim PERIODS_XX As String

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW0 As String
    Dim RYW1 As String
    Dim Periods As Integer
    Dim ITEM_CODE As String
    Dim US_STATES() As String
    Dim USmap As MapLayer

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        Set_cmbYW("RYW0", ASCMAIN1.CYW, -3 * 52, 0, -13)
        'Set_cmbYW_Child("RYW1", ASCMAIN1.CYW, "RYW0", 0)
        Set_cmbYW("RYW1", ASCMAIN1.CYW, -3 * 52, 0, 0)

        With dst
            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE, ITEM_DESC DESC_VALUE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATISLS1", "**", 0, False)
            With .Tables("SATISLS1")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To 120
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE, '0' YEAR, 'XX' DATA_TYPE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATISLS1_DTL", "**", 0, False)
            With .Tables("SATISLS1_DTL")
                For P As Integer = 0 To 120
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            .Relations.Add("SATISLS1_SATISLS1_DTL" _
                           , New DataColumn() {.Tables("SATISLS1").Columns("CODE_VALUE")} _
                           , New DataColumn() {.Tables("SATISLS1_DTL").Columns("CODE_VALUE")})

            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE_PARENT, ITEM_CODE CODE_VALUE, ITEM_DESC DESC_VALUE from ICTITEM1"
            Create_TDA(.Tables.Add, "SATISLS2", "**", 0, False)
            With .Tables("SATISLS2")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To 120
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            .Relations.Add("SATISLS1_SATISLS2" _
           , New DataColumn() {.Tables("SATISLS1").Columns("CODE_VALUE")} _
           , New DataColumn() {.Tables("SATISLS2").Columns("CODE_VALUE_PARENT")})



            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" _
            & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO" _
            & ", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH2.OPS_YYYYWW" _
            & ", SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO" _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') CUST_STORE_LOCATION" _
            & ", SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE" _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" _
            & " from SOTINVH2,ICTITEM1,ARTCUST2,SOTINVH1 " _
            & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE " _
            & " and ARTCUST2.CUST_CODE (+) = SOTINVH2.CUST_CODE " _
            & " and ARTCUST2.CUST_STORE_NO (+) = SOTINVH2.CUST_STORE_NO " _
            & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE " _
            & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            sqlSOTINVHX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False)

            Create_TDA(.Tables.Add, "TATSTATE", "*", 0, False)

            With .Tables.Add("SATISLSS")
                .Columns.Add("STATE_CODE")
                .Columns.Add("STATE_NAME")
                .Columns.Add("SALES", GetType(System.Int32))
            End With

        End With

        Fill_Records("TATSTATE")

        grdSATISLS1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
        grdSATISLS1.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdSATISLS1.DisplayLayout.MaxBandDepth = 1
        grdSATISLS1.DataSource = dst.Tables("SATISLS1")

        grdSATISLS2.DataSource = dst.Tables("SATISLS2")

        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")

        Dim dvw As DataView = dst.Tables("SATISLSS").DefaultView
        dvw.RowFilter = "SALES <> 0"
        grdSATISLSS.DataSource = dvw

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_SHIP")
        Create_Summary(grdSOTINVHX, "ORDR_AMT_SHIP")

        Create_Summary(grdSATISLSS, "STATE_CODE", "Count")
        Create_Summary(grdSATISLSS, "SALES")

        Create_Summary(grdSATISLS1, "CODE_VALUE", "Count")
        For P As Integer = 0 To 120
            Create_Summary(grdSATISLS1, "P" & Format(P, "00"))
        Next
        Create_Summary(grdSATISLS1, "PXX")

        Create_Summary(grdSATISLS2, "CODE_VALUE", "Count")
        For P As Integer = 0 To 120
            Create_Summary(grdSATISLS2, "P" & Format(P, "00"))
        Next
        Create_Summary(grdSATISLS2, "PXX")

        With grdSATISLS1.DisplayLayout.Bands("SATISLS1")
            .Columns("CODE_VALUE").Header.Fixed = True
            .Columns("DESC_VALUE").Header.Fixed = True
            .Columns("SUB_CODE_VALUE1").Header.Fixed = True
            .Columns("SUB_CODE_VALUE2").Header.Fixed = True
            .Columns("SUB_CODE_VALUE3").Header.Fixed = True
            .Columns("SUB_CODE_VALUE4").Header.Fixed = True
            .Columns("SUB_CODE_VALUE5").Header.Fixed = True
            .Columns("RTL_PRICE").Header.Fixed = True
            .Columns("WSL_PRICE").Header.Fixed = True
            .Columns("P00").Header.Fixed = True
            .Columns("PXX").Header.Fixed = True
        End With

        With grdSATISLS2.DisplayLayout.Bands("SATISLS2")
            .Columns("CODE_VALUE").Header.Fixed = True
            .Columns("DESC_VALUE").Header.Fixed = True
            .Columns("SUB_CODE_VALUE1").Header.Fixed = True
            .Columns("SUB_CODE_VALUE2").Header.Fixed = True
            .Columns("SUB_CODE_VALUE3").Header.Fixed = True
            .Columns("SUB_CODE_VALUE4").Header.Fixed = True
            .Columns("SUB_CODE_VALUE5").Header.Fixed = True
            .Columns("RTL_PRICE").Header.Fixed = True
            .Columns("WSL_PRICE").Header.Fixed = True
            .Columns("P00").Header.Fixed = True
            .Columns("PXX").Header.Fixed = True
        End With

        With chtSATISLS1
            .Axis.X.ScrollScale.Visible = True
            .Axis.Y.ScrollScale.Visible = True

            .Axis.X.ScrollScale.Scale = 1 ' 0.25
            .Axis.Y.ScrollScale.Scale = 1 ' 0.25
            Me.trkbrXAxis.Value = .Axis.X.ScrollScale.Scale * 100
            Me.trkbrYAxis.Value = .Axis.Y.ScrollScale.Scale * 100
            .EnableCrossHair = True

            '.ColorModel.ModelStyle = ColorModels.CustomLinear '  CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        End With


        grpWEEK_RANGE.Top = grpPERIOD_RANGE.Top
        grpWEEK_RANGE.Left = grpPERIOD_RANGE.Left

        'CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        'Dim modelStyle As String() = System.Enum.GetNames(GetType(ColorModels))
        'Dim s As String
        'For Each s In modelStyle
        '    Me.comboBox1.Items.Add(s)
        'Next s

        'Me.comboBox1.SelectedItem = Me.comboBox1.Items(Me.comboBox1.FindString(System.Enum.GetName(GetType(ColorModels), chtICTINVAT.ColorModel.ModelStyle), 0))

        'chtICTINVAT.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), Me.comboBox1.SelectedItem.ToString()), ColorModels)

        'Dim colors As Array = System.Enum.GetValues(GetType(Infragistics.UltraChart.Shared.Styles.ColorModels))
        'For i As Integer = 0 To colors.Length
        '    colors(i).ToString()
        'Next
        'cbeColor.DataSource = colors

        Dim modelStyle As String() = System.Enum.GetNames(GetType(ColorModels))
        cbeColor.DataSource = modelStyle
        cbeColor.SelectedItem = cbeColor.Items(cbeColor.FindString(System.Enum.GetName(GetType(ColorModels), chtSATISLS1.ColorModel.ModelStyle), 0))

        'cbeColorBest.DataSource = System.Enum.GetNames(GetType(System.Drawing.Color))
        'cbeColorBest.SelectedItem = cbeColorBest.Items(cbeColor.FindString("Yellow", 0))
        'cbeColorWorst.DataSource = System.Enum.GetNames(GetType(System.Drawing.Color))
        'cbeColorWorst.SelectedItem = cbeColorBest.Items(cbeColor.FindString("Red", 0))

        Setup_Map()

        ' until tested
       optXP.Visible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Call Validate_Code("ITEM_CODE")

                If EMsg = "" Then
                    If Absx1.optFor("RANGE").Value = "P" Then
                        If Absx1.cmbFor("RYP0").Value & "" = "" Then
                            EMsg &= vbCr & "You must Specify a Starting Period"
                        End If
                        If Absx1.cmbFor("RYP1").Value & "" = "" Then
                            EMsg &= vbCr & "You must Specify an Ending Period"
                        End If

                        If EMsg = "" Then
                            RYP0 = Absx1.cmbFor("RYP0").Value
                            RYP1 = Absx1.cmbFor("RYP1").Value
                            Periods = ASCMAIN1.Period_Diff(RYP0, RYP1) + 1
                        End If
                    Else
                        If Absx1.cmbFor("RYW0").Value & "" = "" Then
                            EMsg &= vbCr & "You must Specify a Starting Week"
                        End If
                        If Absx1.cmbFor("RYW1").Value & "" = "" Then
                            EMsg &= vbCr & "You must Specify an Ending Week"
                        End If

                        If EMsg = "" Then
                            RYW0 = Absx1.cmbFor("RYW0").Value
                            RYW1 = Absx1.cmbFor("RYW1").Value
                            Periods = ASCMAIN1.Week_Diff(RYW0, RYW1) + 1
                        End If
                    End If

                    If Periods < 1 Or Periods > 120 Then
                        EMsg &= vbCr & "Total number of Periods must be between 1 and 120"
                    End If

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

            Case "Load"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

            Case "Print Report"
                Call Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Data Options").Visible = tf
                .Groups("Options").Visible = Not tf

                ' UNTIL TESTED
                .Groups("Options").Visible = False ' ASCMAIN1.Running_in_VS

                .Groups("Charts").Visible = False
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATISLS1", "SATISLS1_DTL", "SATISLS2", "SOTINVHX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        'Absx1.txtFor("CUST_CODE").Text = ""
        tabDetails.SelectedTab = tabDetails.Tabs("Details")
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Customer Sales Data")
        Save_Header_Fields(UltraGroupBox1)
        ITEM_CODE = HFs("ITEM_CODE")
        Create_SATISLS1()
        optXP.Items(0).DisplayText = optRANGE.Items(optRANGE.CheckedIndex).DisplayText
        Load_Data()
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATISLS1, "SSSSSSSSS", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5")
        Load_Popup_Menu(grdSATISLS2, "SSSSSSSSS", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5")
        Load_Popup_Menu(grdSOTINVHX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSATISLSS, "CC", "Best", "Worst")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show SUB_CODE_VALUE1") Then
            For I As Integer = 1 To 5
                Dim COLUMN_NAME As String = "SUB_CODE_VALUE" & CStr(I)
                tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
                tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            Next

            'COLUMN_NAME = "RTL_PRICE"
            'tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
            'tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
            'tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            'COLUMN_NAME = "WSL_PRICE"
            'tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
            'tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
            'tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else

            Select Case e.SourceControl.Name
                Case "grdSATISLS1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5" ', "Show RTL_PRICE", "Show WSL_PRICE"

                Dim COLUMN_NAME As String = Mid(e.Tool.Key, 6)
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                'With grdSATISLS1.DisplayLayout.Bands(1).Columns("DATA_TYPE")
                '    If tlb_sbt.Checked Then
                '        .ColSpan += 1
                '    Else
                '        .ColSpan -= 1
                '    End If
                'End With

            Case "Best"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorEnd

            Case "Worst"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorBegin
        End Select
    End Sub

    Public Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)

        Select Case e.Tool.Key
            Case "Best"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorEnd = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                'grdSATISLSS.DataBind()
                Application.DoEvents()
                grdSATISLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

            Case "Worst"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                'grdSATISLSS.DataBind()
                Application.DoEvents()
                grdSATISLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        End Select

    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            'Case "DTE0", "DTE1"
            '    If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select

    End Sub

#End Region

    Sub Create_SATISLS1()

        If SATISLS2 = "" Then
            SATISLS1 = ASCMAIN1.Temp_Table("Select ITEM_CODE from ICTITEM1 where ROWNUM < 1")
            SATISLS2 = ASCMAIN1.Temp_Table("Select ITEM_CODE from ICTITEM1 where ROWNUM < 1")
        End If
        ASCDATA1.ExecuteSQL("Drop Table " & SATISLS1)
        ASCDATA1.ExecuteSQL("Drop Table " & SATISLS2)

        Dim PX As String = IIf(optRANGE.Value = "P", "OPS_YYYYPP", "OPS_YYYYWW")
        Dim PX0 As String = IIf(optRANGE.Value = "P", RYP0, RYW0)
        Dim PX1 As String = IIf(optRANGE.Value = "P", RYP1, RYW1)

        sqlSATISLS1 = ""
        sqlSATISLS2 = ""
        Dim SQL() As String = New String() {"", ""}
        Dim P As Integer
        PERIODS_XX = ""

        'SQL = ""
        For P = 1 To Periods
            Dim PXP As String = IIf(optRANGE.Value = "P", ASCMAIN1.Period_Calc(RYP0, P - 1), ASCMAIN1.Week_Calc(RYW0, P - 1))
            SQL(0) &= ", Sum (Decode(" & PX & ",'" & PXP & "',RSTRETL1.QTY_SOLD,0)) P" & Format(P, "00") & vbCrLf
            SQL(1) &= ", Sum (Decode(" & PX & ",'" & ASCMAIN1.Period_Calc(PXP, -12) & "',RSTRETL1.QTY_SOLD,0)) P" & Format(P, "00") & vbCrLf
            sqlSATISLS1 &= ", P" & Format(P, "00")
            sqlSATISLS2 &= ", SUM (P" & Format(P, "00") & ") P" & Format(P, "00")
            PERIODS_XX &= "+P" & Format(P, "00")
        Next

        Dim YEAR_max As Int32 = 0
        If chkPriorYear.Checked Then
            YEAR_max = 1
        End If

        For YEAR As Int32 = 0 To YEAR_max

            PX = IIf(optRANGE.Value = "P", "OPS_YYYYPP", "OPS_YYYYWW")

            PX0 = IIf(optRANGE.Value = "P", RYP0, RYW0)
            PX1 = IIf(optRANGE.Value = "P", RYP1, RYW1)
            If YEAR = 1 Then
                If optRANGE.Value = "P" Then
                    PX0 = ASCMAIN1.Period_Calc(PX0, -12)
                    PX1 = ASCMAIN1.Period_Calc(PX1, -12)
                Else
                    PX0 = ASCMAIN1.Week_Calc(PX0, -52)
                    PX1 = ASCMAIN1.Week_Calc(PX1, -52)
                End If
            End If


            Dim sqla As String = "Select 'TU' DATA_TYPE, '" & CStr(YEAR) & "' YEAR" & vbCrLf _
            & ", RSTRETL1.ITEM_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & SQL(YEAR) & vbCrLf _
            & " from RSTRETL1,GLTPARM3  " & vbCrLf _
            & " where RSTRETL1.ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
            & " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf _
            & " and RSTRETL1." & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & " group by RSTRETL1.ITEM_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO"

            Dim SQL_ORIG As String = sqla

            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & SATISLS2 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & SATISLS2 _
                & " Add Primary Key (DATA_TYPE, YEAR, ITEM_CODE, CUST_CODE, CUST_STORE_NO)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)
            End If
            sqla = Replace(sqla, "'TU' DATA_TYPE", "'TD' DATA_TYPE")
            sqla = Replace(sqla, "QTY_SOLD", "AMT_SOLD")
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)

            sqla = Replace(SQL_ORIG, "'TU' DATA_TYPE", "'HU' DATA_TYPE")
            sqla = Replace(sqla, " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW", " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf & IIf(optRANGE.Value = "P", " and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK", ""))
            sqla = Replace(sqla, "QTY_SOLD", "QTY_EOW")
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)

            sqla = Replace(SQL_ORIG, "'TU' DATA_TYPE", "'HD' DATA_TYPE")
            sqla = Replace(sqla, " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW", " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf & IIf(optRANGE.Value = "P", " and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK", ""))
            sqla = Replace(sqla, "from RSTRETL1", "from RSTRETL1,ICTITEM1")
            sqla = Replace(sqla, "group by", " and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE group by")
            sqla = Replace(sqla, "QTY_SOLD", "QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE") ' SB USING ICTRETLA
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)

            PX = IIf(optRANGE.Value = "P", "ORDR_YYYYPP_UPDATED", "OPS_YYYYWW")
            sqla = ""
            For P = 1 To Periods
                Dim PXP As String = IIf(optRANGE.Value = "P", _
                                       ASCMAIN1.Period_Calc(RYP0, P - 1 - 12 * YEAR), _
                                       ASCMAIN1.Week_Calc(RYW0, P - 1 - 52 * YEAR))
                sqla &= ", Sum (Decode(" & PX & ",'" & PXP & "',SOTINVH2.ORDR_QTY_SHIP,0)) P" & Format(P, "00") & vbCrLf
            Next
            sqla = "Select DECODE(INV_TYPE,'I','S','C','R') || 'U' DATA_TYPE" & vbCrLf _
            & ", '" & CStr(YEAR) & "' YEAR" & vbCrLf _
            & ", ITEM_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
            & sqla & vbCrLf _
            & " from SOTINVH2 " & vbCrLf _
            & " where ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
            & " and " & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & " group by DECODE(INV_TYPE,'I','S','C','R') || 'U', ITEM_CODE, CUST_CODE, CUST_STORE_NO"

            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)
            sqla = Replace(sqla, "'U' DATA_TYPE", "'D' DATA_TYPE")
            sqla = Replace(sqla, " || 'U',", " || 'D',")
            sqla = Replace(sqla, "ORDR_QTY_SHIP", "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS2 & " " & sqla)

            sqla = "Select 'I' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", ITEM_CODE, CUST_CODE CODE_VALUE" _
            & sqlSATISLS2 & " from " & SATISLS2 _
            & " group by DATA_TYPE, ITEM_CODE, CUST_CODE"
            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & SATISLS1 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & SATISLS1 & " Add Primary Key (SI, DATA_TYPE, YEAR, ITEM_CODE, CODE_VALUE)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SATISLS1 & " " & sqla)
            End If

            sqla = "Select 'S' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", ITEM_CODE, CUST_CODE CODE_VALUE" & sqlSATISLS2 _
            & " from " & SATISLS2 & " group by DATA_TYPE, ITEM_CODE, CUST_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & SATISLS1 & " " & sqla)
        Next YEAR

        ASCDATA1.ExecuteSQL("Alter Table " & SATISLS1 & " Add PXX NUMBER (10,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SATISLS2 & " Add PXX NUMBER (10,0)")

        ASCMAIN1.sql = "Update " & SATISLS1 & " X SET PXX = (Select P" & Format(Periods, "00") & " from " & SATISLS1 & " " _
        & " where SI = X.SI and DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and ITEM_CODE = X.ITEM_CODE and CODE_VALUE = X.CODE_VALUE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & SATISLS2 & " X SET PXX = (Select P" & Format(Periods, "00") & " from " & SATISLS2 & " " _
        & " where DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and ITEM_CODE = X.ITEM_CODE and CUST_STORE_NO = X.CUST_STORE_NO AND CUST_CODE = X.CUST_CODE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        STORES = New List(Of String)
        STORES_XX = ""
        Dim SQLX As String = ""

        SQLX = "Select Distinct CUST_STORE_NO from " & SATISLS2 & " order by CUST_STORE_NO"
        For Each row As DataRow In ASCDATA1.GetDataTable(SQLX).Rows
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            STORES.Add(CUST_STORE_NO)
        Next

        SQLX = ""
        For S As Integer = 1 To STORES.Count
            SQLX &= ", Sum (Decode(CUST_STORE_NO,'" & STORES(S - 1) & "'," & Mid(PERIODS_XX, 2) & ",0)) P" & Format(S, "00") & vbCrLf
            STORES_XX &= "+P" & Format(S, "00")
        Next
        sqlSATISLS1_STORES = SQLX

    End Sub

    Sub Print_Report()
        Call Print_Report_Begin()

        Dim SUBT As String = ""
        Dim RecordSelectionFormula As String = ""
        Generate_Report("SARCSLS1", "", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        grpWEEK_RANGE.Visible = (optRANGE.Value = "W")
    End Sub

    Private Sub optSI_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Load_Data()
    End Sub

    Sub Setup_grd()

        Dim CAPTION As String = optType1.Text & " (" & optType2.Text & ") for " & ITEM_CODE
        CAPTION &= ", by Customer"
        grdSATISLS1.Text = CAPTION

        Dim g1 As UltraWinGrid.UltraGrid
        Dim g2 As UltraWinGrid.UltraGrid
        g1 = grdSATISLS1
        g2 = grdSATISLS2


        With g1.DisplayLayout.Bands(0)
            .Columns("CODE_VALUE").Header.Caption = "Customer"
            .Columns("DESC_VALUE").Header.Caption = "Name"
            .Columns("SUB_CODE_VALUE1").Header.Caption = "Rep"
            .Columns("SUB_CODE_VALUE2").Header.Caption = "State"
            .Columns("SUB_CODE_VALUE3").Header.Caption = "Class"
            .Columns("SUB_CODE_VALUE4").Header.Caption = "City"
            .Columns("SUB_CODE_VALUE5").Header.Caption = "Zip"
            .Columns("RTL_PRICE").Header.Caption = "Retail"
            .Columns("WSL_PRICE").Header.Caption = "WhSale"

            .Columns("CODE_VALUE").Width = 80
            .Columns("DESC_VALUE").Width = 140
            .Columns("SUB_CODE_VALUE1").Width = 50
            .Columns("SUB_CODE_VALUE2").Width = 50
            .Columns("SUB_CODE_VALUE3").Width = 50
            .Columns("SUB_CODE_VALUE4").Width = 50
            .Columns("SUB_CODE_VALUE5").Width = 50
            .Columns("RTL_PRICE").Width = 65
            .Columns("WSL_PRICE").Width = 65
        End With

        With g2.DisplayLayout.Bands(0)
            .Columns("CODE_VALUE").Header.Caption = "Store"
            .Columns("DESC_VALUE").Header.Caption = "Location"
            .Columns("SUB_CODE_VALUE1").Header.Caption = "Rep"
            .Columns("SUB_CODE_VALUE2").Header.Caption = "State"
            .Columns("SUB_CODE_VALUE3").Header.Caption = "Group"
            .Columns("SUB_CODE_VALUE4").Header.Caption = "City"
            .Columns("SUB_CODE_VALUE5").Header.Caption = "Zip"
            .Columns("RTL_PRICE").Header.Caption = "Retail"
            .Columns("WSL_PRICE").Header.Caption = "WhSale"

            .Columns("CODE_VALUE").Width = 80
            .Columns("DESC_VALUE").Width = 140
            .Columns("SUB_CODE_VALUE1").Width = 50
            .Columns("SUB_CODE_VALUE2").Width = 50
            .Columns("SUB_CODE_VALUE3").Width = 50
            .Columns("SUB_CODE_VALUE4").Width = 50
            .Columns("SUB_CODE_VALUE5").Width = 50
            .Columns("RTL_PRICE").Width = 65
            .Columns("WSL_PRICE").Width = 65
        End With

        For Each G As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
        {grdSATISLS1, grdSATISLS2}
            With G.DisplayLayout.Bands(0)
                If G.Name = "grdSATSLSC2" Then
                    .Columns("CODE_VALUE_PARENT").Hidden = True
                End If
                .Columns("CODE_VALUE").Hidden = False
                .Columns("DESC_VALUE").Hidden = False
                .Columns("SUB_CODE_VALUE1").Hidden = False
                .Columns("SUB_CODE_VALUE2").Hidden = False
                .Columns("SUB_CODE_VALUE3").Hidden = True
                .Columns("SUB_CODE_VALUE4").Hidden = True
                .Columns("SUB_CODE_VALUE5").Hidden = True
                .Columns("RTL_PRICE").Hidden = True
                .Columns("WSL_PRICE").Hidden = True
            End With
        Next

        If chkExtendedData.Checked Then
            g1.DisplayLayout.MaxBandDepth = 2
        Else
            g1.DisplayLayout.MaxBandDepth = 1
        End If

        For Each G As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {g1, g2}
            Dim BMAX As Int32 = 0
            If G Is g1 And chkExtendedData.Checked Then
                BMAX = 1
            End If
            For B As Int32 = 0 To BMAX

                With G.DisplayLayout.Bands(B)
                    If B = 1 Then
                        .Columns("DATA_TYPE").Hidden = False
                        .Columns("DATA_TYPE").ColSpan = 3
                        .Columns("DATA_TYPE").Header.Caption = "Data Type"
                        .Columns("YEAR").Hidden = False
                        .Columns("YEAR").Header.Caption = "Year"
                        .Columns("YEAR").Width = 100
                        .RowLayoutStyle = UltraWinGrid.RowLayoutStyle.None
                        .Override.AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized
                    End If
                    For P As Integer = 0 To 120
                        COLUMN_NAME = "P" & Format(P, "00")
                        If optXP.Value = "S" And G.Name = grdSATISLS1.Name Then
                            .Columns(COLUMN_NAME).Hidden = (P > STORES.Count)
                            If P <= STORES.Count Then
                                Dim LEGEND As String
                                If P = 0 Then
                                    LEGEND = "Total"
                                    .Columns(COLUMN_NAME).Width = 80
                                Else
                                    .Columns(COLUMN_NAME).Width = 70
                                    LEGEND = STORES(P - 1)
                                End If
                                .Columns(COLUMN_NAME).Header.Caption = LEGEND
                            End If
                        Else
                            .Columns(COLUMN_NAME).Hidden = (P > Periods)
                            If P <= Periods Then
                                Dim LEGEND As String
                                If P = 0 Then
                                    LEGEND = "Total"
                                    .Columns(COLUMN_NAME).Width = 80
                                Else
                                    .Columns(COLUMN_NAME).Width = 70
                                    If optRANGE.Value = "P" Then
                                        LEGEND = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP0, P - 1))
                                        LEGEND = Mid(LEGEND, 10, 6)
                                    Else
                                        LEGEND = ASCMAIN1.Get_Legend_Wk(ASCMAIN1.Week_Calc(RYW0, P - 1))
                                        LEGEND = Mid(LEGEND, 10, 7)
                                    End If
                                End If
                                .Columns(COLUMN_NAME).Header.Caption = LEGEND
                            End If
                        End If
                    Next

                    .Columns("PXX").Hidden = Not (optType1.Value = "T")
                    .Columns("PXX").Header.Caption = "O/H"

                End With
            Next
        Next

    End Sub

    Private Sub grdSATISLS1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATISLS1.AfterRowActivate
        Setup_grdSATISLS2()
    End Sub

    Private Sub grdSATISLS1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSATISLS1.InitializeLayout

    End Sub

    Sub Setup_grdSATISLS2()

        If grdSATISLS1.ActiveRow Is Nothing OrElse Not grdSATISLS1.ActiveRow.IsDataRow Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
            Exit Sub
        Else
            chkShowDetails.Enabled = True
        End If

        Dim DATA_TYPE As String = optType1.Value & optType2.Value
        Dim CODE_VALUE_PARENT As String = grdSATISLS1.ActiveRow.Cells("CODE_VALUE").Text

        Load_SATISLS2(DATA_TYPE, CODE_VALUE_PARENT, False)
        Sort_grdColumns(grdSATISLS2, "CODE_VALUE")

        Dim CAPTION As String = optType1.Text & " (" & optType2.Text & ") for " & ITEM_CODE
        CAPTION &= ", by Customer"
        grdSATISLS2.Text = CAPTION

        Dim sql As String = ""
        sql = sqlSOTINVHX & " and SOTINVH2.ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
            & " and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf
        sql &= " and SOTINVH2.CUST_CODE = '" & CODE_VALUE_PARENT & "'" & vbCrLf
        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("CUST_STORE_NO").Hidden = False
            .Columns("CUST_STORE_LOCATION").Hidden = False
            .Columns("ITEM_CODE").Hidden = True
            .Columns("ITEM_DESC").Hidden = True
        End With
        Fill_Records("SOTINVHX", "", True, sql)
        grdSOTINVHX.Text = "Sales Documents for " & ITEM_CODE & " - Customer " & CODE_VALUE_PARENT
        grdSOTINVHX.DisplayLayout.CaptionVisible = DefaultableBoolean.True

    End Sub

    Sub Load_SATISLS2(ByVal DATA_TYPE As String, ByVal CODE_VALUE_PARENT As String, ByVal all_parents As Boolean)
        Dim sql As String = ""
        sql = "Select SATISLS2.ITEM_CODE CODE_VALUE_PARENT " & vbCrLf _
          & ", SATISLS2.CUST_STORE_NO CODE_VALUE" & vbCrLf _
          & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') DESC_VALUE" & vbCrLf _
          & ", ARTCUST2.SELL_CODE SUB_CODE_VALUE1" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_STATE SUB_CODE_VALUE2" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_GROUP SUB_CODE_VALUE3" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_CITY SUB_CODE_VALUE4" & vbCrLf _
          & ", ARTCUST2.CUST_STORE_ZIP_CODE SUB_CODE_VALUE5" & vbCrLf _
          & sqlSATISLS1 & ",PXX  from ARTCUST2," & SATISLS2 & " SATISLS2 " & vbCrLf _
          & " where ARTCUST2.CUST_CODE (+) = SATISLS2.CUST_CODE " & vbCrLf _
          & " and ARTCUST2.CUST_STORE_NO (+) = SATISLS2.CUST_STORE_NO" & vbCrLf _
          & " and DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
          & IIf(all_parents, "", " and SATISLS2.CUST_CODE = '" & CODE_VALUE_PARENT & "'")
        'dst.Tables("SATISLS1").Rows.Clear()
        Fill_Records("SATISLS2", "", True, sql)
    End Sub

    Private Sub chkNoDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = Not chkShowDetails.Checked
    End Sub

    Sub Load_Data()

        dst.EnforceConstraints = False

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim DATA_TYPE As String = optType1.Value & optType2.Value

        optXP.Visible = True

        Dim sql As String = ""

        If optXP.Value = "S" Then

            Dim SQLX As String = sqlSATISLS1_STORES
            If Mid(DATA_TYPE, 1, 1) = "H" Then
                SQLX = Replace(sqlSATISLS1_STORES, Mid(PERIODS_XX, 2), "P" & Format(Periods, "00"))
            End If

            sql = "Select SATISLS2.ITEM_CODE CODE_VALUE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC DESC_VALUE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE SUB_CODE_VALUE2" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
            & ", ICTITEM1.STYLE_CODE SUB_CODE_VALUE4" & vbCrLf _
            & ", ICTITEM1.DEPT_CODE SUB_CODE_VALUE5" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE RTL_PRICE" & vbCrLf _
            & ", ICTITEM1.ITEM_PRICE WSL_PRICE" & vbCrLf _
            & SQLX & "  from ICTITEM1," & SATISLS2 & " SATISLS2 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = SATISLS2.ITEM_CODE " & vbCrLf _
            & " and SATISLS2.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
            & " and SATISLS2.YEAR = '0'" & vbCrLf _
            & " group by SATISLS2.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
            & ", ICTITEM1.STYLE_CODE" & vbCrLf _
            & ", ICTITEM1.DEPT_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
            & ", ICTITEM1.ITEM_PRICE" & vbCrLf
        Else
            sql = "Select SATISLS1.CODE_VALUE CODE_VALUE" & vbCrLf _
            & ", ARTCUST1.CUST_NAME DESC_VALUE" & vbCrLf _
            & ", ARTCUST1.SREP_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ARTCUST1.CUST_STATE SUB_CODE_VALUE2" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
            & ", ARTCUST1.CUST_CITY SUB_CODE_VALUE4" & vbCrLf _
            & ", ARTCUST1.CUST_ZIP_CODE SUB_CODE_VALUE5" & vbCrLf _
            & ", 0 RTL_PRICE" & vbCrLf _
            & ", 0 WSL_PRICE" & vbCrLf _
            & sqlSATISLS1 & ",PXX from ARTCUST1," & SATISLS1 & " SATISLS1 " & vbCrLf _
            & " where ARTCUST1.CUST_CODE (+) = SATISLS1.CODE_VALUE " & vbCrLf _
            & " and SATISLS1.SI = '" & "I" & "'" _
            & " and SATISLS1.DATA_TYPE = '" & DATA_TYPE & "'" _
            & " and SATISLS1.YEAR = '0'" & vbCrLf
        End If

        dst.Tables("SATISLS1").Rows.Clear()
        dst.Tables("SATISLS1_DTL").Rows.Clear()
        dst.Tables("SATISLS2").Rows.Clear()

        If optXP.Value = "S" Then
            dst.Tables("SATISLS1").Columns("P00").Expression = Mid(STORES_XX, 2)
            dst.Tables("SATISLS1_DTL").Columns("P00").Expression = Mid(STORES_XX, 2)
        Else
            If optType1.Value = "H" Then
                dst.Tables("SATISLS1").Columns("P00").Expression = Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
            Else
                dst.Tables("SATISLS1").Columns("P00").Expression = Mid(PERIODS_XX, 2)
            End If
            dst.Tables("SATISLS1_DTL").Columns("P00").Expression = "IIF(DATA_TYPE='HU' OR DATA_TYPE='HD'," & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3) & "," & Mid(PERIODS_XX, 2) & ")"
        End If

        If optType1.Value = "H" Then
            dst.Tables("SATISLS2").Columns("P00").Expression = Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
        Else
            dst.Tables("SATISLS2").Columns("P00").Expression = Mid(PERIODS_XX, 2)
        End If

        Fill_Records("SATISLS1", "", True, sql)
        Sort_grdColumns(grdSATISLS1, "CODE_VALUE")

        If chkExtendedData.Checked Then
            sql = "Select SATISLS1.CODE_VALUE CODE_VALUE, SATISLS1.YEAR, SATISLS1.DATA_TYPE" _
            & sqlSATISLS1 _
            & " from " & SATISLS1 & " SATISLS1 " _
            & " where SATISLS1.SI = '" & "I" & "'" _
            & " and SATISLS1.DATA_TYPE LIKE '%" & Mid(DATA_TYPE, 2, 1) & "'" _
            & " and (YEAR = '1' OR SATISLS1.DATA_TYPE <> '" & DATA_TYPE & "')"

            Fill_Records("SATISLS1_DTL", "", True, sql)
            Sort_grdColumns(grdSATISLS1, "DATA_TYPE", , 1)
        End If

        If grdSATISLS1.Rows.Count = 0 Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
        Else
            chkShowDetails.Enabled = True
        End If

        'dst.EnforceConstraints = True

        Setup_grd()


        tabDetails.Tabs("Sales Documents").Visible = (optType1.Value <> "T")

        If tabDetails.SelectedTab.Key = "Map" Then
            tabDetails.SelectedTab = tabDetails.Tabs("Details")
        End If
        tabDetails.Tabs("Map").Visible = False

        CreateGraph_SATISLS1()
        CreateGraph_SATISLS1_X()
        chtSATISLS1.Visible = True
        chtSATISLS1_X.Visible = True

       If grdSATISLS1.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE1") Then
            grdSATISLS1.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE1")
        End If
        If grdSATISLS2.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE2") Then
            grdSATISLS2.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE2")
        End If

        ASCMAIN1.Add_Value_List(grdSATISLS1, "SUB_CODE_VALUE2", , , , "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
        ASCMAIN1.Add_Value_List(grdSATISLS2, "SUB_CODE_VALUE1", , , , "Select SELL_CODE, SELL_NAME from SOTSELL1")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub optType1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optType1.ValueChanged

        If SELECTION_NO = 0 Then Exit Sub

        Load_Data()

    End Sub

    Private Sub optType2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optType2.ValueChanged

        If SELECTION_NO = 0 Then Exit Sub
        Load_Data()
    End Sub

    Sub CreateGraph_SATISLS1()

        Dim chtIsVisible As Boolean = chtSATISLS1.Visible
        chtSATISLS1.Visible = False

        chtSATISLS1.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String

        chtSATISLS1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATISLS1.LabelHash = labelHash

        chtSATISLS1.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATISLS1.Tooltips.FormatString = "<HIGHLOW>"

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Decimal))
        End With

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SATISLS1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SATISLS1").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item("CODE_VALUE"), row.Item("P00")})
        Next
        'chtSATISLS1.Data.SetRowLabels(RL)
        'chtSATISLS1.Data.SetColumnLabels(CL)

        'chtSATISLS1.DataSource = dst.Tables("SATISLS1")
        chtSATISLS1.DataSource = DTY
        chtSATISLS1.PieChart.ColumnIndex = -1
        chtSATISLS1.PieChart.OthersCategoryPercent = 2
        'chtSATISLS1.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATISLS1.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATISLS1.Data.IncludeColumn("P00", True)


        chtSATISLS1.DataBind()

        chtSATISLS1.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Sub CreateGraph_SATISLS1_X()

        Dim chtIsVisible As Boolean = chtSATISLS1_X.Visible
        chtSATISLS1_X.Visible = False

        chtSATISLS1_X.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String
        ReDim CL(Periods)

        'this will be necessary for line graph
        'For i As Integer = MOSMAX To 0 Step -1
        '    Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
        '    CL(MOSMAX - i) = Mid(L, 10, 6)
        '    grdSATISLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        'Next
        For i As Integer = 1 To Periods
            'Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
            CL(i - 1) = grdSATISLS1.DisplayLayout.Bands(0).Columns("P" & Format(i, "00")).Header.Caption
            'grdSATISLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        Next

        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtSATISLS1_X.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATISLS1_X.LabelHash = labelHash

        chtSATISLS1_X.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATISLS1_X.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To Periods
            DT.Columns.Add("P" & Format(P, "00"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SATISLS1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SATISLS1").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1

            Dim rowDT As DataRow = DT.NewRow
            rowDT.Item("CODE_VALUE") = row("CODE_VALUE")
            rowDT.Item("DESC_VALUE") = row("DESC_VALUE")
            For P As Integer = 1 To Periods
                rowDT.Item("P" & Format(P, "00")) = row("P" & Format(P, "00"))
            Next
            DT.Rows.Add(rowDT)
        Next
        chtSATISLS1_X.Data.SetRowLabels(RL)
        chtSATISLS1_X.Data.SetColumnLabels(CL)

        chtSATISLS1_X.DataSource = DT
        'chtSATISLS1_X.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATISLS1_X.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATISLS1_X.Data.IncludeColumn("P00", False)

        chtSATISLS1_X.DataBind()

        chtSATISLS1_X.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Private Sub trkbrXAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrXAxis.Scroll
        chtSATISLS1_X.Axis.X.ScrollScale.Scale = Me.trkbrXAxis.Value / 100.0
    End Sub

    Private Sub trkbrYAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrYAxis.Scroll
        chtSATISLS1_X.Axis.Y.ScrollScale.Scale = Me.trkbrYAxis.Value / 100.0
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Charts").Visible = (tabDetails.SelectedTab.Key = "Charts")
    End Sub

    Private Sub optTotalsChartType_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTotalsChartType.ValueChanged
        Set_Totals_ChartType()
    End Sub

    Private Sub chkTotalsChart3D_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTotalsChart3D.CheckedChanged
        Set_Totals_ChartType()
    End Sub

    Sub Set_Totals_ChartType()
        If Not chkTotalsChart3D.Checked Then
            chtSATISLS1_X.ChartType = ChartType.LineChart
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATISLS1.ChartType = ChartType.PieChart
                Case "DoughnutChart"
                    chtSATISLS1.ChartType = ChartType.DoughnutChart
            End Select
        Else
            chtSATISLS1_X.ChartType = ChartType.LineChart3D
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATISLS1.ChartType = ChartType.PieChart3D
                Case "DoughnutChart"
                    chtSATISLS1.ChartType = ChartType.DoughnutChart3D
            End Select

        End If
    End Sub

    Private Sub cbeColor_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeColor.ValueChanged
        'chtSATISLS1.ColorModel.ModelStyle = cbeColor.ValueMember
        'chtSATISLS1_X.ColorModel.ModelStyle = Infragistics.UltraChart.Shared.Styles.ColorModels.PureRandom
        chtSATISLS1.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), cbeColor.SelectedItem.ToString()), ColorModels)
        chtSATISLS1_X.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), cbeColor.SelectedItem.ToString()), ColorModels)

    End Sub

    Sub Setup_Map()
        '' create the layer
        Dim points As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), ASCMAIN1.Folders("Images") & "ABS\UsMap\US_STATES.xml")
        'Dim USmap As New MapLayer(points)
        USmap = New MapLayer(points)

        dst.Tables("SATISLSS").Rows.Clear()
        US_STATES = USmap.STATES
        For i As Integer = 0 To USmap.STATES.Length - 1
            dst.Tables("SATISLSS").Rows.Add(New Object() {"", USmap.STATES(i), 0})
        Next

        '' set the layer
        Me.UltraChart1.ChartType = ChartType.Composite
        Me.UltraChart1.CompositeChart.ChartAreas.Add(New ChartArea())
        Me.UltraChart1.UserLayerIndex = New String() {"USMap"}
        Me.UltraChart1.Layer.Add("USMap", USmap)

        '' set the tooltip.
        Dim labelRenderers As New Hashtable()
        labelRenderers.Add("USMap", New USMapLabelRenderer(dst.Tables("SATISLSS")))
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
                Dim rows() As DataRow = dst.Tables("SATISLSS").Select("STATE_NAME = '" & US_STATES(I) & "'")
                Dim SALES As Int32 = 0
                If rows.Length = 1 Then
                    SALES = Val(rows(0).Item("SALES") & "")
                End If
                StatesDataFromDataSource(I) = New StateDataInfo(US_STATES(I), SALES, "")
            Next
        End If
        'StatesDataFromDataSource(0) = New StateExpenseViewInfo("Alabama", 1915560.96, "")
        Return StatesDataFromDataSource
    End Function
#End Region

    Private Sub grdSATISLSS_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSATISLSS.DoubleClickRow
        Show_Filter(grdSATISLS1, True)
        grdSATISLS1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdSATISLS1.Rows.ColumnFilters("SUB_CODE_VALUE2").FilterConditions.Add _
        (Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Equals, e.Row.Cells("STATE_CODE").Text)
        chkShowDetails.Checked = True
    End Sub

    Private Sub grdSATISLSS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATISLSS.InitializeRow
        If USmap.COLORS.ContainsKey(e.Row.Cells("STATE_NAME").Text) Then
            e.Row.Cells("SALES").Appearance.ForeColor = USmap.COLORS(e.Row.Cells("STATE_NAME").Text)
        End If
    End Sub

    Private Sub optXP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optXP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_Data()
    End Sub

    Private Sub grdSATISLSS_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSATISLSS.InitializeLayout

    End Sub

    Private Sub chkExtendedData_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkExtendedData.CheckedChanged
        If Not chkExtendedData.Checked Then
            chkPriorYear.Checked = False
        End If
        chkPriorYear.Enabled = chkExtendedData.Checked
    End Sub

End Class