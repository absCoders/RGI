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

Public Class SAFCSLS1
    Dim SATCSLS1 As String
    Dim SATCSLS2 As String
    Dim sqlSATCSLS1 As String
    Dim sqlSATCSLS1_sum As String
    Dim sqlSATCSLS2 As String
    Dim sqlSOTINVHX As String
    Dim sqlSATCSLS1_STORES As String
    Dim STORES As List(Of String)
    Dim STORES_XX As String
    Dim PERIODS_XX As String

    Dim SOTINVHO As String

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW0 As String
    Dim RYW1 As String
    Dim Periods As Integer
    Dim CUST_CODE As String
    Dim US_STATES() As String
    Dim USmap As MapLayer
    Dim Stores_Max As Integer = 120

    Dim summary_generated As Boolean = False
    Dim summary_parameters As String = ""
    Dim SATCSLS1_summary As String
    Dim SATCSLS2_summary As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        Set_cmbYW("RYW0", ASCMAIN1.CYW, -3 * 52, 0, -13)
        Set_cmbYW("RYW1", ASCMAIN1.CYW, -3 * 52, 0, 0)

        With dst
            ASCMAIN1.sql = "Select STYLE_CODE CODE_VALUE, STYLE_DESC DESC_VALUE from ICTSTYL1"
            Create_TDA(.Tables.Add, "SATCSLS1", "**", 0, False)
            With .Tables("SATCSLS1")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
                .Columns.Add("CGS", GetType(System.Int32))
                .Columns.Add("GPA", GetType(System.Int32), "P00 - CGS")
                .Columns.Add("GPP", GetType(System.Decimal), "IIF(P00=0,0,100 * GPA/P00)")
            End With

            ASCMAIN1.sql = "Select STYLE_CODE CODE_VALUE, '0' YEAR, 'XX' DATA_TYPE from ICTSTYL1"
            Create_TDA(.Tables.Add, "SATCSLS1_DTL", "**", 0, False)
            With .Tables("SATCSLS1_DTL")
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            .Relations.Add("SATCSLS1_SATCSLS1_DTL" _
                           , New DataColumn() {.Tables("SATCSLS1").Columns("CODE_VALUE")} _
                           , New DataColumn() {.Tables("SATCSLS1_DTL").Columns("CODE_VALUE")})

            ASCMAIN1.sql = "Select STYLE_CODE CODE_VALUE_PARENT, STYLE_CODE CODE_VALUE, STYLE_DESC DESC_VALUE from ICTSTYL1"
            Create_TDA(.Tables.Add, "SATCSLS2", "**", 0, False)
            With .Tables("SATCSLS2")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            Create_Relation("SATCSLS1", "SATCSLS2", "CODE_VALUE", "CODE_VALUE_PARENT")

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" _
                & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO" _
                & ", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.ORDR_NO" _
                & ", SOTINVH2.CUST_CODE, NVL(SOTINVH1.CUST_STORE_NO,'000000') CUST_STORE_NO" _
                & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION" _
                & ", SOTINVH2.STYLE_CODE, ICTSTYL1.STYLE_DESC" _
                & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_UNIT_COST" _
                & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" _
                & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST CGS" _
                & " from SOTINVH2,ICTSTYL1,ARTCUST2,SOTINVH1 " _
                & " where ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE " _
                & " and ARTCUST2.CUST_CODE (+) = SOTINVH1.CUST_CODE " _
                & " and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' " _
                & " and ARTCUST2.CUST_ADDR_CODE (+) = SOTINVH1.CUST_STORE_NO " _
                & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE " _
                & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            sqlSOTINVHX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False)
            .Tables("SOTINVHX").Columns.Add("GPA", GetType(System.Decimal), "ISNULL(ORDR_AMT_SHIP,0) - ISNULL(CGS,0)")
            .Tables("SOTINVHX").Columns.Add("GPP", GetType(System.Decimal), "IIF(ISNULL(ORDR_AMT_SHIP,0)=0,0,100 * ISNULL(GPA,0)/ISNULL(ORDR_AMT_SHIP,0))")

            Create_TDA(.Tables.Add, "TATSTATE", "*", 0, False)

            With .Tables.Add("SATCSLSS")
                .Columns.Add("STATE_CODE")
                .Columns.Add("STATE_NAME")
                .Columns.Add("SALES", GetType(System.Int32))
            End With

        End With

        Fill_Records("TATSTATE")

        grdSATCSLS1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
        grdSATCSLS1.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdSATCSLS1.DisplayLayout.MaxBandDepth = 1
        grdSATCSLS1.DataSource = dst.Tables("SATCSLS1")

        grdSATCSLS2.DataSource = dst.Tables("SATCSLS2")

        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdEDT852T1.DataSource = dst.Tables("EDT852T1")

        Dim dvw As DataView = dst.Tables("SATCSLSS").DefaultView
        dvw.RowFilter = "SALES <> 0"
        grdSATCSLSS.DataSource = dvw

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, New String() {"ORDR_QTY_SHIP", "ORDR_AMT_SHIP", "CGS", "GPA"})
        Create_Summary(grdSOTINVHX, "GPP", "Custom")

        Create_Summary(grdSATCSLSS, "STATE_CODE", "Count")
        Create_Summary(grdSATCSLSS, "SALES")



        With grdSATCSLS1.DisplayLayout.Bands("SATCSLS1")
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
            .Columns("CGS").Header.Fixed = True
            .Columns("GPA").Header.Fixed = True
            .Columns("GPP").Header.Fixed = True

            .Columns("CGS").Header.Caption = "CGS"
            .Columns("GPA").Header.Caption = "$GP"
            .Columns("GPP").Header.Caption = "GP%"
            .Columns("GPP").Format = "#.0"
            .Columns("GPP").Width = 70
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSATCSLS1.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Color.White
            gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            gcol.Header.Appearance.BackColor2 = Color.LightGray
            If New String() {"CGS", "GPA", "GPP"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = Color.Gold
            ElseIf gcol.Key.StartsWith("P") And gcol.Key.Length = 3 Then
                gcol.Header.Appearance.BackColor2 = Color.LightGreen
            End If
        Next
        grdSATCSLS1.DisplayLayout.Override.ActiveRowAppearance.BackColor = Color.LightGreen

        Create_Summary(grdSATCSLS1, "CODE_VALUE", "Count")
        For P As Integer = 0 To Stores_Max
            Create_Summary(grdSATCSLS1, "P" & Format(P, "00"))
        Next
        Create_Summary(grdSATCSLS1, New String() {"PXX", "CGS", "GPA"})
        Create_Summary(grdSATCSLS1, "GPP", "Custom")

        With grdSATCSLS2.DisplayLayout.Bands("SATCSLS2")
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

        Create_Summary(grdSATCSLS2, "CODE_VALUE", "Count")
        For P As Integer = 0 To Stores_Max
            Create_Summary(grdSATCSLS2, "P" & Format(P, "00"))
        Next
        Create_Summary(grdSATCSLS2, "PXX")

        With chtSATCSLS1
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
        cbeColor.SelectedItem = cbeColor.Items(cbeColor.FindString(System.Enum.GetName(GetType(ColorModels), chtSATCSLS1.ColorModel.ModelStyle), 0))

        'cbeColorBest.DataSource = System.Enum.GetNames(GetType(System.Drawing.Color))
        'cbeColorBest.SelectedItem = cbeColorBest.Items(cbeColor.FindString("Yellow", 0))
        'cbeColorWorst.DataSource = System.Enum.GetNames(GetType(System.Drawing.Color))
        'cbeColorWorst.SelectedItem = cbeColorBest.Items(cbeColor.FindString("Red", 0))

        Setup_Map()

        optSI.Tag = "*"
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                Validate_Code("CUST_CODE")
                Validate_Code("ORDR_TYPE_CODE", , True)

                If EMsg = "" Then

                    If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                        If cdr.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                And Not TAC.TACMAIN1.SREP_CODEs.Contains(cdr.Item("SREP_CODE") & "") Then


                            Dim found_store As Boolean = False
                            ASCMAIN1.sql = "Select Distinct SREP_CODE from ARTCUST2 where CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                            ASCMAIN1.sql &= " UNION "
                            ASCMAIN1.sql &= "Select Distinct SELL_CODE from ARTCUST2 where CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                            For Each rowARTCUST2_SREP As DataRow In ASCDATA1.GetDataTable.Select("")
                                If rowARTCUST2_SREP.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                    And Not TAC.TACMAIN1.SREP_CODEs.Contains(rowARTCUST2_SREP.Item("SREP_CODE") & "") Then
                                Else
                                    found_store = True
                                End If
                            Next

                            If Not found_store Then
                                EMsg &= vbCr & "Customer " & Absx1.txtFor("CUST_CODE").Text & " is not connected to Sales Rep code " & TAC.TACMAIN1.SREP_CODE
                            End If
                        End If
                    End If

                    Validate_Range(EMsg)


                End If

            Case "Load Summary"
                Validate_Code("ORDR_TYPE_CODE", , True)
                validate_Range(EMsg)

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
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print Report"
                Print_Report()

            Case "Load Summary"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Sales Summary by Customer")

                Create_SATCSLS1()

                SATCSLS1_summary = ASCMAIN1.Temp_Table("Select * from " & SATCSLS1)
                SATCSLS2_summary = ASCMAIN1.Temp_Table("Select * from " & SATCSLS2)

                Load_Summary()

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

                summary_generated = True
                summary_parameters = optRANGE.Value & Absx1.cmbFor("RYP0").Text & Absx1.cmbFor("RYP0").Text
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Load Summary").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Load Summary").Visible = Not (Trim(ASCMAIN1.USER_CODES) = "FS")
                .Groups("Data Options").Visible = tf
                .Groups("Options").Visible = Not tf
                .Groups("Charts").Visible = False
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = tf
        grpSummary.Visible = Not tf

        If ScreenMode Then
            grdSATCSLS1.Parent = SplitContainer1.Panel1
            optSI.Visible = True
            optXP.Visible = True
            chkShowDetails.Visible = True
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATCSLS1", "SATCSLS1_DTL", "SATCSLS2", "SOTINVHX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        'Absx1.txtFor("CUST_CODE").Text = ""
        tabDetails.SelectedTab = tabDetails.Tabs("Details")

        If summary_parameters <> optRANGE.Value & Absx1.cmbFor("RYP0").Text & Absx1.cmbFor("RYP0").Text Then
            summary_generated = False
        End If
        If summary_generated Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Restoring Summary")
            'ASCDATA1.ExecuteSQL("Truncate Table " & SATCSLS1)
            'ASCDATA1.ExecuteSQL("Truncate Table " & SATCSLS2)
            'ASCDATA1.ExecuteSQL("Insert into " & SATCSLS1 & " Select * from " & SATCSLS1_summary)
            'ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " Select * from " & SATCSLS2_summary)
            SATCSLS1 = SATCSLS1_summary
            SATCSLS2 = SATCSLS2_summary
            Load_Summary()
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Customer Sales Data")
        Save_Header_Fields(UltraGroupBox1)
        CUST_CODE = HFs("CUST_CODE")
        Create_SATCSLS1()
        optXP.Items(0).DisplayText = optRANGE.Items(optRANGE.CheckedIndex).DisplayText
        Load_Data()
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "CUST_CODE"

                'If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                '    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                '        sql_where = " and ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')"
                '    Else
                '        sql_where = " and ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'"
                '    End If
                'End If
                If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                        sql_where = " and (ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')))"
                    Else
                        sql_where = " and (ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'))"
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSATCSLS1, "SSSSSSSSS", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE")
        Call Load_Popup_Menu(grdSATCSLS2, "SSSSSSSSS", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE")
        Call Load_Popup_Menu(grdEDT852T1, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdSOTINVHX, "SSBBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Show Invoice", "Show Costing")
        Call Load_Popup_Menu(grdSATCSLSS, "CC", "Best", "Worst")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_btn As UltraWinToolbars.ButtonTool
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show SUB_CODE_VALUE1") Then
            For I As Integer = 1 To 5
                Dim COLUMN_NAME As String = "SUB_CODE_VALUE" & CStr(I)
                tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
                tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            Next

            COLUMN_NAME = "RTL_PRICE"
            tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = (EntryMode = "E")
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
            tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            COLUMN_NAME = "WSL_PRICE"
            tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = (EntryMode = "E")
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
            tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
        End If

        If tlb_pop.Tools.Exists("Show Costing") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Show Costing"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = False
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSOTINVHX"
                    Dim ORDR_TYPE_CODE As String = grd.ActiveRow.Cells("ORDR_TYPE_CODE").Value & ""
                    If ORDR_TYPE_CODE = "BTB" Then
                        Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                        tlb_btn = DirectCast(tlb_pop.Tools("Show Costing"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = True
                        tlb_btn.Tag = ORDR_NO
                    End If

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

            Case "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE"

                Dim COLUMN_NAME As String = Mid(e.Tool.Key, 6)
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                If grdSATCSLS1.DisplayLayout.Bands.Count > 1 Then
                    With grdSATCSLS1.DisplayLayout.Bands(1).Columns("DATA_TYPE")
                        If tlb_sbt.Checked Then
                            .ColSpan += 1
                        Else
                            .ColSpan -= 1
                        End If
                    End With
                End If

            Case "Best"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorEnd

            Case "Worst"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorBegin

            Case "Show Costing"
                Dim tlb_btn As UltraWinToolbars.ButtonTool = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Dim ORDR_NO As String = tlb_btn.Tag
                ASCMAIN1.sql = "Select * from POTSHIP2 where ORDR_NO = '" & ORDR_NO & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If row IsNot Nothing Then
                    Dim PO_SHIPMENT_NO As String = row.Item("PO_SHIPMENT_NO")
                    Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPC")
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Show Invoice"
                Dim FILENAME As String = ""
                If grd.ActiveRow IsNot Nothing Then
                    If Not grd.ActiveRow.Selected Then
                        grd.Selected.Rows.Clear()
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.Selected Then
                    Exit Sub
                End If

                'Dim INV_TYPE As String = grd.ActiveRow.Cells("INV_TYPE").Value & ""
                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value & ""

                'If INV_TYPE <> "I" And INV_TYPE <> "C" Then
                '    Exit Sub
                'End If
                FILENAME = TAC.SOCMAIN1.Create_Invoice(Me, INV_NO)

                Show_Document(FILENAME)

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
                'grdSATCSLSS.DataBind()
                Application.DoEvents()
                grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

            Case "Worst"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                'grdSATCSLSS.DataBind()
                Application.DoEvents()
                grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        End Select

    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub
#End Region

    Sub Create_SATCSLS1()

        If SATCSLS1 = SATCSLS1_summary Then
            SATCSLS1 = ""
            SATCSLS2 = ""
        End If
        If SATCSLS2 = "" Then
            SATCSLS1 = ASCMAIN1.Temp_Table("Select CUST_CODE from ARTCUST1 where ROWNUM < 1")
            SATCSLS2 = ASCMAIN1.Temp_Table("Select CUST_CODE from ARTCUST1 where ROWNUM < 1")
        End If
        ASCDATA1.ExecuteSQL("Drop Table " & SATCSLS1)
        ASCDATA1.ExecuteSQL("Drop Table " & SATCSLS2)

        Dim PX As String = "ORDR_YYYYPP_UPDATED"
        Dim PX0 As String = RYP0
        Dim PX1 As String = RYP1

        ' PROBABLY JUST SHOULD PUT STYLE_CODE_SUB, ORDR_NO, ORDR_LNO into SOTINVH2
        If SOTINVHO <> "" Then ASCDATA1.ExecuteSQL("Drop Table " & SOTINVHO)
        Dim sqlx0 As String = ""
        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            sqlx0 = " AND ROWNUM < 1"
        End If

        Dim ORDR_TYPE_CODE As String = Absx1.txtFor("ORDR_TYPE_CODE").Text
        SOTINVHO = ASCMAIN1.Temp_Table("Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH2.INV_LNO" & vbCrLf _
                                       & ", SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO, SOTORDR2.STYLE_CODE_SUB, SOTINVH2.STYLE_CODE" & vbCrLf _
                                       & " from SOTINVH2,SOTINVH1,SOTPICK2,SOTORDR2" & vbCrLf _
                                       & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                                       & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                                       & IIf(ORDR_TYPE_CODE <> "", "  and SOTINVH1.ORDR_TYPE_CODE = '" & ORDR_TYPE_CODE & "'" & vbCrLf, "") _
                                       & "   and SOTPICK2.PICK_NO = SOTINVH1.PICK_NO" & vbCrLf _
                                       & "   and SOTPICK2.PICK_LNO = SOTINVH2.INV_LNO" & vbCrLf _
                                       & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                                       & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                                       & "   and SOTORDR2.STYLE_CODE_SUB is Not Null" & sqlx0)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHO & " Add Primary Key (INV_TYPE,INV_NO,INV_LNO)")

        sqlSATCSLS1 = ""
        sqlSATCSLS1_sum = ""
        sqlSATCSLS2 = ""
        Dim SQL() As String = New String() {"", ""}
        Dim P As Integer
        PERIODS_XX = ""

        'SQL = ""
        For P = 1 To Periods
            Dim PXP As String = ASCMAIN1.Period_Calc(RYP0, P - 1)
            SQL(0) &= ", Sum (Decode(SOTINVH2." & PX & ",'" & PXP & "',SOTINVH2.ORDR_QTY_SHIP,0)) P" & Format(P, "00") & vbCrLf
            SQL(1) &= ", Sum (Decode(SOTINVH2." & PX & ",'" & ASCMAIN1.Period_Calc(PXP, -12) & "',SOTINVH2.ORDR_QTY_SHIP,0)) P" & Format(P, "00") & vbCrLf
            sqlSATCSLS1 &= ", P" & Format(P, "00")
            sqlSATCSLS1_sum &= ", Sum (P" & Format(P, "00") & ") P" & Format(P, "00")
            sqlSATCSLS2 &= ", SUM (P" & Format(P, "00") & ") P" & Format(P, "00")
            PERIODS_XX &= "+P" & Format(P, "00")
        Next

        Dim YEAR_max As Int32 = 0
        If EntryMode = "E" Then
            If chkPriorYear.Checked Then
                YEAR_max = 1
            End If
        End If

        For YEAR As Int32 = 0 To YEAR_max

            PX = "OPS_YYYYPP"

            PX0 = RYP0
            PX1 = RYP1
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
            & ", SOTINVH1.CUST_CODE, NVL(SOTINVH1.CUST_STORE_NO,'000000') CUST_STORE_NO, SOTINVH2.STYLE_CODE" & vbCrLf _
            & SQL(YEAR) & vbCrLf _
            & " from SOTINVH1,SOTINVH2" & vbCrLf _
            & " where SOTINVH1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & IIf(ORDR_TYPE_CODE <> "", "  and SOTINVH1.ORDR_TYPE_CODE = '" & ORDR_TYPE_CODE & "'" & vbCrLf, "") _
            & " and SOTINVH1.ORDR_YYYYPP_UPDATED Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & " and ROWNUM < 1" & vbCrLf _
            & " group by SOTINVH1.CUST_CODE, NVL(SOTINVH1.CUST_STORE_NO,'000000'), SOTINVH2.STYLE_CODE"

            Dim SQL_ORIG As String = sqla

            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & SATCSLS2 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS2 _
                & " Add Primary Key (DATA_TYPE, YEAR, CUST_CODE, CUST_STORE_NO, STYLE_CODE)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)
            End If

            PX = "ORDR_YYYYPP_UPDATED"
            sqla = ""
            For P = 1 To Periods
                Dim PXP As String = ASCMAIN1.Period_Calc(RYP0, P - 1 - 12 * YEAR)
                sqla &= ", Sum (Decode(SOTINVH2." & PX & ",'" & PXP & "',SOTINVH2.ORDR_QTY_SHIP,0)) P" & Format(P, "00") & vbCrLf
            Next
            sqla = "Select DECODE(SOTINVH2.INV_TYPE,'I','S','C','R') || 'U' DATA_TYPE" & vbCrLf _
            & ", '" & CStr(YEAR) & "' YEAR" & vbCrLf _
            & ", SOTINVH2.CUST_CODE, NVL(SOTINVH1.CUST_STORE_NO,'000000') CUST_STORE_NO" & vbCrLf _
            & ", NVL(SOTINVHO.STYLE_CODE_SUB,SOTINVH2.STYLE_CODE)" & vbCrLf _
            & sqla & vbCrLf _
            & " from SOTINVH2,SOTINVH1," & SOTINVHO & " SOTINVHO " & vbCrLf _
            & " where SOTINVH2." & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & IIf(ORDR_TYPE_CODE <> "", "  and SOTINVH1.ORDR_TYPE_CODE = '" & ORDR_TYPE_CODE & "'" & vbCrLf, "") _
            & IIf(EntryMode = "E", " and SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf, "") _
            & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & " and SOTINVHO.INV_TYPE (+) = SOTINVH2.INV_TYPE and SOTINVHO.INV_NO (+) = SOTINVH2.INV_NO and SOTINVHO.INV_LNO (+) = SOTINVH2.INV_LNO" & vbCrLf _
            & " group by DECODE(SOTINVH2.INV_TYPE,'I','S','C','R') || 'U', SOTINVH2.CUST_CODE" & vbCrLf _
            & ", NVL(SOTINVH1.CUST_STORE_NO,'000000')" & vbCrLf _
            & ", NVL(SOTINVHO.STYLE_CODE_SUB,SOTINVH2.STYLE_CODE)" & vbCrLf
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)

            Dim SQLA_ORIG As String = sqla

            sqla = Replace(sqla, "'U' DATA_TYPE", "'D' DATA_TYPE")
            sqla = Replace(sqla, " || 'U',", " || 'D',")
            sqla = Replace(sqla, "ORDR_QTY_SHIP", "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)

            sqla = SQLA_ORIG
            sqla = Replace(sqla, "'U' DATA_TYPE", "'P' DATA_TYPE")
            sqla = Replace(sqla, " || 'U',", " || 'P',")
            ' sqla = Replace(sqla, "ORDR_QTY_SHIP", "ORDR_QTY_SHIP * (ORDR_UNIT_PRICE - ORDR_UNIT_COST)")
            sqla = Replace(sqla, "SOTINVH2.ORDR_QTY_SHIP", "NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0)")
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS2 & " " & sqla)

            sqla = "Select 'I' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", CUST_CODE, STYLE_CODE CODE_VALUE" _
            & sqlSATCSLS2 & " from " & SATCSLS2 _
            & " group by DATA_TYPE, CUST_CODE, STYLE_CODE"
            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & SATCSLS1 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS1 & " Add Primary Key (SI, DATA_TYPE, YEAR, CUST_CODE, CODE_VALUE)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SATCSLS1 & " " & sqla)
            End If

            sqla = "Select 'S' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", CUST_CODE, CUST_STORE_NO CODE_VALUE" & sqlSATCSLS2 _
            & " from " & SATCSLS2 & " group by DATA_TYPE, CUST_CODE, CUST_STORE_NO"
            ASCDATA1.ExecuteSQL("Insert into " & SATCSLS1 & " " & sqla)

        Next YEAR

        ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS1 & " Add PXX NUMBER (10,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS2 & " Add PXX NUMBER (10,0)")

        ASCMAIN1.sql = "Update " & SATCSLS1 & " X SET PXX = (Select P" & Format(Periods, "00") & " from " & SATCSLS1 & " " _
        & " where SI = X.SI and DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and CUST_CODE = X.CUST_CODE and CODE_VALUE = X.CODE_VALUE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & SATCSLS2 & " X SET PXX = (Select P" & Format(Periods, "00") & " from " & SATCSLS2 & " " _
        & " where DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and CUST_CODE = X.CUST_CODE and CUST_STORE_NO = X.CUST_STORE_NO AND STYLE_CODE = X.STYLE_CODE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        STORES = New List(Of String)
        STORES_XX = ""
        Dim SQLX As String = ""

        If EntryMode = "E" Then
            SQLX = "Select Distinct CUST_STORE_NO from " & SATCSLS2 & " order by CUST_STORE_NO"
            For Each row As DataRow In ASCDATA1.GetDataTable(SQLX).Rows
                Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                STORES.Add(CUST_STORE_NO)
            Next

        End If

        SQLX = ""
        If EntryMode = "E" Then
            For S As Integer = 1 To STORES.Count
                SQLX &= ", Sum (Decode(CUST_STORE_NO,'" & STORES(S - 1) & "'," & Mid(PERIODS_XX, 2) & ",0)) P" & Format(S, "00") & vbCrLf
                STORES_XX &= "+P" & Format(S, "00")
            Next
        End If
        sqlSATCSLS1_STORES = SQLX

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

    Private Sub optSI_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSI.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        optSI.Tag = "*"
        Load_Data()
    End Sub

    Sub Setup_grd()

        Dim CAPTION As String = optType1.Text & " (" & optType2.Text & ") for " & CUST_CODE
        If optSI.Value = "S" Then
            CAPTION &= ", by Store"
        Else
            CAPTION &= ", by Item"
        End If
        If EntryMode = "E" Then
            grdSATCSLS1.Text = CAPTION
        End If
        If Absx1.txtFor("ORDR_TYPE_CODE").Text <> "" Then
            grdSATCSLS1.Text &= "- Order Type " & Absx1.txtFor("ORDR_TYPE_CODE").Text & " only"
        End If

        Dim g1 As UltraWinGrid.UltraGrid
        Dim g2 As UltraWinGrid.UltraGrid
        If optSI.Value = "S" Then
            g1 = grdSATCSLS1
            g2 = grdSATCSLS2
        Else
            g1 = grdSATCSLS2
            g2 = grdSATCSLS1
        End If

        If optSI.Tag = "*" Then

            optSI.Tag = ""

            With g1.DisplayLayout.Bands(0)
                .Columns("CODE_VALUE").Header.Caption = "Store"
                .Columns("DESC_VALUE").Header.Caption = "Location"
                .Columns("SUB_CODE_VALUE1").Header.Caption = "DC"
                .Columns("SUB_CODE_VALUE2").Header.Caption = "State"
                .Columns("SUB_CODE_VALUE3").Header.Caption = "City"
                .Columns("SUB_CODE_VALUE4").Header.Caption = "Zip"
                .Columns("SUB_CODE_VALUE5").Header.Caption = "Group"
                .Columns("RTL_PRICE").Header.Caption = "Retail"
                .Columns("WSL_PRICE").Header.Caption = "WhSale"

                .Columns("CODE_VALUE").Width = 60
                .Columns("DESC_VALUE").Width = 140
                .Columns("SUB_CODE_VALUE1").Width = 40
                .Columns("SUB_CODE_VALUE2").Width = 50
                .Columns("SUB_CODE_VALUE3").Width = 50
                .Columns("SUB_CODE_VALUE4").Width = 50
                .Columns("SUB_CODE_VALUE5").Width = 50
                .Columns("RTL_PRICE").Width = 65
                .Columns("WSL_PRICE").Width = 65
            End With

            With g2.DisplayLayout.Bands(0)
                .Columns("CODE_VALUE").Header.Caption = "Style"
                .Columns("DESC_VALUE").Header.Caption = "Description"
                .Columns("SUB_CODE_VALUE1").Header.Caption = "Division"
                .Columns("SUB_CODE_VALUE2").Header.Caption = "Group"
                .Columns("SUB_CODE_VALUE3").Header.Caption = "Class"
                .Columns("SUB_CODE_VALUE4").Header.Caption = "Customer"
                .Columns("SUB_CODE_VALUE5").Header.Caption = "Supplier"
                .Columns("RTL_PRICE").Header.Caption = "Retail"
                .Columns("WSL_PRICE").Header.Caption = "WhSale"

                .Columns("CODE_VALUE").Width = 120
                .Columns("DESC_VALUE").Width = 180
                .Columns("SUB_CODE_VALUE1").Width = 80
                .Columns("SUB_CODE_VALUE2").Width = 60
                .Columns("SUB_CODE_VALUE3").Width = 60
                .Columns("SUB_CODE_VALUE4").Width = 60
                .Columns("SUB_CODE_VALUE5").Width = 60
                .Columns("RTL_PRICE").Width = 65
                .Columns("WSL_PRICE").Width = 65
            End With

            For Each G As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdSATCSLS1, grdSATCSLS2}
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

        End If

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
                    For P As Integer = 0 To Stores_Max
                        COLUMN_NAME = "P" & Format(P, "00")
                        If optSI.Value = "I" And optXP.Value = "S" And G.Name = grdSATCSLS1.Name Then
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

    Private Sub grdSATCSLS1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATCSLS1.AfterRowActivate
        If EntryMode = "E" Then
            Setup_grdSATCSLS2()
        End If
    End Sub

    Sub Setup_grdSATCSLS2()

        If grdSATCSLS1.ActiveRow Is Nothing OrElse Not grdSATCSLS1.ActiveRow.IsDataRow Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
            Exit Sub
        Else
            chkShowDetails.Enabled = True
        End If

        Dim DATA_TYPE As String = optType1.Value & optType2.Value
        Dim CODE_VALUE_PARENT As String = grdSATCSLS1.ActiveRow.Cells("CODE_VALUE").Text

        Load_SATCSLS2(DATA_TYPE, CODE_VALUE_PARENT, False)
        Sort_grdColumns(grdSATCSLS2, "CODE_VALUE")

        Dim CAPTION As String = optType1.Text & " (" & optType2.Text & ") for " & CUST_CODE
        If optSI.Value = "S" Then
            CAPTION &= " - Store " & CODE_VALUE_PARENT & ", by Item"
        Else
            CAPTION &= " - Item " & CODE_VALUE_PARENT & ", by Store"
        End If
        grdSATCSLS2.Text = CAPTION

        Dim sql As String = ""
        sql = Replace(sqlSOTINVHX, "from SOTINVH2", "from SOTINVH2," & SOTINVHO & " SOTINVHO") _
            & " and SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & " and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
            & " and SOTINVHO.INV_TYPE (+) = SOTINVH2.INV_TYPE  and SOTINVHO.INV_NO (+) = SOTINVH2.INV_NO  and SOTINVHO.INV_LNO (+) = SOTINVH2.INV_LNO" & vbCrLf
        If optSI.Value = "I" Then
            sql &= " and NVL(SOTINVHO.STYLE_CODE_SUB,SOTINVH2.STYLE_CODE) = '" & CODE_VALUE_PARENT & "'" & vbCrLf
        Else
            sql &= " and NVL(SOTINVH1.CUST_STORE_NO,'000000') = '" & CODE_VALUE_PARENT & "'" & vbCrLf
        End If
        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("CUST_STORE_NO").Hidden = (optSI.Value = "S")
            .Columns("CUST_STORE_LOCATION").Hidden = (optSI.Value = "S")
            .Columns("STYLE_CODE").Hidden = (optSI.Value = "I")
            .Columns("STYLE_DESC").Hidden = (optSI.Value = "I")
        End With
        Fill_Records("SOTINVHX", "", True, sql)
        grdSOTINVHX.Text = "Sales Documents for " & CUST_CODE & IIf(optSI.Value = "S", " - Store ", " - Item ") & CODE_VALUE_PARENT
        grdSOTINVHX.DisplayLayout.CaptionVisible = DefaultableBoolean.True

    End Sub

    Sub Load_SATCSLS2(ByVal DATA_TYPE As String, ByVal CODE_VALUE_PARENT As String, ByVal all_parents As Boolean)
        Dim sql As String = ""

        If optSI.Value = "I" Then
            sql = "Select SATCSLS2.STYLE_CODE CODE_VALUE_PARENT " & vbCrLf _
            & ", SATCSLS2.CUST_STORE_NO CODE_VALUE" & vbCrLf _
            & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) DESC_VALUE" & vbCrLf _
            & ", ARTCUST2.CUST_DC_NO SUB_CODE_VALUE1" & vbCrLf _
            & ", ARTCUST2.CUST_STATE SUB_CODE_VALUE2" & vbCrLf _
            & ", ARTCUST2.CUST_CITY SUB_CODE_VALUE3" & vbCrLf _
            & ", ARTCUST2.CUST_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
            & ", ARTCUST2.CUST_ADDR_GROUP SUB_CODE_VALUE5" & vbCrLf _
            & sqlSATCSLS1 & ",PXX  from ARTCUST2," & SATCSLS2 & " SATCSLS2 " & vbCrLf _
            & " where ARTCUST2.CUST_CODE (+) = SATCSLS2.CUST_CODE " & vbCrLf _
            & " and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK'" & vbCrLf _
            & " and ARTCUST2.CUST_ADDR_CODE (+) = SATCSLS2.CUST_STORE_NO" & vbCrLf _
            & " and DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
            & IIf(all_parents, "", " and SATCSLS2.STYLE_CODE = '" & CODE_VALUE_PARENT & "'")
        Else
            sql = "Select SATCSLS2.CUST_STORE_NO CODE_VALUE_PARENT " & vbCrLf _
            & ", SATCSLS2.STYLE_CODE CODE_VALUE" & vbCrLf _
            & ", ICTSTYL1.STYLE_DESC DESC_VALUE" & vbCrLf _
            & ", ICTSTYL1.SALES_DIVISION_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ICTSTYL1.STYLE_GROUP_CODE SUB_CODE_VALUE2" & vbCrLf _
            & ", ICTSTYL1.STYLE_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
            & ", ICTSTYL1.CUST_CODE SUB_CODE_VALUE4" & vbCrLf _
            & ", ICTSTYL1.VEND_CODE SUB_CODE_VALUE5" & vbCrLf _
            & ", ICTSTYL1.STYLE_RETAIL RTL_PRICE" & vbCrLf _
            & ", ICTSTYL1.STYLE_PRICE WSL_PRICE" & vbCrLf _
            & sqlSATCSLS1 & ",PXX  from ICTSTYL1," & SATCSLS2 & " SATCSLS2 " & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE (+) = SATCSLS2.STYLE_CODE " & vbCrLf _
            & " and DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
            & IIf(all_parents, "", " and SATCSLS2.CUST_STORE_NO = '" & CODE_VALUE_PARENT & "'")
        End If
        'dst.Tables("SATCSLS1").Rows.Clear()
        Fill_Records("SATCSLS2", "", True, sql)
    End Sub

    Private Sub chkNoDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = Not chkShowDetails.Checked
    End Sub

    Sub Load_Data()

        Dim SI As String = optSI.Value
        If EntryMode = "" Then SI = "S"

        dst.EnforceConstraints = False

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim DATA_TYPE As String = optType1.Value & optType2.Value

        optXP.Visible = (optSI.Value = "I")

        Dim sql As String = ""

        If SI = "S" Then
            If EntryMode = "E" Then
                sql = "Select SATCSLS1.CODE_VALUE CODE_VALUE" & vbCrLf _
                & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) DESC_VALUE" & vbCrLf _
                & ", ARTCUST2.CUST_DC_NO SUB_CODE_VALUE1" & vbCrLf _
                & ", ARTCUST2.CUST_STATE SUB_CODE_VALUE2" & vbCrLf _
                & ", ARTCUST2.CUST_CITY SUB_CODE_VALUE3" & vbCrLf _
                & ", ARTCUST2.CUST_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ARTCUST2.CUST_ADDR_GROUP SUB_CODE_VALUE5" & vbCrLf _
                & sqlSATCSLS1 & ",PXX from ARTCUST2," & SATCSLS1 & " SATCSLS1 " & vbCrLf _
                & " where ARTCUST2.CUST_CODE (+) = SATCSLS1.CUST_CODE " & vbCrLf _
                & " and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK'" & vbCrLf _
                & " and ARTCUST2.CUST_ADDR_CODE (+) = SATCSLS1.CODE_VALUE" & vbCrLf _
                & " and SATCSLS1.SI = '" & SI & "' and SATCSLS1.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and SATCSLS1.YEAR = '0'"
            Else
                sql = "Select SATCSLS1.CUST_CODE CODE_VALUE" & vbCrLf _
                & ", ARTCUST1.CUST_NAME DESC_VALUE" & vbCrLf _
                & ", ARTCUST1.SREP_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ARTCUST1.CUST_STATE SUB_CODE_VALUE2" & vbCrLf _
                & ", ARTCUST1.CUST_CITY SUB_CODE_VALUE3" & vbCrLf _
                & ", ARTCUST1.CUST_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ARTCUST1.TRADE_CLASS_CODE SUB_CODE_VALUE5" & vbCrLf _
                & sqlSATCSLS1_sum & ", Sum (PXX) PXX from ARTCUST1," & SATCSLS1 & " SATCSLS1 " & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = SATCSLS1.CUST_CODE " & vbCrLf _
                & " and SATCSLS1.SI = '" & SI & "' and SATCSLS1.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and SATCSLS1.YEAR = '0'" & vbCrLf _
                & " group by SATCSLS1.CUST_CODE" & vbCrLf _
                & ", ARTCUST1.CUST_NAME" & vbCrLf _
                & ", ARTCUST1.SREP_CODE" & vbCrLf _
                & ", ARTCUST1.CUST_STATE" & vbCrLf _
                & ", ARTCUST1.CUST_CITY" & vbCrLf _
                & ", ARTCUST1.CUST_ZIP_CODE" & vbCrLf _
                & ", ARTCUST1.TRADE_CLASS_CODE"
            End If
        Else
            If optXP.Value = "S" Then

                Dim SQLX As String = sqlSATCSLS1_STORES
                If Mid(DATA_TYPE, 1, 1) = "H" Then
                    SQLX = Replace(sqlSATCSLS1_STORES, Mid(PERIODS_XX, 2), "P" & Format(Periods, "00"))
                End If

                sql = "Select SATCSLS2.STYLE_CODE CODE_VALUE" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC DESC_VALUE" & vbCrLf _
                & ", ICTSTYL1.SALES_DIVISION_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ICTSTYL1.STYLE_GROUP_CODE SUB_CODE_VALUE2" & vbCrLf _
                & ", ICTSTYL1.STYLE_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
                & ", ICTSTYL1.CUST_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ICTSTYL1.VEND_CODE SUB_CODE_VALUE5" & vbCrLf _
                & ", ICTSTYL1.STYLE_RETAIL RTL_PRICE" & vbCrLf _
                & ", ICTSTYL1.STYLE_PRICE WSL_PRICE" & vbCrLf _
                & SQLX & "  from ICTSTYL1," & SATCSLS2 & " SATCSLS2 " & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE (+) = SATCSLS2.STYLE_CODE " & vbCrLf _
                & " and SATCSLS2.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and SATCSLS2.YEAR = '0'" & vbCrLf _
                & " group by SATCSLS2.STYLE_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC" & vbCrLf _
                & ", ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_GROUP_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                & ", ICTSTYL1.CUST_CODE" & vbCrLf _
                & ", ICTSTYL1.VEND_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_RETAIL" & vbCrLf _
                & ", ICTSTYL1.STYLE_PRICE" & vbCrLf
            Else
                sql = "Select SATCSLS1.CODE_VALUE CODE_VALUE" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC DESC_VALUE" & vbCrLf _
                & ", ICTSTYL1.SALES_DIVISION_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ICTSTYL1.STYLE_GROUP_CODE SUB_CODE_VALUE2" & vbCrLf _
                & ", ICTSTYL1.STYLE_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
                & ", ICTSTYL1.CUST_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ICTSTYL1.VEND_CODE SUB_CODE_VALUE5" & vbCrLf _
                & ", ICTSTYL1.STYLE_RETAIL RTL_PRICE" & vbCrLf _
                & ", ICTSTYL1.STYLE_PRICE WSL_PRICE" & vbCrLf _
                & sqlSATCSLS1 & ",PXX from ICTSTYL1," & SATCSLS1 & " SATCSLS1 " & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE (+) = SATCSLS1.CODE_VALUE " & vbCrLf _
                & " and SATCSLS1.SI = '" & SI & "'" _
                & " and SATCSLS1.DATA_TYPE = '" & DATA_TYPE & "'" _
                & " and SATCSLS1.YEAR = '0'" & vbCrLf
            End If
        End If

        dst.Tables("SATCSLS1").Rows.Clear()
        dst.Tables("SATCSLS1_DTL").Rows.Clear()
        dst.Tables("SATCSLS2").Rows.Clear()

        If SI = "I" And optXP.Value = "S" Then
            dst.Tables("SATCSLS1").Columns("P00").Expression = Mid(STORES_XX, 2)
            dst.Tables("SATCSLS1_DTL").Columns("P00").Expression = Mid(STORES_XX, 2)
        Else
            If optType1.Value = "H" Then
                dst.Tables("SATCSLS1").Columns("P00").Expression = Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
            Else
                dst.Tables("SATCSLS1").Columns("P00").Expression = Mid(PERIODS_XX, 2)
            End If
            dst.Tables("SATCSLS1_DTL").Columns("P00").Expression = "IIF(DATA_TYPE='HU' OR DATA_TYPE='HD'," & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3) & "," & Mid(PERIODS_XX, 2) & ")"
        End If

        If optType1.Value = "H" Then
            dst.Tables("SATCSLS2").Columns("P00").Expression = Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3)
        Else
            dst.Tables("SATCSLS2").Columns("P00").Expression = Mid(PERIODS_XX, 2)
        End If

        Fill_Records("SATCSLS1", "", True, sql)
        Sort_grdColumns(grdSATCSLS1, "CODE_VALUE")


        Dim hide_GP As Boolean = True
        If optType2.Value = "D" Then hide_GP = False
        With grdSATCSLS1.DisplayLayout.Bands(0)
            .Columns("CGS").Hidden = hide_GP
            .Columns("GPA").Hidden = hide_GP
            .Columns("GPP").Hidden = hide_GP
        End With

        If Not hide_GP Then
            Dim TBL As DataTable = ASCDATA1.GetDataTable(Replace(sql, "SATCSLS1.DATA_TYPE = 'SD'", "SATCSLS1.DATA_TYPE = 'SP'"))
            For Each ROW As DataRow In TBL.Select("")
                Dim P00 As Decimal = 0
                For P As Integer = 1 To Periods
                    P00 += Val(ROW.Item("P" & Format(P, "00")) & "")
                Next
                Dim ROW1 As DataRow = dst.Tables("SATCSLS1").Rows.Find(ROW.Item(0))
                ROW1.Item("CGS") = P00
            Next
        End If

        If chkExtendedData.Checked Then
            sql = "Select SATCSLS1.CODE_VALUE CODE_VALUE, SATCSLS1.YEAR, SATCSLS1.DATA_TYPE" _
            & sqlSATCSLS1 _
            & " from " & SATCSLS1 & " SATCSLS1 " _
            & " where SATCSLS1.SI = '" & SI & "'" _
            & " and SATCSLS1.DATA_TYPE LIKE '%" & Mid(DATA_TYPE, 2, 1) & "'" _
            & " and (YEAR = '1' OR SATCSLS1.DATA_TYPE <> '" & DATA_TYPE & "')"

            Fill_Records("SATCSLS1_DTL", "", True, sql)
            Sort_grdColumns(grdSATCSLS1, "DATA_TYPE", , 1)
        End If

        If grdSATCSLS1.Rows.Count = 0 Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
        Else
            chkShowDetails.Enabled = True
        End If

        'dst.EnforceConstraints = True

        Setup_grd()
        grdSATCSLS1.DisplayLayout.Bands(0).Columns("P00").Header.Caption = optType2.Text

        tabDetails.Tabs("EDI Documents").Visible = (optType1.Value = "T")
        tabDetails.Tabs("Sales Documents").Visible = (optType1.Value <> "T")

        If SI = "I" Then
            If tabDetails.SelectedTab.Key = "Map" Then
                tabDetails.SelectedTab = tabDetails.Tabs("Details")
            End If
        Else
            'dst.Tables("SATCSLSS").Rows.Clear()
            For Each ROW As DataRow In dst.Tables("SATCSLSS").Rows
                ROW.Item("SALES") = 0
            Next
            For Each row As DataRow In ASCDATA1.SelectDistinct("SATCSLS1", "SUB_CODE_VALUE2").Rows
                Dim rowTATSTATE As DataRow = dst.Tables("TATSTATE").Rows.Find(row(0))
                If rowTATSTATE IsNot Nothing Then
                    Dim rows() As DataRow = dst.Tables("SATCSLSS").Select _
                    ("STATE_NAME = '" & rowTATSTATE("STATE_NAME") & "'")
                    If rows.Length = 1 Then
                        Dim SALES As Decimal = dst.Tables("SATCSLS1").Compute("SUM (P00)", "SUB_CODE_VALUE2 = '" & row(0) & "'")
                        rows(0).Item("STATE_CODE") = rowTATSTATE("STATE_CODE")
                        rows(0).Item("SALES") = SALES
                    End If
                End If

                'dst.Tables("SATCSLSS").Rows.Add(New Object() _
                '    {row("STATE_CODE"), row("STATE_NAME"), Val(row("SALES") & "")})
            Next
            Me.UltraChart1.Data.DataSource = StatesData()
            Me.UltraChart1.Data.DataBind()
        End If
        tabDetails.Tabs("Map").Visible = (SI = "S")

        'CreateGraph_SATCSLS1()
        'CreateGraph_SATCSLS1_X()
        chtSATCSLS1.Visible = True
        chtSATCSLS1_X.Visible = True

        If SI = "S" Then
            If grdSATCSLS1.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE2") Then
                grdSATCSLS1.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE2")
            End If
            If grdSATCSLS2.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE1") Then
                grdSATCSLS2.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE1")
            End If

            ASCMAIN1.Add_Value_List(grdSATCSLS1, "SUB_CODE_VALUE1", , , , "Select CUST_ADDR_CODE, CUST_NAME from ARTCUST2 where CUST_ADDR_TYPE = 'DC'")
            ASCMAIN1.Add_Value_List(grdSATCSLS2, "SUB_CODE_VALUE2", , , , "Select STYLE_CLASS_CODE, STYLE_CLASS_DESC from ICTCLAS1")
        Else
            If grdSATCSLS1.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE1") Then
                grdSATCSLS1.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE1")
            End If
            If grdSATCSLS2.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE2") Then
                grdSATCSLS2.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE2")
            End If

            ASCMAIN1.Add_Value_List(grdSATCSLS1, "SUB_CODE_VALUE2", , , , "Select STYLE_CLASS_CODE, STYLE_CLASS_DESC from ICTCLAS1")
            ASCMAIN1.Add_Value_List(grdSATCSLS2, "SUB_CODE_VALUE1", , , , "Select CUST_ADDR_CODE, CUST_NAME from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "' and CUST_ADDR_TYPE = 'DC'")
        End If

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

    Sub CreateGraph_SATCSLS1()

        Dim chtIsVisible As Boolean = chtSATCSLS1.Visible
        chtSATCSLS1.Visible = False

        chtSATCSLS1.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String

        chtSATCSLS1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATCSLS1.LabelHash = labelHash

        chtSATCSLS1.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATCSLS1.Tooltips.FormatString = "<HIGHLOW>"

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Decimal))
        End With

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SATCSLS1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SATCSLS1").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") ' & ":" & row("DESC_VALUE")
            RLi += 1
            DTY.Rows.Add(New Object() {row.Item("CODE_VALUE"), row.Item("P00")})
        Next
        'chtSATCSLS1.Data.SetRowLabels(RL)
        'chtSATCSLS1.Data.SetColumnLabels(CL)

        'chtSATCSLS1.DataSource = dst.Tables("SATCSLS1")
        chtSATCSLS1.DataSource = DTY
        chtSATCSLS1.PieChart.ColumnIndex = -1
        chtSATCSLS1.PieChart.OthersCategoryPercent = 2
        'chtSATCSLS1.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATCSLS1.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATCSLS1.Data.IncludeColumn("P00", True)


        chtSATCSLS1.DataBind()

        chtSATCSLS1.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Sub CreateGraph_SATCSLS1_X()

        Dim chtIsVisible As Boolean = chtSATCSLS1_X.Visible
        chtSATCSLS1_X.Visible = False

        chtSATCSLS1_X.DataSource = Nothing

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
        '    grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        'Next
        For i As Integer = 1 To Periods
            'Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
            CL(i - 1) = grdSATCSLS1.DisplayLayout.Bands(0).Columns("P" & Format(i, "00")).Header.Caption
            'grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        Next

        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtSATCSLS1_X.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATCSLS1_X.LabelHash = labelHash

        chtSATCSLS1_X.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATCSLS1_X.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To Periods
            DT.Columns.Add("P" & Format(P, "00"), GetType(System.Decimal))
        Next

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("SATCSLS1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("SATCSLS1").Select("", "CODE_VALUE")
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
        chtSATCSLS1_X.Data.SetRowLabels(RL)
        chtSATCSLS1_X.Data.SetColumnLabels(CL)

        chtSATCSLS1_X.DataSource = DT
        'chtSATCSLS1_X.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("P00", False)

        chtSATCSLS1_X.DataBind()

        chtSATCSLS1_X.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Private Sub trkbrXAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrXAxis.Scroll
        chtSATCSLS1_X.Axis.X.ScrollScale.Scale = Me.trkbrXAxis.Value / 100.0
    End Sub

    Private Sub trkbrYAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrYAxis.Scroll
        chtSATCSLS1_X.Axis.Y.ScrollScale.Scale = Me.trkbrYAxis.Value / 100.0
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
            chtSATCSLS1_X.ChartType = ChartType.LineChart
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATCSLS1.ChartType = ChartType.PieChart
                Case "DoughnutChart"
                    chtSATCSLS1.ChartType = ChartType.DoughnutChart
            End Select
        Else
            chtSATCSLS1_X.ChartType = ChartType.LineChart3D
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATCSLS1.ChartType = ChartType.PieChart3D
                Case "DoughnutChart"
                    chtSATCSLS1.ChartType = ChartType.DoughnutChart3D
            End Select

        End If
    End Sub

    Private Sub cbeColor_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeColor.ValueChanged
        'chtSATCSLS1.ColorModel.ModelStyle = cbeColor.ValueMember
        'chtSATCSLS1_X.ColorModel.ModelStyle = Infragistics.UltraChart.Shared.Styles.ColorModels.PureRandom
        chtSATCSLS1.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), cbeColor.SelectedItem.ToString()), ColorModels)
        chtSATCSLS1_X.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), cbeColor.SelectedItem.ToString()), ColorModels)

    End Sub

    Sub Setup_Map()
        '' create the layer
        Dim points As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Application.ExecutablePath), ASCMAIN1.Folders("Images") & "ABS\UsMap\US_STATES.xml")
        'Dim USmap As New MapLayer(points)
        USmap = New MapLayer(points)

        dst.Tables("SATCSLSS").Rows.Clear()
        US_STATES = USmap.STATES
        For i As Integer = 0 To USmap.STATES.Length - 1
            dst.Tables("SATCSLSS").Rows.Add(New Object() {"", USmap.STATES(i), 0})
        Next

        '' set the layer
        Me.UltraChart1.ChartType = ChartType.Composite
        Me.UltraChart1.CompositeChart.ChartAreas.Add(New ChartArea())
        Me.UltraChart1.UserLayerIndex = New String() {"USMap"}
        Me.UltraChart1.Layer.Add("USMap", USmap)

        '' set the tooltip.
        Dim labelRenderers As New Hashtable()
        labelRenderers.Add("USMap", New USMapLabelRenderer(dst.Tables("SATCSLSS")))
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
                Dim rows() As DataRow = dst.Tables("SATCSLSS").Select("STATE_NAME = '" & US_STATES(I) & "'")
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

    Private Sub grdSATCSLSS_ClickCell(sender As Object, e As UltraWinGrid.ClickCellEventArgs) Handles grdSATCSLSS.ClickCell
        Dim COLUMN_NAME As String = e.Cell.Column.Key
    End Sub

    Private Sub grdSATCSLSS_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSATCSLSS.DoubleClickRow
        Show_Filter(grdSATCSLS1, True)
        grdSATCSLS1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdSATCSLS1.Rows.ColumnFilters("SUB_CODE_VALUE2").FilterConditions.Add _
        (Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Equals, e.Row.Cells("STATE_CODE").Text)
        chkShowDetails.Checked = True
    End Sub

    Private Sub grdSATCSLSS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATCSLSS.InitializeRow
        If USmap.COLORS.ContainsKey(e.Row.Cells("STATE_NAME").Text) Then
            e.Row.Cells("SALES").Appearance.ForeColor = USmap.COLORS(e.Row.Cells("STATE_NAME").Text)
        End If
    End Sub

    Private Sub optXP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optXP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If optXP.Value = "S" Then
            If STORES.Count > Stores_Max Then
                MsgBox("Too Many Stores (" & STORES.Count & ") for this option.  Max is " & CStr(Stores_Max))
                optXP.Value = "P"
                Exit Sub
            End If
        End If
        Load_Data()
    End Sub

    Private Sub chkExtendedData_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkExtendedData.CheckedChanged
        If Not chkExtendedData.Checked Then
            chkPriorYear.Checked = False
        End If
        chkPriorYear.Enabled = chkExtendedData.Checked
    End Sub

    Sub Validate_Range(ByRef EMsg As String)

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
            'If Absx1.dteFor("RYW0").Value > Absx1.dteFor("RYW1").Value Then
            '    EMsg &= vbCr & "Starting Week cannot be later than Ending Week"
            'End If

            If EMsg = "" Then
                RYW0 = Absx1.cmbFor("RYW0").Value
                RYW1 = Absx1.cmbFor("RYW1").Value
                Periods = ASCMAIN1.Week_Diff(RYW0, RYW1) + 1
            End If
        End If

        If EMsg = "" Then
            If Periods < 1 Or Periods > Stores_Max Then
                EMsg &= vbCr & "Total number of Periods must be between 1 and " & CStr(Stores_Max)
            End If
        End If

    End Sub

    Private Sub grdSATCSLS1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSATCSLS1.DoubleClickRow
        If EntryMode = "" Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CODE_VALUE").Value
            Click_Command("View")
        End If
    End Sub


    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdSATCSLS1"
                Dim KEY As String = summarySettings.Key
                If KEY = "GPP" Then
                    TOTALS.Add("P00", 0)
                    TOTALS.Add("GPA", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("P00") <> 0 Then CustomValue = 100 * TOTALS("GPA") / TOTALS("P00")
                Else
                    Stop
                End If

            Case "grdSOTINVHX"
                Dim KEY As String = summarySettings.Key
                If KEY = "GPP" Then
                    TOTALS.Add("ORDR_AMT_SHIP", 0)
                    TOTALS.Add("GPA", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("ORDR_AMT_SHIP") <> 0 Then CustomValue = 100 * TOTALS("GPA") / TOTALS("ORDR_AMT_SHIP")
                Else
                    Stop
                End If
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Public Overrides Function CustomStringSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As String, _
        ByVal grd As UltraWinGrid.UltraGrid) As String

        Select Case grd.Name
            Case "grdSATSIST1"
                Dim KEY As String = summarySettings.Key
                CustomValue = "Totals"
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub CustomSummary_Calculate_Totals( _
       ByVal rows As UltraWinGrid.RowsCollection, _
       ByRef TOTALS As Dictionary(Of String, Decimal), _
       ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                CustomSummary_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY = "GPP" Then
                    Dim SI As String = "SIQTY_P" & Mid(KEY, 8)
                    Dim ST As String = "STQTY_P" & Mid(KEY, 8)

                    If TOTALS.ContainsKey("ORDR_AMT_SHIP") Then
                        TOTALS("ORDR_AMT_SHIP") += Val(grow2.Cells("ORDR_AMT_SHIP").Value & "")
                    Else
                        TOTALS("P00") += Val(grow2.Cells("P00").Value & "")
                    End If

                    TOTALS("GPA") += Val(grow2.Cells("GPA").Value & "")
                ElseIf KEY = "TRADE_CLASS_CODE" Then
                    '  TOTALS(KEY) = "Totals"
                End If
            End If
        Next
    End Sub

    Sub Load_Summary()
        optXP.Items(0).DisplayText = optRANGE.Items(optRANGE.CheckedIndex).DisplayText

        Load_Data()
        grdSATCSLS1.Parent = grpSummary

        grdSATCSLS1.Text = "Sales Summary by Customer"
        If Absx1.txtFor("ORDR_TYPE_CODE").Text <> "" Then
            grdSATCSLS1.Text &= "- Order Type " & Absx1.txtFor("ORDR_TYPE_CODE").Text & " only"
        End If
        With grdSATCSLS1.DisplayLayout.Bands(0)
            .Columns("CODE_VALUE").Header.Caption = "Customer"
            .Columns("DESC_VALUE").Header.Caption = "Customer Name"
        End With

        UltraExplorerBar1.Groups("Data Options").Visible = True
        optSI.Visible = False
        optXP.Visible = False
        chkShowDetails.Visible = False
    End Sub
     
End Class