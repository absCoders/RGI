Public Class SOFWALM1

    Dim SOTWALM1 As String
    Dim SLS_AVG As Decimal = 0
    Dim ORD_AVG As Decimal = 0
    Dim VOL_GROUP_min As String = ""
    Dim SLS_FACTOR_min As Decimal = 0
    Dim working_with_imported_order As Boolean = False
    Dim UNITS_PER_INNER As Integer = 1 ' 12
    Dim INNERS_PER_CASE As Integer = 12 ' 5

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' Get_PARM("POTPARM1")

        With dst

            'ASCMAIN1.sql = "Select ARTCUST2.CUST_ADDR_CODE CUST_STORE_NO, ARTCUST2.CUST_NAME CUST_STORE_NAME, ARTCUST2.CUST_DC_NO" _
            '    & " from ARTCUST2 where CUST_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "ARTWALMX", "**", 0, False, "V", 1)
            With .Tables.Add("ARTWALMX")
                .Columns.Add("WALMART_STYLE", GetType(System.String))
                .Columns.Add("CUST_DC_NO", GetType(System.String))
                .Columns.Add("CUST_STORE_NO", GetType(System.String))
                .Columns.Add("CUST_STORE_NAME", GetType(System.String))
                .Columns.Add("VOL_GROUP", GetType(System.String))
                .Columns.Add("VOL_QTY", GetType(System.Int64))
                .Columns.Add("SLS_LIF", GetType(System.Int64))
                .Columns.Add("SLS_PTD", GetType(System.Int64))
                .Columns.Add("SUP_STR", GetType(System.Int64))
                .Columns.Add("SUP_XIT", GetType(System.Int64))
                .Columns.Add("SUP_WHS", GetType(System.Int64))
                .Columns.Add("SUP_OPO", GetType(System.Int64))
                .Columns.Add("SUP_TOT", GetType(System.Int64), "ISNULL(SUP_STR,0)+ISNULL(SUP_XIT,0)+ISNULL(SUP_WHS,0)+ISNULL(SUP_OPO,0)")
                .Columns.Add("WKS_SUP", GetType(System.Decimal))
                .Columns.Add("SLS_FACTOR", GetType(System.Decimal))
                .Columns.Add("ORD_FACTOR", GetType(System.Decimal))
                .Columns.Add("ORDR_QTY_CALC", GetType(System.Decimal))
                .Columns.Add("ORDR_QTY_ROUND", GetType(System.Int64))
                .Columns.Add("ORDR_QTY", GetType(System.Int64))
                .Columns.Add("SEL", GetType(System.String))
                .Columns("SEL").DefaultValue = "0"
                .Columns.Add("TRAITED", GetType(System.String))
                .Columns("TRAITED").DefaultValue = "0"
                .Columns.Add("VALID", GetType(System.String))
                .Columns("VALID").DefaultValue = "0"
                .PrimaryKey = New DataColumn() {.Columns("WALMART_STYLE"), .Columns("CUST_DC_NO"), .Columns("CUST_STORE_NO")}
            End With

            With .Tables.Add("ARTWALMY")
                .Columns.Add("WALMART_STYLE")
                .Columns.Add("CUST_DC_NO")
                .Columns.Add("CARTON_PACK_QTY", GetType(System.Int64))
                .Columns.Add("INNER_PACK_QTY", GetType(System.Int64))
                .Columns.Add("SLS_LIF", GetType(System.Int64))
                .Columns.Add("SLS_PTD", GetType(System.Int64))
                .Columns.Add("SUP_TOT", GetType(System.Int64))
                .Columns.Add("STORES", GetType(System.Int64))
                .Columns.Add("STORES_WITH_SALES", GetType(System.Int64))
                .Columns.Add("SLS_AVG", GetType(System.Decimal))
                .Columns.Add("ORDR_QTY", GetType(System.Int64))
                .Columns.Add("CASES", GetType(System.Int64))
                .Columns.Add("ORDR_QTY_OFF", GetType(System.Int64))
                .PrimaryKey = New DataColumn() {.Columns("WALMART_STYLE"), .Columns("CUST_DC_NO")}
            End With

            Create_Relation("ARTWALMY", "ARTWALMX", "WALMART_STYLE,CUST_DC_NO")

            With .Tables("ARTWALMY")
                .Columns("SLS_LIF").Expression = "SUM(CHILD.SLS_LIF)"
                .Columns("SLS_PTD").Expression = "SUM(CHILD.SLS_PTD)"
                .Columns("SUP_TOT").Expression = "SUM(CHILD.SUP_TOT)"
                .Columns("STORES").Expression = "COUNT(CHILD.CUST_STORE_NO)"
                .Columns("ORDR_QTY").Expression = "SUM(CHILD.ORDR_QTY)"
                .Columns("CASES").Expression = "ORDR_QTY / 1"
                .Columns("ORDR_QTY_OFF").Expression = "ORDR_QTY - CASES * 1"
            End With

            Create_TDA(.Tables.Add, "ARTWALM2", "*")
            Create_TDA(.Tables.Add, "ARTWALM3", "*")

            ASCMAIN1.sql = "Select STORE_NBR, STORE_NAME, REGIONAL_DC, SPECIALTY_DC from ARTWALM2"
            Create_TDA(.Tables.Add, "ARTWALMS", "**", 0, False, "", 1)
            With .Tables("ARTWALMS")
                .Columns.Add("SALES_IND")
                .Columns.Add("VOL_GROUP", GetType(System.String))
                .Columns.Add("VOL_QTY", GetType(System.Int64))
                .Columns.Add("SLS_LIF", GetType(System.Int64))
                .Columns.Add("SLS_PTD", GetType(System.Int64))
                .Columns.Add("SLS_FACTOR", GetType(System.Decimal))
                .Columns.Add("ORD_FACTOR", GetType(System.Decimal))
            End With


            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_DESC WALMART_STYLE" _
                & ", ICTSTYL1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY" _
                & " from ICTSTYL1,ICTSTYC1 where ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE"
            Create_TDA(.Tables.Add, "ARTWALMI", "**", 0, False, "", 1)

            Create_Relation("ARTWALMI", "ARTWALMY", "WALMART_STYLE")

            With .Tables("ARTWALMI")
                .Columns.Add("MIN_QTY", GetType(System.Int64))
                .Columns.Add("MAX_QTY", GetType(System.Int64))
                .Columns.Add("ORDR_QTY", GetType(System.Int64), "SUM(CHILD.ORDR_QTY)")
                .Columns.Add("PO_QTY", GetType(System.Int64))
                .Columns("COLOR_CODE").AllowDBNull = True
                .Columns("STYLE_CODE").AllowDBNull = True
            End With

            With .Tables.Add("ARTWALMV")
                .Columns.Add("VOL_GROUP")
                .Columns.Add("MIN_FACTOR", GetType(System.Decimal))
                .Columns.Add("QTY", GetType(System.Int64))
                .Columns.Add("STORES", GetType(System.Int64))
                .PrimaryKey = New DataColumn() {.Columns("VOL_GROUP")}
            End With


            With .Tables.Add("ARTWALMO")
                .Columns.Add("WALMART_STYLE", GetType(System.String))
                .Columns.Add("CUST_STORE_NO", GetType(System.String))
                .Columns.Add("CUST_STORE_NAME", GetType(System.String))
                .Columns.Add("CUST_DC_NO", GetType(System.String))
                .Columns.Add("ORDR_QTY", GetType(System.Int64))
                .Columns.Add("ORDR_QTY_ORIG", GetType(System.Int64))
                .PrimaryKey = New DataColumn() {.Columns("WALMART_STYLE"), .Columns("CUST_STORE_NO")}
            End With

            With .Tables.Add("ARTWALMR")
                .Columns.Add("WALMART_STYLE", GetType(System.String))
                .Columns.Add("CUST_DC_NO", GetType(System.String))
                .PrimaryKey = New DataColumn() {.Columns("WALMART_STYLE"), .Columns("CUST_DC_NO")}
            End With

            Create_Relation("ARTWALMR", "ARTWALMO", "WALMART_STYLE,CUST_DC_NO")
            With .Tables("ARTWALMR")
                .Columns.Add("ORDR_QTY", GetType(System.Int64), "SUM(CHILD.ORDR_QTY)")
                .Columns.Add("ORDR_QTY_ORIG", GetType(System.Int64), "SUM(CHILD.ORDR_QTY_ORIG)")
                .Columns.Add("DIFF", GetType(System.Int64), "ORDR_QTY - ORDR_QTY_ORIG")
                .Columns.Add("STRS", GetType(System.Int64), "COUNT(CHILD.CUST_STORE_NO)")
                .Columns.Add("CASES", GetType(System.Int64), "ORDR_QTY/" & CStr(INNERS_PER_CASE))
                .Columns.Add("OVER", GetType(System.Int64), "ORDR_QTY - CASES * " & CStr(INNERS_PER_CASE))
                .Columns.Add("CPR", GetType(System.Int64), "-1 * OVER")
            End With
        End With

        grdARTWALMO.DataSource = dst.Tables("ARTWALMO")
        grdARTWALMR.DataSource = dst.Tables("ARTWALMR")
        grdARTWALMX.DataSource = dst.Tables("ARTWALMX")
        grdARTWALMY.DataSource = dst.Tables("ARTWALMY")
        grdARTWALMS.DataSource = dst.Tables("ARTWALMS")
        grdARTWALMI.DataSource = dst.Tables("ARTWALMI")
        grdARTWALMV.DataSource = dst.Tables("ARTWALMV")


        With grdARTWALMX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "ORDR_QTY" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"CUST_STORE_NO", "CUST_STORE_NAME"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If

                If New String() {"SUP_STR", "SUP_XIT", "SUP_WHS", "SUP_OPO", "SUP_TOT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"SLS_LIF", "SLS_PTD"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"WKS_SUP", "SLS_FACTOR", "ORD_FACTOR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY_CALC", "ORDR_QTY_ROUND", "ORDR_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightCyan
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If

                If gcol.Key = "SUP_TOT" Then
                    gcol.CellAppearance.BackColor = gcol.Header.Appearance.BackColor2
                End If
            Next
        End With

        With grdARTWALMV.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "QTY" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        With grdARTWALMS.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"STORE_NBR"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If
                If New String() {"SLS_LIF", "SLS_PTD"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"SLS_FACTOR", "ORD_FACTOR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        With grdARTWALMY.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"CUST_DC_NO"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If

                If New String() {"SUP_TOT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"SLS_LIF", "SLS_PTD"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"WKS_SUP", "SLS_FACTOR", "ORD_FACTOR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY_OFF", "CASES", "ORDR_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightCyan
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        With grdARTWALMI.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"WALMART_STYLE", "STYLE_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Fixed = True
                End If

                If New String() {"MIN_QTY", "MAX_QTY", "PO_QTY"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightCyan
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            .Columns("COLOR_CODE").Hidden = True
        End With

        Create_Summary(grdARTWALMO, "CUST_STORE_NO", "Count")
        Create_Summary(grdARTWALMO, New String() {"ORDR_QTY"})

        Create_Summary(grdARTWALMR, "CUST_DC_NO", "Count")
        Create_Summary(grdARTWALMR, New String() {"ORDR_QTY", "ORDR_QTY_ORIG", "STRS", "CASES", "OVER", "CPR"})

        Create_Summary(grdARTWALMI, "WALMART_STYLE", "Count")
        Create_Summary(grdARTWALMI, New String() {"ORDR_QTY", "PO_QTY"})

        Create_Summary(grdARTWALMX, "CUST_STORE_NO", "Count")
        Create_Summary(grdARTWALMX, New String() {"SEL", "SLS_PTD", "SLS_LIF", "SUP_STR", "SUP_XIT", "SUP_WHS", "SUP_OPO", "SUP_TOT", "ORDR_QTY_CALC", "ORDR_QTY"})

        Create_Summary(grdARTWALMY, "CUST_DC_NO", "Count")
        Create_Summary(grdARTWALMY, New String() {"SLS_PTD", "SLS_LIF", "STORES", "STORES_WITH_SALES", "SUP_TOT", "ORDR_QTY", "CASES", "ORDR_QTY_OFF"})
        Create_Summary(grdARTWALMY, "SLS_AVG", "Custom")

        spl.Panel1Collapsed = True

        Populate_Volume_Groups()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Plan Orders"
                If dst.Tables("ARTWALMS").Rows.Count = 0 Or dst.Tables("ARTWALMI").Rows.Count = 0 Then
                    EMsg &= vbCr & "No Stores or Items loaded to Plan Orders" _
                        & vbCr & " - You must first Load Sales in order to Plan Orders"
                Else
                    Dim DC_TYPE As String = IIf(optDC.Value = "R", "REGIONAL_DC", "SPECIALTY_DC")
                    If dst.Tables("ARTWALMS").Select("SALES_IND = '1' and ISNULL(" & DC_TYPE & ",'')=''").Length <> 0 Then

                        If Not chkSkipNoDC.Checked Then
                            EMsg &= vbCr & "Some Stores with Sales are Missing their associated DC Code" _
                                & vbCr & " - You must first Update Stores & DCs"

                            Using F As New ASFMSGBF
                                Dim DT As DataTable = New DataView(dst.Tables("ARTWALMS"), "SALES_IND = '1' and ISNULL(" & DC_TYPE & ",'')=''", "", DataViewRowState.CurrentRows).ToTable
                                F.Show_grd(DT, Me, "Stores without DCs", "Stores without " & optDC.Text & " DC")
                            End Using
                        End If

                    End If
                    If dst.Tables("ARTWALMI").Select("ISNULL(STYLE_CODE,'')=''").Length <> 0 Then
                        EMsg &= vbCr & "Some Walmart Items are not mapped to Styles" _
                            & vbCr & " - You must first add Walmart Styles to the Customer Style Cross-Reference"
                    End If
                    If dst.Tables("ARTWALMI").Select("ISNULL(STYLE_CODE,'')<>'' and (ISNULL(CARTON_PACK_QTY,0)=0 or ISNULL(INNER_PACK_QTY,0)=0 or ISNULL(CARTON_PACK_QTY,0) < ISNULL(INNER_PACK_QTY,0))").Length <> 0 Then
                        EMsg &= vbCr & "Some Styles do Not have Valid Case or Inner Pack Qtys" _
                            & vbCr & " - You must first update the Style Master with valid Case and Inner Pack Qtys"
                    End If

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
            Case "Load Sales"
                Mode_Settings(False)

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Importing Wal-Mart Sales from Excel Workbook")
                Try
                    Load_Sales()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Trying to Import Wal-Mart Sales from Excel Workbook")
                End Try
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Load Order"

                ' Mode_Settings(False)

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Importing Wal-Mart Order (Simple Format) from Excel Workbook")
                Try
                    Load_Order()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Trying to Import Wal-Mart Order from Excel Workbook")
                End Try
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Plan Orders"
                Load_ARTWALMY()
                Mode_Settings(True)
                working_with_imported_order = False

            Case "Load Stores"
                Mode_Settings(False)

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Importing Wal-Mart Stores from Excel Workbook")
                Try
                    Load_Stores()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Trying to Import Wal-Mart Stores from Excel Workbook")
                End Try
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Update"

            Case "Cancel"
                Mode_Settings(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load Stores").Settings.Enabled = not_iScreenMode
                    .Items("Load Sales").Settings.Enabled = not_iScreenMode
                    .Items("Load Order").Settings.Enabled = not_iScreenMode
                    .Items("Plan Orders").Settings.Enabled = not_iScreenMode

                    .Items("Load Stores").Visible = Not ScreenMode
                    .Items("Load Sales").Visible = Not ScreenMode
                    .Items("Load Order").Visible = Not ScreenMode

                    .Items("Update").Visible = ScreenMode
                    .Items("Cancel").Visible = ScreenMode
                End With
                .Groups("Parameters").Visible = ScreenMode
                .Groups("Plan Orders").Visible = False
                .Groups("Volume Groups").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabPlan.Visible = ScreenMode
        SplitContainer1.Visible = False
        SplitContainer2.Visible = False

        With grdARTWALMI.DisplayLayout.Bands(0)
            .Columns("MIN_QTY").Hidden = Not ScreenMode
            .Columns("MAX_QTY").Hidden = Not ScreenMode
            .Columns("ORDR_QTY").Hidden = Not ScreenMode
            .Columns("PO_QTY").Hidden = Not ScreenMode
        End With

        With grdARTWALMS.DisplayLayout.Bands(0)
            .Columns("VOL_GROUP").Hidden = Not ScreenMode
            .Columns("VOL_QTY").Hidden = Not ScreenMode
            .Columns("SLS_LIF").Hidden = Not ScreenMode
            .Columns("SLS_PTD").Hidden = Not ScreenMode
            .Columns("SLS_FACTOR").Hidden = Not ScreenMode
            .Columns("ORD_FACTOR").Hidden = Not ScreenMode
        End With

        If ScreenMode Then
            grdARTWALMI.Parent = splPlanItemDC.Panel1
            grdARTWALMS.Parent = tabPlan.Tabs("Summary by Store").TabPage
            grdARTWALMS.Text = "Stores Order Recap by Style"

        Else
            Clear_Record()

            tabPlan.SelectedTab = tabPlan.Tabs("Order Planning Worksheet")

            grdARTWALMI.Parent = splSalesStoreItem.Panel2
            grdARTWALMS.Parent = splSalesStoreItem.Panel1
            grdARTWALMS.Text = "Stores with Sales"
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"ARTWALMX", "ARTWALMY", "ARTWALM2", "ARTWALMS", "ARTWALMI"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Clear_ARTWALMS()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        ' Sort_grdColumns(grdPOTSHIP3, "PO_SHIPMENT_LNO")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Stop
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Preparing " & Me.Text)

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report("PORWREC2")
        Print_Report_End()

        '    Me.Cursor = Cursors.Default
        '    ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PO_SHIPMENT_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTWALMX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSales, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdARTWALMS, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdARTWALMI, "B", "Replace Order with Spreadsheet")
        Load_Popup_Menu(grdARTWALMV, "B", "Set as Minimum Volume Group")
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdPOTORDRX"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked

            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Replace Order with Spreadsheet"
                Load_Order_Replacement()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Set as Minimum Volume Group"
                VOL_GROUP_min = grd.ActiveRow.Cells("VOL_GROUP").Value
                grdARTWALMV.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

    Sub Load_ARTWALMY()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")


        Dim MIN_FACTORs() As Decimal
        Dim QTYs() As Int32
        Dim STOREs() As Int32
        Dim VOL_GROUPs() As String
        Dim VOL_GROUP_count As Int64 = dst.Tables("ARTWALMV").Rows.Count
        ReDim MIN_FACTORs(VOL_GROUP_count - 1)
        ReDim QTYs(VOL_GROUP_count - 1)
        ReDim STOREs(VOL_GROUP_count - 1)
        ReDim VOL_GROUPs(VOL_GROUP_count - 1)
        Dim v As Integer = -1
        For Each rowARTWALMV As DataRow In dst.Tables("ARTWALMV").Select("", "MIN_FACTOR")
            v += 1
            MIN_FACTORs(v) = Val(rowARTWALMV.Item("MIN_FACTOR") & "")
            QTYs(v) = Val(rowARTWALMV.Item("QTY") & "")
            VOL_GROUPs(v) = rowARTWALMV.Item("VOL_GROUP")

        Next

        EnforceConstraints(False)

        Dim WKS As Int64 = numPeriod1Weeks.Value
        grdARTWALMX.DisplayLayout.Bands(0).Columns("SLS_PTD").Header.Caption = numPeriod1Weeks.Value & "Wks"
        grdARTWALMY.DisplayLayout.Bands(0).Columns("SLS_PTD").Header.Caption = numPeriod1Weeks.Value & "Wks"
        grdARTWALMS.DisplayLayout.Bands(0).Columns("SLS_PTD").Header.Caption = numPeriod1Weeks.Value & "Wks"

        Load_ARTWALMX()

        For Each rowARTWALMS As DataRow In dst.Tables("ARTWALMS").Select("SALES_IND = '1'")
            Dim CUST_STORE_NO As String = rowARTWALMS.Item("STORE_NBR")
            Dim sqlw As String = "CUST_STORE_NO = '" & CUST_STORE_NO & "'"
            Dim SLS_LIF As Int64 = Val(dst.Tables("ARTWALMX").Compute("SUM(SLS_LIF)", sqlw) & "")
            Dim SLS_PTD As Int64 = Val(dst.Tables("ARTWALMX").Compute("SUM(SLS_PTD)", sqlw) & "")
            rowARTWALMS.Item("SLS_LIF") = SLS_LIF
            rowARTWALMS.Item("SLS_PTD") = SLS_PTD
        Next

        Dim SLS_LIF_TOTAL As Int64 = Val(dst.Tables("ARTWALMS").Compute("SUM(SLS_LIF)", "SLS_LIF > 0") & "")
        Dim SLS_PTD_TOTAL As Int64 = Val(dst.Tables("ARTWALMS").Compute("SUM(SLS_PTD)", "SLS_PTD > 0") & "")
        Dim STR_LIF_TOTAL As Int64 = Val(dst.Tables("ARTWALMS").Compute("COUNT(STORE_NBR)", "SLS_LIF > 0") & "")
        Dim STR_PTD_TOTAL As Int64 = Val(dst.Tables("ARTWALMS").Compute("COUNT(STORE_NBR)", "SLS_PTD > 0") & "")
        SLS_AVG = IIf(STR_LIF_TOTAL = 0, 0, SLS_LIF_TOTAL / STR_LIF_TOTAL)

        For Each rowARTWALMS As DataRow In dst.Tables("ARTWALMS").Select("SALES_IND = '1'")
            Dim CUST_STORE_NO As String = rowARTWALMS.Item("STORE_NBR")
            Dim sqlw As String = "CUST_STORE_NO = '" & CUST_STORE_NO & "'"

            Dim SLS_PTD As Int64 = Val(rowARTWALMS.Item("SLS_LIF") & "")
            Dim SLS_FACTOR As Decimal = SLS_PTD / SLS_AVG
            If SLS_FACTOR < 0 Then SLS_FACTOR = 0

            For v = 0 To VOL_GROUP_count - 1
                If v = VOL_GROUP_count - 1 OrElse SLS_FACTOR < MIN_FACTORs(v + 1) Then Exit For
            Next
            Dim VOL_GROUP As String = VOL_GROUPs(v)

            rowARTWALMS.Item("VOL_GROUP") = VOL_GROUP
         '   Dim VOL_QTY As Int64 = QTYs(v)
            STOREs(v) += 1
            '  rowARTWALMS.Item("VOL_QTY") = VOL_QTY
            rowARTWALMS.Item("SLS_FACTOR") = SLS_FACTOR

            For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select(sqlw)
                rowARTWALMX.Item("VOL_GROUP") = VOL_GROUP
                'rowARTWALMX.Item("VOL_QTY") = VOL_QTY
                rowARTWALMX.Item("SLS_FACTOR") = SLS_FACTOR
            Next
        Next

        For v = 0 To VOL_GROUP_count - 1
            Dim VOL_GROUP As String = VOL_GROUPs(v)
            Dim rowARTWALMV As DataRow = dst.Tables("ARTWALMV").Rows.Find(VOL_GROUP)
            rowARTWALMV.Item("STORES") = STOREs(v)
        Next

        dst.Tables("ARTWALMY").Rows.Clear()

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ARTWALMX"), New String() {"WALMART_STYLE", "CUST_DC_NO"}).Rows
            Dim WALMART_STYLE As String = row.Item("WALMART_STYLE") & ""
            Dim CUST_DC_NO As String = row.Item("CUST_DC_NO") & ""
            dst.Tables("ARTWALMY").Rows.Add(New String() {WALMART_STYLE, CUST_DC_NO})
        Next

        EnforceConstraints(True)

        Calculate()

        Dim cols As New List(Of String)

        dst.Tables("ARTWALMS").Columns.Add("QTY_TOTAL", GetType(System.Int64))
        cols.Add("QTY_TOTAL")
        grdARTWALMS.DisplayLayout.Bands(0).Columns("QTY_TOTAL").Header.Caption = "Total"

        Dim TOTAL As String = ""
        For Each row As DataRow In dst.Tables("ARTWALMI").Select("", "WALMART_STYLE")
            Dim WALMART_STYLE As String = row.Item("WALMART_STYLE")
            dst.Tables("ARTWALMS").Columns.Add("QTY_" & WALMART_STYLE, GetType(System.Int64))
            grdARTWALMS.DisplayLayout.Bands(0).Columns("QTY_" & WALMART_STYLE).Header.Caption = WALMART_STYLE
            TOTAL &= "+ISNULL(QTY_" & WALMART_STYLE & ",0)"
            cols.Add("QTY_" & WALMART_STYLE)
        Next
        dst.Tables("ARTWALMS").Columns("QTY_TOTAL").Expression = Mid(TOTAL, 2)

        For Each COLUMN_NAME As String In cols
            With grdARTWALMS.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .Header.Appearance.BackColor2 = Drawing.Color.Gainsboro
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Hidden = False
                .Width = 90
                .CellAppearance.TextHAlign = HAlign.Right
                .Format = "#,##0"
            End With
            Create_Summary(grdARTWALMS, COLUMN_NAME)
        Next

        Setup_grdARTWALMY()
        Setup_grdARTWALMX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdCalculate_Click(sender As System.Object, e As System.EventArgs) Handles cmdReCalculate.Click
        Calculate()
    End Sub

    Sub Setup_grdARTWALMY()

        If grdARTWALMI.ActiveRow Is Nothing Then
            grdARTWALMY.Visible = False
            Exit Sub
        Else
            grdARTWALMY.Visible = True
        End If
        Dim WALMART_STYLE As String = grdARTWALMI.ActiveRow.Cells("WALMART_STYLE").Value
        Dim sqlw As String = "WALMART_STYLE = '" & WALMART_STYLE & "'"
        Dim dvw As DataView = DirectCast(grdARTWALMY.DataSource, DataTable).DefaultView
        dvw.RowFilter = sqlw

        Sort_grdColumns(grdARTWALMY, "CUST_DC_NO")

        grdARTWALMY.Text = "DC Summary for Walmart Style " & WALMART_STYLE
    End Sub

    Sub Setup_grdARTWALMX()

        If grdARTWALMY.ActiveRow Is Nothing Then
            grdARTWALMX.Visible = False
            Exit Sub
        Else
            grdARTWALMX.Visible = True
        End If
        Dim CUST_DC_NO As String = grdARTWALMY.ActiveRow.Cells("CUST_DC_NO").Value
        Dim WALMART_STYLE As String = grdARTWALMY.ActiveRow.Cells("WALMART_STYLE").Value
        Dim sqlw As String = "CUST_DC_NO = '" & CUST_DC_NO & "' and WALMART_STYLE = '" & WALMART_STYLE & "'"
        Dim dvw As DataView = DirectCast(grdARTWALMX.DataSource, DataTable).DefaultView
        dvw.RowFilter = sqlw

        Sort_grdColumns(grdARTWALMX, "CUST_STORE_NO")

        grdARTWALMX.Text = "Store Order Planning for DC " & CUST_DC_NO & ", Walmart Style " & WALMART_STYLE
    End Sub

    Private Sub grdARTWALMY_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdARTWALMY.AfterRowActivate
        If ScreenMode Then Setup_grdARTWALMX()
    End Sub


    Private Sub grdARTWALMX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTWALMX.InitializeRow
        If Val(e.Row.Cells("ORDR_QTY").Value & "") <> Val(e.Row.Cells("ORDR_QTY_ROUND").Value & "") Then
            e.Row.Cells("ORDR_QTY").Appearance.BackColor = Drawing.Color.LightGreen
        Else
            e.Row.Cells("ORDR_QTY").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Load_Sales()


        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME = "" Then Exit Sub

        Dim f As String = "C:\Users\wjz\Desktop\rtl.xlsx"
        f = FILENAME

        ASCMAIN1.Progress("Now Loading Worksheet into Memory")

        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)
        Dim g As New GemBox.Spreadsheet.ExcelFile
        ' g.LoadXlsx(f, GemBox.Spreadsheet.XlsxOptions.None)
        g = GemBox.Spreadsheet.ExcelFile.Load(f, New GemBox.Spreadsheet.XlsLoadOptions With {.PreserveOptions = GemBox.Spreadsheet.XlsOptions.None})


        Dim ws As GemBox.Spreadsheet.ExcelWorksheet = Nothing

        For i As Integer = 0 To g.Worksheets.Count - 1
            If g.Worksheets(i).Name = "Fashion Data Tab" Then
                ws = g.Worksheets(i)
                Exit For
            End If
        Next

        If ws Is Nothing Then
            MsgBox("Cannot find Fashion Data Tab in selected Workbook")
            Exit Sub
        End If


        ASCMAIN1.Progress("Now Creating Working Datatable")

        Dim t As Integer = 0
        Do

            If ws.Rows(t).Cells(0).Value = "Store Nbr" Then
                Exit Do
            End If
            t += 1
            If t > 100 Then
                MsgBox("Cannot find expected heading")
                Exit Sub
            End If
        Loop


        Dim tbl As New DataTable
        Dim c As Integer = 0
        Do
            Dim col As String = ws.Rows(t).Cells(c).Value & ""
            If col = "" Then
                Exit Do
            Else
                If col.StartsWith("Range") Then
                    tbl.Columns.Add(col, GetType(System.Int64))
                Else
                    tbl.Columns.Add(col)
                End If
                c += 1
            End If
        Loop

        Do
            t += 1
            If ws.Rows(t).Cells(0).Value & "" = "" Then
                Exit Do
            Else
                Dim row As DataRow = tbl.NewRow
                For i As Integer = 0 To c - 1
                    row.Item(i) = ws.Rows(t).Cells(i).Value
                Next
                row.Item(0) = CStr(row.Item(0) & "").PadLeft(4, "0")
                tbl.Rows.Add(row)
            End If
        Loop

        grdSales.DataSource = Nothing
        grdSales.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grdSales.DataSource = tbl
        ASCMAIN1.grdInitializeLayout(grdSales, Me)

        Fill_Records("ARTWALMS")
        For Each row As DataRow In ASCDATA1.SelectDistinct(tbl, New String() {"Store Nbr"}).Select("")
            Dim STORE_NBR As String = row.Item("Store Nbr")
            Dim rowARTWALMS As DataRow = dst.Tables("ARTWALMS").Rows.Find(STORE_NBR)
            If rowARTWALMS Is Nothing Then
                rowARTWALMS = dst.Tables("ARTWALMS").Rows.Add(New String() {STORE_NBR})
            End If
            rowARTWALMS.Item("SALES_IND") = "1"
        Next

        Dim EMsg As String = ""
        Dim DC_TYPE As String = IIf(optDC.Value = "R", "REGIONAL_DC", "SPECIALTY_DC")
        Dim rows() As DataRow = dst.Tables("ARTWALMS").Select("ISNULL(" & DC_TYPE & ",'') = ''")
        If rows.Length > 0 Then
            EMsg &= vbCr & "There are " & CStr(rows.Length) & " stores with Sales History" _
                & vbCr & " that are not in Store Master (ie Store " & rows(0).Item(0) & ")"
        End If

        Sort_grdColumns(grdARTWALMS, "STORE_NBR")


        dst.Tables("ARTWALMI").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(tbl, New String() {"Item Nbr"}).Select("")
            Dim ITEM_NBR As String = row.Item("Item Nbr")

            If dst.Tables("ARTWALMI").Rows.Find(ITEM_NBR) Is Nothing Then
                Dim rowARTWALMI As DataRow = dst.Tables("ARTWALMI").Rows.Add(New String() {ITEM_NBR})
                Dim rowSOTCSTY1 As DataRow = LookUp("SOTCSTY1", New String() {"WALMART", ITEM_NBR})
                If rowSOTCSTY1 IsNot Nothing Then
                    rowARTWALMI.Item("STYLE_CODE") = rowSOTCSTY1.Item("STYLE_CODE")
                    rowARTWALMI.Item("COLOR_CODE") = rowSOTCSTY1.Item("COLOR_CODE")
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTCSTY1.Item("STYLE_CODE"))
                    rowARTWALMI.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                    rowARTWALMI.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")
                    rowARTWALMI.Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
                End If
            End If
        Next

        Sort_grdColumns(grdARTWALMI, "WALMART_STYLE")

        If grdSales.DisplayLayout.Bands(0).Summaries.Count > 0 Then
            grdSales.DisplayLayout.Bands(0).Summaries.Clear()
        End If
        '  Create_Summary(grdSales, "Store Nbr")

        SplitContainer2.Visible = True

        UltraExplorerBar1.Groups("Plan Orders").Visible = True

    End Sub

    Sub Load_Order_Replacement()

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME = "" Then Exit Sub

        Dim f As String = "C:\Users\wjz\Desktop\stores.xlsx"
        f = FILENAME

        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)
        Dim g As New GemBox.Spreadsheet.ExcelFile
        'g.LoadXlsx(f, GemBox.Spreadsheet.XlsxOptions.None)
        g = GemBox.Spreadsheet.ExcelFile.Load(f, New GemBox.Spreadsheet.XlsLoadOptions With {.PreserveOptions = GemBox.Spreadsheet.XlsOptions.None})
        Dim ws As GemBox.Spreadsheet.ExcelWorksheet = g.Worksheets(0)

        Dim t As Integer = 0
        'Do

        '    If ws.Rows(t).Cells(0).Value = "Store Nbr" And ws.Rows(t).Cells(1).Value = "Item Nbr" And ws.Rows(t).Cells(2).Value = "Qty" Then
        '        Exit Do
        '    End If
        '    t += 1
        '    If t > 100 Then
        '        MsgBox("Did not find expected headings")
        '        Exit Sub
        '    End If
        'Loop

        Dim WALMART_ITEMs As New Dictionary(Of Integer, String)

        If ws.Rows(t).Cells(0).Value & "" = "Store Nbr" And ws.Rows(t).Cells(1).Value & "" <> "" Then
            Dim c As Integer = 0
            Do While ws.Rows(t).Cells(c + 1).Value & "" <> ""
                c += 1
                Dim WALMART_STYLE As String = ws.Rows(t).Cells(c).Value

                Dim rowARTWALMI As DataRow = dst.Tables("ARTWALMI").Rows.Find(WALMART_STYLE)
                If rowARTWALMI Is Nothing Then
                    MsgBox("Cannot find Walmart Style " & WALMART_STYLE, _
                             MsgBoxStyle.OkOnly, _
                             "Walmart Styles in Spreadsheet must Exactly Match Styles in Sales File")
                    Exit Sub
                Else
                    WALMART_ITEMs.Add(c, WALMART_STYLE)
                End If

                ' maybe we should take a run down ARTWALMI and make sure all styles are represented
            Loop

        Else
            MsgBox("Did not find expected headings")
            Exit Sub
        End If
       

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Replacing Order Qtys")

        dst.Tables("ARTWALMY").Columns("ORDR_QTY").Expression = ""
        dst.Tables("ARTWALMI").Columns("ORDR_QTY").Expression = ""

        Dim WALMART_STYLEs As New List(Of String)
        Dim CUST_STORE_NOs_not_found As New List(Of String)
        Dim WALMART_STYLEs_not_found As New List(Of String)
        Do
            t += 1
            If ws.Rows(t).Cells(0).Value & "" = "" Then
                Exit Do
            Else
                Dim CUST_STORE_NO As String = ws.Rows(t).Cells(0).Value
                CUST_STORE_NO = CUST_STORE_NO.PadLeft(4, "0")

                For Each iw As Integer In WALMART_ITEMs.Keys

                    Dim WALMART_STYLE As String = WALMART_ITEMs(iw)

                    Dim rowARTWALMI As DataRow = dst.Tables("ARTWALMI").Rows.Find(WALMART_STYLE)
                    If Not WALMART_STYLEs.Contains(WALMART_STYLE) Then
                        For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select("WALMART_STYLE = '" & WALMART_STYLE & "'")
                            rowARTWALMX.Item("ORDR_QTY") = 0
                        Next
                        WALMART_STYLEs.Add(WALMART_STYLE)
                    End If
                    Dim QTY As String = Val(ws.Rows(t).Cells(iw).Value & "")

                    Dim rowARTWALMS As DataRow = dst.Tables("ARTWALMS").Rows.Find(CUST_STORE_NO)
                    If rowARTWALMS Is Nothing Then
                        If Not CUST_STORE_NOs_not_found.Contains(CUST_STORE_NO) Then
                            CUST_STORE_NOs_not_found.Add(CUST_STORE_NO)
                            MsgBox("Cannot find Store No " & CUST_STORE_NO, MsgBoxStyle.OkOnly, "Any Records with this value will be skipped")
                        End If
                    Else
                        Dim DC_TYPE As String = IIf(optDC.Value = "R", "REGIONAL_DC", "SPECIALTY_DC")
                        Dim CUST_DC_NO As String = rowARTWALMS.Item(DC_TYPE)
                        Dim rowARTWALMX As DataRow = dst.Tables("ARTWALMX").Rows.Find(New String() {WALMART_STYLE, CUST_DC_NO, CUST_STORE_NO})
                        If rowARTWALMX Is Nothing Then
                            rowARTWALMX = dst.Tables("ARTWALMX").Rows.Add(New String() {WALMART_STYLE, CUST_DC_NO, CUST_STORE_NO})
                        End If

                        rowARTWALMX.Item("ORDR_QTY") = QTY
                    End If
                    ASCMAIN1.Progress("-", CUST_STORE_NO & ":" & WALMART_STYLE)
                Next
            End If
        Loop

        dst.Tables("ARTWALMY").Columns("ORDR_QTY").Expression = "SUM(CHILD.ORDR_QTY)"
        dst.Tables("ARTWALMI").Columns("ORDR_QTY").Expression = "SUM(CHILD.ORDR_QTY)"

        grdARTWALMI.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

        MsgBox("Import of Replacement Order Qtys Completed Successfully", _
               MsgBoxStyle.OkOnly, "Verification")

        working_with_imported_order = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Load_Stores()

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME = "" Then Exit Sub

        Dim f As String = "C:\Users\wjz\Desktop\stores.xlsx"
        f = FILENAME

        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)
        Dim g As New GemBox.Spreadsheet.ExcelFile
        'g.LoadXlsx(f, GemBox.Spreadsheet.XlsxOptions.None)
        g = GemBox.Spreadsheet.ExcelFile.Load(f, New GemBox.Spreadsheet.XlsLoadOptions With {.PreserveOptions = GemBox.Spreadsheet.XlsOptions.None})
        Dim ws As GemBox.Spreadsheet.ExcelWorksheet = g.Worksheets(0)

        If ws.Name <> "Wal-Mart Stores" Then
            MsgBox("Cannot find Wal-Mart Stores Sheet (should be 1st sheet in Workbook)")
            Exit Sub
        End If

        If g.Worksheets.Count < 3 OrElse g.Worksheets(2).Name <> "Wal-Mart DC Receiving" Then
            MsgBox("Cannot find Wal-Mart DCs Sheet (should be 3rd sheet in Workbook)")
            Exit Sub
        End If

        Load_Walmart(ws, grdStores, "Stores", "ARTWALM2")
        ws = g.Worksheets(2)
        Load_Walmart(ws, grdDCs, "DCs", "ARTWALM3")

        SplitContainer1.Visible = True
    End Sub

    Sub Load_Walmart( _
                    ws As GemBox.Spreadsheet.ExcelWorksheet, _
                    grd As UltraWinGrid.UltraGrid, _
                    legend As String, _
                    TABLE_NAME As String)

        Dim icolStore As Integer = 0
        If TABLE_NAME = "ARTWALM3" Then icolStore = 1

        Dim t As Integer = 0
        Do

            If ws.Rows(t).Cells(icolStore).Value = "Store Nbr" Then
                Exit Do
            End If
            t += 1
            If t > 100 Then
                MsgBox(legend & " - Cannot find expected heading")
                Exit Sub
            End If
        Loop

        Dim tbl As New DataTable
        Dim c As Integer = 0
        Do
            Dim col As String = ws.Rows(t).Cells(c).Value & ""
            If col = "" Then
                Exit Do
            Else
                tbl.Columns.Add(col)
                c += 1
            End If
        Loop

        Do
            t += 1
            If ws.Rows(t).Cells(0).Value & "" = "" Then
                Exit Do
            Else
                Dim row As DataRow = tbl.NewRow
                For i As Integer = 0 To c - 1
                    row.Item(i) = ws.Rows(t).Cells(i).Value
                Next
                tbl.Rows.Add(row)

                ASCMAIN1.Progress("-", ws.Rows(t).Cells(icolStore).Value)
            End If
        Loop

        grd.DataSource = Nothing
        grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grd.DataSource = tbl
        ASCMAIN1.grdInitializeLayout(grd, Me)

        ASCMAIN1.Progress("Now Transferring Imported Data to Data Table")

        Dim PCOL As String = ""
        If TABLE_NAME = "ARTWALM2" Then
            PCOL = "PHONE_NUMBER"
        Else
            PCOL = "TELEPHONE"
        End If

        dst.Tables(TABLE_NAME).Rows.Clear()
        For Each row As DataRow In tbl.Select("")
            Dim rowARTWALMX As DataRow = dst.Tables(TABLE_NAME).NewRow
            rowARTWALMX.ItemArray = row.ItemArray
            'rowARTWALM2.Item("PHONE_NUMBER") = Replace(rowARTWALM2.Item("PHONE_NUMBER") & "", " ", "")
            For i As Integer = 0 To tbl.Columns.Count - 1
                Dim DATA_VALUE As String = Trim(row.Item(i) & "")
                If DATA_VALUE <> "" Then rowARTWALMX.Item(i) = DATA_VALUE
            Next
            rowARTWALMX.Item("STORE_NBR") = Format(Val(rowARTWALMX.Item("STORE_NBR") & ""), "0000")

            rowARTWALMX.Item(PCOL) = Replace(Replace(rowARTWALMX.Item(PCOL) & "", " ", ""), "-", "")

            ASCMAIN1.Progress("-", rowARTWALMX.Item("STORE_NBR"))
            dst.Tables(TABLE_NAME).Rows.Add(rowARTWALMX)
        Next

        ASCMAIN1.Progress("Now Updating Database")

        BeginTrans()
        ASCMAIN1.sql = "Delete from " & TABLE_NAME
        ASCDATA1.ExecuteSQL()
        Update_Record_TDA(TABLE_NAME)
        CommitTrans()

        MsgBox("Import of " & CStr(dst.Tables(TABLE_NAME).Rows.Count) & " " & legend & " Completed Successfully", _
               MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub Load_Order()

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME = "" Then Exit Sub

        Dim f As String = "C:\Users\wjz\Desktop\stores.xlsx"
        f = FILENAME

        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)
        Dim g As New GemBox.Spreadsheet.ExcelFile
        'g.LoadXlsx(f, GemBox.Spreadsheet.XlsxOptions.None)
        g = GemBox.Spreadsheet.ExcelFile.Load(f, New GemBox.Spreadsheet.XlsLoadOptions With {.PreserveOptions = GemBox.Spreadsheet.XlsOptions.None})
        Dim ws As GemBox.Spreadsheet.ExcelWorksheet

        If g.Worksheets.Count < 1 OrElse g.Worksheets(0).Name <> "SSO Submission Form" Then
            MsgBox("Cannot find SSO Submission Form Sheet (should be 1st sheet in Workbook)")
            Exit Sub
        End If

        If g.Worksheets.Count < 3 OrElse g.Worksheets(2).Name <> "SSO Script" Then
            MsgBox("Cannot find SSO Script Sheet (should be 3rd sheet in Workbook)")
            Exit Sub
        End If

        ws = g.Worksheets(0)
        Load_Walmart_styles(ws, Nothing, "Styles", "ARTWALMI")

        ws = g.Worksheets(2)
        Load_Walmart_Order(ws, Nothing, "Orders", "ARTWALMO")


        grdARTWALMI.Parent = splCasePack.Panel1
        grdARTWALMI.DisplayLayout.Bands(0).Columns("CARTON_PACK_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
        '   grdARTWALMI.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        Sort_grdColumns(grdARTWALMO, "WALMART_STYLE,CUST_STORE_NO")
        Sort_grdColumns(grdARTWALMR, "WALMART_STYLE,CUST_DC_NO")

        SplitContainer1.Visible = False
        splCasePack.Visible = True
    End Sub

    Sub Load_Walmart_styles( _
                 ws As GemBox.Spreadsheet.ExcelWorksheet, _
                 grd As UltraWinGrid.UltraGrid, _
                 legend As String, _
                 TABLE_NAME As String)

        Dim icolStore As Integer = 0

        Dim t As Integer = 0
        Do

            If ws.Rows(t).Cells(icolStore).Value & "" = "Vendor Stock #" Then
                Exit Do
            End If
            t += 1
            If t > 100 Then
                MsgBox(legend & " - Cannot find expected heading")
                Exit Sub
            End If
        Loop

        Dim tbl As New DataTable
        Dim c As Integer = 0
        Do
            Dim col As String = ws.Rows(t).Cells(c).Value & ""
            If col = "" Or col = "Updated Vendor Pack Cost" Then
                Exit Do
            Else
                tbl.Columns.Add(col)
                c += 1
            End If
        Loop

        Do
            t += 1
            If ws.Rows(t).Cells(0).Value & "" = "" Then
                Exit Do
            Else
                Dim row As DataRow = tbl.NewRow
                For i As Integer = 0 To c - 1
                    row.Item(i) = ws.Rows(t).Cells(i).Value
                Next
                tbl.Rows.Add(row)

                ASCMAIN1.Progress("-", ws.Rows(t).Cells(icolStore).Value)
            End If
        Loop

        If grd IsNot Nothing Then
            grd.DataSource = Nothing
            grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grd.DataSource = tbl
            ASCMAIN1.grdInitializeLayout(grd, Me)
        End If

        ASCMAIN1.Progress("Now Transferring Imported Data to Data Table")

        dst.Tables(TABLE_NAME).Rows.Clear()
        For Each row As DataRow In tbl.Select("")
            Dim rowARTWALMI As DataRow = dst.Tables("ARTWALMI").NewRow
            rowARTWALMI.Item("STYLE_CODE") = row.Item("Vendor Stock #")
            rowARTWALMI.Item("STYLE_DESC") = row.Item("Item Description")
            rowARTWALMI.Item("INNER_PACK_QTY") = row.Item("Units per pack")
            rowARTWALMI.Item("CARTON_PACK_QTY") = 0
            rowARTWALMI.Item("WALMART_STYLE") = row.Item("Item Number")
            rowARTWALMI.Item("COLOR_CODE") = "AST"
            dst.Tables("ARTWALMI").Rows.Add(rowARTWALMI)
        Next
          
        'ASCMAIN1.Progress("Now Updating Database")

        'BeginTrans()
        'ASCMAIN1.sql = "Delete from " & TABLE_NAME
        'ASCDATA1.ExecuteSQL()
        'Update_Record_TDA(TABLE_NAME)
        'CommitTrans()

        'MsgBox("Import of " & CStr(dst.Tables(TABLE_NAME).Rows.Count) & " " & legend & " Completed Successfully", _
        '       MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub Load_Walmart_Order( _
                    ws As GemBox.Spreadsheet.ExcelWorksheet, _
                    grd As UltraWinGrid.UltraGrid, _
                    legend As String, _
                    TABLE_NAME As String)

        Dim icolStore As Integer = 1


        Dim t As Integer = 0
        Do

            If ws.Rows(t).Cells(icolStore).Value & "" = "STR #" Then
                Exit Do
            End If
            t += 1
            If t > 100 Then
                MsgBox(legend & " - Cannot find expected heading")
                Exit Sub
            End If
        Loop

        Dim tbl As New DataTable
        Dim c As Integer = 0
        Do
            Dim col As String = ws.Rows(t).Cells(c).Value & ""
            If col = "" Then
                Exit Do
            Else
                tbl.Columns.Add(col)
                c += 1
            End If
        Loop

        Do
            t += 1
            If ws.Rows(t).Cells(0).Value & "" = "" Then
                Exit Do
            Else
                Dim row As DataRow = tbl.NewRow
                For i As Integer = 0 To c - 1
                    row.Item(i) = ws.Rows(t).Cells(i).Value
                Next
                tbl.Rows.Add(row)

                ASCMAIN1.Progress("-", ws.Rows(t).Cells(icolStore).Value)
            End If
        Loop

        If grd IsNot Nothing Then
            grd.DataSource = Nothing
            grd.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grd.DataSource = tbl
            ASCMAIN1.grdInitializeLayout(grd, Me)
        End If

        ASCMAIN1.Progress("Now Transferring Imported Data to Data Table")

        EnforceConstraints(False)

        dst.Tables("ARTWALMR").Rows.Clear()

        ASCMAIN1.sql = "Select Distinct STORE_NBR, STORE_NAME, REGIONAL_DC, SPECIALTY_DC from ARTWALM2"
        Fill_Records("ARTWALMS", "", True, ASCMAIN1.sql)
        'dst.Tables("ARTWALMS").Rows.Clear()
        'Fill_Records("ARTWALM2")
        'For Each rowARTWALM2 As DataRow In dst.Tables("ARTWALM2").Select("")
        '    Dim rowARTWALMS As DataRow = dst.Tables("ARTWALMS").NewRow
        '    rowARTWALMS.Item("STORE_NBR") = rowARTWALM2.Item("STORE_NBR")
        '    rowARTWALMS.Item("REGIONAL_DC") = rowARTWALM2.Item("REGIONAL_DC")
        '    rowARTWALMS.Item("STORE_NAME") = rowARTWALM2.Item("STORE_NAME")
        '    dst.Tables("ARTWALMS").Rows.Add(rowARTWALMS)
        'Next

        dst.Tables(TABLE_NAME).Rows.Clear()
        For Each row As DataRow In tbl.Select("")
            Dim rowARTWALMO As DataRow = dst.Tables(TABLE_NAME).NewRow
            rowARTWALMO.Item("WALMART_STYLE") = row.Item(0)
            rowARTWALMO.Item("CUST_STORE_NO") = row.Item(1)
            rowARTWALMO.Item("ORDR_QTY") = row.Item(2)
            rowARTWALMO.Item("ORDR_QTY_ORIG") = row.Item(2)
            Dim CUST_DC_NO As String = "0000"
            Dim CUST_STORE_NAME As String = "?"
            Dim rowS As DataRow = dst.Tables("ARTWALMS").Rows.Find(Format(Val(row.Item(1) & ""), "0000"))
            If rowS IsNot Nothing Then
                CUST_DC_NO = rowS.Item("REGIONAL_DC")
                CUST_STORE_NAME = rowS.Item("STORE_NAME")
            Else
                Stop
            End If
            rowARTWALMO.Item("CUST_STORE_NAME") = CUST_STORE_NAME
            rowARTWALMO.Item("CUST_DC_NO") = CUST_DC_NO

            Dim rowARTWALMR As DataRow = dst.Tables("ARTWALMR").Rows.Find(New String() {row.Item(0), CUST_DC_NO})
            If rowARTWALMR Is Nothing Then
                rowARTWALMR = dst.Tables("ARTWALMR").NewRow
                rowARTWALMR.Item("WALMART_STYLE") = row.Item(0)
                rowARTWALMR.Item("CUST_DC_NO") = CUST_DC_NO
                dst.Tables("ARTWALMR").Rows.Add(rowARTWALMR)
            End If
            ASCMAIN1.Progress("-", rowARTWALMO.Item("CUST_STORE_NO"))
            dst.Tables(TABLE_NAME).Rows.Add(rowARTWALMO)
        Next

        EnforceConstraints(True)

        'ASCMAIN1.Progress("Now Updating Database")

        'BeginTrans()
        'ASCMAIN1.sql = "Delete from " & TABLE_NAME
        'ASCDATA1.ExecuteSQL()
        'Update_Record_TDA(TABLE_NAME)
        'CommitTrans()

        'MsgBox("Import of " & CStr(dst.Tables(TABLE_NAME).Rows.Count) & " " & legend & " Completed Successfully", _
        '       MsgBoxStyle.OkOnly, "Verification")

    End Sub


    Private Sub grdARTWALMI_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdARTWALMI.AfterRowActivate
        If ScreenMode Then Setup_grdARTWALMY()
    End Sub

    Sub Load_ARTWALMX()
        Dim TBL As DataTable = DirectCast(grdSales.DataSource, DataTable)

        Dim CUST_STORE_NO As String = ""
        Dim CUST_DC_NO As String = ""
        Dim CUST_STORE_NAME As String = ""

        Dim sqlw As String = ""
        For Each row As DataRow In TBL.Select(sqlw, "Store Nbr")
            If CUST_STORE_NO <> row.Item("Store Nbr") Then
                CUST_STORE_NO = row.Item("Store Nbr")
                Dim rowARTWALMS As DataRow = dst.Tables("ARTWALMS").Rows.Find(CUST_STORE_NO)
                If rowARTWALMS Is Nothing Then
                    CUST_DC_NO = ""
                    CUST_STORE_NAME = ""
                Else
                    Dim DC_TYPE As String = IIf(optDC.Value = "R", "REGIONAL_DC", "SPECIALTY_DC")
                    CUST_DC_NO = rowARTWALMS.Item(DC_TYPE) & ""
                    CUST_STORE_NAME = rowARTWALMS.Item("STORE_NAME") & ""
                End If
            End If

            If CUST_DC_NO = "" Then
                ' skip
            Else

                Dim rowARTWALMX As DataRow = dst.Tables("ARTWALMX").NewRow
                rowARTWALMX.Item("CUST_DC_NO") = CUST_DC_NO
                rowARTWALMX.Item("WALMART_STYLE") = row.Item("Item Nbr")
                rowARTWALMX.Item("CUST_STORE_NO") = CUST_STORE_NO
                rowARTWALMX.Item("CUST_STORE_NAME") = CUST_STORE_NAME

                rowARTWALMX.Item("SLS_LIF") = row.Item("Range 2 POS Qty")
                rowARTWALMX.Item("SLS_PTD") = row.Item("Range 1 POS Qty")
                rowARTWALMX.Item("SUP_STR") = row.Item("Range 1 Curr Str On Hand Qty")
                rowARTWALMX.Item("SUP_XIT") = row.Item("Range 1 Curr Str In Transit Qty")
                rowARTWALMX.Item("SUP_WHS") = row.Item("Range 1 Curr Str In Whse Qty")
                rowARTWALMX.Item("SUP_OPO") = row.Item("Range 1 Curr Str On Order Qty")
                rowARTWALMX.Item("TRAITED") = row.Item("Range 1 Curr Traited Store/Item Comb.")
                rowARTWALMX.Item("VALID") = row.Item("Range 1 Curr Valid Store/Item Comb.")

                If chkOnlyTraited.Checked And (rowARTWALMX.Item("TRAITED") & "" <> "1" Or rowARTWALMX.Item("VALID") & "" <> "1") Then
                Else
                    dst.Tables("ARTWALMX").Rows.Add(rowARTWALMX)
                End If
            End If
        Next
    End Sub

    Sub Clear_ARTWALMS()
        grdARTWALMS.DisplayLayout.Bands(0).Summaries.Clear()
        If dst.Tables("ARTWALMS").Columns.Contains("QTY_TOTAL") Then
            dst.Tables("ARTWALMS").Columns("ORD_FACTOR").Expression = ""
            dst.Tables("ARTWALMS").Columns.Remove("QTY_TOTAL")
        End If

        For I As Integer = dst.Tables("ARTWALMS").Columns.Count - 1 To 0 Step -1
            Dim dc As DataColumn = dst.Tables("ARTWALMS").Columns(I)
            If dc.ColumnName.StartsWith("QTY_") Then
                dst.Tables("ARTWALMS").Columns.Remove(dc)
            End If
        Next
        Create_Summary(grdARTWALMS, "STORE_NBR", "Count")
        Create_Summary(grdARTWALMS, New String() {"SALES_IND", "SLS_LIF", "SLS_PTD"})

    End Sub

    Private Sub cmdDCCases_Click(sender As System.Object, e As System.EventArgs) Handles cmdDCCases.Click

        Balance_DC_Cases()
        If chkProrate.Checked Then
            If dst.Tables("ARTWALMI").Select("PO_QTY <> 0 AND PO_QTY <> ORDR_QTY").Length <> 0 Then
                Balance_DC_Cases()
            End If
        End If


        grdARTWALMI.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

    End Sub

    Sub Balance_DC_Cases()

        If chkRoundUp.Checked And chkProrate.Checked Then
            MsgBox("Cannot Round Up when Prorating to PO Qty", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Re-Balancing Qtys by Case")
 
        For Each rowARTWALMI As DataRow In dst.Tables("ARTWALMI").Select("", "")

            Dim WALMART_STYLE As String = rowARTWALMI.Item("WALMART_STYLE")

            Dim CARTON_PACK_QTY As Int64 = Val(rowARTWALMI.Item("CARTON_PACK_QTY") & "")
            Dim INNER_PACK_QTY As Int64 = Val(rowARTWALMI.Item("INNER_PACK_QTY") & "")

            Dim MAX_QTY As Int64 = Val(rowARTWALMI.Item("MAX_QTY") & "") * INNER_PACK_QTY
            Dim MIN_QTY As Int64 = Val(rowARTWALMI.Item("MIN_QTY") & "") * INNER_PACK_QTY

            If Not chkMinMax.Checked Then
                MAX_QTY = 0
                MIN_QTY = 0
            End If


            Dim sqlI As String = "WALMART_STYLE = '" & WALMART_STYLE & "'"

            For Each rowARTWALMY As DataRow In dst.Tables("ARTWALMY").Select(sqlI & " and ORDR_QTY_OFF <> 0", "")
                Dim CUST_DC_NO As String = rowARTWALMY.Item("CUST_DC_NO")

                'If ASCMAIN1.Running_in_VS Then
                '    If WALMART_STYLE = "551003015" And CUST_DC_NO = "6010" Then Stop
                'End If

                Dim sqlwDC As String = " and CUST_DC_NO = '" & CUST_DC_NO & "'"
                Dim sqlw As String = sqlI & sqlwDC
                With dst.Tables("ARTWALMX")
                    For Each row As DataRow In .Select(sqlw & " and ISNULL(ORDR_QTY,0) <> 0 and ISNULL(ORDR_QTY,0) <> " & CStr(INNER_PACK_QTY), "")
                        Dim ORDR_QTY As Int64 = Val(row.Item("ORDR_QTY") & "")
                        If ORDR_QTY Mod INNER_PACK_QTY <> 0 Then
                            ORDR_QTY += INNER_PACK_QTY - (ORDR_QTY Mod INNER_PACK_QTY)
                            row.Item("ORDR_QTY") = ORDR_QTY
                        End If
                    Next
                End With
            Next

            Dim PO_QTY As Int64 = Val(rowARTWALMI.Item("PO_QTY") & "")
            Dim ORDR_QTY_style As Int64 = Val(rowARTWALMI.Item("ORDR_QTY") & "")
            Dim ADD_UNITS As Int64 = PO_QTY - ORDR_QTY_style
            Dim ADD_CASES As Int64 = System.Math.Sign(ADD_UNITS) * (System.Math.Abs(ADD_UNITS) + CARTON_PACK_QTY - 1) \ CARTON_PACK_QTY

            Dim sqlQ As String = " and ORDR_QTY_OFF <> 0"
            If chkProrate.Checked And PO_QTY <> 0 Then
                sqlQ = ""
            Else
                ADD_CASES = 0
            End If

            Dim ADD_CASES_balance As Int64 = ADD_CASES

            Dim TOT_CASES As Int64 = Val(dst.Tables("ARTWALMY").Compute("SUM(CASES)", sqlI) & "")

            Dim adC As String = "CASES"
            If ADD_CASES < 0 Then adC = "CASES DESC"


            For Each rowARTWALMY As DataRow In dst.Tables("ARTWALMY").Select(sqlI & sqlQ, adC)

                Dim CASES As Int64 = Val(rowARTWALMY.Item("CASES") & "")
                Dim CASES_to_add As Int64 = 0
                If TOT_CASES <> 0 Then
                    CASES_to_add = ADD_CASES * CASES / TOT_CASES + System.Math.Sign(ADD_CASES)
                End If

                CASES_to_add = -1 * CASES_to_add ' ORDR_QTY_OFF is qty over, so you need a negative to add

                If ADD_CASES_balance = 0 Then
                    CASES_to_add = 0
                Else
                    If System.Math.Abs(CASES_to_add) > System.Math.Abs(ADD_CASES_balance) Then
                        CASES_to_add = -1 * ADD_CASES_balance
                    End If
                    ADD_CASES_balance += CASES_to_add
                End If

                Dim CUST_DC_NO As String = rowARTWALMY.Item("CUST_DC_NO")
                Dim sqlwDC As String = " and CUST_DC_NO = '" & CUST_DC_NO & "'"

                If ASCMAIN1.Running_in_VS Then
                    '  If WALMART_STYLE = "551003015" And CUST_DC_NO = "6010" Then Stop
                End If


                Dim ad As String = ""
                Dim sqlw As String = sqlI & sqlwDC

                Dim QTY As Int64 = INNER_PACK_QTY
                Dim ORDR_QTY_OFF As Int64 = Val(rowARTWALMY.Item("ORDR_QTY_OFF") & "") + CASES_to_add * CARTON_PACK_QTY

                If ORDR_QTY_OFF <> 0 Then

                    If ORDR_QTY_OFF > 0 And chkRoundUp.Checked Then
                        ORDR_QTY_OFF -= CARTON_PACK_QTY
                    End If

                    If ORDR_QTY_OFF > 0 Then

                        ' if we have some 0s, 12s, and 24s, and we need to take away 12, do we take one of the 12s or one of the 24s?

                        If MIN_QTY <> 0 Then sqlw &= " and ORDR_QTY > " & CStr(MIN_QTY)
                        sqlw &= " and ORDR_QTY >= " & CStr(QTY)
                        QTY = -1 * QTY
                        ad = "SLS_FACTOR"
                    ElseIf ORDR_QTY_OFF < 0 Then
                        If MAX_QTY <> 0 Then sqlw &= " and ORDR_QTY < " & CStr(MAX_QTY)

                        If working_with_imported_order Then
                            ' NOTE - THE NEXT LINE OF CODE WAS TAKEN OUT IN ORDER TO GET CASE PACK ROUNDING TO WORK ON AN IMPORTED ORDER
                        Else
                            sqlw &= " and WKS_SUP < " & numSupplyWeeks.Value
                        End If

                        sqlw &= " and ORDR_QTY >= 0" ' TO PREVENT CASE PACK ROUNDING FROM GIVING TO STORES WHO DID NOT GET IN THE 1ST PLACE - MAYBE THIS SHOULD BE TYPED TO THE DISTR METHOD
                        ad = "ORDR_QTY DESC, SLS_FACTOR DESC"
                    End If

                    ' If WALMART_STYLE = "550959444" And CUST_DC_NO = "6009" Then Stop

                    With dst.Tables("ARTWALMY")
                        .Columns("ORDR_QTY").Expression = ""
                        .Columns("CASES").Expression = ""
                        .Columns("ORDR_QTY_OFF").Expression = ""
                    End With

                    With dst.Tables("ARTWALMX")
                        For Each row As DataRow In .Select(sqlw & " and ISNULL(SLS_FACTOR,0) >= " & CStr(SLS_FACTOR_min), ad)
                            Dim ORDR_QTY As Int64 = Val(row.Item("ORDR_QTY") & "")
                            ORDR_QTY += QTY
                            row.Item("ORDR_QTY") = ORDR_QTY
                            ORDR_QTY_OFF += QTY
                            If ORDR_QTY_OFF = 0 Then Exit For
                        Next
                    End With

                    With dst.Tables("ARTWALMY")
                        .Columns("ORDR_QTY").Expression = "SUM(CHILD.ORDR_QTY)"
                        .Columns("CASES").Expression = "ORDR_QTY / CARTON_PACK_QTY"
                        .Columns("ORDR_QTY_OFF").Expression = "ORDR_QTY - CASES * CARTON_PACK_QTY"
                    End With
                End If

            Next
        Next

        grdARTWALMI.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
        grdARTWALMY.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Calculate()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Re-Calculating")

        For Each rowARTWALMS As DataRow In dst.Tables("ARTWALMS").Select("SALES_IND = '1'")
            Dim CUST_STORE_NO As String = rowARTWALMS.Item("STORE_NBR")
            Dim VOL_GROUP As String = rowARTWALMS.Item("VOL_GROUP")
            Dim rowARTWALMV As DataRow = dst.Tables("ARTWALMV").Rows.Find(VOL_GROUP)
            Dim sqlw As String = "CUST_STORE_NO = '" & CUST_STORE_NO & "'"

            Dim SLS_PTD As Int64 = Val(rowARTWALMS.Item("SLS_LIF") & "")
            Dim SLS_FACTOR As Decimal = SLS_PTD / SLS_AVG
            If SLS_FACTOR < 0 Then SLS_FACTOR = 0
            Dim VOL_QTY As Int64 = Val(rowARTWALMV.Item("QTY") & "")
            rowARTWALMS.Item("VOL_QTY") = VOL_QTY

            For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select(sqlw)
                rowARTWALMX.Item("VOL_QTY") = VOL_QTY
            Next
        Next


        Dim SupplyWeeks As Int32 = numSupplyWeeks.Value

        Dim WKS As Int64 = numPeriod1Weeks.Value

        Dim rowARTWALMV_min As DataRow = dst.Tables("ARTWALMV").Rows.Find(VOL_GROUP_min)
        SLS_FACTOR_min = Val(rowARTWALMV_min.Item("MIN_FACTOR") & "")
        Dim VOL_GROUPs_zero As New List(Of String)
        For Each rowARTWALMV_min In dst.Tables("ARTWALMV").Select("MIN_FACTOR < " & CStr(SLS_FACTOR_min))
            VOL_GROUPs_zero.Add(rowARTWALMV_min.Item("VOL_GROUP"))
        Next

        With dst.Tables("ARTWALMY")
            .Columns("ORDR_QTY").Expression = ""
            .Columns("CASES").Expression = ""
            .Columns("ORDR_QTY_OFF").Expression = ""
        End With

        For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select("")
            rowARTWALMX.Item("ORDR_QTY_ROUND") = 0
            rowARTWALMX.Item("ORDR_QTY") = 0
        Next

        For Each rowARTWALMI As DataRow In dst.Tables("ARTWALMI").Select("", "")
            Dim WALMART_STYLE As String = rowARTWALMI.Item("WALMART_STYLE")
            Dim CARTON_PACK_QTY As Int64 = Val(rowARTWALMI.Item("CARTON_PACK_QTY") & "")
            Dim INNER_PACK_QTY As Int64 = Val(rowARTWALMI.Item("INNER_PACK_QTY") & "")
            Dim PO_QTY As Int64 = Val(rowARTWALMI.Item("PO_QTY") & "")
            Dim PO_QTY_GIVEN As Int64 = 0
            Dim STRS_GIVEN As Int64 = 0
            Dim MAX_QTY As Int64 = Val(rowARTWALMI.Item("MAX_QTY") & "") * INNER_PACK_QTY
            Dim MIN_QTY As Int64 = Val(rowARTWALMI.Item("MIN_QTY") & "") * INNER_PACK_QTY

            If Not chkMinMax.Checked Then
                MAX_QTY = 0
                MIN_QTY = 0
            End If

            Dim sqlI As String = "WALMART_STYLE = '" & WALMART_STYLE & "'"

            Select Case optMethod.Value

                Case "B"
                    For Each rowARTWALMY As DataRow In dst.Tables("ARTWALMY").Select(sqlI, "")
                        rowARTWALMY.Item("CARTON_PACK_QTY") = CARTON_PACK_QTY
                        rowARTWALMY.Item("INNER_PACK_QTY") = INNER_PACK_QTY
                    Next

                    If PO_QTY <> 0 Then
                        For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select(sqlI, "SLS_FACTOR DESC")
                            Dim SLSX_PTD As Int64 = Val(rowARTWALMX.Item("SLS_PTD") & "")
                            Dim SLSX_LIF As Int64 = Val(rowARTWALMX.Item("SLS_LIF") & "")
                            Dim SUP_TOT As Int64 = Val(rowARTWALMX.Item("SUP_TOT") & "")
                            ' Dim SLSX As Int64 = SLSX_PTD
                            Dim WKS_SUP As Decimal = IIf(SLSX_PTD = 0, 0, SUP_TOT / (SLSX_PTD / WKS))
                            rowARTWALMX.Item("WKS_SUP") = WKS_SUP

                            Dim CUST_STORE_NO As String = rowARTWALMX.Item("CUST_STORE_NO")
                            'If CUST_STORE_NO = "3271" And WALMART_STYLE = "550959445" Then Stop
                            'If ASCMAIN1.Running_in_VS AndAlso WALMART_STYLE = "550959445" Then Stop
                            Dim SLS_FACTOR As Decimal = Val(rowARTWALMX.Item("SLS_FACTOR") & "")
                            Dim ADD_QTY As Int64 = Val(rowARTWALMX.Item("VOL_QTY") & "") * INNER_PACK_QTY - SUP_TOT
                            If ADD_QTY < 0 Then ADD_QTY = 0

                            Dim ORDR_QTY_CALC As Decimal = SupplyWeeks * (SLSX_PTD / WKS)
                            rowARTWALMX.Item("ORDR_QTY_CALC") = ORDR_QTY_CALC

                            Dim ORDR_QTY_ROUND As Int64 = INNER_PACK_QTY * ((ORDR_QTY_CALC + INNER_PACK_QTY - 1) \ INNER_PACK_QTY) + ADD_QTY

                            If MAX_QTY <> 0 And ORDR_QTY_ROUND > MAX_QTY Then ORDR_QTY_ROUND = MAX_QTY
                            If MIN_QTY <> 0 And ORDR_QTY_ROUND < MIN_QTY Then ORDR_QTY_ROUND = MIN_QTY

                            If SLS_FACTOR < SLS_FACTOR_min Then ORDR_QTY_ROUND = 0

                            If ORDR_QTY_ROUND < 0 Then ORDR_QTY_ROUND = 0
                            If ORDR_QTY_ROUND Mod INNER_PACK_QTY <> 0 Then ORDR_QTY_ROUND = INNER_PACK_QTY * ((ORDR_QTY_ROUND + INNER_PACK_QTY - 1) \ INNER_PACK_QTY)

                            PO_QTY_GIVEN += ORDR_QTY_ROUND
                            rowARTWALMX.Item("ORDR_QTY_ROUND") = ORDR_QTY_ROUND
                            rowARTWALMX.Item("ORDR_QTY") = ORDR_QTY_ROUND
                            ' If ORDR_QTY_ROUND < 0 Or ORDR_QTY_ROUND Mod INNER_PACK_QTY <> 0 Then Stop
                            If ORDR_QTY_ROUND <> 0 Then
                                STRS_GIVEN += 1
                            End If
                            If PO_QTY <> 0 And PO_QTY_GIVEN > PO_QTY Then
                                Exit For
                            End If
                        Next
                    End If

                Case "D"

                    For Each rowARTWALMY As DataRow In dst.Tables("ARTWALMY").Select(sqlI, "")
                        Dim CUST_DC_NO As String = rowARTWALMY.Item("CUST_DC_NO")
                        Dim SLS_PTD As Int64 = Val(rowARTWALMY.Item("SLS_PTD") & "")
                        Dim SLS_LIF As Int64 = Val(rowARTWALMY.Item("SLS_LIF") & "")

                        Dim sqlw = sqlI & " and CUST_DC_NO = '" & CUST_DC_NO & "'"
                        Dim STORES_WITH_SALES As Int64 = Val(dst.Tables("ARTWALMX").Compute("COUNT(CUST_STORE_NO)", sqlw & " and " & SLS_PTD & " <> 0") & "")
                        rowARTWALMY.Item("STORES_WITH_SALES") = STORES_WITH_SALES
                        Dim SLS_AVG As Decimal = IIf(STORES_WITH_SALES = 0, 0, SLS_PTD / STORES_WITH_SALES)
                        rowARTWALMY.Item("SLS_AVG") = SLS_AVG
                        rowARTWALMY.Item("CARTON_PACK_QTY") = CARTON_PACK_QTY
                        rowARTWALMY.Item("INNER_PACK_QTY") = INNER_PACK_QTY

                        For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select(sqlw)
                            Dim SLSX_PTD As Int64 = Val(rowARTWALMX.Item("SLS_PTD") & "")
                            Dim SLSX_LIF As Int64 = Val(rowARTWALMX.Item("SLS_LIF") & "")
                            Dim SUP_TOT As Int64 = Val(rowARTWALMX.Item("SUP_TOT") & "")
                            ' Dim SLSX As Int64 = SLSX_PTD
                            Dim WKS_SUP As Decimal = IIf(SLSX_PTD = 0, 0, SUP_TOT / (SLSX_PTD / WKS))
                            rowARTWALMX.Item("WKS_SUP") = WKS_SUP

                            Dim CUST_STORE_NO As String = rowARTWALMX.Item("CUST_STORE_NO")
                            '  If CUST_STORE_NO = "3271" And WALMART_STYLE = "550959445" Then Stop

                            Dim SLS_FACTOR As Decimal = Val(rowARTWALMX.Item("SLS_FACTOR") & "")
                            Dim ADD_QTY As Int64 = Val(rowARTWALMX.Item("VOL_QTY") & "") * INNER_PACK_QTY - SUP_TOT
                            If ADD_QTY < 0 Then ADD_QTY = 0

                            Dim ORDR_QTY_CALC As Decimal = SupplyWeeks * (SLSX_PTD / WKS)
                            rowARTWALMX.Item("ORDR_QTY_CALC") = ORDR_QTY_CALC

                            Dim ORDR_QTY_ROUND As Int64 = INNER_PACK_QTY * ((ORDR_QTY_CALC + INNER_PACK_QTY - 1) \ INNER_PACK_QTY) + ADD_QTY

                            If MAX_QTY <> 0 And ORDR_QTY_ROUND > MAX_QTY Then ORDR_QTY_ROUND = MAX_QTY
                            If MIN_QTY <> 0 And ORDR_QTY_ROUND < MIN_QTY Then ORDR_QTY_ROUND = MIN_QTY

                            If SLS_FACTOR < SLS_FACTOR_min Then ORDR_QTY_ROUND = 0

                            If ORDR_QTY_ROUND < 0 Then ORDR_QTY_ROUND = 0
                            If ORDR_QTY_ROUND Mod INNER_PACK_QTY <> 0 Then ORDR_QTY_ROUND = INNER_PACK_QTY * ((ORDR_QTY_ROUND + INNER_PACK_QTY - 1) \ INNER_PACK_QTY)

                            rowARTWALMX.Item("ORDR_QTY_ROUND") = ORDR_QTY_ROUND
                            rowARTWALMX.Item("ORDR_QTY") = ORDR_QTY_ROUND
                            '  If ORDR_QTY_ROUND < 0 Or ORDR_QTY_ROUND Mod INNER_PACK_QTY <> 0 Then Stop
                        Next
                    Next

            End Select


            If chkProrate.Checked Then


                If PO_QTY <> 0 Then

                    Dim ORDR_QTY_total As Decimal = Val(dst.Tables("ARTWALMX").Compute("SUM(ORDR_QTY)", sqlI) & "")
                    ' Dim ORDR_QTY_CALC_total As Decimal = Val(dst.Tables("ARTWALMX").Compute("SUM(ORDR_QTY_CALC)", sqlI) & "")
                    Dim ROUND_FACTOR As Decimal = PO_QTY / ORDR_QTY_total

                    If ROUND_FACTOR > 2 Or ROUND_FACTOR < 0.5 Then
                        For Each rowARTWALMY As DataRow In dst.Tables("ARTWALMY").Select(sqlI, "")
                            Dim CUST_DC_NO As String = rowARTWALMY.Item("CUST_DC_NO")
                            Dim sqlw = sqlI & " and CUST_DC_NO = '" & CUST_DC_NO & "'"
                            For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select(sqlw)
                                'Dim ORDR_QTY_CALC As Decimal = Val(rowARTWALMX.Item("ORDR_QTY_CALC") & "")
                                'Dim VOL_QTY As Int64 = Val(rowARTWALMX.Item("VOL_QTY") & "")
                                Dim ORDR_QTY As Decimal = Val(rowARTWALMX.Item("ORDR_QTY") & "")

                                Dim ORDR_QTY_ROUND As Int64 = INNER_PACK_QTY * ((ORDR_QTY * ROUND_FACTOR + INNER_PACK_QTY - 1) \ INNER_PACK_QTY)

                                If MAX_QTY <> 0 And ORDR_QTY_ROUND > MAX_QTY Then ORDR_QTY_ROUND = MAX_QTY
                                If MIN_QTY <> 0 And ORDR_QTY_ROUND < MIN_QTY Then ORDR_QTY_ROUND = MIN_QTY

                                Dim SLS_FACTOR As Decimal = Val(rowARTWALMX.Item("SLS_FACTOR") & "")
                                If SLS_FACTOR < SLS_FACTOR_min Then ORDR_QTY_ROUND = 0

                                rowARTWALMX.Item("ORDR_QTY_ROUND") = ORDR_QTY_ROUND
                                rowARTWALMX.Item("ORDR_QTY") = ORDR_QTY_ROUND

                                '  If ORDR_QTY_ROUND < 0 Or ORDR_QTY_ROUND Mod INNER_PACK_QTY <> 0 Then Stop
                            Next
                        Next
                    End If

                End If
            End If
        Next


        With dst.Tables("ARTWALMY")
            .Columns("ORDR_QTY").Expression = "SUM(CHILD.ORDR_QTY)"
            .Columns("CASES").Expression = "ORDR_QTY / CARTON_PACK_QTY"
            .Columns("ORDR_QTY_OFF").Expression = "ORDR_QTY - CASES * CARTON_PACK_QTY"
        End With

        grdARTWALMI.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
        grdARTWALMY.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

        If chkProrate.Checked Then
            If dst.Tables("ARTWALMI").Select("PO_QTY <> 0 AND PO_QTY <> ORDR_QTY").Length <> 0 Then
                Balance_DC_Cases()
            End If
        End If
        Balance_DC_Cases()

        grdARTWALMI.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Summary()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Calculating Summary by Store")

        Dim CUST_STORE_NO As String = ""
        Dim rowARTWALMS As DataRow = Nothing
        For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select("", "CUST_STORE_NO")
            If CUST_STORE_NO <> rowARTWALMX.Item("CUST_STORE_NO") Then
                CUST_STORE_NO = rowARTWALMX.Item("CUST_STORE_NO")
                rowARTWALMS = dst.Tables("ARTWALMS").Rows.Find(CUST_STORE_NO)
            End If
            Dim WALMART_STYLE As String = rowARTWALMX.Item("WALMART_STYLE")
            rowARTWALMS.Item("QTY_" & WALMART_STYLE) = rowARTWALMX.Item("ORDR_QTY")
        Next

        Dim QTY_TOTAL As Int64 = Val(dst.Tables("ARTWALMS").Compute("SUM(QTY_TOTAL)", "") & "")
        Dim STORES As Int64 = Val(dst.Tables("ARTWALMS").Compute("COUNT(STORE_NBR)", "QTY_TOTAL > 0") & "")
        ORD_AVG = IIf(STORES = 0, 0, QTY_TOTAL / STORES)

        dst.Tables("ARTWALMS").Columns("ORD_FACTOR").Expression = IIf(ORD_AVG = 0, "0", "QTY_TOTAL / " & CStr(ORD_AVG))

        CUST_STORE_NO = ""
        For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select("", "CUST_STORE_NO")
            If CUST_STORE_NO <> rowARTWALMX.Item("CUST_STORE_NO") Then
                CUST_STORE_NO = rowARTWALMX.Item("CUST_STORE_NO")
                rowARTWALMS = dst.Tables("ARTWALMS").Rows.Find(CUST_STORE_NO)
            End If
            rowARTWALMX.Item("ORD_FACTOR") = rowARTWALMS.Item("ORD_FACTOR")
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Populate_Volume_Groups()
        Dim F As Decimal = 0.25
        Dim F1 As Decimal = 1 + F
        Dim F2 As Decimal = 1 - F

        With dst.Tables("ARTWALMV")
            .Rows.Add("S", 1 * F1 * F1 * F1 * F1 * F1 * F1)
            .Rows.Add("A+++", 1 * F1 * F1 * F1 * F1 * F1)
            .Rows.Add("A++", 1 * F1 * F1 * F1 * F1)
            .Rows.Add("A+", 1 * F1 * F1 * F1)
            .Rows.Add("A", 1 * F1 * F1)
            .Rows.Add("B", 1 * F1)
            .Rows.Add("C", 1)
            .Rows.Add("D", 1 * F2)
            .Rows.Add("E", 1 * F2 * F2)
            .Rows.Add("F", 0)
            VOL_GROUP_min = "F"
        End With
        Sort_grdColumns(grdARTWALMV, "MIN_FACTOR".ToLower)
    End Sub

    Private Sub grdARTWALMI_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTWALMI.InitializeLayout

    End Sub

    Private Sub grdARTWALMI_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTWALMI.InitializeRow
        Dim INNER_PACK_QTY As Int64 = Val(e.Row.Cells("INNER_PACK_QTY").Value & "")
        Dim CARTON_PACK_QTY As Int64 = Val(e.Row.Cells("CARTON_PACK_QTY").Value & "")
        Dim PO_QTY As Int64 = Val(e.Row.Cells("PO_QTY").Value & "")
        If PO_QTY <> 0 AndAlso CARTON_PACK_QTY <> 0 AndAlso PO_QTY Mod CARTON_PACK_QTY <> 0 Then
            e.Row.Cells("PO_QTY").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PO_QTY").ToolTipText = "PO Qty not an even multiple of Case Qty"
        Else
            e.Row.Cells("PO_QTY").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("PO_QTY").ToolTipText = ""
        End If

        e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Empty
        e.Row.Cells("ORDR_QTY").ToolTipText = ""
        If PO_QTY <> 0 Then
            Dim ORDR_QTY As Int64 = Val(e.Row.Cells("ORDR_QTY").Value & "")
            If PO_QTY <> ORDR_QTY Then
                e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("ORDR_QTY").ToolTipText = "Order Qty does not match PO Qty"
            End If
        End If
    End Sub

    Private Sub tabPlan_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabPlan.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        If tabPlan.SelectedTab.Key = "Summary by Store" Then Setup_Summary()

        If ScreenMode Then
            With UltraExplorerBar1
                .Groups("Parameters").Visible = (tabPlan.SelectedTab.Key = "Order Planning Worksheet")
                .Groups("Volume Groups").Visible = (tabPlan.SelectedTab.Key = "Order Planning Worksheet")
            End With
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
            Case "grdARTWALMY"
                Dim KEY As String = summarySettings.Key
                If KEY = "SLS_AVG" Then
                    CustomValue = SLS_AVG
                End If

                'Case "grdARTWALMS"
                '    Dim KEY As String = summarySettings.Key
                '    If KEY = "SLS_AVG" Then
                '        CustomValue = AVG
                '    End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Private Sub grdARTWALMV_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTWALMV.InitializeRow
        If VOL_GROUP_min = "" Then VOL_GROUP_min = "F"
        If e.Row.Cells("VOL_GROUP").Value = VOL_GROUP_min Then
            e.Row.Appearance.BackColor = Drawing.Color.MediumAquamarine
            e.Row.ToolTipText = "Minimum Volume Group - right click to change"
        Else
            e.Row.Appearance.BackColor = Drawing.Color.Empty
            e.Row.ToolTipText = ""
        End If
    End Sub

    Private Sub optMethod_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optMethod.ValueChanged

    End Sub

    Private Sub btnQtyCheck_Click(sender As System.Object, e As System.EventArgs) Handles btnQtyCheck.Click
        If MsgBox("This function will check for Negative Qtys" _
                  & vbCrLf & " and for Qtys that are not evenly divisible by the Inner Pack" _
                  & vbCrLf & vbCrLf & "OK to Proceed?", _
                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
        For Each rowARTWALMX As DataRow In dst.Tables("ARTWALMX").Select("")
            Dim ORDR_QTY As Decimal = Val(rowARTWALMX.Item("ORDR_QTY") & "")
            Dim INNER_PACK_QTY As Decimal = Val(rowARTWALMX.GetParentRow("ARTWALMY_ARTWALMX").Item("INNER_PACK_QTY") & "")
            If ORDR_QTY < 0 Or ORDR_QTY Mod INNER_PACK_QTY <> 0 Then
                MsgBox("Found Qty = " & CStr(ORDR_QTY) & " for Store " & rowARTWALMX.Item("CUST_STORE_NO") & ", Item " & rowARTWALMX.Item("WALMART_STYLE"), MsgBoxStyle.OkOnly, "Found One")
                Exit Sub
            End If
        Next
        MsgBox("Found Nothing negative or not divisible by Inner pack Qty", MsgBoxStyle.OkOnly, "Congratulations")
    End Sub
     
    Private Sub cmdCPR_Click(sender As System.Object, e As System.EventArgs) Handles cmdCPR.Click

        Stop
        ' WE ARE USING HARD-CODED VALUES FOR UNITS_PER_INNER_ AND INNERS_PER_CASE
        ' WE SHOULD BE ABLE TO RELY ON VALUES FROM THE STYLE MASTER TABLE
        ' AND THEN WE SHOULD SET THE VALUES FOR THESE VARABLES IN THE WALMI LOOP


        For Each rowARTWALMR As DataRow In dst.Tables("ARTWALMR").Select("CPR <> 0", "CUST_DC_NO")
            Dim CUST_DC_NO As String = rowARTWALMR.Item("CUST_DC_NO") & ""
            Dim WALMART_STYLE As String = rowARTWALMR.Item("WALMART_STYLE") & ""
            Dim CPR As Integer = Val(rowARTWALMR.Item("CPR") & "")
            Dim sqlsort As String = "ORDR_QTY"
            Dim G As Integer = 1
            If CPR < 0 Then
                sqlsort &= " DESC"
                G = -1
            End If
            Dim sqlx As String = "WALMART_STYLE = '" & WALMART_STYLE & "' AND CUST_DC_NO = '" & CUST_DC_NO & "'"
            For Each rowARTWALMO As DataRow In dst.Tables("ARTWALMO").Select(sqlx, sqlsort)
                Dim ORDR_QTY As Integer = Val(rowARTWALMO.Item("ORDR_QTY") & "")
                ORDR_QTY += G
                rowARTWALMO.Item("ORDR_QTY") = ORDR_QTY
                CPR -= G
                If CPR = 0 Then Exit For
            Next
        Next

        'Dim ORDR_QTY As Int64 = Val(dst.Tables("ARTWALMR").Compute("SUM(ORDR_QTY)", "") & "")
        'Dim ORDR_QTY_ORIG As Int64 = Val(dst.Tables("ARTWALMR").Compute("SUM(ORDR_QTY_ORIG)", "") & "")

        If 1 <> 1 Then
            For Each rowARTWALMI As DataRow In dst.Tables("ARTWALMI").Select("")
                Dim WALMART_STYLE As String = rowARTWALMI.Item("WALMART_STYLE") & ""

                Dim DIFF As Int64 = Val(dst.Tables("ARTWALMR").Compute("SUM(DIFF)", "WALMART_STYLE = '" & WALMART_STYLE & "'") & "")
                'DIFF = DIFF - 3 ' HARD CODED FOR 10/09 SSO
                'DIFF = DIFF - 6 ' HARD CODED FOR 01/30/14 SSO 12 INNERS IN A CASE, INNER QTY = 1
                'DIFF = DIFF - CInt(0.5 + INNERS_PER_CASE / 2)
                If DIFF <> 0 Then
                    Dim sqls As String = "DIFF"
                    Dim sqlsort As String = "ORDR_QTY"
                    Dim G As Integer = 1
                    If DIFF > 0 Then
                        sqls &= " DESC"
                        sqlsort &= " DESC"
                        G = -1
                    End If
                    For Each rowARTWALMR As DataRow In dst.Tables("ARTWALMR").Select("WALMART_STYLE = '" & WALMART_STYLE & "'", sqls)
                        Dim CUST_DC_NO As String = rowARTWALMR.Item("CUST_DC_NO") & ""

                        Dim S As Integer = 0
                        Dim sqlx As String = "WALMART_STYLE = '" & WALMART_STYLE & "' AND CUST_DC_NO = '" & CUST_DC_NO & "'"
                        For Each rowARTWALMO As DataRow In dst.Tables("ARTWALMO").Select(sqlx, sqlsort)
                            Dim ORDR_QTY As Integer = Val(rowARTWALMO.Item("ORDR_QTY") & "")
                            ORDR_QTY += G
                            rowARTWALMO.Item("ORDR_QTY") = ORDR_QTY
                            DIFF += G
                            S += 1
                            If S = INNERS_PER_CASE Then Exit For
                        Next
                        If DIFF = 0 Then Exit For
                    Next
                End If
            Next
        End If
    End Sub

    Private Sub grdARTWALMR_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTWALMR.InitializeLayout

    End Sub

    Private Sub grdARTWALMR_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTWALMR.InitializeRow
        If e.Row.Band.Key = "" Then
            Dim ORDR_QTY As Integer = Val(e.Row.Cells("ORDR_QTY").Value & "")
            Dim ORDR_QTY_ORIG As Integer = Val(e.Row.Cells("ORDR_QTY_ORIG").Value & "")
            If ORDR_QTY <> ORDR_QTY_ORIG Then
                e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Private Sub cmdExportSSO_Click(sender As System.Object, e As System.EventArgs) Handles cmdExportSSO.Click
        Dim EMSG As String = ""
        If dst.Tables("ARTWALMR").Select("CPR <> 0").Length <> 0 Then
            EMSG &= vbCr & "There are some DC's with non-0 CPR values"
        End If
        If dst.Tables("ARTWALMO").Select("ORDR_QTY <= 0").Length <> 0 Then
            EMSG &= vbCr & "There are some Stores with 0 order qtys"
        End If
        If dst.Tables("ARTWALMO").Select("CUST_DC_NO = '0000'").Length <> 0 Then
            EMSG &= vbCr & "There are some Stores with an invalid DC"
        End If
        If EMSG <> "" Then
            MsgBox(Mid(EMSG, 2), vbOKOnly, "Cannot Export the SSO for the reasons noted")
        Else
            grdARTWALMO.DisplayLayout.Bands(0).Columns("CUST_DC_NO").Hidden = True
            grdARTWALMO.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME").Hidden = True
            Export_to_Excel(grdARTWALMO)
            grdARTWALMO.DisplayLayout.Bands(0).Columns("CUST_DC_NO").Hidden = False
            grdARTWALMO.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME").Hidden = False
        End If
    End Sub
End Class