
Public Class WBFTSPO1
    Dim InquiryOnly As Boolean = False
    Dim FromDate As Date
    Dim ToDate As Date
    'Dim RankOption As String = "R"

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("S2.STYLE_CODE,")
            SQLs.AppendLine("S2.COLOR_CODE,")
            SQLs.AppendLine("I1.STYLE_DESC,")
            SQLs.AppendLine("I1.STYLE_CLASS_CODE,")
            SQLs.AppendLine("V1.VEND_SUPPLIER_ID,")
            SQLs.AppendLine("NULL AS THEME_DESC,")
            SQLs.AppendLine("I1.COUNTRY_CODE,")
            SQLs.AppendLine("0 WHSE_QTY_ON_HAND, 0 WHSE_QTY_PICK, 0 OPEN_TO_SELL, 0 WHSE_QTY_TRAN, 0 WHSE_QTY_ON_ORDER, 0 WHSE_QTY_OPEN, 0 FUT_AVAIL,")
            SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0) - NVL(S2.ORDR_QTY_CANC,0)) AS ORDER_QTY,")
            SQLs.AppendLine("SUM((NVL(S2.ORDR_QTY,0) - NVL(S2.ORDR_QTY_CANC,0)) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
            SQLs.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1, ICTSTYL1 I1, APTVEND1 V1")
            SQLs.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
            SQLs.AppendLine("AND S2.STYLE_CODE = I1.STYLE_CODE")
            SQLs.AppendLine("AND I1.VEND_CODE (+) = V1.VEND_CODE")
            SQLs.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
            SQLs.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
            SQLs.AppendLine("AND S1.ORDR_DATE >= '01-May-2025'")
            SQLs.AppendLine("AND S1.ORDR_DATE < '01-Jun-2025'")
            SQLs.AppendLine("GROUP BY S2.STYLE_CODE, S2.COLOR_CODE, I1.STYLE_DESC, I1.STYLE_CLASS_CODE, V1.VEND_SUPPLIER_ID, I1.COUNTRY_CODE")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTHORNT", "**", 0, False, "", 2)
            With .Tables("WBTHORNT").Columns
                .Add("SEL", GetType(System.String))
            End With

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("S1.SREP_CODE,")
            SQLs.AppendLine("R1.SREP_NAME,")
            SQLs.AppendLine("S1.ORDR_DATE,")
            SQLs.AppendLine("S1.CUST_CODE,")
            SQLs.AppendLine("S1.CUST_NAME,")
            SQLs.AppendLine("S1.ORDR_NO,")
            SQLs.AppendLine("S1.ORDR_NO_WEB,")
            SQLs.AppendLine("DECODE(NVL(S1.ORDR_NO_WEB,'0'),'0','0','1') AS WEB_ORDR,")
            SQLs.AppendLine("S1.ORDR_GROUP_NO,")
            SQLs.AppendLine("S1.ORDR_CUST_PO,")
            SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
            SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
            SQLs.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
            SQLs.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
            SQLs.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
            SQLs.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
            SQLs.AppendLine("GROUP BY")
            SQLs.AppendLine("S1.SREP_CODE,")
            SQLs.AppendLine("R1.SREP_NAME,")
            SQLs.AppendLine("S1.ORDR_DATE,")
            SQLs.AppendLine("S1.CUST_CODE,")
            SQLs.AppendLine("S1.CUST_NAME,")
            SQLs.AppendLine("S1.ORDR_NO,")
            SQLs.AppendLine("S1.ORDR_NO_WEB,")
            SQLs.AppendLine("DECODE(NVL(S1.ORDR_NO_WEB,'0'),'0','0','1'),")
            SQLs.AppendLine("S1.ORDR_GROUP_NO,")
            SQLs.AppendLine("S1.ORDR_CUST_PO")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTHORND", "**", 0, False)
            Create_TDA(.Tables.Add, "WBTHORNO", "**", 0, False)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("P2.STYLE_CODE,")
            SQLs.AppendLine("P2.COLOR_CODE,")
            SQLs.AppendLine("P2.PO_ORDER_NO,")
            SQLs.AppendLine("P2.PO_DATE_SHIP_BY,")
            SQLs.AppendLine("P1.PO_STATUS,")
            SQLs.AppendLine("P2.PO_QTY_OPN")
            SQLs.AppendLine("FROM POTORDR1 P1, POTORDR2 P2")
            SQLs.AppendLine("WHERE P1.PO_ORDER_NO = P2.PO_ORDER_NO")
            SQLs.AppendLine("And P2.PO_QTY_OPN > 0")
            SQLs.AppendLine($"And P1.WHSE_CODE = 'MS'")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTHORNP", "**", 0, False)

            .Relations.Add(ASCDATA1.GetRelation(dst, "WBTHORNT", "WBTHORNP", "STYLE_CODE, COLOR_CODE"))
        End With

        grdWBFTSPO1.DataSource = dst.Tables("WBTHORNT")
        grdWBFHORND.DataSource = dst.Tables("WBTHORND")
        grdWBFHORNO.DataSource = dst.Tables("WBTHORNO")

        Create_Summary(grdWBFTSPO1, "ORDER_QTY", "Sum", "", "###,##0")
        Create_Summary(grdWBFTSPO1, "SALES", "Sum", "", "###,##0.00")
        Create_Summary(grdWBFTSPO1, "STYLE_CODE", "Count", "", "###,##0")

        Create_Summary(grdWBFHORND, "ORDER_QTY", "Sum", "", "###,##0")
        Create_Summary(grdWBFHORND, "SALES", "Sum", "", "###,##0.00")

        Create_Summary(grdWBFHORNO, "ORDER_QTY", "Sum", "", "###,##0")
        Create_Summary(grdWBFHORNO, "SALES", "Sum", "", "###,##0.00")

        With grdWBFTSPO1.DisplayLayout.Bands(0)
            .Columns("ORDER_QTY").Format = "###,##0"
            .Columns("SALES").Format = "###,##0.00"
            .Columns("WHSE_QTY_ON_HAND").Format = "###,##0"
            .Columns("WHSE_QTY_PICK").Format = "###,##0"
            .Columns("OPEN_TO_SELL").Format = "###,##0"
            .Columns("WHSE_QTY_TRAN").Format = "###,##0"
            .Columns("WHSE_QTY_ON_ORDER").Format = "###,##0"
            .Columns("WHSE_QTY_OPEN").Format = "###,##0"
            .Columns("FUT_AVAIL").Format = "###,##0"

            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
        End With

        grdWBFHORND.DisplayLayout.Bands(0).Columns("ORDER_QTY").Format = "###,##0"
        grdWBFHORND.DisplayLayout.Bands(0).Columns("SALES").Format = "###,##0.00"

        grdWBFHORNO.DisplayLayout.Bands(0).Columns("ORDER_QTY").Format = "###,##0"
        grdWBFHORNO.DisplayLayout.Bands(0).Columns("SALES").Format = "###,##0.00"

        With grdWBFTSPO1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdWBFTSPO1.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBFTSPO1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        With grdWBFHORND.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdWBFHORND.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBFHORND.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        With grdWBFHORNO.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdWBFHORNO.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBFHORNO.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        dtFROM.Value = CDate(String.Format("{0}/01/{1}", Now.Month, Now.Year))
        dtTO.Value = CDate(String.Format("{0}/{1}/{2}", Now.Month, Date.DaysInMonth(Now.Year, Now.Month), Now.Year))

        Load_Record(False)

        Sort_grdColumns(grdWBFTSPO1, "SALES".ToLower(), False)
        Sort_grdColumns(grdWBFHORND, "SALES".ToLower(), False)
        Sort_grdColumns(grdWBFHORNO, "SALES".ToLower(), False)

        tab.Visible = False

        Dim lstWHSE_CODE As New List(Of String)
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT WHSE_CODE FROM ICTWHSE1")
        Dim tblICTWHSE1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        For Each rowICTWHSE1 As DataRow In tblICTWHSE1.Rows
            lstWHSE_CODE.Add(rowICTWHSE1.Item("WHSE_CODE").ToString & String.Empty)
        Next
        cboWHSE_CODE.DataSource = lstWHSE_CODE
        cboWHSE_CODE.SelectedItem = "MS"

        'grdWBFHORNT.Parent = tab.Parent

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"

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
                Load_Record(True)
            Case "Exit"
                Me.Close()
                'Call Mode_Settings(False)
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
        'UltraExplorerBar1.Groups("E-Commerce").Visible = False
        SetShowOrderDetails()
        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        dst.Tables("WBTHORNT").Rows.Clear()
    End Sub

    Sub Load_Record(Optional showRefreshing As Boolean = False)
        Me.Cursor = Cursors.WaitCursor
        If showRefreshing Then
            ASCMAIN1.Progress("Refreshing Data", "")
        End If
        Application.DoEvents()
        'Call Save_Header_Fields(UltraGroupBox1)
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        EnforceConstraints(False)
        Dim SELECTS As String = ""
        Dim GROUPS As String = "'"

        FromDate = CDate(dtFROM.Value)
        ToDate = CDate(dtTO.Value).AddDays(1)
        Dim SQLE As New Text.StringBuilder With {.Length = 0}

        SQLs.Length = 0
        SQLs.AppendLine("SELECT S2.STYLE_CODE, S2.COLOR_CODE, I1.STYLE_DESC,")
        SQLs.AppendLine("I1.STYLE_CLASS_CODE, V1.VEND_SUPPLIER_ID, T1.THEME_DESC, I1.COUNTRY_CODE,")
        SQLs.AppendLine("W1.WHSE_QTY_ON_HAND, W1.WHSE_QTY_PICK, W1.OPEN_TO_SELL, W1.WHSE_QTY_TRAN, W1.WHSE_QTY_ON_ORDER, W1.WHSE_QTY_OPEN, W1.FUT_AVAIL,")
        If chkRemoveCancelled.Checked Then
            SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0) - NVL(S2.ORDR_QTY_CANC,0)) AS ORDER_QTY,")
            SQLs.AppendLine("SUM((NVL(S2.ORDR_QTY,0) - NVL(S2.ORDR_QTY_CANC,0)) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
        Else
            SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
            SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
        End If
        SQLs.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1, ICTSTYL1 I1, APTVEND1 V1, ICTSTYC1 C1, ICTTHEME T1, ")
        '
        SQLs.AppendLine("(")
        SQLs.AppendLine("   SELECT")
        SQLs.AppendLine("   STYLE_CODE,")
        SQLs.AppendLine("   COLOR_CODE,")
        SQLs.AppendLine("   NVL(WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND,")
        SQLs.AppendLine("   NVL(WHSE_QTY_PICK,0) WHSE_QTY_PICK,")
        SQLs.AppendLine("   (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0)) OPEN_TO_SELL,")
        SQLs.AppendLine("   NVL(WHSE_QTY_TRAN,0) WHSE_QTY_TRAN,")
        SQLs.AppendLine("   NVL(WHSE_QTY_ON_ORDER,0) WHSE_QTY_ON_ORDER,")
        SQLs.AppendLine("   NVL(WHSE_QTY_OPEN,0) WHSE_QTY_OPEN,")
        SQLs.AppendLine("   (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) + NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0)) FUT_AVAIL")
        SQLs.AppendLine("   FROM ICTSTAT2")
        SQLs.AppendLine($"  WHERE WHSE_CODE = '{cboWHSE_CODE.Text.ToString & String.Empty}'")
        'SQLs.AppendLine("   AND (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) + NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0)) <> 0")
        SQLs.AppendLine(") W1")
        SQLs.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
        SQLs.AppendLine("AND S2.STYLE_CODE = I1.STYLE_CODE")
        SQLs.AppendLine("AND I1.VEND_CODE (+) = V1.VEND_CODE")
        SQLs.AppendLine("AND S2.STYLE_CODE = C1.STYLE_CODE (+)")
        SQLs.AppendLine("AND S2.COLOR_CODE = C1.COLOR_CODE (+)")
        SQLs.AppendLine("AND S2.STYLE_CODE = W1.STYLE_CODE (+)")
        SQLs.AppendLine("AND S2.COLOR_CODE = W1.COLOR_CODE (+)")
        SQLs.AppendLine("AND C1.THEME_CODE = T1.THEME_CODE (+)")
        SQLs.AppendLine("AND S1.SREP_CODE = R1.SREP_CODE (+)")
        SQLs.AppendLine("AND S1.ORDR_STATUS <> 'C'")
        SQLs.AppendLine($"AND S1.WHSE_CODE = '{cboWHSE_CODE.Text.ToString & String.Empty}'")
        SQLs.AppendLine(String.Format("AND S1.ORDR_DATE >= '{0}'", Format(FromDate, "dd-MMM-yyyy")))
        SQLs.AppendLine(String.Format("AND S1.ORDR_DATE < '{0}'", Format(ToDate, "dd-MMM-yyyy")))
        If chkStylesInventory.Checked Then
            SQLs.AppendLine("AND (S2.STYLE_CODE, S2.COLOR_CODE) IN ")
            SQLs.AppendLine("(SELECT")
            SQLs.AppendLine("STYLE_CODE,")
            SQLs.AppendLine("COLOR_CODE")
            SQLs.AppendLine("FROM ICTSTAT2")
            SQLs.AppendLine("WHERE (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) + NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0)) > 0")
            SQLs.AppendLine($"AND WHSE_CODE = '{cboWHSE_CODE.Text.ToString & String.Empty}'")
            SQLs.AppendLine("GROUP BY STYLE_CODE,")
            SQLs.AppendLine("COLOR_CODE")
            SQLs.AppendLine(")")
        End If
        SQLs.AppendLine(SQLE.ToString)
        SQLs.AppendLine("GROUP BY S2.STYLE_CODE, S2.COLOR_CODE, I1.STYLE_DESC, I1.STYLE_CLASS_CODE, V1.VEND_SUPPLIER_ID, T1.THEME_DESC, I1.COUNTRY_CODE,")
        SQLs.AppendLine("W1.WHSE_QTY_ON_HAND, W1.WHSE_QTY_PICK, W1.OPEN_TO_SELL, W1.WHSE_QTY_TRAN, W1.WHSE_QTY_ON_ORDER, W1.WHSE_QTY_OPEN, W1.FUT_AVAIL")

        Fill_Records("WBTHORNT", , , SQLs.ToString)

        SQLs.Length = 0
        SQLs.AppendLine("SELECT")
        SQLs.AppendLine("P2.STYLE_CODE,")
        SQLs.AppendLine("P2.COLOR_CODE,")
        SQLs.AppendLine("P2.PO_ORDER_NO,")
        SQLs.AppendLine("P2.PO_DATE_SHIP_BY,")
        SQLs.AppendLine("P1.PO_STATUS,")
        SQLs.AppendLine("P2.PO_QTY_OPN")
        SQLs.AppendLine("FROM POTORDR1 P1, POTORDR2 P2")
        SQLs.AppendLine("WHERE P1.PO_ORDER_NO = P2.PO_ORDER_NO")
        SQLs.AppendLine("And P2.PO_QTY_OPN > 0")
        SQLs.AppendLine($"And P1.WHSE_CODE = '{cboWHSE_CODE.Text.ToString & String.Empty}'")
        Fill_Records("WBTHORNP", , , SQLs.ToString)

        'EnforceConstraints(True)
        'grdWBFTSPO1.Text = "Hot Or Not " & RANKING
        'grdWBFTSPO1.DisplayLayout.Bands(0).Columns.Item("RANK_CODE").Header.Caption = RANK_CODE
        'grdWBFTSPO1.DisplayLayout.Bands(0).Columns.Item("RANK_NAME").Header.Caption = RANK_NAME

        'grdWBFTSPO1.DisplayLayout.Bands(0).Columns.Item("ORDER_QTY").Header.VisiblePosition = 9
        'grdWBFTSPO1.DisplayLayout.Bands(0).Columns.Item("SALES").Header.VisiblePosition = 10

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
        FillDetails()
        If grdWBFTSPO1.Rows.Count > 0 Then
            grdWBFTSPO1.Rows(0).Activate()
        End If

        Me.Cursor = Cursors.Default
        If showRefreshing Then
            ASCMAIN1.Progress("")
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
        Generate_Report("WBRHORNT")
        Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWBFTSPO1, "SSBBBBB", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Style Masterfile", "Expand All", "Collapse All", "Copy To Clipboard")
        Load_Popup_Menu(grdWBFHORND, "SSBB", "Show Filter", "Show GroupBox", "Cust Order Inq", "Sales Order Inq")
        Load_Popup_Menu(grdWBFHORNO, "SSBB", "Show Filter", "Show GroupBox", "Cust Order Inq", "Sales Order Inq")
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
                e.Tool.ToolbarsManager.Tools("Style Status Inquiry").SharedProps.Visible = True
                e.Tool.ToolbarsManager.Tools("Style Masterfile").SharedProps.Visible = True
                e.Tool.ToolbarsManager.Tools("Copy To Clipboard").SharedProps.Visible = False
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
            'Case "Edit Ship To"
            '    If Not InquiryOnly Then
            '        MsgBox("Edit Ship To Feature Coming Soon", MsgBoxStyle.Exclamation, "Waiting For Feature")
            '    End If
            Case "Copy To Clipboard"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Clipboard.SetText(STYLE_CODE)
                MsgBox($"{STYLE_CODE} Copied To Clipboard.", vbOKOnly, "Clipboard")
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
            Case "Style Masterfile"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If
            Case "Cust Order Inq"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Dim FIND_BY As String = CUST_CODE
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Text
                FIND_BY &= ":" & ORDR_GROUP_NO
                Context_Launch("Select", FIND_BY, e.Tool.Key, "SOFCORD1")

            Case "Sales Order Inq"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If
            Case "Select All"
                For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
                    rowECTECOM1_FILTER.Item("SEL") = "1"
                Next
            Case "Select None"
                For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
                    rowECTECOM1_FILTER.Item("SEL") = "0"
                Next
            Case "Expand All"
                Me.Cursor = Cursors.WaitCursor
                Application.DoEvents()
                grd.Rows.ExpandAll(True)
                Me.Cursor = Cursors.Default
                Application.DoEvents()
            Case "Collapse All"
                Me.Cursor = Cursors.WaitCursor
                Application.DoEvents()
                grd.Rows.CollapseAll(True)
                Me.Cursor = Cursors.Default
                Application.DoEvents()
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If
        End Select
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
    Private Sub SetRankOption(OPTR As String)
        'RankOption = "S"
        chkStylesInventory.Visible = True
        chkStylesInventory.Checked = False
        chkStyleColors.Visible = True
        chkStyleColors.Checked = False
        UltraExplorerBar1.Groups("E-Commerce").Visible = True
    End Sub

    Private Sub SetShowOrderDetails()
        SplitContainer2.AutoSize = True
        If chkShowDetails.Checked Then
            SplitContainer2.Panel2.Show()
            SplitContainer2.Panel2Collapsed = False
        Else
            SplitContainer2.Panel2.Hide()
            SplitContainer2.Panel2Collapsed = True
        End If
    End Sub
#End Region

#Region "Form Controls"
    Private Sub chkShowDetails_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowDetails.CheckedChanged
        SetShowOrderDetails()
    End Sub

    Private Sub grdWBFHORNT_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWBFTSPO1.AfterRowActivate
        FillDetails()

    End Sub

    Private Sub FillDetails()
        If chkShowDetails.Checked Then
            If IsNothing(grdWBFTSPO1.ActiveRow) Then
                dst.Tables.Item("WBTHORND").Clear()
                dst.Tables.Item("WBTHORNO").Clear()
                grdWBFHORND.Text = "Please Select A Row Above To See Details"
                grdWBFHORNO.Text = "Please Select A Row Above To See Details"
                Exit Sub
            End If
            If Not grdWBFTSPO1.ActiveRow Is Nothing And grdWBFTSPO1.ActiveRow.IsDataRow Then
                FromDate = CDate(dtFROM.Value)
                ToDate = CDate(dtTO.Value).AddDays(1)
                Dim S As New Text.StringBuilder
                Dim STYLE_CODE As String = grdWBFTSPO1.ActiveRow.Cells("STYLE_CODE").Text
                Dim COLOR_CODE As String = grdWBFTSPO1.ActiveRow.Cells("COLOR_CODE").Text
                Dim SQLE As New Text.StringBuilder With {.Length = 0}
                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("S1.SREP_CODE,")
                S.AppendLine("R1.SREP_NAME,")
                S.AppendLine("S1.ORDR_DATE,")
                S.AppendLine("S1.CUST_CODE,")
                S.AppendLine("S1.CUST_NAME,")
                S.AppendLine("S1.ORDR_NO,")
                S.AppendLine("S1.ORDR_NO_WEB,")
                S.AppendLine("DECODE(NVL(S1.ORDR_NO_WEB,'0'),'0','0','1') AS WEB_ORDR,")
                S.AppendLine("S1.ORDR_GROUP_NO,")
                S.AppendLine("S1.ORDR_CUST_PO,")
                If chkRemoveCancelled.Checked Then
                    S.AppendLine("SUM(NVL(S2.ORDR_QTY,0) - NVL(S2.ORDR_QTY_CANC,0)) AS ORDER_QTY,")
                    S.AppendLine("SUM((NVL(S2.ORDR_QTY,0) - NVL(S2.ORDR_QTY_CANC,0)) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                Else
                    S.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
                    S.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                End If
                S.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
                S.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                S.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                S.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
                S.AppendLine(String.Format("AND S2.STYLE_CODE = '{0}'", STYLE_CODE))
                S.AppendLine(String.Format("AND S2.COLOR_CODE = '{0}'", COLOR_CODE))
                S.AppendLine(String.Format("AND S1.WHSE_CODE = '{0}'", cboWHSE_CODE.Text.ToString & String.Empty))
                S.AppendLine(String.Format("AND S1.ORDR_DATE >= '{0}'", Format(FromDate, "dd-MMM-yyyy")))
                S.AppendLine(String.Format("AND S1.ORDR_DATE < '{0}'", Format(ToDate, "dd-MMM-yyyy")))
                S.AppendLine("GROUP BY")
                S.AppendLine("S1.SREP_CODE,")
                S.AppendLine("R1.SREP_NAME,")
                S.AppendLine("S1.ORDR_DATE,")
                S.AppendLine("S1.CUST_CODE,")
                S.AppendLine("S1.CUST_NAME,")
                S.AppendLine("S1.ORDR_NO,")
                S.AppendLine("S1.ORDR_NO_WEB,")
                S.AppendLine("DECODE(NVL(S1.ORDR_NO_WEB,'0'),'0','0','1'),")
                S.AppendLine("S1.ORDR_GROUP_NO,")
                S.AppendLine("S1.ORDR_CUST_PO")
                grdWBFHORND.Text = $"Details For Style/Color {STYLE_CODE} / {COLOR_CODE}"
                Fill_Records("WBTHORND", , , S.ToString)

                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("S1.SREP_CODE,")
                S.AppendLine("R1.SREP_NAME,")
                S.AppendLine("S1.ORDR_DATE,")
                S.AppendLine("S1.CUST_CODE,")
                S.AppendLine("S1.CUST_NAME,")
                S.AppendLine("S1.ORDR_NO,")
                S.AppendLine("S1.ORDR_NO_WEB,")
                S.AppendLine("DECODE(NVL(S1.ORDR_NO_WEB,'0'),'0','0','1') AS WEB_ORDR,")
                S.AppendLine("S1.ORDR_GROUP_NO,")
                S.AppendLine("S1.ORDR_CUST_PO,")
                S.AppendLine("SUM(NVL(S2.ORDR_QTY_OPEN,0)) AS ORDER_QTY,")
                S.AppendLine("SUM((NVL(S2.ORDR_QTY_OPEN,0)) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                S.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
                S.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                S.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                S.AppendLine("AND  S1.ORDR_STATUS = 'O'")
                S.AppendLine(String.Format("AND S2.STYLE_CODE = '{0}'", STYLE_CODE))
                S.AppendLine(String.Format("AND S2.COLOR_CODE = '{0}'", COLOR_CODE))
                S.AppendLine(String.Format("AND S1.WHSE_CODE = '{0}'", cboWHSE_CODE.Text.ToString & String.Empty))
                S.AppendLine("GROUP BY")
                S.AppendLine("S1.SREP_CODE,")
                S.AppendLine("R1.SREP_NAME,")
                S.AppendLine("S1.ORDR_DATE,")
                S.AppendLine("S1.CUST_CODE,")
                S.AppendLine("S1.CUST_NAME,")
                S.AppendLine("S1.ORDR_NO,")
                S.AppendLine("S1.ORDR_NO_WEB,")
                S.AppendLine("DECODE(NVL(S1.ORDR_NO_WEB,'0'),'0','0','1'),")
                S.AppendLine("S1.ORDR_GROUP_NO,")
                S.AppendLine("S1.ORDR_CUST_PO")
                S.AppendLine("HAVING SUM(NVL(S2.ORDR_QTY_OPEN,0)) <> 0")
                grdWBFHORNO.Text = $"Open Orders For Style/Color {STYLE_CODE} / {COLOR_CODE}"
                Fill_Records("WBTHORNO", , , S.ToString)
            End If
        End If
    End Sub

    Private Sub chkStyleColors_CheckedChanged(sender As Object, e As EventArgs) Handles chkStyleColors.CheckedChanged
        If chkStyleColors.Checked Then
            chkShowDetails.Checked = False
            chkShowDetails.Visible = False
        Else
            chkShowDetails.Visible = True
        End If
        SetShowOrderDetails()
    End Sub
#End Region
End Class