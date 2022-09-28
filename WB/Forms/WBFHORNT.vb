
Public Class WBFHORNT
    Dim InquiryOnly As Boolean = False
    Dim FromDate As Date
    Dim ToDate As Date
    Dim RankOption As String = "R"

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst

            SQLs.Length = 0
            SQLs.AppendLine("SELECT 'ALL' AS RANK_CODE, R1.SREP_NAME AS RANK_NAME,")
            SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
            SQLs.AppendLine("SUM((NVL(S2.ORDR_QTY,0) - NVL(S2.ORDR_QTY_CANC,0)) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
            SQLs.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
            SQLs.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
            SQLs.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
            SQLs.AppendLine("AND S1.ORDR_DATE > '01-JAN-2014'")
            SQLs.AppendLine("AND S1.ORDR_DATE < '01-JAN-2014'")
            SQLs.AppendLine("GROUP BY R1.SREP_NAME")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTHORNT", "**", 0, False, "", 2)

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
        End With

        grdWBFHORNT.DataSource = dst.Tables("WBTHORNT")
        grdWBFHORND.DataSource = dst.Tables("WBTHORND")

        Create_Summary(grdWBFHORNT, "ORDER_QTY", "Sum", "", "###,##0")
        Create_Summary(grdWBFHORNT, "SALES", "Sum", "", "###,##0.00")

        Create_Summary(grdWBFHORND, "ORDER_QTY", "Sum", "", "###,##0")
        Create_Summary(grdWBFHORND, "SALES", "Sum", "", "###,##0.00")

        grdWBFHORNT.DisplayLayout.Bands(0).Columns("ORDER_QTY").Format = "###,##0"
        grdWBFHORNT.DisplayLayout.Bands(0).Columns("SALES").Format = "###,##0.00"

        grdWBFHORND.DisplayLayout.Bands(0).Columns("ORDER_QTY").Format = "###,##0"
        grdWBFHORND.DisplayLayout.Bands(0).Columns("SALES").Format = "###,##0.00"

        With grdWBFHORNT.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdWBFHORNT.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBFHORNT.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        With grdWBFHORND.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdWBFHORND.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBFHORND.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        dtFROM.Value = CDate(String.Format("{0}/01/{1}", Now.Month, Now.Year))
        dtTO.Value = CDate(String.Format("{0}/{1}/{2}", Now.Month, Date.DaysInMonth(Now.Year, Now.Month), Now.Year))

        Load_Record()

        Sort_grdColumns(grdWBFHORNT, "SALES".ToLower(), False)
        Sort_grdColumns(grdWBFHORND, "SALES".ToLower(), False)

        tab.Visible = False

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
                Load_Record()
            Case "Exit"
                Call Mode_Settings(False)
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
        SetShowOrderDetails()
        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        dst.Tables("WBTHORNT").Rows.Clear()
    End Sub

    Sub Load_Record()
        'Call Save_Header_Fields(UltraGroupBox1)
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        EnforceConstraints(False)
        Dim SELECTS As String = ""
        Dim EXTRA1 As String = ""
        Dim EXTRA2 As String = ""
        Dim EXTRA3 As String = ""
        Dim GROUPS As String = "'"
        Dim RANKING As String
        Dim RANK_CODE As String
        Dim RANK_NAME As String

        'FromDate = CDate(dtFROM.Value).AddDays(-1)
        FromDate = CDate(dtFROM.Value)
        ToDate = CDate(dtTO.Value).AddDays(1)

        Select Case RankOption
            Case "R"
                RANKING = "Ranking Sales Reps"
                RANK_CODE = "Rep Code"
                RANK_NAME = "Sales Rep Name"
                SQLs.Length = 0
                SQLs.AppendLine("SELECT S1.SREP_CODE AS RANK_CODE, R1.SREP_NAME AS RANK_NAME,")
                SQLs.AppendLine("NULL AS EXTRA1, NULL AS EXTRA2, NULL AS EXTRA3,")
                SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
                SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                SQLs.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
                SQLs.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                SQLs.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                SQLs.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
                SQLs.AppendLine(String.Format("AND S1.ORDR_DATE >= '{0}'", Format(FromDate, "dd-MMM-yyyy")))
                SQLs.AppendLine(String.Format("AND S1.ORDR_DATE < '{0}'", Format(ToDate, "dd-MMM-yyyy")))
                SQLs.AppendLine("GROUP BY S1.SREP_CODE, R1.SREP_NAME")
            Case "C"
                RANKING = "Ranking Customers"
                RANK_CODE = "Cust Code"
                RANK_NAME = "Customer Name"
                EXTRA1 = "City"
                EXTRA2 = "State"
                EXTRA3 = "Country"
                SQLs.Length = 0
                SQLs.AppendLine("SELECT S1.CUST_CODE AS RANK_CODE, S1.CUST_NAME AS RANK_NAME,")
                SQLs.AppendLine("C1.CUST_CITY AS EXTRA1, NVL(C1.CUST_STATE,'XX') AS EXTRA2, NVL(C1.CUST_COUNTRY,'USA') AS EXTRA3,")
                SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
                SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                SQLs.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1, ARTCUST1 C1")
                SQLs.AppendLine("WHERE S1.CUST_CODE = C1.CUST_CODE")
                SQLs.AppendLine("AND S1.ORDR_NO = S2.ORDR_NO")
                SQLs.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                SQLs.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
                SQLs.AppendLine(String.Format("AND S1.ORDR_DATE >= '{0}'", Format(FromDate, "dd-MMM-yyyy")))
                SQLs.AppendLine(String.Format("AND S1.ORDR_DATE < '{0}'", Format(ToDate, "dd-MMM-yyyy")))
                SQLs.AppendLine("GROUP BY S1.CUST_CODE, S1.CUST_NAME,  C1.CUST_CITY,  C1.CUST_STATE,  C1.CUST_COUNTRY")
            Case "S"
                RANKING = "Ranking Styles"
                RANK_CODE = "Style Code"
                RANK_NAME = "Style Description"
                EXTRA1 = "Class"
                EXTRA2 = "Factory"
                EXTRA3 = ""
                SQLs.Length = 0
                SQLs.AppendLine("SELECT S2.STYLE_CODE AS RANK_CODE, S2.STYLE_DESC AS RANK_NAME,")
                SQLs.AppendLine("I1.STYLE_CLASS_CODE AS EXTRA1, V1.VEND_SUPPLIER_ID AS EXTRA2, NULL AS EXTRA3,")
                SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
                SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                SQLs.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1, ICTSTYL1 I1, APTVEND1 V1")
                SQLs.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                SQLs.AppendLine("AND S2.STYLE_CODE = I1.STYLE_CODE")
                SQLs.AppendLine("AND I1.VEND_CODE (+) = V1.VEND_CODE")
                SQLs.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                SQLs.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
                SQLs.AppendLine(String.Format("AND S1.ORDR_DATE >= '{0}'", Format(FromDate, "dd-MMM-yyyy")))
                SQLs.AppendLine(String.Format("AND S1.ORDR_DATE < '{0}'", Format(ToDate, "dd-MMM-yyyy")))
                If chkStylesInventory.Checked Then
                    SQLs.AppendLine("AND (S2.STYLE_CODE, S2.COLOR_CODE) IN ")
                    SQLs.AppendLine("(SELECT")
                    SQLs.AppendLine("STYLE_CODE,")
                    SQLs.AppendLine("COLOR_CODE")
                    SQLs.AppendLine("FROM ICTSTAT2")
                    SQLs.AppendLine("WHERE (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) + NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0)) > 0")
                    SQLs.AppendLine("GROUP BY STYLE_CODE,")
                    SQLs.AppendLine("COLOR_CODE")
                    SQLs.AppendLine(")")
                End If
                SQLs.AppendLine("GROUP BY S2.STYLE_CODE, S2.STYLE_DESC, I1.STYLE_CLASS_CODE, V1.VEND_SUPPLIER_ID")
            Case Else
                RANKING = "Ranking Sales Reps"
                RANK_CODE = "Rep Code"
                RANK_NAME = "Sales Rep Name"
                SQLs.Length = 0
                SQLs.AppendLine("SELECT 'ALL' AS RANK_CODE, R1.SREP_NAME AS RANK_NAME,")
                SQLs.AppendLine("NULL AS EXTRA1, NULL AS EXTRA2, NULL AS EXTRA3,")
                SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
                SQLs.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                SQLs.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
                SQLs.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                SQLs.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                SQLs.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
                SQLs.AppendLine(String.Format("AND S1.ORDR_DATE >= '{0}'", Format(FromDate, "dd-MMM-yyyy")))
                SQLs.AppendLine(String.Format("AND S1.ORDR_DATE < '{0}'", Format(ToDate, "dd-MMM-yyyy")))
                SQLs.AppendLine("GROUP BY R1.SREP_NAME")
        End Select

        Fill_Records("WBTHORNT", , , SQLs.ToString)

        'EnforceConstraints(True)

        grdWBFHORNT.Text = "Hot Or Not " & RANKING
        grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("RANK_CODE").Header.Caption = RANK_CODE
        grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("RANK_NAME").Header.Caption = RANK_NAME

        If EXTRA1.Length > 0 Then
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA1").Header.Caption = EXTRA1
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA1").Hidden = False
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA1").Header.VisiblePosition = 2
        Else
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA1").Header.Caption = ""
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA1").Hidden = True
        End If
        If EXTRA2.Length > 0 Then
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA2").Header.Caption = EXTRA2
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA2").Hidden = False
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA2").Header.VisiblePosition = 3
        Else
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA2").Header.Caption = ""
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA2").Hidden = True
        End If
        If EXTRA3.Length > 0 Then
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA3").Header.Caption = EXTRA3
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA3").Hidden = False
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA3").Header.VisiblePosition = 4
        Else
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA3").Header.Caption = ""
            grdWBFHORNT.DisplayLayout.Bands(0).Columns.Item("EXTRA3").Hidden = True
        End If

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
        FillDetails()
        If grdWBFHORNT.Rows.Count > 0 Then
            grdWBFHORNT.Rows(0).Activate()
        End If

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
        Load_Popup_Menu(grdWBFHORNT, "SSBBB", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Style Masterfile")
        Load_Popup_Menu(grdWBFHORND, "SSBB", "Show Filter", "Show GroupBox", "Cust Order Inq", "Sales Order Inq")
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
                If optRANKS.Checked Then
                    e.Tool.ToolbarsManager.Tools("Style Status Inquiry").SharedProps.Visible = True
                    e.Tool.ToolbarsManager.Tools("Style Masterfile").SharedProps.Visible = True
                Else
                    e.Tool.ToolbarsManager.Tools("Style Status Inquiry").SharedProps.Visible = False
                    e.Tool.ToolbarsManager.Tools("Style Masterfile").SharedProps.Visible = False
                End If

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
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("RANK_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
            Case "Style Masterfile"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("RANK_CODE").Text
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
        Select Case OPTR
            Case "R"
                RankOption = "R"
                optRANKC.Checked = False
                optRANKS.Checked = False
                chkStylesInventory.Visible = False
                chkStylesInventory.Checked = False
            Case "C"
                RankOption = "C"
                optRANKR.Checked = False
                optRANKS.Checked = False
                chkStylesInventory.Visible = False
                chkStylesInventory.Checked = False
            Case "S"
                RankOption = "S"
                optRANKR.Checked = False
                optRANKC.Checked = False
                chkStylesInventory.Visible = True
                chkStylesInventory.Checked = False
        End Select
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
    Private Sub optRANKC_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optRANKC.CheckedChanged
        If optRANKC.Checked Then
            SetRankOption("C")
        End If
    End Sub

    Private Sub optRANKR_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optRANKR.CheckedChanged
        If optRANKR.Checked Then
            SetRankOption("R")
        End If
    End Sub

    Private Sub optRANKS_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optRANKS.CheckedChanged
        If optRANKS.Checked Then
            SetRankOption("S")
        End If
    End Sub

    Private Sub chkShowDetails_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowDetails.CheckedChanged
        SetShowOrderDetails()
    End Sub

    Private Sub grdWBFHORNT_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWBFHORNT.AfterRowActivate
        FillDetails()

    End Sub

    Private Sub FillDetails()
        If chkShowDetails.Checked Then
            If IsNothing(grdWBFHORNT.ActiveRow) Then
                dst.Tables.Item("WBTHORND").Clear()
                grdWBFHORND.Text = "Please Select A Row Above To See Details"
                Exit Sub
            End If
            If Not grdWBFHORNT.ActiveRow Is Nothing And grdWBFHORNT.ActiveRow.IsDataRow Then
                FromDate = CDate(dtFROM.Value)
                ToDate = CDate(dtTO.Value).AddDays(1)
                Dim S As New Text.StringBuilder
                Select Case RankOption
                    Case "R"
                        Dim SREP_CODE As String = grdWBFHORNT.ActiveRow.Cells("RANK_CODE").Text
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
                        S.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
                        S.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                        S.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
                        S.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                        S.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                        S.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
                        S.AppendLine(String.Format("AND S1.SREP_CODE = '{0}'", SREP_CODE))
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
                        grdWBFHORND.Text = "Details For Sales Rep " & SREP_CODE
                    Case "S"
                        Dim STYLE_CODE As String = grdWBFHORNT.ActiveRow.Cells("RANK_CODE").Text
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
                        S.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
                        S.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                        S.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
                        S.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                        S.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                        S.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
                        S.AppendLine(String.Format("AND S2.STYLE_CODE = '{0}'", STYLE_CODE))
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
                        grdWBFHORND.Text = "Details For Style Code " & STYLE_CODE
                    Case Else
                        Dim CUST_CODE As String = grdWBFHORNT.ActiveRow.Cells("RANK_CODE").Text
                        Dim CUST_NAME As String = grdWBFHORNT.ActiveRow.Cells("RANK_NAME").Text
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
                        S.AppendLine("SUM(NVL(S2.ORDR_QTY,0)) AS ORDER_QTY,")
                        S.AppendLine("SUM(NVL(S2.ORDR_QTY,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS SALES")
                        S.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2, SOTSREP1 R1")
                        S.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
                        S.AppendLine("AND  S1.SREP_CODE = R1.SREP_CODE (+)")
                        S.AppendLine("AND  S1.ORDR_STATUS <> 'C'")
                        S.AppendLine(String.Format("AND S1.CUST_CODE = '{0}'", CUST_CODE))
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
                        grdWBFHORND.Text = "Details For " & CUST_NAME
                End Select

                Fill_Records("WBTHORND", , , S.ToString)
            End If
        End If
    End Sub
#End Region
End Class