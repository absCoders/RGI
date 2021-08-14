Imports Infragistics.Win.UltraWinGrid

Public Class SOFSHPWC
    Dim SQ As New System.Text.StringBuilder
    Dim _WMWEEKS As New List(Of WMWEEKS)
    Dim SplitDist As Int64 = 0
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFSHPWI" Then
            InquiryMode = True
        End If

        MakeWMWEEKS

        dteORDR_DATE_FROM.DateTime = CDate(Now().ToShortDateString)
        dteORDR_DATE_TO.DateTime = CDate(Now().ToShortDateString)

        Check_Form_Options()
        With dst
            SQ.Length = 0
            SQ.AppendLine("SELECT")
            SQ.AppendLine(" 99 as WALMART,")
            SQ.AppendLine(" OX.ORDR_CUST_PO,")
            SQ.AppendLine(" OX.ORDR_DATE,")
            SQ.AppendLine(" I1.INV_DATE,")
            SQ.AppendLine(" I1.CUST_CODE,")
            SQ.AppendLine(" OX.CUST_NAME,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY_ORIG) AS ORDR_QTY_ORIG,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY) AS ORDR_QTY,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY_ALLO) AS ORDR_QTY_ALLO,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY_CANC) AS ORDR_QTY_CANC,")
            SQ.AppendLine(" SUM(ORDR_QTY_CANC_WHSE) AS ORDR_QTY_CANC_WHSE,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY_SHIP) AS ORDR_QTY_SHIP,")
            SQ.AppendLine(" SUM(NVL(I2.ORDR_QTY_SHIP,0)) AS INVOICED")
            SQ.AppendLine(" FROM SOTINVH1 I1, SOTINVH2 I2,")
            SQ.AppendLine(" (")
            SQ.AppendLine("   SELECT")
            SQ.AppendLine("   O1.ORDR_DATE,")
            SQ.AppendLine("   O1.ORDR_NO,")
            SQ.AppendLine("   O2.ORDR_LNO,")
            SQ.AppendLine("   O1.CUST_CODE,")
            SQ.AppendLine("   C1.CUST_NAME,")
            SQ.AppendLine("   O1.ORDR_CUST_PO,")
            SQ.AppendLine("   O2.ORDR_QTY_ORIG,")
            SQ.AppendLine("   O2.ORDR_QTY,")
            SQ.AppendLine("   O2.ORDR_QTY_CANC,")
            SQ.AppendLine("   O2.ORDR_QTY_ALLO,")
            SQ.AppendLine("   (O2.ORDR_QTY_CANC - (O2.ORDR_QTY - O2.ORDR_QTY_ALLO)) AS ORDR_QTY_CANC_WHSE,")
            SQ.AppendLine("   O2.ORDR_QTY_SHIP")
            SQ.AppendLine("   FROM SOTORDR1 O1, SOTORDR2 O2, ARTCUST1 C1")
            SQ.AppendLine("   WHERE O1.ORDR_NO = O2.ORDR_NO")
            SQ.AppendLine("   AND O1.CUST_CODE = C1.CUST_CODE")
            SQ.AppendLine("   AND O1.CUST_CODE = 'WALMART'")
            SQ.AppendLine("   AND (O1.ORDR_DATE >= :PARM1 AND O1.ORDR_DATE <= :PARM2)")
            SQ.AppendLine(" ) OX")
            SQ.AppendLine(" WHERE I1.INV_TYPE = I2.INV_TYPE")
            SQ.AppendLine(" AND I1.INV_NO = I2.INV_NO")
            SQ.AppendLine(" AND I1.ORDR_NO (+) = OX.ORDR_NO")
            SQ.AppendLine(" AND I2.INV_LNO (+) = OX.ORDR_LNO")
            SQ.AppendLine(" GROUP BY")
            SQ.AppendLine(" OX.ORDR_CUST_PO,")
            SQ.AppendLine(" OX.ORDR_DATE,")
            SQ.AppendLine(" I1.INV_DATE,")
            SQ.AppendLine(" I1.CUST_CODE,")
            SQ.AppendLine(" OX.CUST_NAME")
            SQ.AppendLine(" ORDER BY")
            SQ.AppendLine(" OX.ORDR_CUST_PO,")
            SQ.AppendLine(" OX.ORDR_DATE,")
            SQ.AppendLine(" I1.INV_DATE,")
            SQ.AppendLine(" I1.CUST_CODE,")
            SQ.AppendLine(" OX.CUST_NAME")
            ASCMAIN1.sql = SQ.ToString
            Create_TDA(.Tables.Add, "SOTSHPWX", "**", 0, False, "DD", 0)

            SQ.Length = 0
            SQ.AppendLine("SELECT")
            SQ.AppendLine(" 99 as WALMART,")
            SQ.AppendLine(" OX.ORDR_CUST_PO,")
            SQ.AppendLine(" OX.ORDR_DATE,")
            SQ.AppendLine(" I1.INV_DATE,")
            SQ.AppendLine(" I1.CUST_CODE,")
            SQ.AppendLine(" OX.CUST_NAME,")
            SQ.AppendLine(" OX.STYLE_CODE,")
            SQ.AppendLine(" OX.COLOR_CODE,")
            SQ.AppendLine(" OX.STYLE_DESC,")
            SQ.AppendLine(" OX.CUST_SKU,")
            SQ.AppendLine(" OX.CUST_COLOR_CODE,")
            SQ.AppendLine(" OX.CUST_SIZE_CODE,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY_ORIG) AS ORDR_QTY_ORIG,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY) AS ORDR_QTY,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY_ALLO) AS ORDR_QTY_ALLO,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY_CANC) AS ORDR_QTY_CANC,")
            SQ.AppendLine(" SUM(ORDR_QTY_CANC_WHSE) AS ORDR_QTY_CANC_WHSE,")
            SQ.AppendLine(" SUM(OX.ORDR_QTY_SHIP) AS ORDR_QTY_SHIP,")
            SQ.AppendLine(" SUM(NVL(I2.ORDR_QTY_SHIP,0)) AS INVOICED")
            SQ.AppendLine(" FROM SOTINVH1 I1, SOTINVH2 I2,")
            SQ.AppendLine(" (")
            SQ.AppendLine("   SELECT")
            SQ.AppendLine("   O1.ORDR_DATE,")
            SQ.AppendLine("   O1.ORDR_NO,")
            SQ.AppendLine("   O2.ORDR_LNO,")
            SQ.AppendLine("   O1.CUST_CODE,")
            SQ.AppendLine("   C1.CUST_NAME,")
            SQ.AppendLine("   O1.ORDR_CUST_PO,")
            SQ.AppendLine("   O2.STYLE_CODE,")
            SQ.AppendLine("   O2.COLOR_CODE,")
            SQ.AppendLine("   O2.STYLE_DESC,")
            SQ.AppendLine("   O2.CUST_SKU,")
            SQ.AppendLine("   O2.CUST_COLOR_CODE,")
            SQ.AppendLine("   O2.CUST_SIZE_CODE,")
            SQ.AppendLine("   O2.ORDR_QTY_ORIG,")
            SQ.AppendLine("   O2.ORDR_QTY,")
            SQ.AppendLine("   O2.ORDR_QTY_CANC,")
            SQ.AppendLine("   O2.ORDR_QTY_ALLO,")
            SQ.AppendLine("   (O2.ORDR_QTY_CANC - (O2.ORDR_QTY - O2.ORDR_QTY_ALLO)) AS ORDR_QTY_CANC_WHSE,")
            SQ.AppendLine("   O2.ORDR_QTY_SHIP")
            SQ.AppendLine("   FROM SOTORDR1 O1, SOTORDR2 O2, ARTCUST1 C1")
            SQ.AppendLine("   WHERE O1.ORDR_NO = O2.ORDR_NO")
            SQ.AppendLine("   AND O1.CUST_CODE = C1.CUST_CODE")
            SQ.AppendLine("   AND O1.CUST_CODE = 'WALMART'")
            'SQ.AppendLine("   AND O1.ORDR_CUST_PO = '6804175082'")
            SQ.AppendLine("   AND (O1.ORDR_DATE >= :PARM1 AND O1.ORDR_DATE <= :PARM2)")
            SQ.AppendLine(" ) OX")
            SQ.AppendLine(" WHERE I1.INV_TYPE = I2.INV_TYPE")
            SQ.AppendLine(" AND I1.INV_NO = I2.INV_NO")
            SQ.AppendLine(" AND I1.ORDR_NO (+) = OX.ORDR_NO")
            SQ.AppendLine(" AND I2.INV_LNO (+) = OX.ORDR_LNO")
            SQ.AppendLine(" GROUP BY")
            SQ.AppendLine(" OX.ORDR_CUST_PO,")
            SQ.AppendLine(" OX.ORDR_DATE,")
            SQ.AppendLine(" I1.INV_DATE,")
            SQ.AppendLine(" I1.CUST_CODE,")
            SQ.AppendLine(" OX.CUST_NAME,")
            SQ.AppendLine(" OX.STYLE_CODE,")
            SQ.AppendLine(" OX.COLOR_CODE,")
            SQ.AppendLine(" OX.STYLE_DESC,")
            SQ.AppendLine(" OX.CUST_SKU,")
            SQ.AppendLine(" OX.CUST_COLOR_CODE,")
            SQ.AppendLine(" OX.CUST_SIZE_CODE")
            'SQ.AppendLine(" HAVING SUM(ORDR_QTY_CANC_WHSE) > :PARM1")
            SQ.AppendLine(" ORDER BY")
            SQ.AppendLine(" OX.ORDR_CUST_PO,")
            SQ.AppendLine(" OX.ORDR_DATE,")
            SQ.AppendLine(" I1.INV_DATE,")
            SQ.AppendLine(" I1.CUST_CODE,")
            SQ.AppendLine(" OX.CUST_NAME,")
            SQ.AppendLine(" OX.STYLE_CODE,")
            SQ.AppendLine(" OX.COLOR_CODE")
            ASCMAIN1.sql = SQ.ToString
            Create_TDA(.Tables.Add, "SOTSHPWC", "**", 0, False, "DD", 0)
            '.Tables("SOTORDRS").Columns.Add("ORDR_UNIT_PRICE", GetType(System.Decimal), "IIF(ISNULL(ORDR_QTY,0)=0,0,ISNULL(ORDR_AMT,0) / ISNULL(ORDR_QTY,0))")
        End With

        grdSOTSHPWC.DataSource = dst.Tables("SOTSHPWC")
        grdSOTSHPWX.DataSource = dst.Tables("SOTSHPWX")

        'Sort_grdColumns(grdSOTSHPWC, "ORDR_YYYYPP_BOOKED, ORDR_GROUP_NO", False)

        TABLE_NAME = "SOTSHPWC"

        EntryMode = "E"
        'Call Load_Record()

        With grdSOTSHPWC.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"WALMART", "ORDR_QTY_ORIG", "ORDR_QTY", "ORDR_QTY_ALLO", "ORDR_QTY_CANC", "ORDR_QTY_CANC_WHSE", "ORDR_QTY_SHIP", "INVOICED"}
                .Columns(COLUMN_NAME).Format = "#,###,##0"
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_ORIG", "ORDR_QTY", "ORDR_QTY_ALLO", "ORDR_QTY_CANC", "ORDR_QTY_CANC_WHSE", "ORDR_QTY_SHIP", "INVOICED"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"CUST_SKU", "CUST_COLOR_CODE", "CUST_SIZE_CODE"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_DATE", "INV_DATE"}
                .Columns(COLUMN_NAME).Format = "MM/dd/yy"
            Next
        End With

        With grdSOTSHPWX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"WALMART", "ORDR_QTY_ORIG", "ORDR_QTY", "ORDR_QTY_ALLO", "ORDR_QTY_CANC", "ORDR_QTY_CANC_WHSE", "ORDR_QTY_SHIP", "INVOICED"}
                .Columns(COLUMN_NAME).Format = "#,###,##0"
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_ORIG", "ORDR_QTY", "ORDR_QTY_ALLO", "ORDR_QTY_CANC", "ORDR_QTY_CANC_WHSE", "ORDR_QTY_SHIP", "INVOICED"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_DATE", "INV_DATE"}
                .Columns(COLUMN_NAME).Format = "MM/dd/yy"
            Next
        End With

        For Each COLUMN_NAME As String In New String() {"ORDR_QTY_ORIG", "ORDR_QTY", "ORDR_QTY_ALLO", "ORDR_QTY_CANC", "ORDR_QTY_CANC_WHSE", "ORDR_QTY_SHIP", "INVOICED"}
            Create_Summary(grdSOTSHPWC, COLUMN_NAME)
        Next

        For Each COLUMN_NAME As String In New String() {"ORDR_QTY_ORIG", "ORDR_QTY", "ORDR_QTY_ALLO", "ORDR_QTY_CANC", "ORDR_QTY_CANC_WHSE", "ORDR_QTY_SHIP", "INVOICED"}
            Create_Summary(grdSOTSHPWX, COLUMN_NAME)
        Next


        'Bind_Controls(grpHeader, "SOTSHPWH")

        'ASCMAIN1.Add_Value_List(grdSOTSHPWC, "XXXXXXX", , New String() {":", "A:AAAAAAA", "B:BBBBBBB"})

        Call Mode_Settings(True)

    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Sub Check_Form_Options()
        'With UltraExplorerBar1.Groups("Screen Control")
        '    .Items("New").Visible = (Me.Name = "PMFVIST1")
        'End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Update"
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Done"
                Call Mode_Settings(False)
                Me.Close()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        'Absx1.txtFor("ORDR_GROUP_NO").Text = ""

        dst.EnforceConstraints = False
        For Each TABLE_NAME As String In New String() _
            {"SOTSHPWC", "SOTSHPWX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        dst.EnforceConstraints = True
    End Sub

    Sub Load_Record()

        'tab.Visible = ScreenMode

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        'dst.Tables("SOTROYLI").Rows.Clear()

        dst.EnforceConstraints = False

        dst.EnforceConstraints = True

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()
        'BeginTrans()
        'INIT_LAST("PMTVIST1", True, "", True)
        'Update_Record_TDA("PMTVIST1")
        'CommitTrans("Update Complete")
    End Sub

    Sub Setup_Summary()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOTSHPWX, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdSOTSHPWC, "SS", "Show Filter", "Show GroupBox")
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
        'Dim tlb_btn As UltraWinToolbars.ButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""

                    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                        e.Cancel = True
                    End If
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
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"
    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

#End Region

#Region "Custom Methods"
    Private Sub CalculateWMWeeks()
        For Each TABLES As String In New String() {"SOTSHPWX", "SOTSHPWC"}
            For Each row As DataRow In dst.Tables(TABLES).Select()
                For Each W As WMWEEKS In _WMWEEKS
                    If CDate(row.Item("ORDR_DATE").ToString & String.Empty) >= W.BWEEK And CDate(row.Item("ORDR_DATE").ToString & String.Empty) <= W.EWEEK Then
                        row.Item("WALMART") = W.WEEKNO
                        Exit For
                    End If
                Next
            Next
        Next
    End Sub

    Private Sub filterSelected()
        Dim Fltr As String = ""
        Dim Fltrx As String = ""

        Dim dvw As DataView = DirectCast(grdSOTSHPWC.DataSource, DataTable).DefaultView
        Dim dvwx As DataView = DirectCast(grdSOTSHPWX.DataSource, DataTable).DefaultView

        If chkOnlyWhse.Checked Then
            Fltr = "ORDR_QTY_CANC_WHSE > 0"
            Fltrx = "ORDR_QTY_CANC_WHSE > 0"
        End If

        If chkShowAll.Checked Then
            dvw.RowFilter = String.Format(Fltr)
        Else
            If Not IsNothing(grdSOTSHPWX.ActiveRow) Then
                Dim ORDR_CUST_PO As String = grdSOTSHPWX.ActiveRow.Cells.Item("ORDR_CUST_PO").Text & String.Empty
                If Fltr.Length > 0 Then
                    Fltr = Fltr & $" AND ORDR_CUST_PO = '{ORDR_CUST_PO}'"
                Else
                    Fltr = $"ORDR_CUST_PO = '{ORDR_CUST_PO}'"
                End If
                dvw.RowFilter = String.Format(Fltr)
                dvwx.RowFilter = String.Format(Fltrx)
            Else
                dvw.RowFilter = String.Format(Fltr)
                dvwx.RowFilter = String.Format(Fltrx)
            End If
        End If
    End Sub

    Private Sub MakeWMWEEKS()
        _WMWEEKS.Clear()
        Dim YRS As New Dictionary(Of Int64, Int64)
        YRS.Add(2019, 53)
        YRS.Add(2020, 52)
        YRS.Add(2021, 52)
        YRS.Add(2022, 52)
        YRS.Add(2023, 53)
        YRS.Add(2024, 52)
        Dim BD As Date = DateSerial(2019, 1, 26)
        Dim ED As Date = BD.AddDays(6)
        For Each Y As KeyValuePair(Of Int64, Int64) In YRS
            For wk As Int64 = 1 To Y.Value
                Dim newWMWEEK As New WMWEEKS
                newWMWEEK.YEAR = Y.Key
                newWMWEEK.WEEKNO = wk
                newWMWEEK.BWEEK = BD
                newWMWEEK.EWEEK = ED
                _WMWEEKS.Add(newWMWEEK)
                BD = BD.AddDays(7)
                ED = ED.AddDays(7)
            Next
        Next

    End Sub
#End Region

#Region "Form Controls"

    Private Sub btnFETCH_Click(sender As Object, e As EventArgs) Handles btnFETCH.Click
        Me.Cursor = Cursors.WaitCursor
        If dteORDR_DATE_FROM.DateTime > dteORDR_DATE_TO.DateTime Then
            MsgBox("Invalid From / To Dates", vbExclamation, "Check Your Dates")
            Exit Sub
        End If
        If DateDiff(DateInterval.Day, dteORDR_DATE_FROM.DateTime, dteORDR_DATE_TO.DateTime) > 30 Then
            MsgBox("You Can Not Request More Than 30 Days Of Data.", vbExclamation, "Too Much!!")
            Exit Sub
        End If

        Dim FDate As Date = CDate(dteORDR_DATE_FROM.DateTime.Date.ToShortDateString)
        Dim TDate As String = CDate(dteORDR_DATE_TO.DateTime.Date.ToShortDateString)

        Fill_Records("SOTSHPWX", New String() {FDate, TDate})
        Fill_Records("SOTSHPWC", New String() {FDate, TDate})

        CalculateWMWeeks()

        filterSelected()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub chkOnlyWhse_CheckedChanged(sender As Object, e As EventArgs) Handles chkOnlyWhse.CheckedChanged
        filterSelected()
    End Sub

    Private Sub chkShowAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAll.CheckedChanged
        If chkShowAll.Checked Then
            SplitDist = SplitContainer1.SplitterDistance
            SplitContainer1.Panel1.Visible = False
            SplitContainer1.SplitterDistance = 0
        Else
            SplitContainer1.Panel1.Visible = True
            SplitContainer1.SplitterDistance = SplitDist
        End If
        filterSelected()
    End Sub

    Private Sub grdSOTSHPWX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTSHPWX.AfterRowActivate
        filterSelected()
    End Sub
#End Region

End Class

Public Class WMWEEKS
    Public YEAR As Int64
    Public WEEKNO As Int64
    Public BWEEK As Date
    Public EWEEK As Date
End Class