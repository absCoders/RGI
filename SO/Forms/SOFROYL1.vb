Imports System.Drawing

Public Class SOFROYL1
    Dim sqls As String = ""
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFROYLI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst
            Dim SQLB As New System.Text.StringBuilder
            SQLB.Length = 0
            SQLB.AppendLine("SELECT")
            SQLB.AppendLine("'NONE' AS SEASON,")
            SQLB.AppendLine("SUBSTR(I1.ORDR_YYYYPP_UPDATED,1,4) AS SHIP_YEAR,")
            SQLB.AppendLine("I2.STYLE_CODE,")
            SQLB.AppendLine("I2.ORDR_UNIT_PRICE,")
            SQLB.AppendLine("S1.STYLE_DESC,")
            SQLB.AppendLine("C1.CUST_NAME,")
            SQLB.AppendLine("I1.INV_DATE,")
            SQLB.AppendLine("NVL(A1.CUST_COUNTRY,'USA') AS COUNTRY,")
            SQLB.AppendLine("SUM(I2.ORDR_QTY_SHIP) AS ORDR_QTY_SHIP,")
            SQLB.AppendLine("SUM(I2.ORDR_UNIT_PRICE * I2.ORDR_QTY_SHIP) AS INVOICED")
            SQLB.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ICTSTYL1 S1, ARTCUST1 C1, ARTCUST1 A1")
            SQLB.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
            SQLB.AppendLine("AND I1.CUST_CODE = A1.CUST_CODE")
            SQLB.AppendLine("AND I1.INV_NO = I2.INV_NO")
            SQLB.AppendLine("AND I2.STYLE_CODE = S1.STYLE_CODE")
            SQLB.AppendLine("AND I1.CUST_CODE = C1.CUST_CODE")
            SQLB.AppendLine("AND S1.SALES_DIVISION_CODE = :PARM1")
            SQLB.AppendLine("AND I1.ORDR_YYYYPP_UPDATED >= :PARM2")
            SQLB.AppendLine("AND I1.ORDR_YYYYPP_UPDATED <= :PARM3")
            SQLB.AppendLine("AND I1.INV_TYPE = :PARM4")
            SQLB.AppendLine("AND I2.ORDR_QTY_SHIP <> 0")
            SQLB.AppendLine("GROUP BY")
            SQLB.AppendLine("SUBSTR(I1.ORDR_YYYYPP_UPDATED,1,4),")
            SQLB.AppendLine("I2.STYLE_CODE,")
            SQLB.AppendLine("I2.ORDR_UNIT_PRICE,")
            SQLB.AppendLine("S1.STYLE_DESC,")
            SQLB.AppendLine("C1.CUST_NAME,")
            SQLB.AppendLine("I1.INV_DATE,")
            SQLB.AppendLine("NVL(A1.CUST_COUNTRY,'USA')")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTROYLI", "**", 0, False, "VVVV")
            With .Tables("SOTROYLI").Columns
                .Add("RET_AMT", GetType(System.Decimal))
                .Add("RET_QTY", GetType(System.Decimal))
                .Add("ALLOWANCE", GetType(System.Decimal))
                .Add("ROYALTY_PAID", GetType(System.Decimal))
                .Add("ROYALTY_PCT", GetType(System.Decimal))
                .Add("COLOR_TYPE", GetType(System.String))
                .Add("FOB", GetType(System.String))
            End With


            SQLB.Length = 0
            SQLB.AppendLine("SELECT * FROM SOTSDIV1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTSDIV1", "**", 0, False)

        End With

        grdSOTROYL1.DataSource = dst.Tables("SOTROYLI")

        Fill_Records("SOTSDIV1")
        Dim salesCodes As New List(Of String)
        For Each rowSOTSDIV1 As DataRow In dst.Tables("SOTSDIV1").Select()
            salesCodes.Add(rowSOTSDIV1.Item("SALES_DIVISION_CODE").ToString())
        Next
        cboSALES_DIVISION_CODE.DataSource = salesCodes

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -72, 0, 0)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -72, 0, 0)

        Sort_grdColumns(grdSOTROYL1, "INV_DATE", False)

        grdSOTROYL1.DisplayLayout.Bands(0).Columns("ORDR_QTY_SHIP").Format = "###,##0"

        'ASCMAIN1.Add_Value_List(grdSOTROYL1, "REPORT_TYPE", , New String() {":", "I:Initial", "A:Amended", "S:Subsequent", "R:Revised"})

        Create_Summary(grdSOTROYL1, "ORDR_QTY_SHIP")
        Create_Summary(grdSOTROYL1, "INVOICED")

        TABLE_NAME = "SOTROYLI"

        EntryMode = "E"
        'Call Load_Record()
        Call Mode_Settings(True)

        txtPct.Text = Format(7.0, "###,##0.00")
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
            '.Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        'dst.EnforceConstraints = False
        'dst.Tables("PMTVIST1").Rows.Clear()
        'dst.Tables("PMTVISTH").Rows.Clear()

        'Dim dvw As DataView = DirectCast(grdPMTVIST1.DataSource, DataTable).DefaultView
        'dvw.RowStateFilter = DataViewRowState.CurrentRows

        'Fill_Records("PMTVIST1")
        'Process_SVRs()

        'Sort_grdColumns(grdPMTVIST1, "DATE_VISITED".ToLower)
        'Sort_grdColumns(grdPMTVISTH, "DATE_VISITED".ToLower)
        'dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        dst.Tables("SOTROYLI").Rows.Clear()

        dst.EnforceConstraints = False

        Fill_Records("SOTROYLI")
        'dst.EnforceConstraints = True

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
        ReLoadData()
        grdSOTROYL1.Update()
        grdSOTROYL1.Refresh()
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
        Call Load_Popup_Menu(grdSOTROYL1, "SSB", "Show Filter", "Show GroupBox")
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
        'Select Case Absx1.GetABSColumnName(sender)
        'Case "EMPLOYEE_CODE"
        '    If e.KeyCode = Windows.Forms.Keys.Enter Then
        '        Setup_Summary()
        '    End If
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        'Select Case COLUMN_NAME
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

#End Region

#Region "Custom Methods"

    Private Sub CalculateFields(ByVal DataType As String)
        Dim pct As Double = Val(txtPct.Text)
        Select Case DataType
            Case Is = "I"
                For Each rowSOTROYLI As DataRow In dst.Tables("SOTROYLI").Select()
                    If rowSOTROYLI.Item("COLOR_TYPE").ToString() = "" Then
                        rowSOTROYLI.Item("COLOR_TYPE") = "G"
                    End If
                    rowSOTROYLI.Item("ROYALTY_PCT") = Format(pct, "###,##0.00")
                    Dim INVOICED As Double = Val(rowSOTROYLI.Item("INVOICED") & "")
                    Dim ROYALTY_PAID As Double = INVOICED * (pct / 100)
                    rowSOTROYLI.Item("ROYALTY_PAID") = Format(ROYALTY_PAID, "###,##0.00")
                Next
        End Select
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTROYL1.Rows
            If grow.Cells.Item("COLOR_TYPE").Text = "G" Then
                For Each ColName As String In New String() {"SEASON", _
                                                            "SHIP_YEAR", _
                                                            "STYLE_CODE", _
                                                            "ORDR_UNIT_PRICE", _
                                                            "STYLE_DESC", _
                                                            "CUST_NAME", _
                                                            "INV_DATE", _
                                                            "COUNTRY", _
                                                            "ORDR_QTY_SHIP", _
                                                            "INVOICED", _
                                                            "FOB"}
                    grow.Cells.Item(ColName).Appearance.BackColor = Color.LightBlue
                Next
            End If
        Next
    End Sub

    Private Sub ReLoadData()

    End Sub

#End Region

#Region "Form Controls"

    Private Sub btnFETCH_Click(sender As Object, e As EventArgs) Handles btnFETCH.Click
        If Not IsNumeric(txtPct.Text) Then
            MsgBox("Default % Is Not Numeric", vbOKOnly, "Problem")
            Exit Sub
        End If

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        Dim SALES_DIVISION_CODE As String = cboSALES_DIVISION_CODE.Text
        Dim ORDR_YYYYPP_FROM As String = Absx1.cmbFor("RYP0").Text
        Dim ORDR_YYYYPP_TO As String = Absx1.cmbFor("RYP1").Text
        If ORDR_YYYYPP_FROM.Length >= 7 Then
            ORDR_YYYYPP_FROM = ORDR_YYYYPP_FROM.Substring(0, 4) & ORDR_YYYYPP_FROM.Substring(5, 2)
        Else
            Exit Sub
        End If
        If ORDR_YYYYPP_TO.Length >= 7 Then
            ORDR_YYYYPP_TO = ORDR_YYYYPP_FROM.Substring(0, 4) & ORDR_YYYYPP_TO.Substring(5, 2)
        Else
            Exit Sub
        End If

        Dim INV_TYPE As String = "I"
        Fill_Records("SOTROYLI", New Object() {SALES_DIVISION_CODE, ORDR_YYYYPP_FROM, ORDR_YYYYPP_TO, INV_TYPE}, True)
        CalculateFields(INV_TYPE)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cboSALES_DIVISION_CODE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSALES_DIVISION_CODE.SelectedIndexChanged
        If cboSALES_DIVISION_CODE.Text <> "" Then
            Dim filter As String = String.Format("SALES_DIVISION_CODE = '{0}'", cboSALES_DIVISION_CODE.Text)
            txtSALES_DIVISION_NAME.Text = dst.Tables.Item("SOTSDIV1").Select(filter).FirstOrDefault().Item("SALES_DIVISION_NAME").ToString()
        Else
            txtSALES_DIVISION_NAME.Text = ""
        End If
    End Sub

    Private Sub txtPct_LostFocus(sender As Object, e As EventArgs) Handles txtPct.LostFocus
        If IsNumeric(txtPct.Text) Then
            txtPct.Text = Format(Val(txtPct.Text), "###,##0.00")
        Else
            MsgBox("Default % Is Not Numeric", vbOKOnly, "Problem")
            txtPct.Text = Format(0, "###,##0.00")
        End If

    End Sub

#End Region
End Class