Public Class WBFPACKS
    Dim S As New System.Text.StringBuilder With {.Length = 0}
    Dim TEMP_TABLE As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "WBFPACKI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        MakeTempTable(False)

        With dst

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("S1.STYLE_CODE,")
            S.AppendLine("S1.STYLE_DESC,")
            S.AppendLine("S1.CUST_CODE,")
            S.AppendLine("S1.DUTY_RATE_CODE,")
            S.AppendLine("S1.STYLE_SO_QTY_MIN,")
            S.AppendLine("S1.INIT_DATE,")
            S.AppendLine("NVL(S1.CARTON_PACK_QTY,0) AS CARTON_PACK_QTY,")
            S.AppendLine("NVL(S1.INNER_PACK_QTY,0) AS INNER_PACK_QTY,")
            S.AppendLine("S1.VEND_CODE,")
            S.AppendLine("S1.STYLE_MATL_DESC,")
            S.AppendLine("LENGTH_CTN,")
            S.AppendLine("WIDTH_CTN,")
            S.AppendLine("HEIGHT_CTN,")
            S.AppendLine("WEIGHT_CTN,")
            S.AppendLine("LENGTH_INR,")
            S.AppendLine("WIDTH_INR,")
            S.AppendLine("HEIGHT_INR,")
            S.AppendLine("WEIGHT_INR,")
            S.AppendLine("LENGTH_IT,")
            S.AppendLine("WIDTH_IT,")
            S.AppendLine("HEIGHT_IT,")
            S.AppendLine("WEIGHT_IT")
            S.AppendLine(String.Format("FROM ICTSTYL1 S1, {0} P1", TEMP_TABLE))
            S.AppendLine("WHERE S1.STYLE_CODE = P1.STYLE_CODE (+)")
            S.AppendLine("AND S1.STYLE_STATUS = 'A'")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "WBTPACKS", "**", 0, False)
            With .Tables("WBTPACKS").Columns
                .Add("LAST_SALE", GetType(Date))
            End With

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("SOTINVH2.STYLE_CODE,")
            S.AppendLine("MAX(INV_DATE) AS LAST_SALE")
            S.AppendLine("FROM SOTINVH1, SOTINVH2")
            S.AppendLine("WHERE SOTINVH1.INV_NO = SOTINVH2.INV_NO")
            S.AppendLine("and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
            S.AppendLine("GROUP BY SOTINVH2.STYLE_CODE")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "WBTPACKD", "**", 0, False)
            Fill_Records("WBTPACKD")

        End With

        grdWBTPACKS.DataSource = dst.Tables("WBTPACKS")

        'ASCMAIN1.Add_Value_List(grdSOFCSTMX, "REPORT_TYPE", , New String() {":", "I:Initial", "A:Amended", "S:Subsequent", "R:Revised"})

        Create_Summary(grdWBTPACKS, "STYLE_CODE", "Count")

        Sort_grdColumns(grdWBTPACKS, "STYLE_CODE", False)

        With grdWBTPACKS.DisplayLayout.Bands(0)
            For Each COL_NAME As String In New String() {"STYLE_CODE"}
                .Columns(COL_NAME).Header.Fixed = True
            Next
        End With

        grdWBTPACKS.DisplayLayout.Bands(0).Columns("CARTON_PACK_QTY").Format = "###,##0"
        grdWBTPACKS.DisplayLayout.Bands(0).Columns("INNER_PACK_QTY").Format = "###,##0"

        TABLE_NAME = "WBTPACKS"

        EntryMode = "E"
        'Call Load_Record()
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

        Setup_Summary()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()

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
        Call Load_Popup_Menu(grdWBTPACKS, "SS", "Show Filter", "Show GroupBox")
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

        'Select Case e.Tool.Key
        '    Case "Project Center"
        '        Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
        '        Context_Launch("Edit", Column_Values("JOB_NO", JOB_NO), e.Tool.Key, "PMFJOBM1")
        '    Case "Show Report"
        '        Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\randfromdrc\RandInvoices\310 West 52nd Street - 30760.pdf"
        '        Show_Document(FILENAME)

        'End Select
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

    Sub Setup_Summary()
        Dim sqlwhere As String = ""
        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        dst.Tables("WBTPACKS").Rows.Clear()

        MakeTempTable(True)

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("S1.STYLE_CODE,")
        S.AppendLine("S1.STYLE_DESC,")
        S.AppendLine("S1.CUST_CODE,")
        S.AppendLine("S1.DUTY_RATE_CODE,")
        S.AppendLine("S1.STYLE_SO_QTY_MIN,")
        S.AppendLine("S1.INIT_DATE,")
        S.AppendLine("NVL(S1.CARTON_PACK_QTY,0) AS CARTON_PACK_QTY,")
        S.AppendLine("NVL(S1.INNER_PACK_QTY,0) AS INNER_PACK_QTY,")
        S.AppendLine("S1.VEND_CODE,")
        S.AppendLine("S1.STYLE_MATL_DESC,")
        S.AppendLine("LENGTH_CTN,")
        S.AppendLine("WIDTH_CTN,")
        S.AppendLine("HEIGHT_CTN,")
        S.AppendLine("WEIGHT_CTN,")
        S.AppendLine("LENGTH_INR,")
        S.AppendLine("WIDTH_INR,")
        S.AppendLine("HEIGHT_INR,")
        S.AppendLine("WEIGHT_INR,")
        S.AppendLine("LENGTH_IT,")
        S.AppendLine("WIDTH_IT,")
        S.AppendLine("HEIGHT_IT,")
        S.AppendLine("WEIGHT_IT")
        S.AppendLine(String.Format("FROM ICTSTYL1 S1, {0} P1", TEMP_TABLE))
        S.AppendLine("WHERE S1.STYLE_CODE = P1.STYLE_CODE (+)")
        S.AppendLine("AND S1.STYLE_STATUS = 'A'")

        Fill_Records("WBTPACKS")

        FillLastSalesDate()

        RemoveRecords()

        dst.EnforceConstraints = False

        ASCMAIN1.Progress("")
        grdWBTPACKS.Update()
        grdWBTPACKS.Refresh()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        Setup_Summary()
    End Sub

    Private Sub RemoveRecords()
        For Each rowWBTPACKS As DataRow In dst.Tables("WBTPACKS").Select()
            Dim STYLE_CODE As String = rowWBTPACKS.Item("STYLE_CODE").ToString & String.Empty
            Dim CUST_CODE As String = rowWBTPACKS.Item("CUST_CODE").ToString & String.Empty
            Dim LAST_SALE As Date = DateSerial(2100, 1, 1)
            If IsDate(rowWBTPACKS.Item("LAST_SALE").ToString & String.Empty) Then
                LAST_SALE = CDate(rowWBTPACKS.Item("LAST_SALE").ToString & String.Empty)
            End If

            Dim removeRow As Boolean = False

            If chkNOMTB.Checked Then
                If STYLE_CODE.Length > 3 Then
                    If STYLE_CODE.Substring(0, 3) = "MTB" Then
                        removeRow = True
                    End If
                End If
            End If

            If chkSTOCK_ONLY.Checked Then
                If CUST_CODE.Length > 0 Then
                    removeRow = True
                End If
            End If

            If chkNQKS.Checked Then
                If STYLE_CODE.Length > 0 Then
                    If STYLE_CODE.EndsWith("N") Or STYLE_CODE.EndsWith("Q") Or STYLE_CODE.EndsWith("K") Or STYLE_CODE.EndsWith("S") Then
                        removeRow = True
                    End If
                End If
            End If

            If chkSALES_YEARS.Checked Then
                Dim LAST_DATE_AS As Date = Now().AddYears(Val(txtSALES_YEARS.Text) * -1)
                If LAST_SALE <= LAST_DATE_AS Then
                    removeRow = True
                End If
            End If

            If removeRow Then
                rowWBTPACKS.Delete()
            End If
        Next
        dst.Tables("WBTPACKS").AcceptChanges()
    End Sub

    Private Sub MakeTempTable(ByVal CreateEmpty As Boolean)
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("STYLE_CODE,")
        S.AppendLine("SUM(LENGTH_CTN) AS LENGTH_CTN,")
        S.AppendLine("SUM(WIDTH_CTN) AS WIDTH_CTN,")
        S.AppendLine("SUM(HEIGHT_CTN) AS HEIGHT_CTN,")
        S.AppendLine("SUM(WEIGHT_CTN) AS WEIGHT_CTN,")
        S.AppendLine("SUM(LENGTH_INR) AS LENGTH_INR,")
        S.AppendLine("SUM(WIDTH_INR) AS WIDTH_INR,")
        S.AppendLine("SUM(HEIGHT_INR) AS HEIGHT_INR,")
        S.AppendLine("SUM(WEIGHT_INR) AS WEIGHT_INR,")
        S.AppendLine("SUM(LENGTH_IT) AS LENGTH_IT,")
        S.AppendLine("SUM(WIDTH_IT) AS WIDTH_IT,")
        S.AppendLine("SUM(HEIGHT_IT) AS HEIGHT_IT,")
        S.AppendLine("SUM(WEIGHT_IT) AS WEIGHT_IT")
        S.AppendLine("FROM")
        S.AppendLine("(")
        S.AppendLine("SELECT")
        S.AppendLine("STYLE_CODE,")
        S.AppendLine("LENGTH AS LENGTH_CTN,")
        S.AppendLine("WIDTH AS WIDTH_CTN,")
        S.AppendLine("HEIGHT AS HEIGHT_CTN,")
        S.AppendLine("WEIGHT AS WEIGHT_CTN,")
        S.AppendLine("0 AS LENGTH_INR,")
        S.AppendLine("0 AS WIDTH_INR,")
        S.AppendLine("0 AS HEIGHT_INR,")
        S.AppendLine("0 AS WEIGHT_INR,")
        S.AppendLine("0 AS LENGTH_IT,")
        S.AppendLine("0 AS WIDTH_IT,")
        S.AppendLine("0 AS HEIGHT_IT,")
        S.AppendLine("0 AS WEIGHT_IT")
        S.AppendLine("FROM ICTSTYLD ")
        S.AppendLine("WHERE PACK_CODE = 'CTN'")
        S.AppendLine("UNION")
        S.AppendLine("SELECT")
        S.AppendLine("STYLE_CODE,")
        S.AppendLine("0 AS LENGTH_CTN,")
        S.AppendLine("0 AS WIDTH_CTN,")
        S.AppendLine("0 AS HEIGHT_CTN,")
        S.AppendLine("0 AS WEIGHT_CTN,")
        S.AppendLine("LENGTH AS LENGTH_INR,")
        S.AppendLine("WIDTH AS WIDTH_INR,")
        S.AppendLine("HEIGHT AS HEIGHT_INR,")
        S.AppendLine("WEIGHT AS WEIGHT_INR,")
        S.AppendLine("0 AS LENGTH_IT,")
        S.AppendLine("0 AS WIDTH_IT,")
        S.AppendLine("0 AS HEIGHT_IT,")
        S.AppendLine("0 AS WEIGHT_IT")
        S.AppendLine("FROM ICTSTYLD")
        S.AppendLine("WHERE PACK_CODE = 'INR'")
        S.AppendLine("UNION")
        S.AppendLine("SELECT")
        S.AppendLine("STYLE_CODE,")
        S.AppendLine("0 AS LENGTH_CTN,")
        S.AppendLine("0 AS WIDTH_CTN,")
        S.AppendLine("0 AS HEIGHT_CTN,")
        S.AppendLine("0 AS WEIGHT_CTN,")
        S.AppendLine("0 AS LENGTH_INR,")
        S.AppendLine("0 AS WIDTH_INR,")
        S.AppendLine("0 AS HEIGHT_INR,")
        S.AppendLine("0 AS WEIGHT_INR,")
        S.AppendLine("LENGTH AS LENGTH_IT,")
        S.AppendLine("WIDTH AS WIDTH_IT,")
        S.AppendLine("HEIGHT AS HEIGHT_IT,")
        S.AppendLine("WEIGHT AS WEIGHT_IT")
        S.AppendLine("FROM ICTSTYLD")
        S.AppendLine("WHERE PACK_CODE = 'IT'")
        S.AppendLine(")")
        If CreateEmpty Then
            S.AppendLine("WHERE ROWNUM < 0")
        End If
        S.AppendLine("GROUP BY STYLE_CODE")
        ASCMAIN1.sql = S.ToString
        TEMP_TABLE = ASCMAIN1.Temp_Table
    End Sub

    Private Sub txtSALES_YEARS_ValueChanged(sender As Object, e As EventArgs) Handles txtSALES_YEARS.ValueChanged
        If Not IsNumeric(txtSALES_YEARS.Text) Then
            MsgBox("Years Must Be A Number", vbOKOnly, "Numbers")
            txtSALES_YEARS.Value = 2
        End If
    End Sub

    Private Sub FillLastSalesDate()
        For Each rowWBTPACKS As DataRow In dst.Tables("WBTPACKS").Select()
            Dim STYLE_CODE As String = rowWBTPACKS.Item("STYLE_CODE").ToString & String.Empty

            Dim FILTER As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            Dim rowWBTPACKD As DataRow = dst.Tables.Item("WBTPACKD").Select(FILTER).FirstOrDefault
            If Not IsNothing(rowWBTPACKD) Then
                Dim LAST_SALE As String = rowWBTPACKD.Item("LAST_SALE").ToString & String.Empty
                If IsDate(LAST_SALE) Then
                    rowWBTPACKS.Item("LAST_SALE") = LAST_SALE
                End If
            End If
        Next
    End Sub
End Class