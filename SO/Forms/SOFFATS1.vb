Imports System.Text

Public Class SOFFATS1
    Dim SOTDEMD1 As String
    Dim SOTSUPP1 As String
    Dim S As New StringBuilder With {.Length = 0}
    Dim edi850cust As List(Of String)
    Dim AllocationMaxPeriods As Integer = 3

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFFATSI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst
            S.Length = 0
            S.AppendLine("Select ICTSTYL1.STYLE_CODE,")
            S.AppendLine("ICTSTAT2.COLOR_CODE,")
            S.AppendLine("ICTSTYL1.STYLE_DESC,")
            S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) AS WHSE_QTY_ON_HAND,")
            S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0)) AS WHSE_QTY_ON_ORDER,")
            S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_OPEN,0)) AS WHSE_QTY_OPEN,")
            S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_PICK,0)) AS WHSE_QTY_PICK,")
            S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0)) AS WHSE_QTY_TRAN,")
            S.AppendLine("SUM(NVL(SOTRSRV2.RSRV_QTY_OPEN,0)) AS RSRV_QTY_OPEN")
            S.AppendLine("FROM ICTSTYL1, ICTSTAT2, SOTRSRV2")
            S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE")
            S.AppendLine("AND ICTSTYL1.STYLE_CODE = SOTRSRV2.STYLE_CODE (+)")
            S.AppendLine("GROUP BY")
            S.AppendLine("ICTSTYL1.STYLE_CODE,")
            S.AppendLine("ICTSTAT2.COLOR_CODE,")
            S.AppendLine("ICTSTYL1.STYLE_DESC")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOFFATS1", "**", 0, False)
            .Tables.Item("SOFFATS1").Columns.Add("UPC_CODE", GetType(System.String))
            .Tables.Item("SOFFATS1").Columns.Add("CUST_STYLE_CODE", GetType(System.String))
            .Tables.Item("SOFFATS1").Columns.Add("CUST_UPC", GetType(System.String))
            For i As Integer = 1 To AllocationMaxPeriods
                Dim iFormat As String = Format(i, "00")
                .Tables.Item("SOFFATS1").Columns.Add("ALLOC_DATE_" & iFormat, GetType(System.DateTime))
                .Tables.Item("SOFFATS1").Columns.Add("ALLOC_QTY_" & iFormat, GetType(System.Double))
            Next

            SOTSUPP1 = ASCMAIN1.Temp_Table("Select * from SOTSUPP1")
            ASCMAIN1.sql = "Select * from " & SOTSUPP1
            Create_TDA(.Tables.Add, "SOTSUPP1", "**", 0, False)

            SOTDEMD1 = ASCMAIN1.Temp_Table("Select * from SOTDEMD1")
            ASCMAIN1.sql = "Select * from " & SOTDEMD1
            Create_TDA(.Tables.Add, "SOTDEMD1", "**", 0, False)

            S.Length = 0
            S.AppendLine("Select * from ICVLUPC1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ICVLUPC1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT * FROM SOTCSTY1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTCSTY1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("STYLE_CODE")
            S.AppendLine("FROM ICTSTYL1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "SOTSTYLES", "**", 0, False)
        End With

        grdSOTFATS1.DataSource = dst.Tables("SOFFATS1")
        grdICVLUPC1.DataSource = dst.Tables("ICVLUPC1")
        grdSOTCSTY1.DataSource = dst.Tables("SOTCSTY1")
        grdSTYLES.DataSource = dst.Tables("SOTSTYLES")

        'Sort_grdColumns(grdSOTFATS1, "ORDR_CUST_PO, LNO", False)

        With grdSOTFATS1.DisplayLayout.Bands(0)
            .Columns("WHSE_QTY_ON_HAND").Format = "###,##0"
            .Columns("WHSE_QTY_ON_ORDER").Format = "###,##0"
            .Columns("WHSE_QTY_OPEN").Format = "###,##0"
            .Columns("WHSE_QTY_PICK").Format = "###,##0"
            .Columns("WHSE_QTY_TRAN").Format = "###,##0"
            .Columns("RSRV_QTY_OPEN").Format = "###,##0"
            For i As Integer = 1 To AllocationMaxPeriods
                Dim iFormat As String = Format(i, "00")
                .Columns("ALLOC_QTY_" & iFormat).Format = "###,##0"
                .Columns("ALLOC_DATE_" & iFormat).Format = "MM/dd/yy"
            Next
        End With


        'Create_Summary(grdSOTCANG1, "ORDR_QTY")
        'Create_Summary(grdSOTCANG1, "ORDR_QTY_CANC")

        TABLE_NAME = "SOFFATS1"

        edi850cust = TAC.SOCMAIN1.Get_EDI_Custs("850")

        EntryMode = "E"
        'Call Load_Record()

        FilterSubGrids()

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
        dst.Tables("SOFFATS1").Rows.Clear()
        dst.Tables("ICVLUPC1").Rows.Clear()
        dst.Tables("SOTCSTY1").Rows.Clear()

        dst.EnforceConstraints = False

        Dim lstStyles As String = ""

        For Each rowSOTSTYLES As DataRow In dst.Tables("SOTSTYLES").Select()
            lstStyles = lstStyles & "'" & rowSOTSTYLES.Item("STYLE_CODE").ToString() & "',"
        Next
        lstStyles = lstStyles.Substring(0, lstStyles.Length - 1)

        S.Length = 0
        S.AppendLine("Select ICTSTYL1.STYLE_CODE,")
        S.AppendLine("NVL(ICTSTAT2.COLOR_CODE,'NONE') AS COLOR_CODE,")
        S.AppendLine("ICTSTYL1.STYLE_DESC,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) AS WHSE_QTY_ON_HAND,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0)) AS WHSE_QTY_ON_ORDER,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_OPEN,0)) AS WHSE_QTY_OPEN,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_PICK,0)) AS WHSE_QTY_PICK,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0)) AS WHSE_QTY_TRAN,")
        S.AppendLine("SUM(NVL(SOTRSRV2.RSRV_QTY_OPEN,0)) AS RSRV_QTY_OPEN")
        S.AppendLine("FROM ICTSTYL1, ICTSTAT2, SOTRSRV2")
        S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE (+)")
        S.AppendLine("AND ICTSTYL1.STYLE_CODE = SOTRSRV2.STYLE_CODE (+)")
        S.AppendLine("AND ICTSTYL1.STYLE_CODE IN (")
        'S.AppendLine(lstStyles)
        S.AppendLine("Select STYLE_CODE FROM")
        S.AppendLine("ICTSTYL1 WHERE CUST_CODE = 'WALMART'")
        S.AppendLine("And INIT_DATE >= '01-JAN-2013'")
        S.AppendLine(")")
        S.AppendLine("GROUP BY")
        S.AppendLine("ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTSTYL1.STYLE_DESC")
        Fill_Records("SOFFATS1",, True, S.ToString)

        S.Length = 0
        S.AppendLine("Select * from ICVLUPC1")
        S.AppendLine("WHERE ICVLUPC1.STYLE_CODE IN (")
        S.AppendLine(lstStyles)
        S.AppendLine(")")
        Fill_Records("ICVLUPC1",, True, S.ToString)

        S.Length = 0
        S.AppendLine("SELECT * FROM SOTCSTY1")
        S.AppendLine("WHERE CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text.ToString & "'")
        S.AppendLine("AND SOTCSTY1.STYLE_CODE IN (")
        S.AppendLine(lstStyles)
        S.AppendLine(")")
        Fill_Records("SOTCSTY1",, True, S.ToString)
        'dst.EnforceConstraints = True

        FillExtraFields()
        grdSOTFATS1.UpdateData()
        grdSOTFATS1.Update()

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
        grdSOTFATS1.Update()
        grdSOTFATS1.Refresh()
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
        Call Load_Popup_Menu(grdSOTFATS1, "SSB", "Show Filter", "Show GroupBox", "Style Status Inquiry")
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

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value & ""
                Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")

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

    Private Sub AllocateStyles()
        Dim TABLE_NAMEs As Dictionary(Of String, String) = Nothing

        TABLE_NAMEs = TAC.SOCMAIN1.Allocation_Initialization(Me,
          "",
          False,
          True,
          False,
          "", Now.Date, "")

        Dim newStyle As Boolean = True
        Dim lastStyle As String = ""
        Dim Zeros As Double()
        ReDim Zeros(8)
        For i As Integer = 0 To 8
            Zeros(i) = 0
        Next
        For Each rowSOFFATS1 As DataRow In dst.Tables("SOFFATS1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowSOFFATS1.Item("STYLE_CODE").ToString()
            Dim COLOR_CODE As String = rowSOFFATS1.Item("COLOR_CODE").ToString()
            If STYLE_CODE = lastStyle Then
                newStyle = False
            Else
                newStyle = True
                ASCMAIN1.Progress("Now Allocating Style ", STYLE_CODE)
            End If
            lastStyle = STYLE_CODE

            'If STYLE_CODE = "66114WM" Then Stop
            If newStyle Then
                Dim totAlloc As Int64 = 0
                If StyleShouldAllocate(STYLE_CODE) Then
                    Dim Allocations As Boolean = MakeAllocationTable(STYLE_CODE, TABLE_NAMEs, newStyle)
                    Dim DQFilter As String = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"
                    Dim iCnt As Integer = 0
                    For Each rowICTSTDQ1 As DataRow In dst.Tables("ICTSTDQ1").Select(DQFilter, "STATUS_DATE")
                        iCnt += 1
                        If iCnt <= AllocationMaxPeriods Then
                            Dim iFormat As String = Format(iCnt, "00")
                            If IsDate(rowICTSTDQ1.Item("STATUS_DATE").ToString) Then
                                rowSOFFATS1.Item("ALLOC_DATE_" & iFormat) = rowICTSTDQ1.Item("STATUS_DATE").ToString
                            End If
                            rowSOFFATS1.Item("ALLOC_QTY_" & iFormat) = Val(rowICTSTDQ1.Item("QTY_ATS_CUM").ToString)
                        End If
                    Next
                End If
            End If
        Next
    End Sub

    Private Sub FillExtraFields()
        For Each rowSOFFATS1 As DataRow In dst.Tables("SOFFATS1").Select()
            Dim STYLE_CODE As String = rowSOFFATS1.Item("STYLE_CODE").ToString
            Dim COLOR_CODE As String = rowSOFFATS1.Item("COLOR_CODE").ToString
            Dim Filter As String = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"
            Select Case dst.Tables.Item("ICVLUPC1").Select(Filter).Count
                Case 0
                    rowSOFFATS1.Item("UPC_CODE") = "None"
                Case 1
                    rowSOFFATS1.Item("UPC_CODE") = dst.Tables.Item("ICVLUPC1").Select(Filter).FirstOrDefault.Item("UPC_CODE").ToString
                Case Else
                    rowSOFFATS1.Item("UPC_CODE") = "Multiple"
            End Select

            Select Case dst.Tables.Item("SOTCSTY1").Select(Filter).Count
                Case 0
                    rowSOFFATS1.Item("CUST_UPC") = "None"
                    rowSOFFATS1.Item("CUST_STYLE_CODE") = "None"
                Case 1
                    rowSOFFATS1.Item("CUST_UPC") = dst.Tables.Item("SOTCSTY1").Select(Filter).FirstOrDefault.Item("CUST_UPC").ToString
                    rowSOFFATS1.Item("CUST_STYLE_CODE") = dst.Tables.Item("SOTCSTY1").Select(Filter).FirstOrDefault.Item("CUST_STYLE_CODE").ToString
                Case Else
                    rowSOFFATS1.Item("CUST_UPC") = "Multiple"
                    rowSOFFATS1.Item("CUST_STYLE_CODE") = "Multiple"
            End Select
        Next
    End Sub

    Private Sub FilterSubGrids()
        Dim dvw1 As DataView = DirectCast(grdICVLUPC1.DataSource, DataTable).DefaultView
        Dim dvw2 As DataView = DirectCast(grdSOTCSTY1.DataSource, DataTable).DefaultView

        If IsNothing(grdSOTFATS1.ActiveRow) Then
            Dim DefaultFilter As String = "STYLE_CODE = 'XXXXXXXX'"
            dvw1.RowFilter = DefaultFilter
            dvw2.RowFilter = DefaultFilter
            grdICVLUPC1.Text = "UPC Codes"
            grdSOTCSTY1.Text = "Customer Codes"
        Else
            Dim STYLE_CODE As String = grdSOTFATS1.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdSOTFATS1.ActiveRow.Cells("COLOR_CODE").Value
            Dim DefaultFilter As String = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"
            dvw1.RowFilter = DefaultFilter
            dvw2.RowFilter = DefaultFilter
            grdICVLUPC1.Text = "UPC Codes for " & STYLE_CODE & " - " & COLOR_CODE
            grdSOTCSTY1.Text = "Customer Codes for " & STYLE_CODE & " - " & COLOR_CODE
        End If

    End Sub

    Private Function MakeAllocationTable(ByVal STYLE_CODE As String,
                                    ByVal TABLE_NAMEs As Dictionary(Of String, String),
                                    ByVal NEWSTYLE As Boolean) As Boolean
        Dim RetVal As Boolean = True

        If NEWSTYLE Then
            Dim SOTORDR0 As String = TABLE_NAMEs("SOTORDR0")
            Dim SOTORDR1 As String = TABLE_NAMEs("SOTORDR1")
            Dim SOTORDR2 As String = TABLE_NAMEs("SOTORDR2")
            Dim SOTRSRV1 As String = TABLE_NAMEs("SOTRSRV1")
            Dim SOTRSRV2 As String = TABLE_NAMEs("SOTRSRV2")
            Dim ARTCUST1 As String = TABLE_NAMEs("ARTCUST1")

            For Each TABLE_NAME As String In New String() {"SOTORDR1", "SOTORDR0", "ARTCUST1", "ICTSTDQ1", "SOTORDR2", "SOTRSRV1", "SOTRSRV2"}
                ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAMEs(TABLE_NAME))
            Next

            For Each sql As String In TABLE_NAMEs.Keys
                If sql.StartsWith("sql") Then
                    Dim sqlstmt As String = Replace(TABLE_NAMEs(sql), "'STYLE_CODE'", "'" & STYLE_CODE & "'")
                    ASCDATA1.ExecuteSQL(sqlstmt)
                End If
            Next

            dst.Tables("SOTSUPP0").Rows.Clear()
            dst.Tables("SOTSUPPI").Rows.Clear()
            dst.Tables("SOTORDR7").Rows.Clear()
            dst.Tables("ICTSTDQ1").Rows.Clear()
            dst.Tables("ICTSTDQ2").Rows.Clear()

            TAC.SOCMAIN1.Allocation(Me,
                False,
                True,
                 "",
                 "", edi850cust,
                SOTSUPP1, SOTDEMD1, TABLE_NAMEs, True, True, STYLE_CODE, , , , False)
        End If
        Return RetVal
    End Function

    Private Function StyleShouldAllocate(ByVal STYLE_CODE As String) As Boolean
        Dim retVal As Boolean = True
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT SUM(NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)+NVL(WHSE_QTY_OPEN,0)+NVL(WHSE_QTY_PICK,0)) TOT")
        SQLS.AppendLine("FROM ICTSTAT2")
        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim TOT As Int64 = Val(ASCDATA1.GetDataValue)
        If TOT = 0 Then
            retVal = False
        End If
        Return retVal
    End Function
#End Region

#Region "Form Controls"

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        Dim Proceed As Boolean = False
        If Absx1.txtFor("CUST_CODE").Text.Length > 0 And dst.Tables.Item("SOTSTYLES").Rows.Count > 0 Then
            Proceed = True
        End If
        If Proceed Then
            Load_Record()
            ASCMAIN1.Progress("Now Allocating")
            AllocateStyles()
            ASCMAIN1.Progress("Now Allocating")
        Else
            MsgBox("You Must Select A Customer And Some Styles Before Loading", vbExclamation, "Selection")
        End If

    End Sub

    Private Sub btnStyles_Click(sender As Object, e As EventArgs) Handles btnStyles.Click
        If Absx1.txtFor("CUST_CODE").Text = "" Then
            MsgBox("You Must Select A Customer First")
        Else
            dst.Tables("SOTSTYLES").Clear()
            Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
            SQLS.AppendLine(" STYLE_CODE IN (SELECT STYLE_CODE FROM ICTSTYL1 WHERE CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "')")
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE", , SQLS.ToString)
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.SelectedCodes.Count > 0 Then
                    For Each SelCode As String In ASCMAIN1.CodeSelector.SelectedCodes
                        Dim newRec As DataRow = dst.Tables("SOTSTYLES").NewRow
                        newRec.Item("STYLE_CODE") = SelCode
                        dst.Tables("SOTSTYLES").Rows.Add(newRec)
                    Next
                End If
                F.Dispose()
            End If
        End If
    End Sub

    Private Sub grdSOTFATS1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTFATS1.AfterRowActivate
        FilterSubGrids()
    End Sub
#End Region
End Class