Imports System.Text
Imports Infragistics.Win.UltraWinGrid
Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel

Public Class SOFCOMTS
    Dim SQL As New System.Text.StringBuilder
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFCOMTI" Then
            InquiryMode = True
        End If

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, -12)

        Check_Form_Options()
        Dim SQL As New System.Text.StringBuilder

        With dst
            'SQL.Length = 0
            'SQL.AppendLine("SELECT")
            'ASCMAIN1.sql = SQL.ToString
            'Dim TABLES As String() = {"SOTSHPWA", "SOTSHPWH"}



            'Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")


        End With

        'grdSOTCOMTS.DataSource = dst.Tables("SOFCOMTS")

        'Sort_grdColumns(grdSOTCOMTS, "ORDR_YYYYPP_BOOKED, ORDR_GROUP_NO", False)
        'Create_Summary(grdSOTCOMTS, "CUST_STORE_NO", "Count")

        TABLE_NAME = "SOTCOMTS"

        EntryMode = "E"

        'With grdSOTCOMTS.DisplayLayout.Bands(0)
        '    .Columns("ST_EA_ORDR").Format = "#,###,##0"
        '    .Columns("ST_EACH_RCD").Format = "#,###,##0"
        '    .Columns("ORDR_QTY_SHIP").Format = "#,###,##0"
        '    .Columns("VARIANCE").Format = "#,###,##0"
        '    .Columns("EXCEL_LINE").Format = "######0"
        'End With

        'Call Load_Record()


        'Bind_Controls(grpHeader, "SOTSHPWH")

        'ASCMAIN1.Add_Value_List(grdSOTCOMTS, "ORDR_STATUS", , New String() {":", "C:Cancelled", "D:Deleted"})

        Call Mode_Settings(True)

        'SplitContainer2.SplitterDistance = 120

        'Fill_Records("STOREPOS")

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
                'Me.Close()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        'TabControl1.Visible = Not tf

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        For Each TABLE_NAME As String In New String() _
            {"SOTCOMTS"}
            'dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        'Fill_Records("PMTVIST1")

        dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        'tab.Visible = ScreenMode

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        'dst.Tables("SOTROYLI").Rows.Clear()

        dst.EnforceConstraints = False

        'fillSOTCOMTS()

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
        grdSOTCOMTS.Update()
        grdSOTCOMTS.Refresh()
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

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOTCOMTS, "SS", "Show Filter", "Show GroupBox")
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

#End Region

#Region "Form Controls"

    Private Sub btnFETCH_Click(sender As Object, e As EventArgs) Handles btnFETCH.Click
        Me.Cursor = Cursors.WaitCursor
        Dim RYPLEGEND0 As String = Absx1.cmbFor("RYP0", True).Value
        Dim RYP0 As String = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine("S1.INV_TYPE,")
        SQL.AppendLine("S1.ORDR_YYYYPP_UPDATED,")
        SQL.AppendLine("I1.SALES_DIVISION_CODE,")
        SQL.AppendLine("D1.SALES_DIVISION_NAME,")
        SQL.AppendLine("S1.CUST_CODE,")
        SQL.AppendLine("C1.CUST_NAME,")
        SQL.AppendLine("S1.INV_NO,")
        SQL.AppendLine("S1.INV_DATE,")
        SQL.AppendLine("S2.STYLE_CODE,")
        SQL.AppendLine("S2.COLOR_CODE,")
        SQL.AppendLine("B1.MASTER_BODY_DESC,")
        SQL.AppendLine("S2.ORDR_QTY_SHIP,")
        SQL.AppendLine("S2.ORDR_UNIT_PRICE,")
        SQL.AppendLine("S2.ORDR_UNIT_COST,")
        SQL.AppendLine("(NVL(S2.ORDR_QTY_SHIP,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS TOT_PRICE,")
        SQL.AppendLine("(NVL(S2.ORDR_QTY_SHIP,0) * NVL(S2.ORDR_UNIT_COST,0)) AS TOT_COST")
        SQL.AppendLine("FROM SOTINVH1 S1, SOTINVH2 S2, ICTSTYL1 I1, ARTCUST1 C1, SOTSDIV1 D1, ICTBODY1 B1, ICTBODY2 B2")
        SQL.AppendLine("WHERE S1.INV_TYPE = S2.INV_TYPE")
        SQL.AppendLine("AND S1.INV_NO = S2.INV_NO")
        SQL.AppendLine("AND S2.STYLE_CODE = I1.STYLE_CODE")
        SQL.AppendLine("AND S1.CUST_CODE = C1.CUST_CODE")
        'SQL.AppendLine("AND S1.SALES_DIVISION_CODE = D1.SALES_DIVISION_CODE")
        SQL.AppendLine("AND I1.SALES_DIVISION_CODE = D1.SALES_DIVISION_CODE")
        SQL.AppendLine("AND I1.SUB_BODY_CODE = B2.SUB_BODY_CODE")
        SQL.AppendLine("AND B2.MASTER_BODY_CODE = B1.MASTER_BODY_CODE")
        SQL.AppendLine($"AND S1.ORDR_YYYYPP_UPDATED = '{RYP0}'")
        SQL.AppendLine("AND S1.INV_TYPE = 'I'")
        'SQL.AppendLine("AND S1.CUST_CODE <> 'WALMART'")
        SQL.AppendLine("AND (S1.CUST_CODE <> 'TRANSFERS' AND S1.CUST_CODE <> 'SAMPLES')")
        SQL.AppendLine("ORDER BY")
        SQL.AppendLine("S1.INV_TYPE,")
        SQL.AppendLine("S1.ORDR_YYYYPP_UPDATED,")
        SQL.AppendLine("I1.SALES_DIVISION_CODE,")
        SQL.AppendLine("S1.CUST_CODE,")
        SQL.AppendLine("S1.INV_NO,")
        SQL.AppendLine("S1.INV_DATE")

        'Fill_Records("SOTCOMTS",,, SQL.ToString)
        Dim tbl As DataTable = ASCDATA1.GetDataTable(SQL.ToString(), String.Empty)

        'grdSOTCOMTS.DataSource = tbl 'dst.Tables("SOFCOMTS")
        'grdSOTCOMTS.Refresh()
        'grdSOTCOMTS.DisplayLayout.Bands(0).Columns.Item("CUST_CODE").Hidden = False
        Dim MAX_FILES As Int64 = 6

        Dim COLS As New List(Of String)
        For Each COL As DataColumn In tbl.Columns
            COLS.Add(COL.ColumnName)
        Next

        Dim FILES_STREAM As New List(Of System.IO.StreamWriter)
        Dim FILES_LIST As New List(Of String)
        For i As Int64 = 0 To MAX_FILES - 1
            FILES_LIST.Add($"C:\TST\SALES_{RYP0}_{i}.csv")
        Next
        For Each FL As String In FILES_LIST
            If System.IO.File.Exists(FL) Then
                System.IO.File.Delete(FL)
            End If
        Next

        Dim CURR_ROW As Int64 = 0
        Dim CURR_FILE As Int64 = 0
        FILES_STREAM.Add(My.Computer.FileSystem.OpenTextFileWriter(FILES_LIST(CURR_FILE), False))
        Dim LN As String = ""
        For Each COL As String In COLS
            LN += Chr(34) & COL & Chr(34) & ","
        Next
        LN = LN.Substring(0, LN.Length - 1)
        FILES_STREAM(CURR_FILE).WriteLine(LN)
        For Each row As DataRow In tbl.Rows
            CURR_ROW += 1
            If CURR_ROW = 500000 Then
                FILES_STREAM(CURR_FILE).Close()
                CURR_ROW = 0
                CURR_FILE += 1
                FILES_STREAM.Add(My.Computer.FileSystem.OpenTextFileWriter(FILES_LIST(CURR_FILE), False))
                LN = ""
                For Each COL As String In COLS
                    LN += Chr(34) & COL & Chr(34) & ","
                Next
                LN = LN.Substring(0, LN.Length - 1)
                FILES_STREAM(CURR_FILE).WriteLine(LN)
            End If
            LN = ""
            For Each COL As String In COLS
                LN += Chr(34) & row.Item(COL).ToString & Chr(34) & ","
            Next
            LN = LN.Substring(0, LN.Length - 1)
            FILES_STREAM(CURR_FILE).WriteLine(LN)
        Next
        FILES_STREAM(CURR_FILE).Close()
        tbl = Nothing
        Application.DoEvents()

        'Dim MOS As New List(Of String)
        'For I As Int64 = 1 To 12
        '    MOS.Add("2021" & Format(I, "00"))
        'Next
        'For I As Int64 = 1 To 7
        '    MOS.Add("2022" & Format(I, "00"))
        'Next

        'For Each MO As String In MOS

        'Next

        Me.Cursor = Cursors.Default
        MsgBox(RYP0, vbOKOnly, "Done")
    End Sub
#End Region

End Class