
Imports System.Text
Imports System.IO

Public Class SOFSURCG
    Dim SQL As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        dteFromDate.Value = DateSerial(2021, 6, 3)
        dteEndDate.Value = DateSerial(2021, 6, 3)

        With dst
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("I1.INV_NO,")
            SQL.AppendLine("I1.ORDR_NO,")
            SQL.AppendLine("I1.INV_DATE,")
            SQL.AppendLine("I1.CUST_CODE,")
            SQL.AppendLine("A1.CUST_NAME,")
            SQL.AppendLine("I1.ORDR_CUST_PO,")
            SQL.AppendLine("I1.WHSE_CODE,")
            SQL.AppendLine("CAST(C1.ORDR_TOTAL AS NUMBER(13,2)) ORDR_TOTAL,")
            SQL.AppendLine("I1.INV_SALES,")
            SQL.AppendLine("I1.INV_FREIGHT,")
            SQL.AppendLine("I1.INV_MISC_CHG,")
            SQL.AppendLine("I1.INV_TOTAL_AMOUNT,")
            SQL.AppendLine("OS.OFTSUR,")
            'SQL.AppendLine("(NVL(OFTSUR, 0) / I1.INV_SALES) * 100 OFTPCT,")
            SQL.AppendLine("I1.SREP_CODE,")
            SQL.AppendLine("I1.INIT_DATE,")
            SQL.AppendLine("I1.INIT_OPER")
            SQL.AppendLine("FROM SOTINVH1 I1, ARTCUST1 A1, SOROFSURCHG C1, (SELECT INV_NO, SUM(INV_MISC_CHG) AS OFTSUR FROM SOTINVHM WHERE MISC_CHG_CODE = 'OFTSUR' GROUP BY INV_NO) OS")
            SQL.AppendLine("WHERE I1.ORDR_NO = C1.ORDR_NO")
            SQL.AppendLine("AND I1.CUST_CODE = A1.CUST_CODE")
            SQL.AppendLine("AND I1.INV_NO = OS.INV_NO (+)")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "SOFSURCG", "**", 0, False)
            .Tables("SOFSURCG").Columns.Add("OFTPCT", GetType(System.Decimal))
        End With

        'Fill_Records("SOFSURCG")

        grdSOFSURCG.DataSource = dst.Tables("SOFSURCG")

        Create_Summary(grdSOFSURCG, "INV_NO", "Count", "", "###,##0")
        Create_Summary(grdSOFSURCG, "ORDR_TOTAL", "Sum", "", "###,##0.00")
        Create_Summary(grdSOFSURCG, "INV_SALES", "Sum", "", "###,##0.00")
        Create_Summary(grdSOFSURCG, "INV_FREIGHT", "Sum", "", "###,##0.00")
        Create_Summary(grdSOFSURCG, "INV_MISC_CHG", "Sum", "", "###,##0.00")
        Create_Summary(grdSOFSURCG, "INV_TOTAL_AMOUNT", "Sum", "", "###,##0.00")
        Create_Summary(grdSOFSURCG, "OFTSUR", "Sum", "", "###,##0.00")


        'ASCMAIN1.Add_Value_List(grdSOTQRDR1, "CALC_STATUS", , New String() {":", "I:Imported From Web", "L:Pulled To Laptop", "O:Finalized As Order", "X:Deleted", "M:Marked Complete", "T:Testing"})
        Sort_grdColumns(grdSOFSURCG, "INV_NO", False)

        With grdSOFSURCG.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        With grdSOFSURCG.DisplayLayout.Bands(0).Columns
            For Each COL As String In New String() {"ORDR_TOTAL", "INV_SALES", "INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT", "OFTSUR"}
                .Item(COL).Format = "###,##0.00"
            Next
            .Item("OFTPCT").Format = "###,##0.00%"
        End With

        For i As Integer = 0 To grdSOFSURCG.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOFSURCG.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        Load_Record()

        tab.Visible = False
        isFormLoading = False

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
                'Load_Record()

                RefreshData()
            Case "Exit"
                Call Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Visible = True
                .Groups("Screen Control").Items("Exit").Visible = True
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        'dst.Tables("SOTQRDR1").Rows.Clear()
    End Sub

    Sub Load_Record()
        'Call Save_Header_Fields(UltraGroupBox1)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Refreshing Data", "")

        EnforceConstraints(False)

        'Fill_Records("SOTQRDR1")

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'Update_Record_TDA("SOTQRDR1")
        Call CommitTrans("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    'Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
    '    Print_Report_Begin()
    '    'frm.CR_params.Add("SUBT", "")
    '    'Fill SOTORDRP records
    '    Fill_Records("SOTQRDR5", ORDR_NO, True)
    '    For Each rowSOTQRDR1 As DataRow In dst.Tables("SOTQRDR1").Select()
    '        If rowSOTQRDR1.Item("ORDR_NO") = ORDR_NO Then
    '            rowSOTQRDR1.Item("ERRORS") = "NEW"
    '        Else
    '            rowSOTQRDR1.Item("ERRORS") = ""
    '        End If
    '    Next
    '    'Generate_Report("SORQRDRO")
    '    Generate_Report("WBRWEBQT", "Quotes Imported From Web", "Re-printed From Quote Maint.")
    '    '    Print_Report_End()
    'End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOFSURCG, "SSBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Customer Inquiry")
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
            'Case "grdSOTORDR1"
            '    If Not InquiryOnly Then
            '        e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
            '    End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Selected.Rows.Count = 0 Then
        '    MsgBox("You Must Select One And Only One Row First", vbOKOnly, "Select A Row")
        '    Exit Sub
        'End If

        Select Case e.Tool.Key
            Case "Something"
                'grd.ActiveRow.Cells.Item("ORDR_NO_WEB").Value = ""
        End Select

        Update_Record()
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
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
    Private Sub RefreshData()
        ASCMAIN1.Progress("Refreshing Styles", "")
        Dim BDate As String = Format(CDate(dteFromDate.Value), "dd-MMM-yyyy")
        Dim EDate As String = Format(CDate(dteEndDate.Value), "dd-MMM-yyyy")

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine("I1.INV_NO,")
        SQL.AppendLine("I1.ORDR_NO,")
        SQL.AppendLine("I1.INV_DATE,")
        SQL.AppendLine("I1.CUST_CODE,")
        SQL.AppendLine("A1.CUST_NAME,")
        SQL.AppendLine("I1.ORDR_CUST_PO,")
        SQL.AppendLine("I1.WHSE_CODE,")
        SQL.AppendLine("CAST(C1.ORDR_TOTAL AS NUMBER(13,2)) ORDR_TOTAL,")
        SQL.AppendLine("I1.INV_SALES,")
        SQL.AppendLine("I1.INV_FREIGHT,")
        SQL.AppendLine("I1.INV_MISC_CHG,")
        SQL.AppendLine("I1.INV_TOTAL_AMOUNT,")
        SQL.AppendLine("OS.OFTSUR,")
        SQL.AppendLine("I1.SREP_CODE,")
        SQL.AppendLine("I1.INIT_DATE,")
        SQL.AppendLine("I1.INIT_OPER")
        SQL.AppendLine("FROM SOTINVH1 I1, ARTCUST1 A1, SOROFSURCHG C1, (SELECT INV_NO, INV_TYPE, SUM(NVL(INV_MISC_CHG,0)) AS OFTSUR FROM SOTINVHM WHERE MISC_CHG_CODE = 'OFTSUR' GROUP BY INV_NO, INV_TYPE) OS")
        SQL.AppendLine("WHERE I1.ORDR_NO = C1.ORDR_NO")
        SQL.AppendLine("AND I1.CUST_CODE = A1.CUST_CODE")
        SQL.AppendLine("AND I1.INV_NO = OS.INV_NO (+)")
        SQL.AppendLine("AND I1.INV_TYPE = OS.INV_TYPE (+)")
        SQL.AppendLine($"AND I1.INV_DATE >= '{BDate}'")
        SQL.AppendLine($"AND I1.INV_DATE <= '{EDate}'")
        Fill_Records("SOFSURCG",,, SQL.ToString)

        For Each rowSOFSURCG As DataRow In dst.Tables("SOFSURCG").Select()
            Dim INV_SALES As Decimal = Val(rowSOFSURCG.Item("INV_SALES").ToString & String.Empty)
            Dim OFTSUR As Decimal = Val(rowSOFSURCG.Item("OFTSUR").ToString & String.Empty)
            Dim OFTPCT As Decimal = 0
            If INV_SALES > 0 And OFTSUR > 0 Then
                OFTPCT = (OFTSUR / INV_SALES)
            End If
            rowSOFSURCG.Item("OFTPCT") = OFTPCT
        Next

        ASCMAIN1.Progress("", "")
        'grdECTSZIO1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ResizeAllColumns
    End Sub
#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class