
Imports System.Text
Imports System.IO

Public Class SOFQVCA1
    Dim SQL As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With dst
            'SQL.Length = 0
            'SQL.AppendLine("Select SOTORDR1.* from SOTORDR1 where ORDR_NO = :PARM1")
            'ASCMAIN1.sql = SQL.ToString()
            'Create_TDA(.Tables.Add, "SOTQVCA1", "**", 0, False, "V", 1)
            '.Tables("SOTQVCA1").Columns.Add("ItemID", GetType(System.String))
            SQL.Length = 0
            SQL.AppendLine("Select SOTORDR1.*")
            SQL.AppendLine(" from SOTORDR1")
            SQL.AppendLine(" where CUST_CODE = :PARM1")
            SQL.AppendLine(" and ORDR_STATUS ='O'")
            SQL.AppendLine("")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "SOTQVCA1", "**", 0, False, "V", 1)
            .Tables("SOTQVCA1").Columns.Add("CUST_CITY", GetType(System.String))
            .Tables("SOTQVCA1").Columns.Add("CUST_STATE", GetType(System.String))
            .Tables("SOTQVCA1").Columns.Add("CUST_COUNTRY", GetType(System.String))
            'ORDR_QTY_OPEN
            .Tables("SOTQVCA1").Columns.Add("ORDR_QTY_OPEN", GetType(System.Decimal))
            .Tables("SOTQVCA1").Columns.Add("ORDR_QTY_PICK", GetType(System.Decimal))
            .Tables("SOTQVCA1").Columns.Add("ORDR_QTY_SHIP", GetType(System.Decimal))
            .Tables("SOTQVCA1").Columns.Add("ORDR_QTY_CANC", GetType(System.Decimal))
            .Tables("SOTQVCA1").Columns.Add("ORDR_TOTAL", GetType(System.Decimal), "ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP")

            'Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)

            SQL.Length = 0
            SQL.AppendLine("Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTSTYL1.CASE_CUBE, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_STATUS, ICTSTYL1.STYLE_ASST_QTY, ICTSTYC1.STYLE_COLOR_STATUS")
            SQL.AppendLine(", ICTSTDQ3.DATE_1, ICTSTDQ3.QTY_1, ICTSTDQ3.DATE_2, ICTSTDQ3.QTY_2, ICTSTDQ3.DATE_3, ICTSTDQ3.QTY_3, ICTSTDQ3.DATE_4, ICTSTDQ3.QTY_4")
            SQL.AppendLine(" from SOTORDR2,ICTCOLR1,ICTSTYL1,ICTSTYC1,ICTSTDQ3")
            'SQL.AppendLine(" where SOTORDR2.ORDR_NO = :PARM1")
            SQL.AppendLine("   where ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE")
            SQL.AppendLine("   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE")
            SQL.AppendLine("   and ICTSTYC1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE")
            SQL.AppendLine("   and ICTSTYC1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE")
            SQL.AppendLine("   and ICTSTDQ3.ORDR_GROUP_NO (+) = SOTORDR2.ORDR_NO")
            SQL.AppendLine("   and ICTSTDQ3.STYLE_CODE (+) = SOTORDR2.STYLE_CODE")
            SQL.AppendLine("   and ICTSTDQ3.COLOR_CODE (+) = SOTORDR2.COLOR_CODE")
            SQL.AppendLine("")
            '        ASCMAIN1.sql = "Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTSTYL1.CASE_CUBE, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_STATUS, ICTSTYL1.STYLE_ASST_QTY, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
            '& ", ICTSTDQ3.DATE_1, ICTSTDQ3.QTY_1, ICTSTDQ3.DATE_2, ICTSTDQ3.QTY_2, ICTSTDQ3.DATE_3, ICTSTDQ3.QTY_3, ICTSTDQ3.DATE_4, ICTSTDQ3.QTY_4" & vbCrLf _
            '& " from SOTORDR2,ICTCOLR1,ICTSTYL1,ICTSTYC1,ICTSTDQ3" & vbCrLf _
            '& " where SOTORDR2.ORDR_NO = :PARM1" & vbCrLf _
            '& "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
            '& "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
            '& "   and ICTSTYC1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf _
            '& "   and ICTSTYC1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
            '& "   and ICTSTDQ3.ORDR_GROUP_NO (+) = SOTORDR2.ORDR_NO" & vbCrLf _
            '& "   and ICTSTDQ3.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf _
            '& "   and ICTSTDQ3.COLOR_CODE (+) = SOTORDR2.COLOR_CODE"
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2)
            .Tables("SOTORDR2").Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(ORDR_QTY,0) / ISNULL(CARTON_PACK_QTY,0))")
            .Tables("SOTORDR2").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")
            .Tables("SOTORDR2").Columns.Add("MU_PCT", GetType(System.Decimal), "IIF(ISNULL(PO_COST,0) = 0, 0, 100 * (ISNULL(ORDR_UNIT_PRICE,0) - ISNULL(PO_COST,0)) / ISNULL(PO_COST,0))")
            .Tables("SOTORDR2").Columns.Add("ORDR_AMT_CURR", GetType(System.Decimal), "ISNULL(ORDR_QTY,0)*ISNULL(ORDR_UNIT_PRICE_CURR,0)")
            If Not dst.Tables("SOTORDR2").Columns.Contains("ORDR_LINE_CANC") Then
                .Tables("SOTORDR2").Columns.Add("ORDR_LINE_CANC", GetType(System.String))
                .Tables("SOTORDR2").Columns("ORDR_LINE_CANC").MaxLength = 1
            End If

            With .Tables("SOTORDR2").Columns
                .Add("RANGE_STYLE_QTY_PER_PP", GetType(System.Int64))
                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_ALLO", GetType(System.Decimal), "ISNULL(ORDR_QTY_ALLO,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_QTY_PICK,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("DISC_AMT", GetType(System.Decimal), "ISNULL(STYLE_PRICE,0)-ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("DISC_PCT", GetType(System.Decimal), "IIF(ISNULL(STYLE_PRICE,0)=0,0,100*DISC_AMT/ISNULL(STYLE_PRICE,0))")

                ' .Add("ORDR_QTY_ALLO_CUR", GetType(System.Int64), "IIF(ORDR_RELEASE_AVAIL IS NULL,ORDR_QTY_ALLO,0)")
                .Add("ORDR_QTY_ALLO_CUR", GetType(System.Int64), "ISNULL(QTY_1,0)")
                .Add("ORDR_AMT_ALLO_CUR", GetType(System.Decimal), "ISNULL(ORDR_QTY_ALLO_CUR,0) * ISNULL(ORDR_UNIT_PRICE,0)")

                .Add("ORDR_UNIT_COST", GetType(System.Decimal))
                .Add("CGS", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_COST,0)")
                .Add("GP_AMT", GetType(System.Decimal), "ISNULL(ORDR_AMT_SHIP,0) - ISNULL(CGS,0)")
                .Add("GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ORDR_AMT_SHIP,0)=0,0,100 * ISNULL(GP_AMT,0) / ISNULL(ORDR_AMT_SHIP,0))")
            End With
            .Tables("SOTORDR2").Columns("ORDR_UNIT_PRICE_MANUAL").DefaultValue = "0"
            '.Tables("SOTORDR2").Columns("DUTY_RATE_CODE").DataType = GetType(System.Double)
            ' .Tables("SOTORDR2").Columns("ORDR_UNIT_PRICE").DataType = GetType(System.Double)

            'If ASCMAIN1.CLIENT = "RGI" Then
            With .Tables("SOTORDR2").Columns
                .Add("AMT_1", GetType(System.Decimal), "ISNULL(QTY_1,0)*ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("AMT_2", GetType(System.Decimal), "ISNULL(QTY_2,0)*ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("AMT_3", GetType(System.Decimal), "ISNULL(QTY_3,0)*ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("AMT_4", GetType(System.Decimal), "ISNULL(QTY_4,0)*ISNULL(ORDR_UNIT_PRICE,0)")
            End With

            Create_Relation("SOTQVCA1", "SOTORDR2", "ORDR_NO")
        End With

        'Fill_Records("SOTQVCA1")

        grdSOTQVCA1.DataSource = dst.Tables("SOTQVCA1")

        Create_Summary(grdSOTQVCA1, "ORDR_NO", "Count")
        Create_Summary(grdSOTQVCA1, "ORDR_QTY_OPEN", "Sum")
        Create_Summary(grdSOTQVCA1, "ORDR_QTY_PICK", "Sum")
        Create_Summary(grdSOTQVCA1, "ORDR_QTY_SHIP", "Sum")
        Create_Summary(grdSOTQVCA1, "ORDR_TOTAL", "Sum")
        Create_Summary(grdSOTQVCA1, "ORDR_QTY_OPEN", "Sum", "SOTQVCA1_SOTORDR2")


        'ASCMAIN1.Add_Value_List(grdSOTQRDR1, "CALC_STATUS", , New String() {":", "I:Imported From Web", "L:Pulled To Laptop", "O:Finalized As Order", "X:Deleted", "M:Marked Complete", "T:Testing"})
        Sort_grdColumns(grdSOTQVCA1, "ORDR_NO", False)

        With grdSOTQVCA1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        For i As Integer = 0 To grdSOTQVCA1.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOTQVCA1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        grdSOTQVCA1.DisplayLayout.UseFixedHeaders = True
        With grdSOTQVCA1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "ORDR_DATE", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        Show_Filter(grdSOTQVCA1, True)
        grdSOTQVCA1.DisplayLayout.GroupByBox.Hidden = False

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
                'Fill_Records("SOTQVCA1")
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
        Load_SOTQVCA1()

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
        Load_Popup_Menu(grdSOTQVCA1, "SS", "Show Filter")
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

        'Fill_Records("SOTQVCA1")
        'Load_SOTQVCA1()
        Load_Record()

        ASCMAIN1.Progress("", "")
        'grdECTSZIO1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ResizeAllColumns
    End Sub

    Sub Load_SOTQVCA1()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor

        'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim CUST_CODE As String = "171659"

        dst.Tables("SOTQVCA1").Rows.Clear()
        dst.Tables("SOTORDR2").Rows.Clear()

        Dim S2 As New StringBuilder With {.Length = 0}

        SQL.Length = 0
        SQL.AppendLine("Select SOTORDR1.*, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY, ARTCUST1.CUST_CREDIT_HOLD")
        SQL.AppendLine(" , (SELECT SUM(ORDR_UNIT_PRICE * NVL(ORDR_QTY_OPEN, 0)) ORDR_TOTAL FROM SOTORDR2 WHERE ORDR_NO = SOTORDR1.ORDR_NO) ORDR_QTY_OPEN")
        SQL.AppendLine(" , (SELECT SUM(ORDR_UNIT_PRICE * NVL(ORDR_QTY_PICK, 0)) ORDR_TOTAL FROM SOTORDR2 WHERE ORDR_NO = SOTORDR1.ORDR_NO) ORDR_QTY_PICK")
        SQL.AppendLine(" , (SELECT SUM(ORDR_UNIT_PRICE * NVL(ORDR_QTY_SHIP, 0)) ORDR_TOTAL FROM SOTORDR2 WHERE ORDR_NO = SOTORDR1.ORDR_NO) ORDR_QTY_SHIP")
        SQL.AppendLine(" , (SELECT SUM(ORDR_UNIT_PRICE * NVL(ORDR_QTY_CANC, 0)) ORDR_TOTAL FROM SOTORDR2 WHERE ORDR_NO = SOTORDR1.ORDR_NO) ORDR_QTY_CANC")
        SQL.AppendLine("  from SOTORDR1, ARTCUST1")
        SQL.AppendLine("  WHERE SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE")

        S2.AppendLine("Select SOTORDR1.ORDR_NO")
        S2.AppendLine(" from SOTORDR1")

        ASCMAIN1.Progress("Now Building List of Sales Orders", "")

        grdSOTQVCA1.Text = "All Open Sales Orders"
        SQL.AppendLine($"  AND ORDR_STATUS = 'O' AND SOTORDR1.CUST_CODE = '{CUST_CODE}'")
        S2.AppendLine($"  WHERE ORDR_STATUS = 'O' AND SOTORDR1.CUST_CODE = '{CUST_CODE}'")

        'If optShowOrders.Value = "A" And CUST_CODE = "" Then
        '    'ASCMAIN1.sql = "Select SOTORDR1.* from SOTORDR1 where ORDR_STATUS = 'O'"
        '    ASCMAIN1.sql &= " and SOTORDR1.ORDR_STATUS = 'O'"
        'grdSOTQVCA1.Text = "All Open Sales Orders"
        'ElseIf optShowOrders.Value = "M" Then
        '    'ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_STATUS = 'O' and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
        '    ASCMAIN1.sql &= " and SOTORDR1.ORDR_STATUS = 'O' and (SOTORDR1.INIT_OPER = '" & ASCMAIN1.USER_ID & "' or SOTORDR1.LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
        '    grdSOTORDRX.Text = "Open Sales Orders entered or modified by Me"
        'ElseIf optShowOrders.Value = "C" Or CUST_CODE <> "" Then
        '    'ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_STATUS = '" & optCustomerOrders.Value & "' and CUST_CODE = '" & CUST_CODE & "'"
        '    ASCMAIN1.sql &= " and SOTORDR1.ORDR_STATUS = '" & optCustomerOrders.Value & "' and SOTORDR1.CUST_CODE = '" & CUST_CODE & "'"
        '    grdSOTORDRX.Text = "Open Sales Orders associated with " & CUST_CODE
        'ElseIf optShowOrders.Value = "D" Then
        '    ASCMAIN1.sql &= " and SOTORDR1.ORDR_DATE BETWEEN '" & dteSearchS.DateTime.ToString("dd-MMM-yyyy") & "' and '" & dteSearchE.DateTime.ToString("dd-MMM-yyyy") & "'"
        '    grdSOTORDRX.Text = "Sales Orders created between " & dteSearchS.DateTime.ToString("MM/dd/yyyy") & " and " & dteSearchE.DateTime.ToString("MM/dd/yyyy")
        'ElseIf optShowOrders.Value = "N" Then
        '    txtCustNameSearch.Text = txtCustNameSearch.Text.Trim
        '    txtCustNameSearch.Text = txtCustNameSearch.Text.Replace("'", "")
        '    If txtCustNameSearch.TextLength = 0 Then
        '        ASCMAIN1.sql &= " and ROWNUM < 1"
        '    Else
        '        ASCMAIN1.sql &= " and UPPER(SOTORDR1.CUST_STORE_NAME) LIKE '%" & txtCustNameSearch.Text.ToUpper & "%'"
        '        ASCMAIN1.sql &= " and SOTORDR1.ORDR_STATUS = '" & optCustomerOrders.Value & "'"
        '        ASCMAIN1.sql &= " and SOTORDR1.ORDR_DATE BETWEEN '" & dteSearchS.DateTime.ToString("dd-MMM-yyyy") & "' and '" & dteSearchE.DateTime.ToString("dd-MMM-yyyy") & "'"
        '    End If
        '    grdSOTORDRX.Text = optCustomerOrders.Text & " Sales Orders for Customer Name like " & txtCustNameSearch.Text
        'Else
        '    'ASCMAIN1.sql = "Select * from SOTORDR1 where ROWNUM < 1"
        '    ASCMAIN1.sql &= " and ROWNUM < 1"
        'End If
        ASCMAIN1.sql = SQL.ToString
        Fill_Records("SOTQVCA1", "", , ASCMAIN1.sql)

        SQL.Length = 0
        SQL.AppendLine("Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTSTYL1.CASE_CUBE, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_STATUS, ICTSTYL1.STYLE_ASST_QTY, ICTSTYC1.STYLE_COLOR_STATUS")
        SQL.AppendLine(", ICTSTDQ3.DATE_1, ICTSTDQ3.QTY_1, ICTSTDQ3.DATE_2, ICTSTDQ3.QTY_2, ICTSTDQ3.DATE_3, ICTSTDQ3.QTY_3, ICTSTDQ3.DATE_4, ICTSTDQ3.QTY_4")
        SQL.AppendLine(" from SOTORDR2,ICTCOLR1,ICTSTYL1,ICTSTYC1,ICTSTDQ3")
        SQL.AppendLine($" where SOTORDR2.ORDR_NO IN ({S2.ToString})")
        SQL.AppendLine("   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE")
        SQL.AppendLine("   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE")
        SQL.AppendLine("   and ICTSTYC1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE")
        SQL.AppendLine("   and ICTSTYC1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE")
        SQL.AppendLine("   and ICTSTDQ3.ORDR_GROUP_NO (+) = SOTORDR2.ORDR_NO")
        SQL.AppendLine("   and ICTSTDQ3.STYLE_CODE (+) = SOTORDR2.STYLE_CODE")
        SQL.AppendLine("   and ICTSTDQ3.COLOR_CODE (+) = SOTORDR2.COLOR_CODE")
        ASCMAIN1.sql = SQL.ToString
        Fill_Records("SOTORDR2", "", , ASCMAIN1.sql)

        Sort_grdColumns(grdSOTQVCA1, "ORDR_NO".ToLower)
        grdSOTQVCA1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        'grdSOTQVCA1.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub
#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class