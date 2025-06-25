Imports Infragistics.Win.UltraWinGrid

Public Class ICFSTYPC
    Dim S As New Text.StringBuilder() With {.Length = 0}
    Dim STY As New Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With dst
            STY.Length = 0
            STY.AppendLine("SELECT")
            STY.AppendLine("ST1.STYLE_CODE,")
            STY.AppendLine("ST1.STYLE_DESC,")
            STY.AppendLine("ST1.STYLE_STATUS,")
            STY.AppendLine("ST1.STYLE_CLASS_CODE,")
            STY.AppendLine("CL1.STYLE_CLASS_DESC,")
            STY.AppendLine("ST1.COUNTRY_CODE,")
            STY.AppendLine("ST1.VEND_CODE,")
            STY.AppendLine("AP1.VEND_NAME,")
            STY.AppendLine("ST1.DUTY_RATE_CODE,")
            STY.AppendLine("DT1.DUTY_RATE_DESC,")
            STY.AppendLine("ST1.ROYALTY_CODE,")
            STY.AppendLine("RY1.ROYALTY_DESC,")
            STY.AppendLine("ST1.STYLE_PRICE,")
            STY.AppendLine("ST1.STYLE_PRICE AS STYLE_PRICE_ORIG")
            STY.AppendLine("FROM ICTSTYL1 ST1, ICTCLAS1 CL1, APTVEND1 AP1, ICTDUTY1 DT1, ICTROYL1 RY1")
            STY.AppendLine("WHERE ST1.STYLE_CLASS_CODE = CL1.STYLE_CLASS_CODE (+)")
            STY.AppendLine("AND ST1.VEND_CODE = AP1.VEND_CODE (+)")
            STY.AppendLine("AND ST1.DUTY_RATE_CODE = DT1.DUTY_RATE_CODE (+)")
            STY.AppendLine("AND ST1.ROYALTY_CODE = RY1.ROYALTY_CODE (+)")
            ASCMAIN1.sql = STY.ToString()
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, True, "",, "STYLE_PRICE")
            .Tables("ICTSTYL1").Columns.Add("SEL", GetType(System.String))
            .Tables("ICTSTYL1").Columns.Add("VAR", GetType(System.Decimal), "STYLE_PRICE - STYLE_PRICE_ORIG")

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("STYLE_CLASS_CODE,")
            S.AppendLine("STYLE_CLASS_DESC")
            S.AppendLine("FROM ICTCLAS1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)
            .Tables("ICTCLAS1").Columns.Add("SEL", GetType(System.String))

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("ATTR_CODE,")
            S.AppendLine("ATTR_DESC")
            S.AppendLine("FROM ICTATTR1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTATTR1", "**", 0, False)
            .Tables("ICTATTR1").Columns.Add("SEL", GetType(System.String))

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("*")
            S.AppendLine("FROM ICTSTYPC")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ICTSTYPC", "**", 0, True)

            'S.Length = 0
            'S.AppendLine("SELECT")
            'S.AppendLine("*")
            'S.AppendLine("FROM ASTAUDT1")
            'ASCMAIN1.sql = S.ToString()
            'Create_TDA(.Tables.Add, "ASTAUDT1", "**", 0, True)
        End With

        grdICTSTYL1.DataSource = dst.Tables("ICTSTYL1")
        grdICTCLAS1.DataSource = dst.Tables("ICTCLAS1")
        grdICTATTR1.DataSource = dst.Tables("ICTATTR1")

        Create_Summary(grdICTSTYL1, "STYLE_CODE", "Count", "", "###,###,##0")
        Create_Summary(grdICTSTYL1, "SEL")

        ASCMAIN1.Add_Value_List(grdICTSTYL1, "STYLE_STATUS", , New String() {":", "A:Active", "N:Do Not Reorder", "D:Discontinued"})
        Sort_grdColumns(grdICTCLAS1, "STYLE_CLASS_CODE", False)
        Sort_grdColumns(grdICTSTYL1, "STYLE_CODE", False)
        Sort_grdColumns(grdICTATTR1, "ATTR_CODE", False)

        With grdICTSTYL1.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
            With .Columns("STYLE_PRICE_ORIG")
                .CellActivation = UltraWinGrid.Activation.AllowEdit
                .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                .Format = "###,##0.00"
                .Header.Appearance.BackColor2 = Drawing.Color.Green
            End With
            With .Columns("STYLE_PRICE")
                .Format = "###,##0.00"
                .Header.Appearance.BackColor2 = Drawing.Color.Green
            End With
            With .Columns("VAR")
                .Format = "###,##0.00"
                .Header.Appearance.BackColor2 = Drawing.Color.Green
            End With
            '.Columns("STYLE_CODE").Header.Fixed = True
            '.Columns("XXXXXXX").Format = "MM/dd/yyyy hh:mm"
        End With

        With grdICTCLAS1.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        With grdICTATTR1.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        Fill_Records("ICTCLAS1")
        Fill_Records("ICTATTR1")

        isFormLoading = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Update Prices"
                Dim eWARN As New Text.StringBuilder With {.Length = 0}
                eWARN.AppendLine("WARNING: This Process Can Not Be")
                eWARN.AppendLine("Run While Users Are Accessing The")
                eWARN.AppendLine("System!!!")
                eWARN.AppendLine("")
                eWARN.AppendLine("If Others Are Processing Data")
                eWARN.AppendLine("Some Of These Changes May Be")
                eWARN.AppendLine("Over-Written!!")
                eWARN.AppendLine("")
                eWARN.AppendLine("Are You Ready?")
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Be Careful!"
                iResult = MsgBox(eWARN.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg = EMsg & vbCrLf & "No Update Done."
                End If
            Case "Save"
                'EMsg = EMsg & vbCrLf & "Feature Not Done Yet."
            Case "Refresh"
                'EMsg = EMsg & vbCrLf & "Feature Not Done Yet."
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update Prices"
                Call Mode_Settings(False)
                Update_Record()
                Clear_Record()
                MsgBox("Prices Updated.  Please Check Masterfiles", vbOKOnly, "Complete")
                Me.Close()
            Case "Refresh"
                Call Mode_Settings(True)
                RefreshData()
            Case "Cancel"
                Call Mode_Settings(False)
                Clear_Record()
                Me.Close()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode  '= iScreenMode
                .Groups("Screen Control").Items("Update Prices").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        dst.Tables("ICTSTYL1").Rows.Clear()
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        dst.Tables.Item("ICTSTYPC").Clear()

        Dim BATCH_NO As String = ASCMAIN1.Next_Control_No("ICTSTYPC.BATCH_NO")
        Dim INIT_OPER As String = ASCMAIN1.USER_ID
        Dim INIT_DATE As Date = DATETIME_STAMP
        For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select()
            Dim STYLE_PRICE_ORIG As Double = Val(rowICTSTYL1.Item("STYLE_PRICE_ORIG").ToString & String.Empty)
            Dim STYLE_PRICE As Double = Val(rowICTSTYL1.Item("STYLE_PRICE").ToString & String.Empty)
            Dim STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty
            If STYLE_PRICE_ORIG <> STYLE_PRICE Then
                Dim rowICTSTYPC As DataRow = dst.Tables("ICTSTYPC").NewRow
                rowICTSTYPC.Item("BATCH_NO") = BATCH_NO
                rowICTSTYPC.Item("STYLE_CODE") = STYLE_CODE
                rowICTSTYPC.Item("STYLE_PRICE_ORIG") = STYLE_PRICE_ORIG
                rowICTSTYPC.Item("STYLE_PRICE") = STYLE_PRICE
                rowICTSTYPC.Item("INIT_DATE") = INIT_DATE
                rowICTSTYPC.Item("INIT_OPER") = INIT_OPER
                dst.Tables.Item("ICTSTYPC").Rows.Add(rowICTSTYPC)

                Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                rowASTAUDT1.Item("TABLE_NAME") = "ICTSTYL1"
                rowASTAUDT1.Item("KEY_VALUE") = STYLE_CODE
                rowASTAUDT1.Item("COLUMN_NAME") = "STYLE_PRICE"
                rowASTAUDT1.Item("USER_ID") = ASCMAIN1.USER_ID
                rowASTAUDT1.Item("INIT_DATE") = DATETIME_STAMP
                rowASTAUDT1.Item("OLD_VALUE") = STYLE_PRICE_ORIG
                rowASTAUDT1.Item("NEW_VALUE") = STYLE_PRICE
                rowASTAUDT1.Item("FM_MODE") = "E"
                rowASTAUDT1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                rowASTAUDT1.Item("SELECTION_NO") = SELECTION_NO
                rowASTAUDT1.Item("XNO") = XNO
                rowASTAUDT1.Item("NOTES") = ""
                rowASTAUDT1.Item("XNO") = XNO
                dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
            End If
        Next
        Update_Record_TDA("ASTAUDT1")
        Update_Record_TDA("ICTSTYPC")
        Update_Record_TDA("ICTSTYL1")
        Call CommitTrans("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYL1, "SSBB", "Show Filter", "Show GroupBox", "Select All", "Select None")
        Load_Popup_Menu(grdICTCLAS1, "SSBB", "Show Filter", "Show GroupBox", "Select All", "Select None")
        Load_Popup_Menu(grdICTATTR1, "SSBB", "Show Filter", "Show GroupBox", "Select All", "Select None")
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
            Case "Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    If Not grow.Hidden Then
                        grow.Cells.Item("SEL").Value = "1"
                        grow.Update()
                    End If
                Next

            Case "Select None"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    If Not grow.Hidden Then
                        grow.Cells.Item("SEL").Value = "0"
                        grow.Update()
                    End If
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Selected.Rows.Count = 0 Then
        '    MsgBox("You Must Select One And Only One Row First", vbOKOnly, "Select A Row")
        '    Exit Sub
        'End If

        Select Case e.Tool.Key

        End Select

        'Update_Record()
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
        Dim SQL As String = STY.ToString
        Dim SQLP As New Text.StringBuilder With {.Length = 0}
        If chkLoadTemp.Checked Then
            SQL = SQL + vbCrLf + $"AND ST1.STYLE_CODE IN (SELECT STYLE_CODE FROM ICTSTYPC WHERE BATCH_NO = '9999999999')"
        Else
            If chkActiveOnly.Checked Then
                SQLP.AppendLine("AND ST1.STYLE_CODE IN (")
                SQLP.AppendLine("SELECT DISTINCT STYLE_CODE FROM")
                SQLP.AppendLine("(")
                SQLP.AppendLine("   SELECT")
                SQLP.AppendLine("   S1.STYLE_CODE,")
                SQLP.AppendLine("   C1.COLOR_CODE,")
                SQLP.AppendLine("   S1.STYLE_STATUS,")
                SQLP.AppendLine("   C1.STYLE_COLOR_STATUS,")
                SQLP.AppendLine("   S1.STYLE_DESC,")
                SQLP.AppendLine("   SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))) AS AVAIL")
                SQLP.AppendLine("   FROM ICTSTYL1 S1, ICTSTYC1 C1, ICTSTAT2 S2")
                SQLP.AppendLine("   WHERE S1.STYLE_CODE = C1.STYLE_CODE")
                SQLP.AppendLine("   AND C1.STYLE_CODE = S2.STYLE_CODE (+)")
                SQLP.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE (+)")
                SQLP.AppendLine("   AND S2.WHSE_CODE = 'MS'")
                SQLP.AppendLine("   GROUP BY")
                SQLP.AppendLine("   S1.STYLE_CODE,")
                SQLP.AppendLine("   C1.COLOR_CODE,")
                SQLP.AppendLine("   S1.STYLE_STATUS,")
                SQLP.AppendLine("   C1.STYLE_COLOR_STATUS,")
                SQLP.AppendLine("   S1.STYLE_DESC")
                SQLP.AppendLine(")")
                SQLP.AppendLine("WHERE (STYLE_COLOR_STATUS = 'A' OR AVAIL > 0)")
                SQLP.AppendLine(")")
                SQL = SQL + vbCrLf + SQLP.ToString
            End If
            Dim CAS As String = getCLASS_ATTR_SQL()
            If CAS.Length > 0 Then
                SQL = SQL + vbCrLf + $"AND ST1.STYLE_CODE IN ( {CAS} )"
            End If
        End If

        Fill_Records("ICTSTYL1",, True, SQL)

        If chkLoadTemp.Checked Then
            Dim s As New Text.StringBuilder With {.Length = 0}
            s.AppendLine("SELECT * FROM ICTSTYPC WHERE BATCH_NO = '9999999999'")
            Dim tblICTSTYPC As DataTable = ASCDATA1.GetDataTable(s.ToString())
            For Each rowICTSTYPC As DataRow In tblICTSTYPC.Rows
                Dim SC As String = rowICTSTYPC.Item("STYLE_CODE").ToString & String.Empty
                Dim rowICTSTYL1 As DataRow = dst.Tables.Item("ICTSTYL1").Select($"STYLE_CODE = '{SC}'").FirstOrDefault
                If Not IsNothing(rowICTSTYL1) Then
                    rowICTSTYL1.Item("STYLE_PRICE") = rowICTSTYPC.Item("STYLE_PRICE")
                    rowICTSTYL1.Item("SEL") = "1"
                End If
            Next
        End If

        ASCMAIN1.Progress("", "")
    End Sub

    Private Function getCLASS_ATTR_SQL() As String
        Dim RETSQL As New Text.StringBuilder With {.Length = 0}
        Dim FLTR As String = "SEL = '1'"

        Dim SCC As String = ""
        For Each rowICTCLAS1 As DataRow In dst.Tables("ICTCLAS1").Select(FLTR)
            SCC = SCC + $"'{rowICTCLAS1.Item("STYLE_CLASS_CODE").ToString}',"
        Next
        If SCC.Length > 0 Then
            SCC = $"({SCC.Substring(0, SCC.Length - 1)})"
        End If

        Dim SAC As String = ""
        For Each rowICTATTR1 As DataRow In dst.Tables("ICTATTR1").Select(FLTR)
            SAC = SAC + $"'{rowICTATTR1.Item("ATTR_CODE").ToString}',"
        Next
        If SAC.Length > 0 Then
            SAC = $"({SAC.Substring(0, SAC.Length - 1)})"
        End If

        If SCC.Length > 0 Or SAC.Length > 0 Then
            RETSQL.AppendLine("SELECT S1.STYLE_CODE")
            RETSQL.AppendLine("FROM ICTSTYL1 S1, ICTSTYL3 S3")
            RETSQL.AppendLine("WHERE S1.STYLE_CODE = S3.STYLE_CODE")
            If SCC.Length > 0 Then
                RETSQL.AppendLine($"AND S1.STYLE_CLASS_CODE IN {SCC}")
            End If
            If SAC.Length > 0 Then
                RETSQL.AppendLine($"AND S3.ATTR_CODE IN {SAC}")
            End If

        End If

        Return RETSQL.ToString
    End Function

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Dim markDownPct As Double = numMarkDownPct.Value
        Dim eMsg As New Text.StringBuilder With {.Length = 0}
        'If markDownPct = 0 Then
        '    eMsg.AppendLine("Markdown Pct Can Not be Zero.")
        'End If
        If markDownPct < -200 Or markDownPct > 200 Then
            eMsg.AppendLine("Markdown Pct Limited to 200%.")
        End If
        If eMsg.Length > 0 Then
            MsgBox(eMsg.ToString, vbCritical, "Can Not Calculate")
        Else
            markDownPct = markDownPct / 100
            For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("SEL = '1'")
                Dim STYLE_PRICE_ORIG As Double = Val(rowICTSTYL1.Item("STYLE_PRICE_ORIG").ToString & String.Empty)
                rowICTSTYL1.Item("STYLE_PRICE") = Math.Round(STYLE_PRICE_ORIG * (1 + markDownPct), 2)
            Next
        End If
    End Sub

#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class