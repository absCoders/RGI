
Imports System.Text
Imports Microsoft.Office.Interop.Word

Public Class WBFWEBQT
    Dim InquiryOnly As Boolean = False
    Dim S As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With dst

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTQRDR1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "SOTQRDR1", "**", 0, True, "", 1)
            .Tables("SOTQRDR1").Columns.Add("CALC_STATUS", GetType(System.String))
            dst.Tables("SOTQRDR1").Columns.Add("ERRORS")
            dst.Tables("SOTQRDR1").Columns.Add("EXCEPTIONS")

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTQRDR2")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "SOTQRDR2", "**", 0, False, "", 2)
            .Tables("SOTQRDR2").Columns.Add("LINE_TOTAL", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            dst.Tables("SOTQRDR2").Columns.Add("ERRORS")
            dst.Tables("SOTQRDR2").Columns.Add("EXCEPTIONS")

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM SOTQRDR5")
            S.AppendLine("WHERE ORDR_NO = :PARM1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "SOTQRDR5", "**", 0, False, "V", 2)

            'Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            'ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
            'Create_TDA(.Tables.Add, "ICTSTYL1", "**", 1, False)
            'Fill_Records("ICTSTYL1", "", , ASCMAIN1.sql)
        End With

        grdSOTQRDR1.DataSource = dst.Tables("SOTQRDR1")
        grdSOTQRDR2.DataSource = dst.Tables("SOTQRDR2")
        Setup_SOTQRDR2()

        Create_Summary(grdSOTQRDR2, "LINE_TOTAL", "Sum", "", "###,##0")

        ASCMAIN1.Add_Value_List(grdSOTQRDR1, "CALC_STATUS", , New String() {":", "I:Imported From Web", "L:Pulled To Laptop", "O:Finalized As Order", "X:Deleted", "M:Marked Complete", "T:Testing"})
        Sort_grdColumns(grdSOTQRDR1, "ORDR_DATE".ToLower(), False)

        With grdSOTQRDR1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdSOTQRDR1.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOTQRDR1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        With grdSOTQRDR2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With
        For i As Integer = 0 To grdSOTQRDR2.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOTQRDR2.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        Load_Record()

        tab.Visible = False
        isFormLoading = False
        filterGrid()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Print Order"
                EMsg &= CheckGridSelection()
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
            Case "Print Order"
                Dim ORDR_NO As String = grdSOTQRDR1.Selected.Rows(0).Cells.Item("ORDR_NO").Text
                Dim CUST_CODE As String = grdSOTQRDR1.Selected.Rows(0).Cells.Item("CUST_CODE").Text
                Print_Record(True, ORDR_NO, CUST_CODE)
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
                .Groups("Screen Control").Items("Refresh").Visible = True
                .Groups("Screen Control").Items("Print Order").Visible = True
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

        Fill_Records("SOTQRDR1")
        Fill_Records("SOTQRDR2")

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        CalculateStatus()

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Update_Record_TDA("SOTQRDR1")
        Call CommitTrans("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        Print_Report_Begin()
        'frm.CR_params.Add("SUBT", "")
        'Fill SOTORDRP records
        Fill_Records("SOTQRDR5", ORDR_NO, True)
        For Each rowSOTQRDR1 As DataRow In dst.Tables("SOTQRDR1").Select()
            If rowSOTQRDR1.Item("ORDR_NO") = ORDR_NO Then
                rowSOTQRDR1.Item("ERRORS") = "NEW"
            Else
                rowSOTQRDR1.Item("ERRORS") = ""
            End If
        Next
        'Generate_Report("SORQRDRO")
        Generate_Report("WBRWEBQT", "Quotes Imported From Web", "Re-printed From Quote Maint.")
        '    Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTQRDR1, "SSBBBBB", "Show Filter", "Show GroupBox", "Release Quote For Re-import", "Mark Quote As Complete", "Mark Quote As Testing", "Re-Assign Quote To New Order", "Delete Quote")
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
            Case "Release Quote For Re-import"
                grd.ActiveRow.Cells.Item("ORDR_NO_WEB").Value = ""
            Case "Mark Quote As Complete"
                grd.ActiveRow.Cells.Item("ORDR_NO_WEB").Value = "COMPLETE"
            Case "Mark Quote As Testing"
                grd.ActiveRow.Cells.Item("ORDR_NO_WEB").Value = "TESTING"
            Case "Delete Quote"
                grd.ActiveRow.Cells.Item("ORDR_NO_WEB").Value = "DELETED"
            Case "Re-Assign Quote To New Order"
                Dim NEW_ORDR As String = InputBox("Please Provide New Order", "New Order No")
                If NEW_ORDR.Length <> 10 Then
                    MsgBox("Must Be 10 Digits", vbCritical, "Sorry, Try Again")
                Else
                    If Not IsNumeric(NEW_ORDR) Then
                        MsgBox("Non-Numeric Characters", vbCritical, "Sorry, Try Again")
                    Else
                        grd.ActiveRow.Cells.Item("ORDR_NO_WEB").Value = NEW_ORDR
                    End If
                End If
        End Select
        CalculateStatus()
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
    Private Sub CalculateStatus()
        For Each rowSOTQRDR1 As DataRow In dst.Tables("SOTQRDR1").Select()
            rowSOTQRDR1.Item("CALC_STATUS") = ""
            Dim ORDR_NO_WEB As String = rowSOTQRDR1.Item("ORDR_NO_WEB").ToString & String.Empty
            Select Case ORDR_NO_WEB
                Case ""
                    rowSOTQRDR1.Item("CALC_STATUS") = "I"
                Case "DELETED"
                    rowSOTQRDR1.Item("CALC_STATUS") = "X"
                Case "COMPLETE"
                    rowSOTQRDR1.Item("CALC_STATUS") = "M"
                Case "TESTING"
                    rowSOTQRDR1.Item("CALC_STATUS") = "T"
                Case Else
                    S.Length = 0
                    S.AppendLine("SELECT COUNT(*)")
                    S.AppendLine("FROM SOTORDR1")
                    S.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO_WEB))
                    ASCMAIN1.sql = S.ToString()
                    Dim ORDR_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                    If ORDR_CNT = 0 Then
                        rowSOTQRDR1.Item("CALC_STATUS") = "L"
                    Else
                        rowSOTQRDR1.Item("CALC_STATUS") = "O"
                    End If
            End Select

        Next
        filterGrid()
    End Sub
    Private Function CheckGridSelection() As String
        Dim EMsg As String = ""
        If grdSOTQRDR1.Selected.Rows.Count = 0 Then
            EMsg &= vbCr & "You Must Select At Least One Order From The Grid"
        End If
        If grdSOTQRDR1.Selected.Rows.Count > 1 Then
            EMsg &= vbCr & "You May Only Select One Order At A Time"
        End If
        Return EMsg
    End Function
    Private Sub filterGrid()
        If Not isFormLoading Then
            Dim dvw As DataView = DirectCast(grdSOTQRDR1.DataSource, DataTable).DefaultView
            Dim filter As New Text.StringBuilder With {.Length = 0}

            If chkFilterI.Checked Then
                filter.AppendLine("CALC_STATUS = 'I'")
            End If

            If chkFilterL.Checked Then
                If filter.Length = 0 Then
                    filter.AppendLine("CALC_STATUS = 'L'")
                Else
                    filter.AppendLine(" OR CALC_STATUS = 'L'")
                End If
            End If

            If chkFilterO.Checked Then
                If filter.Length = 0 Then
                    filter.AppendLine("CALC_STATUS = 'O'")
                Else
                    filter.AppendLine(" OR CALC_STATUS = 'O'")
                End If
            End If

            If chkFilterX.Checked Then
                If filter.Length = 0 Then
                    filter.AppendLine("CALC_STATUS = 'X'")
                Else
                    filter.AppendLine(" OR CALC_STATUS = 'X'")
                End If
            End If

            If chkFilterM.Checked Then
                If filter.Length = 0 Then
                    filter.AppendLine("CALC_STATUS = 'M'")
                Else
                    filter.AppendLine(" OR CALC_STATUS = 'M'")
                End If
            End If

            If chkFilterT.Checked Then
                If filter.Length = 0 Then
                    filter.AppendLine("CALC_STATUS = 'T'")
                Else
                    filter.AppendLine(" OR CALC_STATUS = 'T'")
                End If
            End If

            dvw.RowFilter = String.Format(filter.ToString, "")
        End If
    End Sub
    Sub Setup_SOTQRDR2()
        If grdSOTQRDR1.ActiveRow Is Nothing OrElse (Not grdSOTQRDR1.ActiveRow.IsDataRow Or grdSOTQRDR1.ActiveRow.IsAddRow) Then
            grdSOTQRDR2.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdSOTQRDR2.DataSource, DataTable).DefaultView
            Dim ORDR_NO As String = grdSOTQRDR1.ActiveRow.Cells("ORDR_NO").Value
            dvw.RowFilter = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            grdSOTQRDR2.Text = "Details for Quote " & CStr(ORDR_NO)
            grdSOTQRDR2.Visible = True
        End If
    End Sub
#End Region

#Region "Form Controls"
    Private Sub chkFilterI_CheckedChanged(sender As Object, e As EventArgs) Handles chkFilterI.CheckedChanged
        filterGrid()
    End Sub

    Private Sub chkFilterL_CheckedChanged(sender As Object, e As EventArgs) Handles chkFilterL.CheckedChanged
        filterGrid()
    End Sub

    Private Sub chkFilterO_CheckedChanged(sender As Object, e As EventArgs) Handles chkFilterO.CheckedChanged
        filterGrid()
    End Sub

    Private Sub chkFilterX_CheckedChanged(sender As Object, e As EventArgs) Handles chkFilterX.CheckedChanged
        filterGrid()
    End Sub

    Private Sub chkFilterM_CheckedChanged(sender As Object, e As EventArgs) Handles chkFilterM.CheckedChanged
        filterGrid()
    End Sub

    Private Sub chkFilterT_CheckedChanged(sender As Object, e As EventArgs) Handles chkFilterT.CheckedChanged
        filterGrid()
    End Sub
#Region "Grids"
    Private Sub grdSOTQRDR1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTQRDR1.AfterRowActivate
        Setup_SOTQRDR2()
    End Sub

#End Region
#End Region

End Class