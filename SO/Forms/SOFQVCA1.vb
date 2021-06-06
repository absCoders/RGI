
Imports System.Text
Imports System.IO

Public Class SOFQVCA1
    Dim SQL As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With dst
            SQL.Length = 0
            SQL.AppendLine("Select SOTORDR1.* from SOTORDR1 where ORDR_NO = :PARM1")
            ASCMAIN1.sql = SQL.ToString()
            Create_TDA(.Tables.Add, "SOTQVCA1", "**", 0, False, "V", 1)
            '.Tables("SOTQVCA1").Columns.Add("ItemID", GetType(System.String))

            ASCMAIN1.sql = "Select SOTORDR1.* from SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 1)
            .Tables("SOTORDRX").Columns.Add("CUST_CITY", GetType(System.String))
            .Tables("SOTORDRX").Columns.Add("CUST_STATE", GetType(System.String))
            .Tables("SOTORDRX").Columns.Add("CUST_COUNTRY", GetType(System.String))
            'ORDR_QTY_OPEN
            .Tables("SOTORDRX").Columns.Add("ORDR_QTY_OPEN", GetType(System.Decimal))
            .Tables("SOTORDRX").Columns.Add("ORDR_QTY_PICK", GetType(System.Decimal))
            .Tables("SOTORDRX").Columns.Add("ORDR_QTY_SHIP", GetType(System.Decimal))
            .Tables("SOTORDRX").Columns.Add("ORDR_QTY_CANC", GetType(System.Decimal))
            .Tables("SOTORDRX").Columns.Add("ORDR_TOTAL", GetType(System.Decimal), "ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP")

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)

        End With

        'Fill_Records("SOTQVCA1")

        grdSOTQVCA1.DataSource = dst.Tables("SOTQVCA1")

        Create_Summary(grdSOTQVCA1, "STYLE_CODE", "Count", "", "###,##0")


        'ASCMAIN1.Add_Value_List(grdSOTQRDR1, "CALC_STATUS", , New String() {":", "I:Imported From Web", "L:Pulled To Laptop", "O:Finalized As Order", "X:Deleted", "M:Marked Complete", "T:Testing"})
        Sort_grdColumns(grdSOTQVCA1, "STYLE_CODE", False)

        With grdSOTQVCA1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        For i As Integer = 0 To grdSOTQVCA1.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOTQVCA1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
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
                Fill_Records("SOTQVCA1")
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

        If grd.Selected.Rows.Count = 0 Then
            MsgBox("You Must Select One And Only One Row First", vbOKOnly, "Select A Row")
            Exit Sub
        End If

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

        Fill_Records("SOTQVCA1")

        ASCMAIN1.Progress("", "")
        'grdECTSZIO1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ResizeAllColumns
    End Sub
#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class