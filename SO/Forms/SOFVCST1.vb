Imports System.Text

Public Class SOFVCST1
    Dim Remote As New REMOTE(Me)
    Dim SQL As New System.Text.StringBuilder

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "SOFVCSTI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("SUBSTR(S1.STYLE_CODE,0,8) AS STYLE_CODE,")
            SQL.AppendLine("S2.COLOR_CODE,")
            SQL.AppendLine("MIN(S1.STYLE_CODE) AS STYLE_CODE_FULL,")
            SQL.AppendLine("S1.STYLE_DESC,")
            SQL.AppendLine("C1.COLOR_DESC,")
            SQL.AppendLine("S2.WHSE_CODE,")
            SQL.AppendLine("S1.FABRIC_CODE,")
            SQL.AppendLine("S1.SEASON_CODE,")
            SQL.AppendLine("S1.SUB_BODY_CODE,")
            SQL.AppendLine("S1.FASHION_PROMO,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) WHSE_QTY_ON_HAND,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_ORDER,0)) WHSE_QTY_ON_ORDER,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_OPEN,0)) WHSE_QTY_OPEN,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_PICK,0)) WHSE_QTY_PICK,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_TRAN,0)) WHSE_QTY_TRAN")
            SQL.AppendLine("FROM ICTSTYL1 S1, ICTSTAT2 S2, ICTCOLR1 C1")
            SQL.AppendLine("WHERE S1.STYLE_CODE = S2.STYLE_CODE (+)")
            SQL.AppendLine("AND S2.COLOR_CODE = C1.COLOR_CODE")
            SQL.AppendLine("AND S1.SALES_DIVISION_CODE = '23'")
            SQL.AppendLine("AND S2.WHSE_CODE = 'NJE'")
            SQL.AppendLine("GROUP BY")
            SQL.AppendLine("SUBSTR(S1.STYLE_CODE,0, 8),")
            SQL.AppendLine("S1.STYLE_DESC,")
            SQL.AppendLine("S2.COLOR_CODE,")
            SQL.AppendLine("C1.COLOR_DESC,")
            SQL.AppendLine("S2.WHSE_CODE,")
            SQL.AppendLine("S1.FABRIC_CODE,")
            SQL.AppendLine("S1.SEASON_CODE,")
            SQL.AppendLine("S1.SUB_BODY_CODE,")
            SQL.AppendLine("S1.FASHION_PROMO")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "SOFVCST1", "**", 0, False,, 2)

            Create_TDA(.Tables.Add, "SOFVCST2", "**", 0, False, 3)

            With grdSOFVCST1.DisplayLayout.Bands(0)
                For Each COL_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE"}
                    .Columns(COL_NAME).Header.Fixed = True
                Next
            End With

            'With grdSOFVCST1.DisplayLayout.Bands(1)
            '    For Each COL_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE"}
            '        .Columns(COL_NAME).Header.Fixed = True
            '    Next
            'End With

            'With .Tables("SOTXXXXX").Columns
            '    .Add("BUYER1", GetType(String))
            'End With

            'SQL.Length = 0
            'SQL.AppendLine("SELECT")
            'SQL.AppendLine("  AND S1.ORDR_DATE >= :PARM3")
            'ASCMAIN1.sql = SQL.ToString
            'Create_TDA(.Tables.Add, "SOFVCST1", "**", 0, False, "DDDD")

            Create_Relation("SOFVCST1", "SOFVCST2", "STYLE_CODE,COLOR_CODE")
        End With

        grdSOFVCST1.DataSource = dst.Tables("SOFVCST1")
        'grdSOFVCST2.DataSource = dst.Tables("SOFVCST2")



        Dim Crops As New List(Of Int16)
        For i As Int16 = 1 To 20
            Crops.Add(i)
        Next
        cboCrops.DataSource = Crops
        cboCrops.SelectedIndex = 7


        'ASCMAIN1.Add_Value_List(grdSOFCSTMX, "REPORT_TYPE", , New String() {":", "I:Initial", "A:Amended", "S:Subsequent", "R:Revised"})

        Create_Summary(grdSOFVCST1, "WHSE_QTY_ON_HAND", "Sum")
        Create_Summary(grdSOFVCST1, "WHSE_QTY_ON_ORDER", "Sum")
        Create_Summary(grdSOFVCST1, "WHSE_QTY_TRAN", "Sum")
        Create_Summary(grdSOFVCST1, "WHSE_QTY_PICK", "Sum")
        Create_Summary(grdSOFVCST1, "WHSE_QTY_OPEN", "Sum")

        Sort_grdColumns(grdSOFVCST1, "STYLE_CODE, COLOR_CODE")


        'With grdSOFCSTMX.DisplayLayout.Bands(0)
        '    For Each COL_NAME As String In New String() {"CUST_CODE", "CUST_NAME"}
        '        .Columns(COL_NAME).Header.Fixed = True
        '    Next
        '    .Columns("INIT_DATE").Format = "MM/dd/yy"
        'End With


        TABLE_NAME = "SOTINVH1"
        Setup_SOFVCST2()

        EntryMode = "E"
        Call Load_Record()
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
            Case "Done"
            Case "Refresh"

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
            Case "Refresh"
                Clear_Record()
                Setup_Summary()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        'dst.EnforceConstraints = False
        'dst.Tables("SOTVCST1").Rows.Clear()
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

        Clear_Record()
        Setup_Summary()

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

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOFVCST1, "SS", "Show Filter", "Show GroupBox")
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
                'Case "Customer Master File"
                '    If Not IsNothing(grdSOFCSTMX.ActiveRow) Then
                '        Dim CUST_CODE As String = grdSOFCSTMX.ActiveRow.Cells.Item("CUST_CODE").Text
                '        If CUST_CODE.Length > 0 Then
                '            'Context_Launch("Edit", CUST_CODE, "Customer Master File", "SOTCUST1")
                '            Context_Launch("Edit", CUST_CODE, e.Tool.Key, "SOTCUST1")
                '        Else
                '            Context_Launch("Customer Master File", CUST_CODE, "Customer Master File", "SOTCUST1")
                '        End If
                '    End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Project Center"
                Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
                Context_Launch("Edit", Column_Values("JOB_NO", JOB_NO), e.Tool.Key, "PMFJOBM1")
            Case "Show Report"
                Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\randfromdrc\RandInvoices\310 West 52nd Street - 30760.pdf"
                Show_Document(FILENAME)

        End Select
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
        'dst.Tables("SOTCSTMX").Rows.Clear()

        dst.EnforceConstraints = False

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine(String.Format("SUBSTR(S1.STYLE_CODE,0,{0}) AS STYLE_CODE,", Val(cboCrops.SelectedItem)))
        SQL.AppendLine("S2.COLOR_CODE,")
        SQL.AppendLine("MIN(S1.STYLE_CODE) AS STYLE_CODE_FULL,")
        SQL.AppendLine("MIN(S1.STYLE_DESC) AS STYLE_DESC,")
        SQL.AppendLine("C1.COLOR_DESC,")
        SQL.AppendLine("S2.WHSE_CODE,")
        SQL.AppendLine("S1.FABRIC_CODE,")
        SQL.AppendLine("S1.SEASON_CODE,")
        SQL.AppendLine("S1.SUB_BODY_CODE,")
        SQL.AppendLine("S1.FASHION_PROMO,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) WHSE_QTY_ON_HAND,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_ORDER,0)) WHSE_QTY_ON_ORDER,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_OPEN,0)) WHSE_QTY_OPEN,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_PICK,0)) WHSE_QTY_PICK,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_TRAN,0)) WHSE_QTY_TRAN")
        SQL.AppendLine("FROM ICTSTYL1 S1, ICTSTAT2 S2, ICTCOLR1 C1")
        SQL.AppendLine("WHERE S1.STYLE_CODE = S2.STYLE_CODE (+)")
        SQL.AppendLine("AND S2.COLOR_CODE = C1.COLOR_CODE")
        SQL.AppendLine("AND S1.SALES_DIVISION_CODE = '23'")
        SQL.AppendLine("AND S2.WHSE_CODE = 'NJE'")
        SQL.AppendLine("GROUP BY")
        SQL.AppendLine(String.Format("SUBSTR(S1.STYLE_CODE,0, {0}),", Val(cboCrops.SelectedItem)))
        SQL.AppendLine("S2.COLOR_CODE,")
        SQL.AppendLine("C1.COLOR_DESC,")
        SQL.AppendLine("S2.WHSE_CODE,")
        SQL.AppendLine("S1.FABRIC_CODE,")
        SQL.AppendLine("S1.SEASON_CODE,")
        SQL.AppendLine("S1.SUB_BODY_CODE,")
        SQL.AppendLine("S1.FASHION_PROMO")
        Fill_Records("SOFVCST1",,, SQL.ToString)

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine(String.Format("SUBSTR(S1.STYLE_CODE,0,{0}) AS STYLE_CODE,", Val(cboCrops.SelectedItem)))
        SQL.AppendLine("S2.COLOR_CODE,")
        SQL.AppendLine("S1.STYLE_CODE AS STYLE_CODE_FULL,")
        SQL.AppendLine("S1.STYLE_DESC,")
        SQL.AppendLine("C1.COLOR_DESC,")
        SQL.AppendLine("S2.WHSE_CODE,")
        SQL.AppendLine("S1.FABRIC_CODE,")
        SQL.AppendLine("S1.SEASON_CODE,")
        SQL.AppendLine("S1.SUB_BODY_CODE,")
        SQL.AppendLine("S1.FASHION_PROMO,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) WHSE_QTY_ON_HAND,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_ORDER,0)) WHSE_QTY_ON_ORDER,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_OPEN,0)) WHSE_QTY_OPEN,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_PICK,0)) WHSE_QTY_PICK,")
        SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_TRAN,0)) WHSE_QTY_TRAN")
        SQL.AppendLine("FROM ICTSTYL1 S1, ICTSTAT2 S2, ICTCOLR1 C1")
        SQL.AppendLine("WHERE S1.STYLE_CODE = S2.STYLE_CODE (+)")
        SQL.AppendLine("AND S2.COLOR_CODE = C1.COLOR_CODE")
        SQL.AppendLine("AND S1.SALES_DIVISION_CODE = '23'")
        SQL.AppendLine("AND S2.WHSE_CODE = 'NJE'")
        If chkExclOrig.Checked Then
            SQL.AppendLine(String.Format("AND S1.STYLE_CODE <> SUBSTR(S1.STYLE_CODE,0,{0})", Val(cboCrops.SelectedItem)))
        End If
        SQL.AppendLine("GROUP BY")
        SQL.AppendLine(String.Format("SUBSTR(S1.STYLE_CODE,0, {0}),", Val(cboCrops.SelectedItem)))
        SQL.AppendLine("S2.COLOR_CODE,")
        SQL.AppendLine("S1.STYLE_CODE,")
        SQL.AppendLine("S1.STYLE_DESC,")
        SQL.AppendLine("C1.COLOR_DESC,")
        SQL.AppendLine("S2.WHSE_CODE,")
        SQL.AppendLine("S1.FABRIC_CODE,")
        SQL.AppendLine("S1.SEASON_CODE,")
        SQL.AppendLine("S1.SUB_BODY_CODE,")
        SQL.AppendLine("S1.FASHION_PROMO")
        Fill_Records("SOFVCST2",,, SQL.ToString)

        ASCMAIN1.Progress("")
        grdSOFVCST1.Update()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Sub Setup_SOFVCST2()
        Exit Sub
        If grdSOFVCST1.ActiveRow Is Nothing OrElse (Not grdSOFVCST1.ActiveRow.IsDataRow Or grdSOFVCST1.ActiveRow.IsAddRow) Then
            'grpSOTORDR3.Visible = False
        Else
            'Dim dvw As DataView = DirectCast(grdSOFVCST2.DataSource, DataTable).DefaultView
            Dim STYLE_CODE As String = grdSOFVCST1.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdSOFVCST1.ActiveRow.Cells("COLOR_CODE").Value
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("S1.STYLE_CODE,")
            SQL.AppendLine("S1.STYLE_DESC,")
            SQL.AppendLine("S2.COLOR_CODE,")
            SQL.AppendLine("C1.COLOR_DESC,")
            SQL.AppendLine("S2.WHSE_CODE,")
            SQL.AppendLine("S1.FABRIC_CODE,")
            SQL.AppendLine("S1.SEASON_CODE,")
            SQL.AppendLine("S1.SUB_BODY_CODE,")
            SQL.AppendLine("S1.FASHION_PROMO,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) WHSE_QTY_ON_HAND,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_ORDER,0)) WHSE_QTY_ON_ORDER,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_OPEN,0)) WHSE_QTY_OPEN,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_PICK,0)) WHSE_QTY_PICK,")
            SQL.AppendLine("SUM(NVL(S2.WHSE_QTY_TRAN,0)) WHSE_QTY_TRAN")
            SQL.AppendLine("FROM ICTSTYL1 S1, ICTSTAT2 S2, ICTCOLR1 C1")
            SQL.AppendLine("WHERE S1.STYLE_CODE = S2.STYLE_CODE (+)")
            SQL.AppendLine("AND S2.COLOR_CODE = C1.COLOR_CODE")
            SQL.AppendLine("AND S1.SALES_DIVISION_CODE = '23'")
            SQL.AppendLine("AND S2.WHSE_CODE = 'NJE'")
            SQL.AppendLine(String.Format("AND S1.STYLE_CODE LIKE '{0}%'", STYLE_CODE))
            SQL.AppendLine(String.Format("AND S2.COLOR_CODE = '{0}'", COLOR_CODE))
            SQL.AppendLine("GROUP BY")
            SQL.AppendLine("S1.STYLE_CODE,")
            SQL.AppendLine("S1.STYLE_DESC,")
            SQL.AppendLine("S2.COLOR_CODE,")
            SQL.AppendLine("C1.COLOR_DESC,")
            SQL.AppendLine("S2.WHSE_CODE,")
            SQL.AppendLine("S1.FABRIC_CODE,")
            SQL.AppendLine("S1.SEASON_CODE,")
            SQL.AppendLine("S1.SUB_BODY_CODE,")
            SQL.AppendLine("S1.FASHION_PROMO")
            Fill_Records("SOFVCST2",,, SQL.ToString)
        End If
    End Sub

    Private Sub grdSOFVCST1_AfterRowActivate(sender As Object, e As EventArgs)
        Setup_SOFVCST2()
    End Sub

    Private Sub grdSOFVCST1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOFVCST1.InitializeLayout

    End Sub
End Class