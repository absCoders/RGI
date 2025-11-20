Imports System.Net.Mail
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports Infragistics.Documents.Excel
Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinSchedule

Public Class ARFCUSTM
    Dim S As New System.Text.StringBuilder With {.Length = 0}
    Dim Loading As Boolean = True
    Dim TEMP_SALES As String = ""
    Dim YR1 As String = ""
    Dim YR2 As String = ""
    Dim YR3 As String = ""
    Dim YR4 As String = ""
    Dim YRFR As String = ""
    Dim YRTO As String = ""
    Dim CONTACT_COLS As New Dictionary(Of KeyValuePair(Of String, Int64), List(Of String))


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        MAKE_CONTACT_COLS()

        Check_Form_Options()

        Dim BaseYear As Int64 = Now().Year
        YR1 = (BaseYear).ToString
        YR2 = (BaseYear - 1).ToString
        YR3 = (BaseYear - 2).ToString
        YR4 = (BaseYear - 3).ToString

        YRFR = Now.AddYears(-2).Year.ToString().Substring(2, 2).ToString
        YRTO = Now.Year.ToString().Substring(2, 2).ToString

        'Fill In The Gaps
        'REFRESH_CONTACTS()

        RefreshSalesTempTable()

        With dst
            'S.Length = 0
            'S.AppendLine("SELECT")
            'S.AppendLine("C1.CUST_CODE,")
            'S.AppendLine("C1.CUST_NAME,")
            'S.AppendLine("C1.CUST_CITY,")
            'S.AppendLine("C1.CUST_STATE,")
            'S.AppendLine("C1.CUST_ZIP_CODE,")
            'S.AppendLine("C1.CUST_COUNTRY,")
            'S.AppendLine("C1.CUST_STMT_IND,")
            'S.AppendLine("C1.CUST_STMT_EMAIL,")
            'S.AppendLine("C1.SREP_CODE,")
            'S.AppendLine("C1.INIT_DATE,")
            'S.AppendLine("C1.CUST_XMIT_INV_VIA,")
            'S.AppendLine("C1.CUST_INV_EMAIL,")
            'S.AppendLine("SUM(SALES.YR1) AS YR1,")
            'S.AppendLine("SUM(SALES.YR2) AS YR2,")
            'S.AppendLine("SUM(SALES.YR3) AS YR3,")
            'S.AppendLine("SUM(SALES.YR4) AS YR4")
            'S.AppendLine($"FROM ARTCUST1 C1, {TEMP_SALES} SALES")
            'S.AppendLine("WHERE C1.CUST_CODE = SALES.CUST_CODE (+)")
            'S.AppendLine("AND C1.CUST_STATUS = 'A'")
            'S.AppendLine("GROUP BY")
            'S.AppendLine("C1.CUST_CODE,")
            'S.AppendLine("C1.CUST_NAME,")
            'S.AppendLine("C1.CUST_CITY,")
            'S.AppendLine("C1.CUST_STATE,")
            'S.AppendLine("C1.CUST_ZIP_CODE,")
            'S.AppendLine("C1.CUST_COUNTRY,")
            'S.AppendLine("C1.CUST_STMT_IND,")
            'S.AppendLine("C1.CUST_STMT_EMAIL,")
            'S.AppendLine("C1.SREP_CODE,")
            'S.AppendLine("C1.INIT_DATE,")
            'S.AppendLine("C1.CUST_XMIT_INV_VIA,")
            'S.AppendLine("C1.CUST_INV_EMAIL")
            'ASCMAIN1.sql = S.ToString
            'Create_TDA(.Tables.Add, "ARTLIST", "**", 0, False)
            'With .Tables("ARTLIST").Columns
            '    .Add("YRT", GetType(Double), "YR1 + YR2 + YR3 + YR4")
            'End With

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ARTCUSTD")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, False)
            Fill_Records("ARTCUSTD")

            ASCMAIN1.sql = makeARTLISTC_SQL(True)
            Create_TDA(.Tables.Add, "ARTLISTC", "**", 0, False)
            With .Tables("ARTLISTC").Columns
                .Add("YRT", GetType(Double), "YR1 + YR2 + YR3 + YR4")
                For Each COL_INFO As KeyValuePair(Of KeyValuePair(Of String, Int64), List(Of String)) In CONTACT_COLS
                    Dim COL_TYPE As String = COL_INFO.Key.Key
                    Dim MAX_COLS As Int64 = COL_INFO.Key.Value
                    For i As Int64 = 1 To MAX_COLS
                        For Each COL As String In COL_INFO.Value
                            Dim COL_NAME As String = $"{COL}_{COL_TYPE}_{Format(i, "#0")}"
                            .Add(COL_NAME)
                        Next
                    Next
                Next
            End With
        End With

        grdARTLISTC.DataSource = dst.Tables("ARTLISTC")

        Create_Summary(grdARTLISTC, "CUST_CODE", "Count")

        With grdARTLISTC.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
        End With

        With grdARTLISTC.DisplayLayout.Bands(0)
            .Columns("YR1").Header.Caption = YR1
            .Columns("YR1").Format = "###,##0.00"
            .Columns("YR2").Header.Caption = YR2
            .Columns("YR2").Format = "###,##0.00"
            .Columns("YR3").Header.Caption = YR3
            .Columns("YR3").Format = "###,##0.00"
            .Columns("YR4").Header.Caption = YR4
            .Columns("YR4").Format = "###,##0.00"
        End With

        TABLE_NAME = "ARFCUSTL"

        EntryMode = "E"
        Call Mode_Settings(True)
        'Fill_Records("ARTLISTC")
        'Dim map As Dictionary(Of String, KeyValuePair(Of String, String)) = makeColMaps()
        'ARCCUSTM.ExportToNewWorkbook(dst.Tables.Item("ARTLISTC"), map)
        Loading = False
    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Sub Check_Form_Options()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Cancel"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Cancel?"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("Are You Sure You Want To Cancel?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg += "Cancel Aboorted."
                End If
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
                Call Update_Record("Your Data Is Saved", False)
                Call Mode_Settings(False)
                UltraTabControl1.Tabs.Item("List Maint").Visible = True
                Loading = True
                Me.Close()
                Loading = False
            Case "Cancel"
                Call Mode_Settings(False)
                UltraTabControl1.Tabs.Item("List Maint").Visible = True
                Me.Close()
            Case "Save"
                Update_Record("Your Data Is Saved", False)
                dumpToExcel()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Save").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
        End With

        UltraExplorerBar1.Groups("Options").Visible = False

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        'dst.EnforceConstraints = False
        'dst.Tables("PMTVIST1").Rows.Clear()
        'dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Setup_Summary()

        'Setup_SOTCSTMX()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Sub Setup_Summary()
        txtFileLoc.Text = ""
        btnOpen.Enabled = False
        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        'Update_Record("", True)

        dst.Tables("ARTLISTC").Rows.Clear()

        dst.EnforceConstraints = False

        Fill_Records("ARTLISTC")
        AddContactsToList()
        Dim map As Dictionary(Of String, KeyValuePair(Of String, String)) = makeColMaps()
        txtFileLoc.Text = ARCCUSTM.ExportToNewWorkbook(dst.Tables.Item("ARTLISTC"), map)
        If txtFileLoc.Text.Length > 0 Then
            btnOpen.Enabled = True
        Else
            btnOpen.Enabled = False
        End If
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record(ByVal MsgToShow As String, Optional ByVal AutoSaving As Boolean = False)
        If MsgToShow.Length > 0 Then
            AutoSaving = False
        End If
        If AutoSaving Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Auto-Saving Your Data", "")
            Application.DoEvents()
        End If
        BeginTrans()
        CommitTrans(MsgToShow)
        If AutoSaving Then
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
            Application.DoEvents()
        End If
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
        Call Load_Popup_Menu(grdARTLISTC, "SSB", "Show Filter", "Show GroupBox", "Remove Contact", "Customer Master File", "Customer Inquiry")
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
            Case "Customer Master File"
                If Not IsNothing(grd.ActiveRow) Then
                    Dim CUST_CODE As String = grd.ActiveRow.Cells.Item("CUST_CODE").Text
                    If CUST_CODE.Length > 0 Then
                        Context_Launch("View", Column_Values("CUST_CODE", CUST_CODE), e.Tool.Key, "ARTCUST1")
                    End If
                End If
            Case "Customer Inquiry"
                If Not IsNothing(grd.ActiveRow) Then
                    Dim CUST_CODE As String = grd.ActiveRow.Cells.Item("CUST_CODE").Text
                    If CUST_CODE.Length > 0 Then
                        'Context_Launch("Select Customer", Column_Values("CUST_CODE", CUST_CODE), e.Tool.Key, "ARFCINQ1")
                        Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
                    End If
                End If
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

#Region "Form Controls"
    Private Sub btnRefreshList_Click(sender As Object, e As EventArgs)
        RefreshList()
    End Sub
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Setup_Summary()
        'UltraTabControl1.Tabs.Item("List Maint").Visible = True
    End Sub
    Private Sub chkListActiveOnly_CheckedChanged(sender As Object, e As EventArgs)
        If Not Loading Then
            ListActiveOnly()
        End If
    End Sub
    Private Sub cboCLIST_CODE_SelectedIndexChanged(sender As Object, e As EventArgs)
        If Not Loading Then
            RefreshList()
        End If
    End Sub
    Private Sub grdARTLIST_AfterRowUpdate(sender As Object, e As RowEventArgs)
        grdARTLISTC.Update()
        grdARTLISTC.Refresh()
        Update_Record("", True)
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub AddContactsToList()
        For Each rowARTLISTC As DataRow In dst.Tables("ARTLISTC").Select()
            Dim CUST_CODE As String = rowARTLISTC.Item("CUST_CODE").ToString & String.Empty
            Dim MAX_REC As Int64 = 0

            For Each COL_INFO As KeyValuePair(Of KeyValuePair(Of String, Int64), List(Of String)) In CONTACT_COLS
                Dim COL_TYPE As String = COL_INFO.Key.Key
                Dim MAX_COLS As Int64 = COL_INFO.Key.Value
                Dim CUR_COL As Int64 = 1
                Dim fltrP As String = $"CUST_CODE = '{CUST_CODE}' AND CONTACT_TYPE = '{COL_TYPE}' AND ISNULL(CONTACT_PRIMARY,'0') = '1'"
                For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(fltrP, "CONTACT_NO")
                    If CUR_COL < MAX_COLS Then
                        For Each COL As String In COL_INFO.Value
                            Dim COL_NAME As String = $"{COL}_{COL_TYPE}_{Format(CUR_COL, "#0")}"
                            If COL = "CONTACT_PRIMARY" Then
                                If (rowARTCUSTD.Item(COL).ToString & String.Empty) = "1" Then
                                    rowARTLISTC.Item(COL_NAME) = "Yes"
                                Else
                                    rowARTLISTC.Item(COL_NAME) = "No"
                                End If
                            Else
                                rowARTLISTC.Item(COL_NAME) = rowARTCUSTD.Item(COL)
                            End If
                        Next
                        CUR_COL += 1
                    End If
                Next

                Dim fltrN As String = $"CUST_CODE = '{CUST_CODE}' AND CONTACT_TYPE = '{COL_TYPE}' AND ISNULL(CONTACT_PRIMARY,'0') <> '1'"
                For Each rowARTCUSTD As DataRow In dst.Tables("ARTCUSTD").Select(fltrN, "CONTACT_NO")
                    If CUR_COL < MAX_COLS Then
                        For Each COL As String In COL_INFO.Value
                            Dim COL_NAME As String = $"{COL}_{COL_TYPE}_{Format(CUR_COL, "#0")}"
                            If COL = "CONTACT_PRIMARY" Then
                                If (rowARTCUSTD.Item(COL).ToString & String.Empty) = "1" Then
                                    rowARTLISTC.Item(COL_NAME) = "Yes"
                                Else
                                    rowARTLISTC.Item(COL_NAME) = "No"
                                End If
                            Else
                                rowARTLISTC.Item(COL_NAME) = rowARTCUSTD.Item(COL)
                            End If
                        Next
                        CUR_COL += 1
                    End If
                Next
            Next
        Next
    End Sub
    Private Sub dumpToExcel()
        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing

        Dim xls_file_name As String = "Attributes.xlsx"
        Dim fDialog As New FolderBrowserDialog
        fDialog.Description = "Please Select The Folder To Save File"
        fDialog.ShowDialog()
        xls_file_name = $"{fDialog.SelectedPath}\{xls_file_name}"
        If System.IO.File.Exists(xls_file_name) Then
            Dim iResult As MsgBoxResult = MsgBox("Delete it?", vbYesNo, $"{xls_file_name} Exists!")
            If iResult <> vbYes Then
                Exit Sub
            Else
                System.IO.File.Delete(xls_file_name)
            End If
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Creating File...", "")
        Application.DoEvents()

        oWB = SpreadsheetGear.Factory.GetWorkbook()
        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i
        oSheet = oWB.Worksheets.Add()
        oSheet.Name = "Search Results"

        If oWB.Worksheets.Count = 2 Then
            oWB.Worksheets(0).Delete()
        End If

        ASCMAIN1.Progress("-", oSheet.Name)

        Load_DataTable_into_SGXLS(1, 1, dst.Tables.Item("ARTLISTC"), oSheet, Nothing, Nothing, "", "")

        'Set heading to Grid Heading
        Dim colCnt As Int64 = dst.Tables.Item("ARTLISTC").Columns.Count
        Dim rowCnt As Int64 = dst.Tables.Item("ARTLISTC").Rows.Count
        For i As Int64 = 0 To colCnt
            oSheet.Range(0, i).Select()
            Dim colTitle As String = oSheet.Range(0, i).Text
            If grdARTLISTC.DisplayLayout.Bands(0).Columns.IndexOf(colTitle) <> -1 Then
                Dim grdTitle As String = grdARTLISTC.DisplayLayout.Bands(0).Columns(colTitle).Header.Caption
                oSheet.Range(0, i).Value = grdTitle
            End If
        Next

        'Set Columns to Auto
        oSheet.Range($"A1:ZZ1").EntireColumn.AutoFit()

        'Set Rows to 15
        oSheet.Range($"A1:A{rowCnt + 10}").RowHeight = 15

        oSheet.Range("A1:A1").Select()
        oSheet.WindowInfo.FreezePanes = True
        oSheet.Range("A1:A1").Select()

        oWB.Worksheets(0).Select()

        oWB.SaveAs(xls_file_name, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        'Show_Document(xls_file_name)
        oWB = Nothing

        Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
        MsgBox("File Created", vbOKOnly, "Done")
    End Sub
    Private Sub Fill_Extra_Fields()
        Dim RecTotal As Int64 = dst.Tables("SOTCSTMX").Rows.Count
        Dim OnRow As Int64 = 0
        Dim PCT As String = ""
    End Sub
    Private Function makeARTLISTC_SQL(ByVal genEmpty As Boolean) As String
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("C1.CUST_CODE,")
        S.AppendLine("C1.CUST_NAME,")
        S.AppendLine("C1.CUST_CITY,")
        S.AppendLine("C1.CUST_STATE,")
        S.AppendLine("C1.CUST_ZIP_CODE,")
        S.AppendLine("C1.CUST_COUNTRY,")
        S.AppendLine("C1.CUST_STMT_IND,")
        S.AppendLine("C1.CUST_STMT_EMAIL,")
        S.AppendLine("C1.SREP_CODE,")
        S.AppendLine("C1.INIT_DATE,")
        S.AppendLine("C1.CUST_XMIT_INV_VIA,")
        S.AppendLine("C1.CUST_INV_EMAIL,")
        S.AppendLine("SUM(SALES.YR1) AS YR1,")
        S.AppendLine("SUM(SALES.YR2) AS YR2,")
        S.AppendLine("SUM(SALES.YR3) AS YR3,")
        S.AppendLine("SUM(SALES.YR4) AS YR4")
        S.AppendLine($"FROM ARTCUST1 C1, {TEMP_SALES} SALES")
        S.AppendLine("WHERE C1.CUST_CODE = SALES.CUST_CODE (+)")
        S.AppendLine("AND C1.CUST_STATUS = 'A'")
        'If genEmpty Then
        '    S.AppendLine("AND ROWNUM < 1")
        'End If
        S.AppendLine("GROUP BY")
        S.AppendLine("C1.CUST_CODE,")
        S.AppendLine("C1.CUST_NAME,")
        S.AppendLine("C1.CUST_CITY,")
        S.AppendLine("C1.CUST_STATE,")
        S.AppendLine("C1.CUST_ZIP_CODE,")
        S.AppendLine("C1.CUST_COUNTRY,")
        S.AppendLine("C1.CUST_STMT_IND,")
        S.AppendLine("C1.CUST_STMT_EMAIL,")
        S.AppendLine("C1.SREP_CODE,")
        S.AppendLine("C1.INIT_DATE,")
        S.AppendLine("C1.CUST_XMIT_INV_VIA,")
        S.AppendLine("C1.CUST_INV_EMAIL")
        Return S.ToString
    End Function
    Private Function GetContactList() As String
        Dim RetVal As String = ""
        If chkContactsX.Checked Then
            RetVal = " CONTACT_TYPE = 'X'"
        End If
        If chkContactsB.Checked Then
            If RetVal.Length = 0 Then
                RetVal = " CONTACT_TYPE = 'B'"
            Else
                RetVal = RetVal & " OR CONTACT_TYPE = 'B'"
            End If
        End If
        If chkContactsP.Checked Then
            If RetVal.Length = 0 Then
                RetVal = " CONTACT_TYPE = 'P'"
            Else
                RetVal = RetVal & " OR CONTACT_TYPE = 'P'"
            End If
        End If
        If chkContactsW.Checked Then
            If RetVal.Length = 0 Then
                RetVal = " CONTACT_TYPE = 'W'"
            Else
                RetVal = RetVal & " OR CONTACT_TYPE = 'W'"
            End If
        End If
        If chkContactsM.Checked Then
            If RetVal.Length = 0 Then
                RetVal = " CONTACT_TYPE = 'M'"
            Else
                RetVal = RetVal & " OR CONTACT_TYPE = 'M'"
            End If
        End If
        'Nothing was selected Make Nothing Get Seleted
        If RetVal = "" Then
            RetVal = RetVal & " AND CONTACT_TYPE = 'Z'"
        Else
            RetVal = " AND (" & RetVal & ")"
        End If

        Return RetVal
    End Function
    Private Sub ListActiveOnly()
        Dim filter As String = ""
        Dim dvw As DataView = DirectCast(grdARTLISTC.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Format(filter)
    End Sub
    Private Function makeColMaps() As Dictionary(Of String, KeyValuePair(Of String, String))
        Dim map As New Dictionary(Of String, KeyValuePair(Of String, String)) From
            {
                {"CUST_CODE", New KeyValuePair(Of String, String)("Cust Code", "@")},
                {"CUST_NAME", New KeyValuePair(Of String, String)("Name", "@")},
                {"CUST_CITY", New KeyValuePair(Of String, String)("City", "@")},
                {"CUST_STATE", New KeyValuePair(Of String, String)("State", "@")},
                {"CUST_ZIP_CODE", New KeyValuePair(Of String, String)("Zip Code", "@")},
                {"CUST_COUNTRY", New KeyValuePair(Of String, String)("Country", "@")},
                {"CUST_STMT_IND", New KeyValuePair(Of String, String)("Stmt", "@")},
                {"CUST_STMT_EMAIL", New KeyValuePair(Of String, String)("Stmt E-mail", "@")},
                {"SREP_CODE", New KeyValuePair(Of String, String)("Sales Rep", "@")},
                {"INIT_DATE", New KeyValuePair(Of String, String)("Began", "@")},
                {"CUST_XMIT_INV_VIA", New KeyValuePair(Of String, String)("Inv Via", "@")},
                {"CUST_INV_EMAIL", New KeyValuePair(Of String, String)("Inv E-mail", "@")},
                {"YR1", New KeyValuePair(Of String, String)("Year 1", "#,##0")},
                {"YR2", New KeyValuePair(Of String, String)("Year 2", "#,##0")},
                {"YR3", New KeyValuePair(Of String, String)("Year 3", "#,##0")},
                {"YR4", New KeyValuePair(Of String, String)("Year 4", "#,##0")},
                {"YRT", New KeyValuePair(Of String, String)("Total", "#,##0")}
            }
        For Each COL_INFO As KeyValuePair(Of KeyValuePair(Of String, Int64), List(Of String)) In CONTACT_COLS
            Dim COL_TYPE As String = COL_INFO.Key.Key
            Dim MAX_COLS As Int64 = COL_INFO.Key.Value
            If chkPrimaryOnly.Checked Then
                MAX_COLS = 1
            End If
            For i As Int64 = 1 To MAX_COLS
                For Each COL As String In COL_INFO.Value
                    Dim showCol As Boolean = False

                    Dim COL_TYPENAME As String = ""
                    Select Case COL_TYPE
                        Case "B"
                            COL_TYPENAME = "Buyer"
                            If chkContactsB.Checked Then showCol = True
                        Case "P"
                            COL_TYPENAME = "AP"
                            If chkContactsP.Checked Then showCol = True
                        Case "W"
                            COL_TYPENAME = "Whse"
                            If chkContactsW.Checked Then showCol = True
                        Case "M"
                            COL_TYPENAME = "Misc"
                            If chkContactsM.Checked Then showCol = True
                        Case "X"
                            COL_TYPENAME = "Main"
                            If chkContactsX.Checked Then showCol = True
                    End Select
                    Select Case COL
                        Case "CONTACT_TYPE"
                            COL_TYPENAME = ""
                        Case "CONTACT_PRIMARY"
                            COL_TYPENAME = $"{COL_TYPENAME}{i} {"Primary"}"
                        Case "CONTACT_NAME"
                            COL_TYPENAME = $"{COL_TYPENAME}{i} {"Name"}"
                        Case "CONTACT_TITLE"
                            COL_TYPENAME = $"{COL_TYPENAME}{i} {"Title"}"
                        Case "CONTACT_EMAIL"
                            COL_TYPENAME = $"{COL_TYPENAME}{i} {"Email"}"
                        Case "CONTACT_PHONE"
                            COL_TYPENAME = $"{COL_TYPENAME}{i} {"Phone"}"
                        Case "CONTACT_CELL"
                            COL_TYPENAME = $"{COL_TYPENAME}{i} {"Cell"}"
                    End Select
                    If showCol Then
                        Dim DOLKVP As New KeyValuePair(Of String, String)(COL_TYPENAME, "@")
                        Dim COL_NAME As String = $"{COL}_{COL_TYPE}_{Format(i, "#0")}"
                        If COL_TYPENAME.Length > 0 Then
                            map.Add(COL_NAME, DOLKVP)
                        End If
                    End If
                Next
            Next
        Next

        Return map
    End Function
    Private Sub makegrdARTLISTC()
        If grdARTLISTC.DisplayLayout.Bands(0).Columns.Count = 0 Then
            For Each dc As DataColumn In dst.Tables.Item("ARTLISTC").Columns
                grdARTLISTC.DisplayLayout.Bands(0).Columns.Add(dc.ColumnName)
            Next
        End If
        With grdARTLISTC.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup
            Dim COLS As String() = {"CUST_CODE", "CUST_NAME", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", "CUST_STMT_IND", "CUST_STMT_EMAIL", "SREP_CODE", "INIT_DATE", "CUST_XMIT_INV_VIA", "CUST_INV_EMAIL", "YR1", "YR2", "YR3", "YR4"}
            G = .Groups.Add("Customer", "Customer Information")
            G.Header.Appearance.TextHAlign = HAlign.Center
            G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
            For Each COL As String In COLS
                .Columns(COL).Group = G
            Next
            'G.Header.Fixed = True

            For Each COL_INFO As KeyValuePair(Of KeyValuePair(Of String, Int64), List(Of String)) In CONTACT_COLS
                Dim COL_TYPE As String = COL_INFO.Key.Key
                Dim MAX_COLS As Int64 = COL_INFO.Key.Value
                Dim CUR_COL As Int64 = 0
                Dim COL_TYPENAME As String = "Unknown"
                Select Case COL_TYPE
                    Case "B"
                        COL_TYPENAME = "Buyer"
                    Case "P"
                        COL_TYPENAME = "Accounts Payable"
                    Case "W"
                        COL_TYPENAME = "Warehouse"
                    Case "M"
                        COL_TYPENAME = "Misc"
                    Case "X"
                        COL_TYPENAME = "Main"
                End Select
                For i As Int64 = 1 To MAX_COLS
                    CUR_COL += 1
                    G = .Groups.Add($"{COL_TYPENAME}{Format(CUR_COL, "#0")}", $"{COL_TYPENAME}{Format(CUR_COL, "#0")}")
                    G.Header.Appearance.TextHAlign = HAlign.Center
                    G.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
                    G.Header.Appearance.BackColor2 = Drawing.Color.Transparent
                    For Each COL As String In COL_INFO.Value
                        Dim COL_NAME As String = $"{COL}_{COL_TYPE}_{Format(i, "#0")}"
                        .Columns(COL_NAME).Group = G
                    Next
                Next
            Next
        End With

    End Sub
    Private Sub MAKE_CONTACT_COLS()
        Dim COLS As New List(Of String)
        COLS.Add("CONTACT_TYPE")
        COLS.Add("CONTACT_PRIMARY")
        COLS.Add("CONTACT_NAME")
        COLS.Add("CONTACT_TITLE")
        COLS.Add("CONTACT_EMAIL")
        COLS.Add("CONTACT_PHONE")
        COLS.Add("CONTACT_CELL")

        Dim KVPB As New KeyValuePair(Of String, Int64)("B", 10)
        CONTACT_COLS.Add(KVPB, COLS)

        Dim KVPP As New KeyValuePair(Of String, Int64)("P", 5)
        CONTACT_COLS.Add(KVPP, COLS)

        Dim KVPW As New KeyValuePair(Of String, Int64)("W", 5)
        CONTACT_COLS.Add(KVPW, COLS)

        Dim KVPM As New KeyValuePair(Of String, Int64)("M", 10)
        CONTACT_COLS.Add(KVPM, COLS)

        Dim KVPX As New KeyValuePair(Of String, Int64)("X", 10)
        CONTACT_COLS.Add(KVPX, COLS)

    End Sub
    Private Sub RefreshList()
        ASCMAIN1.Progress("Now Loading List")
        Me.Cursor = Cursors.WaitCursor
        Update_Record("", True)

        dst.EnforceConstraints = False

        ListActiveOnly()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Private Sub RefreshSalesTempTable()
        S.Length = 0
        S.AppendLine("Select CUST_CODE,")
        S.AppendLine("SUM(YR1) As YR1,")
        S.AppendLine("SUM(YR2) As YR2,")
        S.AppendLine("SUM(YR3) As YR3,")
        S.AppendLine("SUM(YR4) As YR4")
        S.AppendLine("FROM")
        S.AppendLine("(")
        S.AppendLine("  Select")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  SUM(INV_SALES) As YR1,")
        S.AppendLine("  0 As YR2,")
        S.AppendLine("  0 As YR3,")
        S.AppendLine("  0 As YR4")
        S.AppendLine("  FROM SOTINVH1")
        S.AppendLine(String.Format("  WHERE EXTRACT(year FROM inv_date) = '{0}'", YR1))
        S.AppendLine("  GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("")
        S.AppendLine("SELECT")
        S.AppendLine("O1.CUST_CODE,")
        S.AppendLine("SUM((NVL(O2.ORDR_QTY_PICK,0) + NVL(O2.ORDR_QTY_OPEN,0)) * O2.ORDR_UNIT_PRICE) AS YR1,")
        S.AppendLine("0 AS YR2,")
        S.AppendLine("0 AS YR3,")
        S.AppendLine("0 AS YR4")
        S.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
        S.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
        S.AppendLine(String.Format("  AND EXTRACT(year FROM ORDR_DATE_RECD) = '{0}'", YR1))
        S.AppendLine("GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("  SELECT")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  0 AS YR1,")
        S.AppendLine("  SUM(INV_SALES) AS YR2,")
        S.AppendLine("  0 AS YR3,")
        S.AppendLine("  0 AS YR4")
        S.AppendLine("  FROM SOTINVH1")
        S.AppendLine(String.Format("  WHERE EXTRACT(year FROM inv_date) = '{0}'", YR2))
        S.AppendLine("  GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("SELECT")
        S.AppendLine("O1.CUST_CODE,")
        S.AppendLine("0 AS YR1,")
        S.AppendLine("SUM((NVL(O2.ORDR_QTY_PICK,0) + NVL(O2.ORDR_QTY_OPEN,0)) * O2.ORDR_UNIT_PRICE) AS YR2,")
        S.AppendLine("0 AS YR3,")
        S.AppendLine("0 AS YR4")
        S.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
        S.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
        S.AppendLine(String.Format("  AND EXTRACT(year FROM ORDR_DATE_RECD) = '{0}'", YR2))
        S.AppendLine("GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("  SELECT")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  0 AS YR1,")
        S.AppendLine("  0 AS YR2,")
        S.AppendLine("  SUM(INV_SALES) AS YR3,")
        S.AppendLine("  0 AS YR4")
        S.AppendLine("  FROM SOTINVH1")
        S.AppendLine(String.Format("  WHERE EXTRACT(year FROM inv_date) = '{0}'", YR3))
        S.AppendLine("  GROUP BY CUST_CODE")
        S.AppendLine("  UNION")
        S.AppendLine("  SELECT")
        S.AppendLine("  CUST_CODE,")
        S.AppendLine("  0 AS YR1,")
        S.AppendLine("  0 AS YR2,")
        S.AppendLine("  0 AS YR3,")
        S.AppendLine("  SUM(INV_SALES) AS YR4")
        S.AppendLine("  FROM SOTINVH1")
        S.AppendLine(String.Format("  WHERE EXTRACT(year FROM inv_date) = '{0}'", YR4))
        S.AppendLine("  GROUP BY CUST_CODE")
        S.AppendLine(") RSLT")
        S.AppendLine("GROUP BY CUST_CODE")
        ASCMAIN1.sql = S.ToString
        TEMP_SALES = ASCMAIN1.Temp_Table

        ASCDATA1.ExecuteSQL("Create Index I_" & TEMP_SALES & "_IND on " & TEMP_SALES & " (CUST_CODE)")
    End Sub
    Private Sub btnUpsert_Click(sender As Object, e As EventArgs) Handles btnUpsert.Click
        MsgBox("This Feature Coming Soon.", vbOKOnly, "Working As Fast As I Can Captain!")
    End Sub
    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        If txtFileLoc.Text.Length > 0 Then
            Show_Document(txtFileLoc.Text)
        End If
    End Sub
#End Region
End Class