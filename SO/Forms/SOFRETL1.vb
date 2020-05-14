Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.IO
Imports System.Threading


<ComVisible(True)> _
Public Class SOFRETL1
    Dim sqls As String = ""
    Dim loggedIn As Boolean = False
    Dim onReportsPage As Boolean = False
    Dim reportsLoaded As Boolean = False
    Dim gotReportFiles As Boolean = False
    Dim downloadedReports As Integer = 0
    Dim skippedReports As Integer = 0
    Dim YYYYWW As String = ASCMAIN1.CYW
    Dim getArchiveData As Boolean = False
    Dim injectedScript As Boolean = False
    Dim resultWorkbook As SpreadsheetGear.IWorkbook = Nothing
    Dim noFileNameError As Boolean = False





#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            With .Tables.Add("SOTRETL1")
                .Columns.Add("YYYYWW")
                .Columns.Add("JOB_ID")
                .Columns.Add("STATUS")
                .Columns.Add("REQUEST_TITLE")
                .Columns.Add("RUN_TIME", GetType(System.DateTime))
                .Columns.Add("FILE_SIZE")
                .Columns.Add("DOWNLOAD_COMPLETE")
                .Columns.Add("PROCESS_REPORT")
                .Columns.Add("REPORT_INDEX")
                .Columns.Add("FILE_NAME")
                .Columns.Add("EXT")
                .Columns.Add("NOTES")
                .PrimaryKey = New DataColumn() {.Columns("JOB_ID")}
            End With
            Create_TDA(.Tables.Add, "TATWMRL1", "*")

            With .Tables.Add("SOTRLWK1")
                .Columns.Add("YYYYWW")
                .Columns.Add("LEGEND")
                .PrimaryKey = New DataColumn() {.Columns("YYYYWW")}
            End With

        End With

        grdSOTRETL1.DataSource = dst.Tables("SOTRETL1")

        Sort_grdColumns(grdSOTRETL1, "JOB_ID".ToLower, False)

        grdSOTRETL1.DisplayLayout.Bands(0).Columns("YYYYWW").Header.Caption = "Week"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("JOB_ID").Header.Caption = "Job Id"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("STATUS").Header.Caption = "Status"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("REQUEST_TITLE").Header.Caption = "Request Title"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("RUN_TIME").Header.Caption = "Run Time"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("FILE_SIZE").Header.Caption = "File Size"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("DOWNLOAD_COMPLETE").Header.Caption = "Downloaded"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("PROCESS_REPORT").Header.Caption = "View"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("EXT").Header.Caption = "Extension"
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("NOTES").Header.Caption = "Notes"
        grdSOTRETL1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ExtendLastColumn
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("EXT").Hidden = True
        Me.wb.ObjectForScripting = Me

        ASCMAIN1.sql = "Select * from TATUSER1 where USER_ID = :PARM1"
        Dim rowTATUSER1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", ASCMAIN1.USER_ID)
        If rowTATUSER1 IsNot Nothing Then
            txtUserID.Text = rowTATUSER1.Item("RETAIL_LINK_USER_ID") & ""
            txtPassword.Text = rowTATUSER1.Item("RETAIL_LINK_PASSWORD") & ""
            chkEmailReport.Checked = (rowTATUSER1.Item("RETAIL_LINK_EMAIL") & "" = "1")
        End If

        cbeReports.DataSource = ASCDATA1.GetDataTable("Select T_CODE REPORT_KEY, T_DESC REPORT_DESC from ASTCODE1 where COLUMN_NAME = 'RETAIL_LINK_REPORTS' and TABLE_NAME = 'TATWMRL1'")

        ASCMAIN1.sql = "Select YYYYWW, LEGEND from GLTPARM3 WHERE WEEK_END_DATE > '" & Format(DateTime.Now.AddYears(-1), "dd-MMM-yyyy") & "'" _
            & " and YYYYWW <= '" & ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -1) & "'"
        Dim tbl As DataTable = ASCDATA1.GetDataTable()

        For Each tRow As DataRow In tbl.Select("")
            Dim weekNo As String = tRow.Item("YYYYWW")
            Dim networkDir As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & weekNo & "\" & cbeReports.Value _
                                              & "\"

            If My.Computer.FileSystem.DirectoryExists(networkDir) Or weekNo = ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -1) Then
                Dim rowSOTRLWK1 As DataRow = dst.Tables("SOTRLWK1").NewRow
                rowSOTRLWK1.Item("YYYYWW") = tRow.Item("YYYYWW")
                rowSOTRLWK1.Item("LEGEND") = tRow.Item("LEGEND")
                dst.Tables("SOTRLWK1").Rows.Add(rowSOTRLWK1)
            End If
        Next

        weekCmb.DataSource = dst.Tables("SOTRLWK1")
        weekCmb.ValueMember = "YYYYWW"
        weekCmb.DisplayMember = "LEGEND"
        weekCmb.Value = ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -1)
        weekCmb.Visible = False
        EntryMode = "E"

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
            Case "Save My Credentials"
                If txtUserID.Text = "" Then
                    EMsg = "Please enter a User ID"
                End If
                If txtPassword.Text = "" Then
                    EMsg = "Please enter a Password"
                End If
            Case "Generate Report"
                If txtUserID.Text = "" Then
                    EMsg = "Please enter a User ID"
                End If
                If txtPassword.Text = "" Then
                    EMsg = "Please enter a Password"
                End If
                If cbeReports.Value & "" = "" Then
                    EMsg = "Please Select a Report Type"
                End If
                If weekCmb.Value & "" = "" Then
                    EMsg = "Please Select a Week"
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
                Call Mode_Settings(False)
                Me.Close()
            Case "Save My Credentials"
                Dim RETAIL_LINK_USER_ID As String = txtUserID.Text
                Dim RETAIL_LINK_PASSWORD As String = txtPassword.Text

                ASCMAIN1.sql = "Select * from TATUSER1 where USER_ID = :PARM1"
                Dim rowTATUSER1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", ASCMAIN1.USER_ID)
                BeginTrans()
                If rowTATUSER1 IsNot Nothing Then
                    ASCMAIN1.sql = "Update TATUSER1 set RETAIL_LINK_USER_ID = '" & RETAIL_LINK_USER_ID & "', RETAIL_LINK_PASSWORD = '" & RETAIL_LINK_PASSWORD & "'" _
                        & " where USER_ID = '" & ASCMAIN1.USER_ID & "'"
                    ASCDATA1.ExecuteSQL()
                    CommitTrans("Credentials Updated")
                Else
                    ASCMAIN1.sql = "Insert into TATUSER1 (USER_ID, RETAIL_LINK_USER_ID, RETAIL_LINK_PASSWORD)" _
                        & " values ('" & ASCMAIN1.USER_ID & "','" & RETAIL_LINK_USER_ID & "','" & RETAIL_LINK_PASSWORD & "')"
                    ASCDATA1.ExecuteSQL()
                    CommitTrans("Credentials Saved")
                End If
            Case "Generate Report"
                ASCMAIN1.Progress("Now Fetching Reports")
                Me.Cursor = Cursors.WaitCursor
                Clear_Record()
                YYYYWW = Absx1.cmbFor("RYW0").Value

                If optMethod.Value & "" = "W" Then

                    If YYYYWW <> ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -1) Then
                        getArchiveData = True
                    End If

                    SplitContainer1.Panel2Collapsed = getArchiveData

                    Try
                        'If getArchiveData Then
                        '    Load_Archived_Spreadsheets()
                        'Else
                        '    DirectCast(wb, Control).Enabled = True
                        '    Dim WebAddress As String = "https://retaillink.wal-mart.com/"
                        '    wb.Navigate(New Uri(WebAddress))
                        'End If

                        DirectCast(wb, Control).Enabled = True
                        Dim WebAddress As String = "https://retaillink.wal-mart.com/"
                        wb.Navigate(New Uri(WebAddress))
                    Catch ex As Exception
                        Dim fubar As String = ""
                    End Try

                Else
                    SplitContainer1.Panel2Collapsed = True
                    Import_AS2_Files()

                End If

                ASCMAIN1.Progress("")
                Me.Cursor = Cursors.Default

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Save My Credentials").Settings.Enabled = not_iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("SOTRETL1").Rows.Clear()
        dst.Tables("TATWMRL1").Rows.Clear()
        loggedIn = False
        onReportsPage = False
        reportsLoaded = False
        gotReportFiles = False
        downloadedReports = 0
        skippedReports = 0
        dst.EnforceConstraints = True
        getArchiveData = False

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        dst.Tables("SOTRETL1").Rows.Clear()
        dst.Tables("TATWMRL1").Rows.Clear()
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
        Call Load_Popup_Menu(grdSOTRETL1, "SSB", "Show Filter", "Show GroupBox")
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

#Region "Custom Methods"

    Private Sub ReLoadData()

    End Sub

#End Region

#Region "Form Controls"



#End Region

    Sub Load_Archived_Scorecard()

        'Dim scorecardFilename As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\ScoreCards ~ Week " & YYYYWW & ".xlsx"
        'If System.IO.File.Exists(scorecardFilename) Then
        '    Dim p As Process = Process.Start(scorecardFilename)
        'Else
        '    MsgBox("Sorry, can't find file.", vbCritical, "Cannot Proceed")
        'End If

    End Sub

    Sub Load_Archived_Spreadsheets()

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        dst.Tables("SOTRETL1").Rows.Clear()
        dst.Tables("TATWMRL1").Rows.Clear()
        Save_Header_Fields(UltraGroupBox1)
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("NOTES").Hidden = True
        ASCMAIN1.sql = "Select * from TATWMRL1 where YYYYWW = '" & YYYYWW & "'"
        Dim tbl As DataTable = ASCDATA1.GetDataTable()

        For Each tRow As DataRow In tbl.Select("")
            Dim JOB_ID As String = tRow.Item("JOB_ID")
            Dim dataFile As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value _
                                              & "\" & JOB_ID & ".xls"
            Dim dataFileSize As String = ""
            If System.IO.File.Exists(dataFile) Then
                Dim dataWorkbookSet As SpreadsheetGear.IWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet()
                Dim dataWorkbook As SpreadsheetGear.IWorkbook = dataWorkbookSet.Workbooks.Open(dataFile)
                Dim infoReader As System.IO.FileInfo
                infoReader = My.Computer.FileSystem.GetFileInfo(dataFile)
                dataFileSize = Math.Round((infoReader.Length * 0.001), 2).ToString & "Kb"
            End If
            Dim rowSOTRETL1 As DataRow = dst.Tables("SOTRETL1").Rows.Find(New Object() {JOB_ID})
            If rowSOTRETL1 Is Nothing Then
                Load_Existing_Record(tRow, dataFileSize)
            End If

        Next

        Load_Archived_Scorecard()

    End Sub

    Private Sub wb_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wb.DocumentCompleted

        If Not loggedIn Then
            Login_Submit()
        Else
            If Not onReportsPage Then
                Navigate_To_Reports()
            Else
                'Thread.Sleep(3000)
                If Not reportsLoaded Then
                    Load_Reports()
                Else
                    'Thread.Sleep(3000)
                    Download_Reports()
                    ASCMAIN1.Progress("")
                    'DirectCast(wb, Control).Enabled = False
                End If


            End If

        End If


    End Sub

    Private Sub Login_Submit()
        Dim theElementCollection As HtmlElementCollection = wb.Document.GetElementsByTagName("input")
        For Each curElement As HtmlElement In theElementCollection
            ASCMAIN1.Progress("Attempting Logon")
            Dim elementName As String = curElement.GetAttribute("name").ToString
            If elementName = "txtUser" Then
                curElement.SetAttribute("value", txtUserID.Text)
            End If
            If elementName = "txtPass" Then
                curElement.SetAttribute("value", txtPassword.Text)
            End If
        Next
        If wb.Document.GetElementById("Login") IsNot Nothing Then
            wb.Document.GetElementById("Login").Focus()
            wb.Document.GetElementById("Login").InvokeMember("click")
        Else
            'ALREADY LOGGED IN?
        End If

        loggedIn = True
    End Sub
    Private Sub Navigate_To_Reports()
        ASCMAIN1.Progress("Locating Reports")
        'no reports to test with
        'Dim WebAddress As String = "https://retaillink.wal-mart.com/decision_support/?ukey=W5741"
        Dim WebAddress As String = "https://retaillink.wal-mart.com/decision_support/Homepage_status.aspx?ApplicationId=300"
        wb.Navigate(New Uri(WebAddress))
        onReportsPage = True
    End Sub

    Private Sub Load_Reports()
        ASCMAIN1.Progress("Loading Reports Into Queue")
        dst.Tables("SOTRETL1").Clear()

        Dim objWebClient As New System.Net.WebClient
        Dim fileNo As Integer = 0
        Dim theElementCollection As HtmlElementCollection = wb.Document.GetElementsByTagName("iframe")
        Dim gotReportsTable As Boolean = False

        For Each curElement As HtmlElement In theElementCollection
            Dim url As String = curElement.GetAttribute("src").ToString
            If curElement.Id = "JobTable" Then
                Dim htmlDoc As HtmlDocument = wb.Document
                For Each htmlWin As HtmlWindow In htmlDoc.Window.Parent.Frames

                    Try
                        Dim FrameDoc As HtmlDocument = htmlWin.Document
                        If FrameDoc.GetElementById("myTable") IsNot Nothing Then
                            Dim tableHTML As String = Trim(FrameDoc.GetElementById("myTable").OuterHtml)
                            If InStr(tableHTML, "<TD") > 0 And Not gotReportsTable Then
                                Dim reportIndex As Integer = 0
                                Dim myReports As DataTable = HTMLTable2DataTable(tableHTML)
                                For Each tRow As DataRow In myReports.Select()
                                    If Include_Report(tRow) Then
                                        If dst.Tables("SOTRETL1").Select("REQUEST_TITLE = '" & tRow.Item(3) & "'").Length > 0 Then
                                            Dim rowSOTRETL1x As DataRow = dst.Tables("SOTRETL1").Select("REQUEST_TITLE = '" & tRow.Item(3) & "'")(0)
                                            If rowSOTRETL1x.Item("RUN_TIME") < CDate(Replace(tRow.Item(4) & "", "&nbsp;", "")) Then
                                                rowSOTRETL1x.Item("YYYYWW") = YYYYWW
                                                rowSOTRETL1x.Item("JOB_ID") = tRow.Item(1) & ""
                                                rowSOTRETL1x.Item("STATUS") = tRow.Item(2) & ""
                                                rowSOTRETL1x.Item("REQUEST_TITLE") = tRow.Item(3) & ""
                                                rowSOTRETL1x.Item("RUN_TIME") = Replace(tRow.Item(4) & "", "&nbsp;", "")
                                                rowSOTRETL1x.Item("FILE_SIZE") = tRow.Item(5) & ""
                                                rowSOTRETL1x.Item("DOWNLOAD_COMPLETE") = "0"
                                                rowSOTRETL1x.Item("REPORT_INDEX") = reportIndex.ToString
                                            End If
                                        Else
                                            Dim rowSOTRETL1 As DataRow = dst.Tables("SOTRETL1").NewRow
                                            rowSOTRETL1.Item("YYYYWW") = YYYYWW
                                            rowSOTRETL1.Item("JOB_ID") = tRow.Item(1) & ""
                                            rowSOTRETL1.Item("STATUS") = tRow.Item(2) & ""
                                            rowSOTRETL1.Item("REQUEST_TITLE") = tRow.Item(3) & ""
                                            rowSOTRETL1.Item("RUN_TIME") = Replace(tRow.Item(4) & "", "&nbsp;", "")
                                            rowSOTRETL1.Item("FILE_SIZE") = tRow.Item(5) & ""
                                            rowSOTRETL1.Item("DOWNLOAD_COMPLETE") = "0"
                                            rowSOTRETL1.Item("REPORT_INDEX") = reportIndex.ToString
                                            rowSOTRETL1.Item("EXT") = "xls"
                                            If Trim(rowSOTRETL1.Item("REQUEST_TITLE")) = "KB Dot Com sales Inc V2 (LW, TW and Last 52 Wks)" _
                                                Or rowSOTRETL1.Item("REQUEST_TITLE") = "KB DotCom Inv LW." Then
                                                rowSOTRETL1.Item("EXT") = "xlsx"
                                            End If
                                            dst.Tables("SOTRETL1").Rows.Add(rowSOTRETL1)
                                            downloadedReports += 1
                                        End If
                                    Else
                                        skippedReports += 1
                                    End If
                                    reportIndex += 1
                                Next
                                reportsLoaded = True
                            End If

                        End If

                    Catch ex As Exception

                    End Try
                Next
            End If
        Next

    End Sub
    Function Include_Report(reportRow As DataRow) As Boolean

        Dim includeReport As Boolean = True
        Dim status As String = reportRow.Item(2) & ""
        Dim requestTitle As String = reportRow.Item(3) & ""

        Select Case status
            Case "Waiting", "No Data Found", "Formatter Error", "System Error", "Formatting", "Delivered AS2"
                Return False
        End Select

        Select Case cbeReports.Value
            Case "KBDOT"
                If InStr(requestTitle.ToUpper(), "KB") = 0 Then
                    Return False
                End If
            Case "SCORE"
                If InStr(requestTitle.ToUpper, "SCORECARD") = 0 Then
                    Return False
                    Debug.Print(requestTitle)
                End If
        End Select

        Return includeReport

    End Function


    Function Locate_Report_Files(Optional jobID As String = "") As String
        Dim reportPath As String = ""
        Dim objWebClient As New System.Net.WebClient

        Dim theElementCollection As HtmlElementCollection = wb.Document.GetElementsByTagName("iframe")
        Dim gotReportFiles As Boolean = False
        Dim c As String = wb.Document.Cookie

        For Each curElement As HtmlElement In theElementCollection
            If curElement.Id = "JobTable" Then
                Dim htmlDoc As HtmlDocument = wb.Document
                For Each htmlWin As HtmlWindow In htmlDoc.Window.Parent.Frames

                    Try
                        Dim FrameDoc As HtmlDocument = htmlWin.Document
                        If FrameDoc.GetElementById("myTable") IsNot Nothing And Not gotReportFiles Then
                            Dim FrameHead As HtmlElement = FrameDoc.GetElementsByTagName("head")(0)
                            Dim fnaScript As HtmlElement = FrameDoc.CreateElement("script")
                            fnaScript.SetAttribute("text", "function fnaGrab(i){return FileNameArray[i]}")
                            FrameHead.AppendChild(fnaScript)

                            For Each row As DataRow In dst.Tables("SOTRETL1").Select()
                                Dim JOB_ID As String = IIf(jobID <> "", jobID, row.Item("JOB_ID") & "")
                                Dim spanCollection As HtmlElementCollection = FrameDoc.GetElementsByTagName("span")
                                For Each spanElement As HtmlElement In spanCollection
                                    Dim spanValue As String = Trim(spanElement.InnerText)
                                    'Debug.Print("Span Value: " & spanValue)
                                    If spanValue = JOB_ID Then
                                        Dim REPORT_INDEX As String = row.Item("REPORT_INDEX")
                                        Dim FILE_NAME As String = htmlWin.Document.InvokeScript("fnaGrab", New String() {REPORT_INDEX})
                                        If FILE_NAME = "" Then
                                            MsgBox("Missing File Name", vbOKOnly, "Error")
                                        End If
                                        row.Item("FILE_NAME") = FILE_NAME
                                        reportPath = FILE_NAME
                                    End If
                                Next
                                row.Item("EXT") = "xls"
                                If Trim(row.Item("REQUEST_TITLE")) = "KB Dot Com sales Inc V2 (LW, TW and Last 52 Wks)" _
                                    Or Trim(row.Item("REQUEST_TITLE")) = "KB DotCom Inv LW." Then
                                    row.Item("EXT") = "xlsx"
                                End If
                            Next
                            dst.Tables("SOTRETL1").AcceptChanges()

                            gotReportFiles = True

                        End If
                    Catch ex As Exception

                    End Try
                Next
            End If
        Next
        Return reportPath

    End Function
    Private Sub Download_Reports()
        Dim objWebClient As New System.Net.WebClient

        Dim theElementCollection As HtmlElementCollection = wb.Document.GetElementsByTagName("iframe")
        Dim gotReportFiles As Boolean = False
        Dim c As String = wb.Document.Cookie
        Dim debugMsg As String = ""
        dst.Tables("TATWMRL1").Clear()

        Dim hasNotes As Boolean = False
        For Each ROW As DataRow In dst.Tables("SOTRETL1").Select()
            Dim reportPath As String = ""
            Dim JOB_ID As String = ROW.Item("JOB_ID")
            ASCMAIN1.Progress("Downloading Job ID " & JOB_ID)
            ASCMAIN1.sql = "Delete from  TATWMRL1 where JOB_ID = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", JOB_ID)
            Dim tempFolder As String = ASCMAIN1.Folders("Temp") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value & "\"
            If Not My.Computer.FileSystem.DirectoryExists(tempFolder) Then
                My.Computer.FileSystem.CreateDirectory(tempFolder)
            End If
            Dim tmpfile As String = tempFolder & JOB_ID & "." & ROW.Item("EXT")
            Dim fileName As String = ROW.Item("FILE_NAME") & ""
            Dim extension As String = ROW.Item("EXT") & ""

            If fileName & "" = "" Then
                debugMsg = "missing FN error: "
                reportPath = Locate_Report_Files(JOB_ID)
            Else
                reportPath = fileName
                debugMsg = fileName & ": "
            End If
            Dim reportUrl As String = "https://retaillink.wal-mart.com" & reportPath & "." & ROW.Item("EXT")
            objWebClient.Headers.Add(HttpRequestHeader.Cookie, wb.Document.Cookie)

            Try
                objWebClient.DownloadFile(reportUrl, tmpfile)

                ROW.Item("DOWNLOAD_COMPLETE") = "1"
            Catch ex As Exception
                ROW.Item("DOWNLOAD_COMPLETE") = "0"
                ROW.Item("NOTES") = debugMsg & ex.Message
                hasNotes = True
            End Try

            Locate_Report_Files(JOB_ID)

            Dim rowTATWMRL1 As DataRow = dst.Tables("TATWMRL1").NewRow
            rowTATWMRL1.Item("YYYYWW") = YYYYWW
            rowTATWMRL1.Item("JOB_ID") = JOB_ID & ""
            rowTATWMRL1.Item("REQUEST_TITLE") = ROW.Item("REQUEST_TITLE")
            rowTATWMRL1.Item("RUN_TIME") = ROW.Item("RUN_TIME")
            rowTATWMRL1.Item("INIT_DATE") = Now
            rowTATWMRL1.Item("INIT_USER") = ASCMAIN1.USER_ID
            dst.Tables("TATWMRL1").Rows.Add(rowTATWMRL1)
        Next
        If hasNotes Then
            grdSOTRETL1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ExtendLastColumn
        Else
            grdSOTRETL1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.None
        End If
        grdSOTRETL1.DisplayLayout.Bands(0).Columns("NOTES").Hidden = Not hasNotes


        Save_To_Network()

        If cbeReports.Value = "SCORE" Then
            Generate_Combined_Scorecard()
        End If

        If cbeReports.Value = "KBDOT" Then
            Generate_KB_Dot_Com_Report()
        End If

    End Sub
    Private Sub Save_To_Network()

        BeginTrans()

        Update_Record_TDA("TATWMRL1")
        Dim tempDir As String = ASCMAIN1.Folders("Temp") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value & "\"
        Dim networkDir As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value & "\"
        'If Not My.Computer.FileSystem.DirectoryExists(tempDir) Then
        '    My.Computer.FileSystem.CreateDirectory(tempDir)
        'End If
        If Not My.Computer.FileSystem.DirectoryExists(networkDir) Then
            My.Computer.FileSystem.CreateDirectory(networkDir)
        Else
            'Dim nd As New DirectoryInfo(networkDir)
            'Dim ndFiles As FileInfo() = nd.GetFiles()
            'For Each ndFileInfo As FileInfo In ndFiles
            '    Dim restoredFile As String = mailboxDir & "\" & as2ArcFI.Name
            '    File.Move(archiveDir & "\" & as2ArcFI.Name, restoredFile)
            'Next
        End If
        My.Computer.FileSystem.CopyDirectory(tempDir, networkDir, True)
        Dim skippedMsg As String = "Reports Skipped: " & skippedReports.ToString
        My.Computer.FileSystem.DeleteDirectory(tempDir, FileIO.DeleteDirectoryOption.DeleteAllContents)
        Dim finishedMsg As String = "Reports Downloaded: " & downloadedReports.ToString & vbCrLf _
                                    & IIf(skippedReports > 0, skippedMsg & vbCrLf, "") _
                                    & vbCrLf & "Click OK to Generate " & cbeReports.Text & " Report"
        'MsgBox(finishedMsg)
        CommitTrans(finishedMsg)

    End Sub
    Sub Email_Report(attachments As Dictionary(Of String, String), attachmentDesc As String, emailBody As String)
        Dim emailForm As New TAFSEND1(Me)

        emailForm.SEND_TOs.Add("rdw@absolution.com", "Robert Wall")
        emailForm.SEND_TOs.Add("dashear@vandale.com", "David Ashear")
        emailForm.SEND_FROM = "dashear@vandale.com"
        emailForm.SEND_BODY = emailBody
        emailForm.EMAIL_KEY = "SOFRETL1"
        emailForm.SEND_SUBJECT = cbeReports.Text & ", Week " & Mid(YYYYWW, 5, 2)
        emailForm.SEND_ATTACHMENTs = attachments

        emailForm.SEND_FROM_SIGNATURE = "" _
        & "     <b>" _
        & "         <i>" _
        & "             <span style='font-size:9.0pt;font-family:&quot;Arial&quot;,sans-serif'>David E. Ashear</span>" _
        & "         </i>" _
        & "     </b>" _
        & "     <span style='font-size:9.0pt;font-family:&quot;Arial&quot;,sans-serif;color:navy'>, Director of Operations &nbsp;</span>" _
        & "     <b>" _
        & "         <span style='font-size:9.0pt;font-family:&quot;Georgia&quot;,serif;color:navy'>• Vandale Industries, Inc.</span>" _
        & "     </b>" _
        & "     <br>" _
        & "     <b>" _
        & "         <span style='font-size:9.0pt;font-family:&quot;Berlin Sans FB Demi&quot;,sans-serif;color:navy'>Corporate Office contact Info</span>" _
        & "     </b>" _
        & "     <span style='font-size:9.0pt;font-family:&quot;Arial&quot;,sans-serif;color:navy'>:180 Madison Ave, New York, NY, 10016 •&nbsp;</span>" _
        & "     <span style='font-size:10.0pt;font-family:Wingdings;color:black'>(</span>" _
        & "     <span style='font-size:9.0pt;font-family:&quot;Arial&quot;,sans-serif;color:navy'>:&nbsp;212-683-8181 ext. 188 • Fax:&nbsp;212-683-1424</span>" _
        & "     <br>" _
        & "     <b>" _
        & "         <span style='font-size:9.0pt;font-family:&quot;Berlin Sans FB Demi&quot;,sans-serif;color:navy'>Warehouse contact Info:</span>" _
        & "     </b>" _
        & "     <span style='font-size:9.0pt;font-family:&quot;Berlin Sans FB Demi&quot;,sans-serif;color:navy'>&nbsp;</span>" _
        & "     <span style='font-size:9.0pt;font-family:&quot;Arial&quot;,sans-serif;color:navy'>40 Executive Ave, Edison , NJ, 08817 &nbsp;•</span>" _
        & "     <span style='font-size:10.0pt;font-family:Wingdings;color:black'>(</span>" _
        & "     <span style='font-size:9.0pt;font-family:&quot;Arial&quot;,sans-serif;color:navy'>: 732-902-2576 &nbsp;• Fax 732-902-2579</span>"

        emailForm.SEND_METHOD = "E"
        emailForm.viewAsHtml = True

        If ASCMAIN1.Running_in_VS Then
            emailForm.Show()
        Else
            emailForm.Send_email_automatically()
            emailForm.Dispose()
            MsgBox("Email has been sent")
        End If
    End Sub

    Sub Import_AS2_Files()

        Dim mailboxDir As String = "\\svr-nyc-edi1\inbox\Wal-MartAS2\ReportsMailbox"
        Dim archiveDir As String = "\\svr-nyc-edi1\inbox\Wal-MartAS2\ReportsMailbox\Archive\" & YYYYWW & "\"

        Dim as2Dir As New DirectoryInfo(mailboxDir)
        Dim as2Files As FileInfo() = as2Dir.GetFiles()
        Dim as2FileCount As Integer = 0

        If as2Files.Count = 0 Then
            If Not My.Computer.FileSystem.DirectoryExists(archiveDir) Then
                MsgBox("No data to import", vbOKOnly + MsgBoxStyle.Critical, "Cannot Proceed")
                Exit Sub
            Else
                Dim noFileMessage As String = "There are no new AS2 files to import." & vbCrLf & vbCrLf _
                                              & "Run report with archived data for week " & YYYYWW & "?"
                If MessageBox.Show(noFileMessage, "No New Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    Dim as2ArcDir As New DirectoryInfo(archiveDir)
                    Dim as2ArcFiles As FileInfo() = as2ArcDir.GetFiles()
                    For Each as2ArcFI As FileInfo In as2ArcFiles
                        Dim restoredFile As String = mailboxDir & "\" & as2ArcFI.Name
                        File.Move(archiveDir & "\" & as2ArcFI.Name, restoredFile)
                    Next

                    as2Files = as2Dir.GetFiles
                Else
                    Exit Sub
                End If
            End If
        End If

        If cbeReports.Value = "KBDOT" Then
            Dim resultWorkbookSet As SpreadsheetGear.IWorkbookSet = Nothing
            Dim tempFolder As String = ASCMAIN1.Folders("Temp") & "RetailLink\"
            If Not My.Computer.FileSystem.DirectoryExists(tempFolder) Then
                My.Computer.FileSystem.CreateDirectory(tempFolder)
            End If
            File.Copy(Path.Combine(ASCMAIN1.Folders("Archive") & "RetailLink\Templates", "KBDotComTemplate.xls"), Path.Combine(tempFolder, "KBDC.xlsx"), True)
            Dim importResult As String = ASCMAIN1.Folders("Temp") & "RetailLink\KBDC.xlsx"
            resultWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet()
            resultWorkbook = resultWorkbookSet.Workbooks.Open(importResult)
        End If

        For Each as2FI As FileInfo In as2Files
            as2FileCount += 1
            Load_File_Into_Spreadsheet(mailboxDir & "\" & as2FI.Name)



            Dim archivedFile As String = archiveDir & as2FI.Name

            If Not My.Computer.FileSystem.DirectoryExists(archiveDir) Then
                My.Computer.FileSystem.CreateDirectory(archiveDir)
            End If

            If File.Exists(archivedFile) Then
                File.Delete(archivedFile)
            End If

            ' Move the file.
            File.Move(mailboxDir & "\" & as2FI.Name, archivedFile)

        Next as2FI

        If cbeReports.Value = "KBDOT" Then

            Dim resultIndexCells As SpreadsheetGear.IRange = resultWorkbook.Worksheets("Program").Cells

            For Each formulaColumn As String In New String() {"E", "M", "N", "O", "Q"}
                For j As Integer = 6 To 600
                    Dim cellReference As String = formulaColumn & j.ToString
                    If Mid(resultIndexCells(cellReference).Formula, 1, 8) = "=VLOOKUP" Then
                        Dim newFormula As String = Replace(resultIndexCells(cellReference).Formula, "+", "")
                        resultIndexCells(cellReference).Formula = newFormula
                    End If
                Next
            Next

            resultIndexCells.Cells("O1").Formula = Mid(YYYYWW, 5, 2)

            Dim resultDirectory As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value & "\"
            Dim resultFilename As String = resultDirectory & "Walmart Dot Com Shapewear Week " & Mid(YYYYWW, 5, 2) & ".xlsx"

            If Not My.Computer.FileSystem.DirectoryExists(resultDirectory) Then
                My.Computer.FileSystem.CreateDirectory(resultDirectory)
            End If

            If System.IO.File.Exists(resultFilename) Then
                My.Computer.FileSystem.DeleteFile(resultFilename)
            End If

            resultWorkbook.SaveAs(resultFilename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)


            If chkEmailReport.Checked Then
                Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(resultFilename)
                Dim attachedReports As New Dictionary(Of String, String)
                attachedReports.Add(FI.Name, resultFilename)
                Dim emailBody As String = Generate_Email_Body(resultWorkbook.Worksheets("Program").Cells("S7:T22"))
                Email_Report(attachedReports, FI.Name, emailBody)
                If Not ASCMAIN1.Running_in_VS Then
                    Dim p As Process = Process.Start(resultFilename)
                End If
            Else
                MsgBox(resultFilename, vbOKOnly, "Done")
                Dim p As Process = Process.Start(resultFilename)
            End If

        End If


    End Sub
    Sub Load_File_Into_Spreadsheet(as2File As String)
        Dim oReader As New StreamReader(as2File)
        Dim sLine As String = Nothing
        Dim reportName As String = ""

        Dim targetWorksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim as2Row As Integer = 0
        Dim worksheetStartRow As Integer = 0
        Dim upcColumn As Integer = 0
        Dim optionsStartRow As Integer = 3

        While Not oReader.EndOfStream
            sLine = oReader.ReadLine()
            If Mid(sLine, 1, 6) = "TITLE:" Then
                reportName = Mid(sLine, 7)
                If InStr(reportName, "KB DotCom") >= 0 Then
                    Dim targetWorksheetName As String = IIf(InStr(reportName, "Inv LW") > 0, "Inv Data", "Sales LW TY")
                    worksheetStartRow = IIf(InStr(reportName, "Inv LW") > 0, 16, 14)
                    upcColumn = IIf(InStr(reportName, "Inv LW") > 0, 3, 0)
                    For i As Integer = 0 To resultWorkbook.Worksheets.Count - 1
                        If resultWorkbook.Worksheets(i).Name = targetWorksheetName Then
                            targetWorksheet = resultWorkbook.Worksheets(i)
                            targetWorksheet.Cells("B1").Formula = Replace(reportName, "AS2", "")
                            targetWorksheet.Cells("B3").Formula = "Report Options"
                        End If
                    Next
                Else

                End If
            End If
            If Mid(sLine, 1, 8) = "REQUEST:" Then
                Dim dataLine As String = Mid(sLine, 9)
                targetWorksheet.Cells(optionsStartRow, 1).Formula = "Requested: " & dataLine
                optionsStartRow += 1
            End If
            If Mid(sLine, 1, 9) = "CRITERIA:" Then

                If Mid(sLine, 10, 16) = "Report Columns :" Then
                    Dim dataLine As String = Mid(sLine, 26)
                    Dim dataArray As String() = dataLine.Split(","c)
                    Dim colNo As Integer = 0
                    Dim cRow As String = "Report Columns :"
                    For Each dv As String In dataArray
                        cRow += dv & ","
                        colNo += 1
                        If Len(cRow) > 90 Or colNo = dataArray.Length Then
                            targetWorksheet.Cells(optionsStartRow, 1).Formula = Mid(cRow, 1, Len(cRow) - 1)
                            cRow = ""
                            colNo = 0
                            optionsStartRow += 1
                        End If
                    Next
                    If cRow <> "" Then
                        targetWorksheet.Cells(optionsStartRow, 1).Formula = Mid(cRow, 1, Len(cRow) - 1)
                        optionsStartRow += 1
                    End If
                Else
                    Dim dataLine As String = Mid(sLine, 10)
                    targetWorksheet.Cells(optionsStartRow, 1).Formula = dataLine
                    optionsStartRow += 1
                End If


            End If
            If Mid(sLine, 1, 7) = "COLUMN:" Then
                Dim dataLine As String = Mid(sLine, 8)
                Dim dataArray As String() = dataLine.Split("|"c)
                Dim colNo As Integer = 0
                For Each dv As String In dataArray
                    targetWorksheet.Cells(worksheetStartRow - 2, colNo).Formula = dv
                    colNo += 1
                Next
            End If
            If Mid(sLine, 1, 5) = "DATA:" Then
                Dim dataLine As String = Mid(sLine, 6)
                Dim dataArray As String() = dataLine.Split("|"c)
                Dim colNo As Integer = 0
                For Each dv As String In dataArray
                    If colNo = upcColumn Then
                        targetWorksheet.Cells(worksheetStartRow - 1, colNo).Formula = "=TEXT(" & dv & ",""0000000000000"")"
                    Else
                        targetWorksheet.Cells(worksheetStartRow - 1, colNo).Formula = dv
                    End If
                    colNo += 1
                Next
                worksheetStartRow += 1

            End If


        End While

        'If InStr(reportName, "KB DotCom") >= 0 Then
        '    resultWorkbook.WindowInfo.ActiveSheet = resultWorkbook.Worksheets("Program")
        '    Dim resultFilename As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\Walmart Dot Com Shapewear Week " & Mid(YYYYWW, 5, 2) & ".xlsx"
        '    If System.IO.File.Exists(resultFilename) Then
        '        My.Computer.FileSystem.DeleteFile(resultFilename)
        '    End If
        '    resultWorkbook.SaveAs(resultFilename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        '    MsgBox(resultFilename, vbOKOnly, "Done")
        '    Dim p As Process = Process.Start(resultFilename)
        'End If


        oReader.Close()
    End Sub
    Function HTMLTable2DataTable(ByVal HTML As String) As DataTable
        ' Declarations  

        Dim dt As DataTable
        Dim dr As DataRow
        'Dim dc As DataColumn
        Dim TableExpression As String = "<TABLE[^>]*>(.*?)</TABLE>"
        Dim HeaderExpression As String = "<TH[^>]*>(.*?)</TH>"
        Dim RowExpression As String = "<TR[^>]*>(.*?)</TR>"
        Dim ColumnExpression As String = "<TD[^>]*>(.*?)</TD>"
        Dim HeadersExist As Boolean = False
        Dim iCurrentColumn As Integer = 0
        Dim iCurrentRow As Integer = 0

        ' Get a match for all the tables in the HTML  
        Dim Tables As MatchCollection = Regex.Matches(HTML, TableExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)


        ' Reset the current row counter and the header flag  
        iCurrentRow = 0
        HeadersExist = False

        ' Add a new table to the DataSet  
        dt = New DataTable
        If Tables.Count = 0 Then
            Return dt
        End If
        Dim Table As Match = Tables(0)
        Dim colCount As Integer = 0
        ' Create the relevant amount of columns for this table (use the headers if they exist, otherwise use default names)  
        If Table.Value.Contains("<th") Then
            ' Set the HeadersExist flag  
            HeadersExist = True

            ' Get a match for all the rows in the table  
            Dim Headers As MatchCollection = Regex.Matches(Table.Value, HeaderExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

            ' Loop through each header element  
            For Each Header As Match In Headers
                Dim cn As String = Header.Groups(1).ToString
                If dt.Columns.Contains(cn) Then
                    cn &= "_" & colCount.ToString
                End If
                dt.Columns.Add(cn)
                colCount += 1
            Next
        Else
            For iColumns As Integer = 1 To Regex.Matches(Regex.Matches(Regex.Matches(Table.Value, TableExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Item(0).ToString, RowExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Item(0).ToString, ColumnExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Count
                Dim cn As String = "Column " & iColumns
                If dt.Columns.Contains(cn) Then
                    cn &= "_" & colCount.ToString
                End If
                dt.Columns.Add(cn)
                colCount += 1
            Next
        End If

        ' Get a match for all the rows in the table  
        Dim Rows As MatchCollection = Regex.Matches(Table.Value, RowExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

        ' Loop through each row element  
        For Each Row As Match In Rows

            ' Only loop through the row if it isn't a header row  
            If Not (iCurrentRow = 0 And HeadersExist = True) Then

                ' Create a new row and reset the current column counter  
                dr = dt.NewRow
                iCurrentColumn = 0

                ' Get a match for all the columns in the row  
                Dim Columns As MatchCollection = Regex.Matches(Row.Value, ColumnExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                ' Loop through each column element  
                For Each Column As Match In Columns

                    ' Add the value to the DataRow  
                    If InStr(Column.Groups(1).ToString, "<INPUT") > 0 Then
                    Else
                        If InStr(Column.Groups(1).ToString, "<SPAN") > 0 Then
                            Dim cellData As MatchCollection = Regex.Matches(Column.Groups(1).Value, "<SPAN[^>]*>(.*?)</SPAN>", RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                            dr(iCurrentColumn) = cellData.Item(0).Groups(1).ToString
                        End If

                    End If


                    ' Increase the current column  
                    iCurrentColumn += 1
                Next

                ' Add the DataRow to the DataTable  
                dt.Rows.Add(dr)

            End If

            ' Increase the current row counter  
            iCurrentRow += 1
        Next

        Return (dt)

    End Function
    Sub Generate_Combined_Scorecard()
        Try
            Dim scorecardTemplate As String = ASCMAIN1.Folders("Archive") & "RetailLink\Templates\ScordCardsTemplate.xlsx"
            Dim combinedScorecard As String = ASCMAIN1.Folders("Archive") & "RetailLink\CombinedScorecard.xlsx"
            Dim templateWorkbookSet As SpreadsheetGear.IWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet()
            Dim templateWorkbook As SpreadsheetGear.IWorkbook = templateWorkbookSet.Workbooks.Open(scorecardTemplate)
            Dim combinedWorkbookSet As SpreadsheetGear.IWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet()
            Dim combinedtWorkbook As SpreadsheetGear.IWorkbook = combinedWorkbookSet.Workbooks.Add()
            Dim yyyyww_Walmart As String = ""

            If System.IO.File.Exists(combinedScorecard) Then
                My.Computer.FileSystem.DeleteFile(combinedScorecard)
            End If
            Dim columnMap As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))
            columnMap.Add(New KeyValuePair(Of String, String)("C", "D"))
            columnMap.Add(New KeyValuePair(Of String, String)("F", "G"))
            columnMap.Add(New KeyValuePair(Of String, String)("I", "J"))
            columnMap.Add(New KeyValuePair(Of String, String)("L", "M"))
            columnMap.Add(New KeyValuePair(Of String, String)("O", "P"))
            columnMap.Add(New KeyValuePair(Of String, String)("R", "S"))
            columnMap.Add(New KeyValuePair(Of String, String)("U", "V"))

            For i As Integer = 0 To templateWorkbook.Worksheets.Count - 1
                Dim templateWorksheet As SpreadsheetGear.IWorksheet = templateWorkbook.Worksheets(i)
                Dim templateRange As SpreadsheetGear.IRange = templateWorksheet.Cells
                Dim templateNames As SpreadsheetGear.INames = templateWorkbook.Names
                Dim combinedWorksheet As SpreadsheetGear.IWorksheet = combinedtWorkbook.Worksheets.Add()
                If templateWorksheet.Name = "Index" Then
                    templateRange.Copy(combinedWorksheet.Cells("A1"), SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
                Else
                    templateRange.Copy(combinedWorksheet.Cells("A1"))
                End If

                combinedWorksheet.Name = templateWorksheet.Name

                If combinedWorksheet.Name <> "Index" Then
                    Dim rlReportName As String = Replace(combinedWorksheet.Name, "D", "Dept ")
                    ASCMAIN1.sql = "Select * from TATWMRL1 where REQUEST_TITLE like '%" & rlReportName & "%' and YYYYWW = '" & YYYYWW & "'"
                    Dim rowTATWMRL1 As DataRow = ASCDATA1.GetDataRow()
                    If rowTATWMRL1 IsNot Nothing Then

                        Dim dataWorksheet As SpreadsheetGear.IWorksheet = Get_Scorecard_Data_For(rlReportName)
                        If dataWorksheet IsNot Nothing Then
                            Dim dataRows As Integer = dataWorksheet.UsedRange.RowCount
                            Dim dataColumns As Integer = dataWorksheet.UsedRange.ColumnCount
                            Dim dataCells As SpreadsheetGear.IRange = dataWorksheet.UsedRange
                            Dim dataRangeToCopy As String = ""
                            Dim dataEnd As String = "W" & dataRows
                            Dim headerRow As Integer = 0

                            For j As Integer = 1 To dataRows
                                Dim dataValue As String = dataCells.Cells("A" & j.ToString).Formula
                                If Len(dataValue) > 0 Then 'should be VENDOR SUMMARY : VANDALE INDUSTRIES INC
                                    dataRangeToCopy = "A" & j.ToString & ":" & dataEnd
                                    headerRow = j - 2
                                    Exit For
                                End If
                            Next
                            Dim dataRange As SpreadsheetGear.IRange = dataWorksheet.Cells(dataRangeToCopy)
                            dataRange.Copy(combinedWorksheet.Cells("B12"), SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)
                            For Each column As KeyValuePair(Of String, String) In columnMap
                                Dim sourceColumn As String = column.Key
                                Dim destColumn As String = column.Value
                                Dim headerValue As String = dataCells.Cells(sourceColumn & headerRow.ToString).Formula
                                combinedWorksheet.Cells(destColumn & "10").Formula = headerValue
                                If destColumn = "D" And yyyyww_Walmart = "" Then
                                    yyyyww_Walmart = Mid(headerValue, 12, 6)
                                End If
                            Next

                            For j As Integer = 1 To 9
                                Dim headerInfo As String = dataCells.Cells("B" & j.ToString).Formula
                                combinedWorksheet.Cells("D" & j.ToString).Formula = headerInfo
                            Next
                            combinedWorksheet.Tab.Color = SpreadsheetGear.Color.FromArgb(55, 86, 35)
                        Else
                            combinedWorksheet.Tab.Color = SpreadsheetGear.Color.FromArgb(255, 0, 0)
                        End If
                    Else
                        combinedWorksheet.Tab.Color = SpreadsheetGear.Color.FromArgb(255, 0, 0)
                    End If
                End If
            Next

            'update index worksheet
            Format_Index_Page(templateWorkbook.Worksheets("Index"), combinedtWorkbook.Worksheets("Index"))

            combinedtWorkbook.Worksheets("Sheet1").Delete()
            combinedtWorkbook.WindowInfo.ActiveSheet = combinedtWorkbook.Worksheets("Index")

            Dim scorecardFilename As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value & "\ScoreCards ~ Week " & yyyyww_Walmart & ".xlsx"
            If System.IO.File.Exists(scorecardFilename) Then
                My.Computer.FileSystem.DeleteFile(scorecardFilename)
            End If
            combinedtWorkbook.SaveAs(scorecardFilename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

            If chkEmailReport.Checked Then
                Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(scorecardFilename)
                Dim attachedReports As New Dictionary(Of String, String)
                attachedReports.Add(FI.Name, scorecardFilename)
                Dim emailBody As String = Generate_Email_Body(combinedtWorkbook.Worksheets("Index").Cells("B2:G7"))
                Email_Report(attachedReports, FI.Name, emailBody)
                If Not ASCMAIN1.Running_in_VS Then
                    Dim p As Process = Process.Start(scorecardFilename)
                End If
            Else
                MsgBox(scorecardFilename, vbOKOnly, "Done")
                Dim p As Process = Process.Start(scorecardFilename)
            End If


        Catch ex As Exception

        End Try

    End Sub
    Function Generate_Email_Body(bodyRange As SpreadsheetGear.IRange) As String
        Dim emailBody As String = ""
        Dim tableWidth As String = IIf(cbeReports.Value = "KBDOT", "400px", "100%")

        emailBody += "<table style='width:" & tableWidth & "'>"
        For i As Integer = 1 To bodyRange.RowCount
            Dim rowFormatting As String = "style='"
            Select Case cbeReports.Value
                Case "SCORE"
                    If i = 1 Then
                        rowFormatting += "background-color:#70AD47;color:#fff;font-family: Arial;text-align:center;padding: 20px 0;height:40px;"
                    Else
                        If (i Mod 2 = 0) Then
                            rowFormatting += "background-color:#E2EFDA;"
                        End If
                        rowFormatting += "color:#000;font-family: Arial;"
                    End If
                Case "KBDOT"
                    If i = 1 Then
                        rowFormatting += "background-color:#4F81BD;color:#fff;font-family: Arial;text-align:center;height:20px;"
                    Else
                        If (i Mod 2 = 0) Then
                            rowFormatting += "background-color:#B8CCE4;"
                        Else
                            rowFormatting += "background-color:#DCE6F1;"

                        End If
                        rowFormatting += "color:#000;font-family: Arial;"
                    End If
            End Select
            rowFormatting += "'"

            emailBody += "<tr " & rowFormatting & ">"
            For j As Integer = 1 To bodyRange.ColumnCount
                Dim cellFormatting As String = "style='"
                Dim spanFormatting As String = "style='"
                Select Case cbeReports.Value
                    Case "SCORE"
                        If i = 1 Then
                            spanFormatting += "margin: 0 20px;"
                        End If
                        If i = 1 And j = bodyRange.ColumnCount Then
                            cellFormatting += "background-color:#ED7D31;"
                        End If
                        If i = bodyRange.RowCount And j > 1 Then
                            cellFormatting += "font-weight:bold;"
                        End If
                        cellFormatting += "text-align:center;"
                    Case "KBDOT"
                        If j = 1 Then
                            cellFormatting += "text-align:right;"
                        Else
                            cellFormatting += "text-align:center;"
                        End If
                        If i = bodyRange.RowCount Then
                            cellFormatting += "background-color:#E6B8B7;"
                        End If
                End Select
                cellFormatting += "'"
                spanFormatting += "'"
                If cellFormatting = "style=''" Then
                    cellFormatting = ""
                End If
                If spanFormatting = "style=''" Then
                    spanFormatting = ""
                End If
                emailBody += "<td " & cellFormatting & " >"
                emailBody += bodyRange(i - 1, j - 1).Text
                emailBody += "</td>"
            Next

            emailBody += "</tr>"
        Next
        emailBody += "</table>"

        Return emailBody

    End Function
    Function Get_Scorecard_Data_For(name As String) As SpreadsheetGear.IWorksheet
        Dim dataWorksheet As SpreadsheetGear.IWorksheet = Nothing
        ASCMAIN1.sql = "Select * from TATWMRL1 where REQUEST_TITLE like '%" & name & "%' and YYYYWW = '" & YYYYWW & "'"
        Dim rowTATWMRL1 As DataRow = ASCDATA1.GetDataRow()
        Dim ext As String = "xls"
        If name = "KB Dot Com sales Inc V2 (LW, TW and Last 52 Wks)" _
    Or name = "KB DotCom Inv LW." Then
            ext = "xlsx"
        End If
        If rowTATWMRL1 IsNot Nothing Then
            Dim JOB_ID As String = rowTATWMRL1.Item("JOB_ID")
            Dim dataFile As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value & "\" & JOB_ID & "." & ext
            Dim dataFileSize As String = ""
            If System.IO.File.Exists(dataFile) Then
                Dim dataWorkbookSet As SpreadsheetGear.IWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet()
                Dim dataWorkbook As SpreadsheetGear.IWorkbook = dataWorkbookSet.Workbooks.Open(dataFile)
                dataWorksheet = dataWorkbook.Worksheets(0)
                Dim infoReader As System.IO.FileInfo
                infoReader = My.Computer.FileSystem.GetFileInfo(dataFile)
                dataFileSize = Math.Round((infoReader.Length * 0.001), 2).ToString & "Kb"
            Else
                Return Nothing
            End If
            Dim rowSOTRETL1 As DataRow = dst.Tables("SOTRETL1").Rows.Find(New Object() {JOB_ID})
            If rowSOTRETL1 Is Nothing Then
                Load_Existing_Record(rowTATWMRL1, dataFileSize)
            End If
        End If
        Return dataWorksheet
    End Function

    Sub Generate_KB_Dot_Com_Report()
        Try
            Dim scorecardTemplate As String = ASCMAIN1.Folders("Archive") & "RetailLink\Templates\KBDotComTemplate.xls"
            File.Copy(Path.Combine(ASCMAIN1.Folders("Archive") & "RetailLink\Templates", "KBDotComTemplate.xls") _
                      , Path.Combine(ASCMAIN1.Folders("Temp") & "RetailLink", "KBDC.xlsx"), True)

            Dim combinedScorecard As String = ASCMAIN1.Folders("Temp") & "RetailLink\KBDC.xlsx"
            Dim templateWorkbookSet As SpreadsheetGear.IWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet()
            Dim templateWorkbook As SpreadsheetGear.IWorkbook = templateWorkbookSet.Workbooks.Open(scorecardTemplate)
            Dim combinedWorkbookSet As SpreadsheetGear.IWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet()
            'Dim combinedtWorkbook As SpreadsheetGear.IWorkbook = combinedWorkbookSet.Workbooks.Add()
            Dim combinedtWorkbook As SpreadsheetGear.IWorkbook = combinedWorkbookSet.Workbooks.Open(combinedScorecard)

            Dim yyyyww_Walmart As String = ""

            If System.IO.File.Exists(combinedScorecard) Then
                My.Computer.FileSystem.DeleteFile(combinedScorecard)
            End If


            For i As Integer = 0 To templateWorkbook.Worksheets.Count - 1
                Dim templateWorksheet As SpreadsheetGear.IWorksheet = templateWorkbook.Worksheets(i)
                Dim templateRange As SpreadsheetGear.IRange = templateWorksheet.Cells
                Dim templateNames As SpreadsheetGear.INames = templateWorkbook.Names
                Dim combinedWorksheet As SpreadsheetGear.IWorksheet = combinedtWorkbook.Worksheets(templateWorksheet.Name)
                templateRange.Copy(combinedWorksheet.Cells("A1"))
                combinedWorksheet.Name = templateWorksheet.Name

                If combinedWorksheet.Name = "Sales LW TY" Or combinedWorksheet.Name = "Inv Data" Then

                    Dim rlReportName As String = "KB Dot Com sales Inc V2 (LW, TW and Last 52 Wks)"
                    If combinedWorksheet.Name = "Inv Data" Then
                        rlReportName = "KB DotCom Inv LW."
                    End If
                    ASCMAIN1.sql = "Select * from TATWMRL1 where REQUEST_TITLE like '%" & rlReportName & "%' and YYYYWW = '" & YYYYWW & "'"
                    Dim rowTATWMRL1 As DataRow = ASCDATA1.GetDataRow()
                    If rowTATWMRL1 IsNot Nothing Then

                        Dim dataWorksheet As SpreadsheetGear.IWorksheet = Get_Scorecard_Data_For(rlReportName)
                        If dataWorksheet IsNot Nothing Then
                            Dim dataRows As Integer = dataWorksheet.UsedRange.RowCount
                            Dim dataColumns As Integer = dataWorksheet.UsedRange.ColumnCount
                            Dim dataCells As SpreadsheetGear.IRange = dataWorksheet.UsedRange
                            Dim dataRangeToCopy As String = ""
                            Dim dataEnd As String = IIf(combinedWorksheet.Name = "Inv Data", "Q", "F") & dataRows
                            Dim headerRow As Integer = 0
                            dataRangeToCopy = "A1:" & dataEnd

                            Dim dataRange As SpreadsheetGear.IRange = dataWorksheet.Cells(dataRangeToCopy)
                            dataRange.Copy(combinedWorksheet.Cells("A1"), SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)

                            'For j As Integer = 1 To 9
                            '    Dim headerInfo As String = dataCells.Cells("B" & j.ToString).Formula
                            '    combinedWorksheet.Cells("D" & j.ToString).Formula = headerInfo
                            'Next
                            combinedWorksheet.Tab.Color = SpreadsheetGear.Color.FromArgb(55, 86, 35)
                        Else
                            combinedWorksheet.Tab.Color = SpreadsheetGear.Color.FromArgb(255, 0, 0)
                        End If
                    Else
                        combinedWorksheet.Tab.Color = SpreadsheetGear.Color.FromArgb(255, 0, 0)
                    End If
                Else
                    If combinedWorksheet.Name = "Program" Then
                        combinedWorksheet.Cells("O1").Formula = Mid(YYYYWW, 5, 2)
                    End If
                End If
            Next

            Update_Formulas(templateWorkbook.Worksheets("Program"), combinedtWorkbook.Worksheets("Program"))


            'combinedtWorkbook.Worksheets("Sheet1").Delete()
            combinedtWorkbook.WindowInfo.ActiveSheet = combinedtWorkbook.Worksheets("Program")
            'Copy of Walmart Dot Com Shapewear Week 43-2016 with Images.xls
            Dim scorecardFilename As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value _
                                              & "\Walmart Dot Com Shapewear Week " & Mid(YYYYWW, 5, 2) & ".xlsx"
            If System.IO.File.Exists(scorecardFilename) Then
                My.Computer.FileSystem.DeleteFile(scorecardFilename)
            End If
            combinedtWorkbook.SaveAs(scorecardFilename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)


            If chkEmailReport.Checked Then
                Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(scorecardFilename)
                Dim attachedReports As New Dictionary(Of String, String)
                attachedReports.Add(FI.Name, scorecardFilename)
                Dim emailBody As String = Generate_Email_Body(combinedtWorkbook.Worksheets("Program").Cells("S7:T22"))
                Email_Report(attachedReports, FI.Name, emailBody)
                If Not ASCMAIN1.Running_in_VS Then
                    Dim p As Process = Process.Start(scorecardFilename)
                End If
            Else
                MsgBox(scorecardFilename, vbOKOnly, "Done")
                Dim p As Process = Process.Start(scorecardFilename)
            End If

        Catch ex As Exception

        End Try

    End Sub
    Sub Update_Formulas(templateIndex As SpreadsheetGear.IWorksheet, combinedIndex As SpreadsheetGear.IWorksheet)
        Dim templateIndexCells As SpreadsheetGear.IRange = templateIndex.Cells
        Dim combinedIndexCells As SpreadsheetGear.IRange = combinedIndex.Cells
        For Each formulaColumn As String In New String() {"E", "M", "N", "O", "Q"}
            For j As Integer = 6 To 600
                Dim cellReference As String = formulaColumn & j.ToString
                If Mid(templateIndexCells(cellReference).Formula, 1, 8) = "=VLOOKUP" Then
                    Dim newFormula As String = Replace(templateIndexCells(cellReference).Formula, "+", "")
                    combinedIndexCells(cellReference).Formula = newFormula
                End If
            Next
        Next
    End Sub
    Private Sub Load_Existing_Record(rowTATWMRL1 As DataRow, fileSize As String)
        Dim rowSOTRETL1 As DataRow = dst.Tables("SOTRETL1").NewRow
        rowSOTRETL1.Item("YYYYWW") = rowTATWMRL1.Item("JOB_ID") & ""
        rowSOTRETL1.Item("JOB_ID") = rowTATWMRL1.Item("JOB_ID") & ""
        rowSOTRETL1.Item("STATUS") = "Existing"
        rowSOTRETL1.Item("REQUEST_TITLE") = rowTATWMRL1.Item("REQUEST_TITLE") & ""
        rowSOTRETL1.Item("RUN_TIME") = rowTATWMRL1.Item("RUN_TIME") & ""
        rowSOTRETL1.Item("FILE_SIZE") = fileSize
        rowSOTRETL1.Item("DOWNLOAD_COMPLETE") = "1"
        dst.Tables("SOTRETL1").Rows.Add(rowSOTRETL1)
    End Sub
    Private Sub grdSOTRETL1_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSOTRETL1.ClickCellButton
        Dim JOB_ID As String = grdSOTRETL1.ActiveCell.Row.Cells("JOB_ID").Value & ""
        Dim dataFile As String = ASCMAIN1.Folders("Archive") & "RetailLink\" & YYYYWW & "\" & cbeReports.Value _
                                              & "\" & JOB_ID & ".xls"
        If System.IO.File.Exists(dataFile) Then
            Dim p As Process = Process.Start(dataFile)
        Else
            MsgBox("Sorry, can't find file.", vbCritical, "Cannot Proceed")
        End If

    End Sub

    Private Sub Format_Index_Page(templateIndex As SpreadsheetGear.IWorksheet, combinedIndex As SpreadsheetGear.IWorksheet)
        Dim templateIndexCells As SpreadsheetGear.IRange = templateIndex.Cells
        Dim combinedIndexCells As SpreadsheetGear.IRange = combinedIndex.Cells

        For Each formulaColumn As String In New String() {"C", "D", "E", "G"}
            For j As Integer = 3 To 7
                Dim cellReference As String = formulaColumn & j.ToString
                Dim newFormula As String = Replace(templateIndexCells(cellReference).Formula, "+", "")
                combinedIndexCells(cellReference).Formula = newFormula
            Next
        Next

        For Each indexHeaderColumn As String In New String() {"B", "C", "D", "E", "F", "G"}
            If indexHeaderColumn = "G" Then
                combinedIndexCells(indexHeaderColumn & "2").Interior.Color = SpreadsheetGear.Color.FromArgb(237, 125, 49)
            Else
                combinedIndexCells(indexHeaderColumn & "2").Interior.Color = SpreadsheetGear.Color.FromArgb(112, 173, 71)
                combinedIndexCells(indexHeaderColumn & "2").Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                combinedIndexCells(indexHeaderColumn & "2").Borders.Color = SpreadsheetGear.Colors.DarkSlateGray
            End If
            combinedIndexCells(indexHeaderColumn & "2").Font.Color = SpreadsheetGear.Color.FromArgb(255, 255, 255)
            For Each oddCell As String In New String() {"3", "5", "7"}
                combinedIndexCells(indexHeaderColumn & oddCell).Interior.Color = SpreadsheetGear.Color.FromArgb(226, 239, 218)
            Next
        Next
        For Each oddCell As String In New String() {"3", "5", "7"}
            With combinedIndexCells.Range("B" & oddCell & ":G" & oddCell)
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.None
                .Borders.Color = SpreadsheetGear.Color.FromArgb(169, 208, 142)
            End With
        Next

        With combinedIndexCells.Range("B2:G7")
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders.Color = SpreadsheetGear.Colors.DarkSlateGray
        End With
    End Sub

    Private Sub grdSOTRETL1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTRETL1.InitializeRow
        Dim FILENAME As String = ASCMAIN1.Get_Filename("XLS")
        e.Row.Cells("PROCESS_REPORT").ButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", FILENAME)

    End Sub

    Private Sub weekCmb_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles weekCmb.InitializeLayout
        weekCmb.DisplayLayout.Bands(0).Columns(0).Hidden = True
        weekCmb.DisplayLayout.Bands(0).Columns(1).Width = UltraExplorerBarContainerControl2.Width
        weekCmb.DisplayLayout.Bands(0).Columns(1).Header.Caption = "Available Weeks"
    End Sub

    Private Sub cbeReports_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles cbeReports.InitializeLayout
        e.Layout.Bands(0).Columns(0).Hidden = True
        e.Layout.Bands(0).Columns(1).Header.Caption = "Report Name"

    End Sub


End Class