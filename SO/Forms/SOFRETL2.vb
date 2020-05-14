Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.IO
Imports System.Threading
Imports System.ComponentModel


<ComVisible(True)> _
Public Class SOFRETL2
    Private Const INTERNET_OPTION_END_BROWSER_SESSION As Integer = 42

    Dim sqls As String = ""
    Dim loggedIn As Boolean = False
    Dim onRoutingStatusPage As Boolean = False
    Dim routingStatusLoaded As Boolean = False
    Dim injectedScript As Boolean = False
    Dim resultWorkbook As SpreadsheetGear.IWorkbook = Nothing
    Dim rlStageDesc As New Dictionary(Of Integer, String)
    Dim rlStageUrls As New Dictionary(Of Integer, String)
    Dim rlStageComplete As New Dictionary(Of Integer, Boolean)
    Dim resultPageLinks As New Dictionary(Of Integer, String)
    Dim resultPagesProcessed As Integer = 0
    Dim htmlResultsTable As String = ""
    Dim rlCurrentStage As Integer = 0
    Dim recordsImported As Integer = 0
    Dim frmLoginDefaultAction As String = ""
    Dim loginErrorMessage As String = ""
    Dim IMPORT_CTL_NO As String = ""
    Dim tblTATWMRLR_dups As DataTable = Nothing
    Dim attempts As Integer = 0



#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        rlStageDesc.Add(0, "Initializing Browser Control")
        rlStageDesc.Add(1, "Enter credentials and ship point to proceed")
        rlStageDesc.Add(2, "Logging Into Retail Link")
        rlStageDesc.Add(3, "Navigate to Routing Status")
        rlStageDesc.Add(4, "Populate Ship Point")
        rlStageDesc.Add(5, "Submit for Status")
        rlStageDesc.Add(6, "Export Results")
        rlStageDesc.Add(7, "Scraping")
        rlStageDesc.Add(8, "Scraping Complete")
        rlStageDesc.Add(9, "Import Routing Status")
        rlStageDesc.Add(10, "Import Finished")

        rlStageUrls.Add(0, "https://retaillink.wal-mart.com/")
        rlStageUrls.Add(1, "")
        rlStageUrls.Add(2, "https://retaillink.wal-mart.com/new_home/")
        rlStageUrls.Add(3, "https://retaillink.wal-mart.com/WTMS/RequestForRouting/View/RoutingStatusView.aspx")
        rlStageUrls.Add(4, "")
        rlStageUrls.Add(5, "")
        rlStageUrls.Add(6, "")
        rlStageUrls.Add(7, "")
        rlStageUrls.Add(8, "")
        rlStageUrls.Add(9, "")
        rlStageUrls.Add(10, "")

        For i As Integer = 0 To 10
            rlStageComplete.Add(i, False)
        Next

        With dst


            Create_TDA(.Tables.Add, "TATWMRLR", "*")


        End With

        Shell("RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 8", vbHide)

        grdTATWMRLR.DataSource = dst.Tables("TATWMRLR")

        grdTATWMRLR.DisplayLayout.Bands(0).Columns("STATUS").Header.Caption = "Status"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("SHIP_ON_DATE").Header.Caption = "Ship On Date"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("CARRIER_PU_DATE").Header.Caption = "Carrie PU Date"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("CARRIER_DUE_DATE").Header.Caption = "Carrie Due Date"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("CARRIER_NAME").Header.Caption = "Carrier Name"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("ROUTING_MODE").Header.Caption = "Mode"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("LOAD_DEST").Header.Caption = "Load dest"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("SHIPPOINT").Header.Caption = "Shippoint"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("LOAD_NO").Header.Caption = "Load #"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("ORDR_NO").Header.Caption = "Order No"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Header.Caption = "Order Group"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("ORDR_CUST_PO").Header.Caption = "PO #"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("CASES").Header.Caption = "Cases"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("WEIGHT").Header.Caption = "Weight"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("PALLETS").Header.Caption = "Palletts"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("CUBE").Header.Caption = "Cube"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("PO_TYPE").Header.Caption = "PO Type"
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("DEPT_NO").Header.Caption = "Dept #"
        grdTATWMRLR.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ExtendLastColumn
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("IMPORT_CTL_NO").Hidden = True
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = True
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Hidden = True
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("LAST_DATE").Hidden = True
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("LAST_USER").Hidden = True
        Me.wb.ObjectForScripting = Me
        Me.wb.Visible = False

        ASCMAIN1.Progress("Connecting to Retail Link...")

        ASCMAIN1.sql = "Select * from TATUSER1 where USER_ID = :PARM1"
        Dim rowTATUSER1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", ASCMAIN1.USER_ID)
        If rowTATUSER1 IsNot Nothing Then
            txtUserID.Text = rowTATUSER1.Item("RETAIL_LINK_USER_ID") & ""
            txtPassword.Text = rowTATUSER1.Item("RETAIL_LINK_PASSWORD") & ""
        End If
        Dim tblShipPoint As DataTable = ASCDATA1.GetDataTable("Select T_CODE SP_KEY, T_DESC SP_DESC from ASTCODE1 where COLUMN_NAME = 'RETAIL_LINK_ROUTING' and TABLE_NAME = 'TATWMRL1'")
        cbeShipPoint.DataSource = tblShipPoint
        If tblShipPoint.Rows.Count > 0 Then
            Dim defaultShipPoint As DataRow = tblShipPoint.Rows(0)
            cbeShipPoint.Value = defaultShipPoint.Item("SP_KEY")
        End If

        EntryMode = "E"

        Call Mode_Settings(True)


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
            Case "Get Routing Status"
                If txtUserID.Text = "" Then
                    EMsg = "Please enter a User ID"
                End If
                If txtPassword.Text = "" Then
                    EMsg = "Please enter a Password"
                End If
                If cbeShipPoint.Value & "" = "" Then
                    EMsg = "Please Select a Ship Point"
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

            Case "Get Routing Status"

                Me.Cursor = Cursors.Default
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Save My Credentials").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Save My Credentials").Visible = False
            .Groups("Screen Control").Items("Get Routing Status").Visible = False
            .Groups("Options").Visible = False
        End With

        SplitContainer1.Panel1Collapsed = True
        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()

            'If ASCMAIN1.DBS_COMPANY = "NYA" Then
            '    If Not wbcInitialized Then

            '    End If
            'End If
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("TATWMRLR").Rows.Clear()
        tblTATWMRLR_dups = Nothing
        resultPagesProcessed = 0
        recordsImported = 0
        htmlResultsTable = ""
        frmLoginDefaultAction = ""
        loggedIn = False
        onRoutingStatusPage = False
        routingStatusLoaded = False
        dst.EnforceConstraints = True
        Set_Current_Stage(0)

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
        Call Load_Popup_Menu(grdTATWMRLR, "SSB", "Show Filter", "Show GroupBox")
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

    Private Sub wb_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wb.DocumentCompleted

        Select Case rlCurrentStage
            Case 0
                Initialize_Browser_Control()
            Case 1
                DirectCast(wb, Control).Enabled = False
                Set_Current_Stage(2)
            Case 2
                If e.Url.AbsoluteUri = DirectCast(sender, WebBrowser).Url.AbsoluteUri Then
                    Set_Current_Stage(3)
                End If
            Case 3
                If e.Url.AbsoluteUri = DirectCast(sender, WebBrowser).Url.AbsoluteUri Then
                    Set_Current_Stage(4)
                    Populate_Ship_Point()
                End If
            Case 4

            Case 5
                Submit_For_Status()
            Case 6
                Export_Results()
            Case 7
                Scrape_Results_Table()
            Case 8
                Build_Results_Datatable()
            Case 9
            Case 10
        End Select

    End Sub
    Sub Initialize_Browser_Control()
        Dim browserInitialized As Boolean = False

        Dim maxAttempts As Integer = 5000
        Inject_Login_Scripts()

        Dim instructions As HtmlElement = wb.Document.GetElementById("directions")
        If instructions IsNot Nothing Then
            instructions.InnerText = "Enter your User ID, Password and Ship Point below."
        End If
        Dim lblShipPoint As HtmlElement = wb.Document.GetElementById("isOnlyForgotPwd")
        If lblShipPoint IsNot Nothing Then
            Dim absHtml As String = "<div id='absLogin' " &
                            " style='float:right;margin-top:20px;color: #FFFFFF;background-color: #2196F3;border-style: solid;" &
                            " border-radius: 5px;border-color: #2196F3;border-width: 1px;padding: 5px 10px;width: 182px;cursor: pointer;font-size: 12px;font-weight: 600;'" &
                            " onclick='absValidateLogin();' >" &
                            "Download Routing Status</div>"
            lblShipPoint.InnerHtml = "<input name=""txtShip"" type=""text"" id=""txtShip"" style=""width: 308px"" tabindex=""3"" autocomplete=""off"">" & absHtml
        End If
        If wb.Document.GetElementById("lblForgotPassLink") IsNot Nothing Then
            Dim lblForgotPW As HtmlElement = wb.Document.GetElementById("lblForgotPassLink")
            If lblForgotPW IsNot Nothing Then
                Dim styleGenerator As StyleGenerator = New StyleGenerator()
                styleGenerator.ParseStyleString(IIf(IsNothing(lblForgotPW.Style), "", lblForgotPW.Style))
                styleGenerator.SetStyle("visibility", "hidden")
                lblForgotPW.Style = styleGenerator.GetStyleString()

            End If
        End If
        Dim registerForm As HtmlElement = wb.Document.GetElementById("lblRegister").Parent().Parent().Parent()
        If registerForm IsNot Nothing Then
            Dim styleGenerator As StyleGenerator = New StyleGenerator()
            styleGenerator.ParseStyleString(IIf(IsNothing(registerForm.Style), "", registerForm.Style))
            styleGenerator.SetStyle("visibility", "hidden")
            registerForm.Style = styleGenerator.GetStyleString()
        End If
        Dim lblScreenResolution As HtmlElement = wb.Document.GetElementById("lblScreenRes").Parent()
        If lblScreenResolution IsNot Nothing Then
            Dim styleGenerator As StyleGenerator = New StyleGenerator()
            styleGenerator.ParseStyleString(IIf(IsNothing(lblScreenResolution.Style), "", lblScreenResolution.Style))
            styleGenerator.SetStyle("visibility", "hidden")
            lblScreenResolution.Style = styleGenerator.GetStyleString()
        End If
        Dim inputElementCollection As HtmlElementCollection = wb.Document.GetElementsByTagName("input")
        For Each curElement As HtmlElement In inputElementCollection
            Dim elementName As String = curElement.GetAttribute("name").ToString
            If elementName = "txtUser" Then
                curElement.SetAttribute("value", txtUserID.Text)
            End If
            If elementName = "txtPass" Then
                curElement.SetAttribute("value", txtPassword.Text)
            End If
            If elementName = "txtShip" Then
                curElement.SetAttribute("value", cbeShipPoint.Value)
            End If
        Next

        If wb.Document.GetElementById("Login") IsNot Nothing Then
            Dim wmSubmitBtn As HtmlElement = wb.Document.GetElementById("Login")
            Dim styleGenerator As StyleGenerator = New StyleGenerator()
            styleGenerator.ParseStyleString(IIf(IsNothing(wmSubmitBtn.Style), "", wmSubmitBtn.Style))
            styleGenerator.SetStyle("visibility", "hidden")
            wmSubmitBtn.Style = styleGenerator.GetStyleString()
            Dim footerImages As HtmlElementCollection = wmSubmitBtn.Parent().Parent().GetElementsByTagName("img")
            For Each footerImage As HtmlElement In footerImages
                footerImage.SetAttribute("HEIGHT", "20")
            Next
            wb.Document.GetElementById("absLogin").Focus()
            browserInitialized = True
        Else
            attempts += 1
            If attempts <= maxAttempts Then
                Initialize_Browser_Control()
            Else
                browserInitialized = True
                Me.wb.Visible = True
            End If

        End If

        If browserInitialized Then
            Me.wb.Visible = True
            Set_Current_Stage(1)
        End If
    End Sub

    Sub Populate_Ship_Point()
        Dim shipPointSet As Boolean = False
        Dim hiddenValueSet As Boolean = False

        Dim theElementCollection As HtmlElementCollection = wb.Document.GetElementsByTagName("input")
        For Each curElement As HtmlElement In theElementCollection
            Dim elementName As String = curElement.GetAttribute("name").ToString
            Dim ii As Integer = InStr(elementName, "$txtShipPoint")
            If ii <> 0 Then
                curElement.SetAttribute("value", cbeShipPoint.Value)
                shipPointSet = True
            End If
            If elementName = "ctl00$ContentPlaceHolder1$hidValue" Then
                curElement.SetAttribute("value", "Add")
                hiddenValueSet = True
            End If
            If shipPointSet And hiddenValueSet Then
                Set_Current_Stage(5)
                If wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnAddItem") IsNot Nothing Then
                    wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnAddItem").Focus()
                    wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnAddItem").InvokeMember("click")
                End If
            End If
        Next
    End Sub
    Sub Submit_For_Status()
        If wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnSubmitForStatus") IsNot Nothing Then
            Set_Current_Stage(6)
            wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnSubmitForStatus").Focus()
            wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnSubmitForStatus").InvokeMember("click")
        End If
    End Sub
    Sub Export_Results()
        Dim pLinks As Integer = 0
        Dim gotLinks As Boolean = False
        resultPageLinks.Clear()

        Dim paginationTable As HtmlElement = wb.Document.GetElementById("Table2")
        For Each tableElement As HtmlElement In paginationTable.All
            If tableElement.TagName = "TR" Then
                Dim tRow As HtmlElement = tableElement.FirstChild
                For Each tdElement As HtmlElement In tRow.All
                    Dim tdElementName As String = tdElement.Id
                    If InStr(tdElementName, "_lnk") > 0 Then
                        resultPageLinks.Add(pLinks, tdElement.Id)
                        pLinks += 1
                        gotLinks = True
                    End If
                Next
                If gotLinks Then
                    Exit For
                End If
            End If
        Next

        Set_Current_Stage(7)
        htmlResultsTable = ""
        Scrape_Results_Table()

        'If wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnExportResults") IsNot Nothing Then

        '    'get result page links
        '    'for each <a></a> tag with id having instr _lnk
        '    wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnExportResults").Focus()
        '    'wb.Document.GetElementById("ctl00_ContentPlaceHolder1_btnExportResults").InvokeMember("click")
        '    Dim FILE_NAME As String = wb.Document.InvokeScript("fnaGrab")
        '    Debug.Print("Injected Function Result: " & FILE_NAME)
        'End If
    End Sub

    Sub Scrape_Results_Table()



        For i As Integer = resultPagesProcessed To resultPageLinks.Count - 1
            resultPagesProcessed += 1

            'get table data here
            Dim resultsTable As HtmlElement = wb.Document.GetElementById("ctl00_ContentPlaceHolder1_dgRoutingStatusResp")

            For Each tableElement As HtmlElement In resultsTable.All

                If tableElement.TagName = "TBODY" Then
                    For n As Integer = 0 To tableElement.Children.Count - 1
                        Dim headerRow As Boolean = (InStr(tableElement.Children(n).GetAttribute("className"), "th") > 0)
                        If resultPagesProcessed > 1 And headerRow Then
                            'do not concat header row
                        Else
                            Dim tableRow As String = tableElement.Children(n).OuterHtml.ToUpper
                            If resultPagesProcessed = 1 And headerRow Then
                                tableRow = Replace(Replace(tableRow, "<TD", "<TH"), "</TD", "</TH")
                            Else
                                recordsImported += 1
                            End If
                            htmlResultsTable &= tableRow
                        End If
                    Next
                End If

            Next

            Dim scrapeComplete As Boolean = (resultPagesProcessed = resultPageLinks.Count)
            Dim pbParm As String = Replace(resultPageLinks(IIf(scrapeComplete, resultPagesProcessed - 1, resultPagesProcessed)), "_", "$")
            If scrapeComplete Then
                Set_Current_Stage(8)
            End If

            wb.Document.InvokeScript("__doPostBack", New String() {pbParm, ""})
            Exit For

        Next


    End Sub

    Sub Build_Results_Datatable()
        Dim htmlTable As String = "<TABLE> " & htmlResultsTable & " </TABLE>"
        Dim tblRouting As DataTable = HTMLTable2DataTable(htmlTable)
        If tblRouting.Rows.Count > 0 Then
            IMPORT_CTL_NO = ASCMAIN1.Next_Control_No("TATWMRLR.IMPORT_CTL_NO")
            Dim i As Integer = 0
            For Each row As DataRow In tblRouting.Select("")
                i += 1
                Dim rowTATWMRLR As DataRow = Write_Routing_Record(row)
                dst.Tables("TATWMRLR").Rows.Add(rowTATWMRLR)
                'Create duplicate records for testing
                'If i = 5 Or i = 10 Then
                '    Dim rowTATWMRLR_dup As DataRow = Write_Routing_Record(row)
                '    dst.Tables("TATWMRLR").Rows.Add(rowTATWMRLR_dup)
                'End If
            Next

            grdTATWMRLR.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
            grdTATWMRLR.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

            SplitContainer1.Panel2Collapsed = True

            If Update_Oracle_Proceed() Then
                Update_Oracle()
                Display_Import_Results(IMPORT_CTL_NO)
            Else
                ASCMAIN1.Progress("Update Cancelled", "")
            End If


        Else
            MsgBox("No Records to Import", vbOKOnly + MsgBoxStyle.Information, "Finished")
        End If
    End Sub
    Function Write_Routing_Record(row As DataRow) As DataRow
        Dim rowTATWMRLR As DataRow = dst.Tables("TATWMRLR").NewRow
        rowTATWMRLR.Item("IMPORT_CTL_NO") = IMPORT_CTL_NO
        rowTATWMRLR.Item("ORDR_CUST_PO") = row.Item("ORDR_CUST_PO")
        rowTATWMRLR.Item("STATUS") = row.Item("STATUS")
        For Each DATE_COLUMN As String In New String() _
            {"SHIP_ON_DATE", "CARRIER_PU_DATE", "CARRIER_DUE_DATE"}
            If IsDate(row.Item(DATE_COLUMN)) Then
                rowTATWMRLR.Item(DATE_COLUMN) = row.Item(DATE_COLUMN)
            End If
        Next
        rowTATWMRLR.Item("CARRIER_NAME") = row.Item("CARRIER_NAME")
        rowTATWMRLR.Item("ROUTING_MODE") = row.Item("MODE")
        rowTATWMRLR.Item("LOAD_DEST") = row.Item("LOAD_DEST")
        rowTATWMRLR.Item("SHIPPOINT") = row.Item("SHIPPOINT")
        rowTATWMRLR.Item("LOAD_NO") = row.Item("LOAD_NO")
        rowTATWMRLR.Item("CASES") = row.Item("CASES")
        rowTATWMRLR.Item("WEIGHT") = row.Item("WEIGHT")
        rowTATWMRLR.Item("PALLETS") = row.Item("PALLETS")
        rowTATWMRLR.Item("CUBE") = row.Item("CUBE")
        rowTATWMRLR.Item("PO_TYPE") = row.Item("PO_TYPE")
        rowTATWMRLR.Item("DEPT_NO") = row.Item("DEPT_NO")
        rowTATWMRLR.Item("INIT_DATE") = DATETIME_STAMP
        rowTATWMRLR.Item("INIT_USER") = ASCMAIN1.USER_ID

        Return rowTATWMRLR

    End Function
    Sub Update_Oracle()
        ASCMAIN1.Progress("Updating Database", IMPORT_CTL_NO)
        Try
            BeginTrans()

            Update_Record_TDA("TATWMRLR")

            Dim sqlTATWMRLR As String = "BEGIN DECLARE CURSOR C1 IS" _
            & " SELECT SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_STATUS, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_ORIG_SHIP_DATE" _
            & " FROM SOTORDR1, (SELECT IMPORT_CTL_NO, ORDR_CUST_PO FROM TATWMRLR WHERE IMPORT_CTL_NO = '" & IMPORT_CTL_NO & "') TATWMRLR" _
            & " WHERE SOTORDR1.ORDR_CUST_PO = TATWMRLR.ORDR_CUST_PO" _
            & " AND SOTORDR1.CUST_CODE = 'WALMART'" _
            & " AND SOTORDR1.ORDR_STATUS IN ('P');" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " UPDATE TATWMRLR SET ORDR_NO = R1.ORDR_NO" _
            & " , ORDR_GROUP_NO = R1.ORDR_GROUP_NO" _
            & "  WHERE" _
            & " ORDR_CUST_PO = R1.ORDR_CUST_PO" _
            & " AND IMPORT_CTL_NO = '" & IMPORT_CTL_NO & "';" _
            & " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sqlTATWMRLR)

            Dim sqlSOTORDR1_ORIG_DATE As String = "UPDATE SOTORDR1 SET ORDR_ORIG_SHIP_DATE = ORDR_SHIP_DATE WHERE ORDR_NO IN (SELECT ORDR_NO FROM TATWMRLR WHERE IMPORT_CTL_NO = '" & IMPORT_CTL_NO & "' AND ORDR_NO IS NOT NULL)" _
            & " AND ORDR_ORIG_SHIP_DATE IS NULL"
            ASCDATA1.ExecuteSQL(sqlSOTORDR1_ORIG_DATE)

            Dim sqlSOTORDR1 As String = "BEGIN DECLARE CURSOR C1 IS" _
            & " SELECT ORDR_NO, ORDR_CUST_PO, SHIP_ON_DATE FROM TATWMRLR " _
            & " WHERE IMPORT_CTL_NO = '" & IMPORT_CTL_NO & "' " _
            & " AND ORDR_NO IS NOT NULL" _
            & " AND SHIP_ON_DATE IS NOT NULL;" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " UPDATE SOTORDR1 SET ORDR_SHIP_DATE = R1.SHIP_ON_DATE, ORDR_CANCEL_DATE = R1.SHIP_ON_DATE" _
            & " WHERE ORDR_NO = R1.ORDR_NO" _
            & " AND ORDR_CUST_PO = R1.ORDR_CUST_PO" _
            & " AND ORDR_STATUS IN ('P')" _
            & " AND CUST_CODE = 'WALMART';" _
            & " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sqlSOTORDR1)

            Dim sqlSOPORDR0_G As String = "BEGIN DECLARE CURSOR C1 IS" _
            & " SELECT DISTINCT ORDR_GROUP_NO FROM TATWMRLR WHERE IMPORT_CTL_NO = '" & IMPORT_CTL_NO & "' AND ORDR_GROUP_NO IS NOT NULL;" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " BEGIN " _
            & " SOPORDR0_G(R1.ORDR_GROUP_NO);" _
            & " END;" _
            & " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sqlSOPORDR0_G)


            Dim sqlSOTSHIP1 As String = "UPDATE SOTSHIP1 SET SHIP_XMIT_FLAG = 'U' WHERE" _
            & " ORDR_GROUP_NO IN (SELECT DISTINCT ORDR_GROUP_NO FROM TATWMRLR WHERE IMPORT_CTL_NO = '" & IMPORT_CTL_NO & "' AND ORDR_GROUP_NO IS NOT NULL)" _
            & " AND SHIP_STATUS = 'P'"
            ASCDATA1.ExecuteSQL(sqlSOTSHIP1)

            ASCMAIN1.Progress("Update Complete")
            CommitTrans(recordsImported & " Records Imported")

        Catch ex As Exception
            Rollback("Error Importing Routing Status: " & ex.Message)
        End Try


    End Sub

    Sub Display_Import_Results(IMPORT_CTL_NO As String)
        Fill_Records("TATWMRLR", "", True, "Select * from TATWMRLR Where IMPORT_CTL_NO = '" & IMPORT_CTL_NO & "'")
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = False
        grdTATWMRLR.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Hidden = False
    End Sub
    Private Function Download_Complete(FilePath As String) As Boolean
        Dim dlComplete As Boolean = False
        Try
            If File.Exists(FilePath) Then
                Using File.OpenRead(FilePath)
                    Return True
                End Using
            Else
                Return False
            End If
        Catch ex As Exception
            Thread.Sleep(1000)
        End Try
        Return dlComplete
    End Function
    Function FindDuplicateRows(dTable As DataTable, colName As String, Optional returnUnique As Boolean = False) As DataTable
        Dim hTable As New Hashtable()
        Dim duplicateList As New ArrayList()
        Dim returnTable As DataTable = Nothing


        'Add list of all the unique item value to hashtable, which stores combination of key, value pair.
        'And add duplicate item value in arraylist.
        For Each drow__1 As DataRow In dTable.Rows
            If hTable.Contains(drow__1(colName)) Then
                duplicateList.Add(drow__1)
            Else
                hTable.Add(drow__1(colName), String.Empty)
            End If
        Next

        If returnUnique Then
            returnTable = dTable
        Else
            returnTable = dTable.Clone
        End If

        'Removing a list of duplicate items from datatable.
        For Each dRow__2 As DataRow In duplicateList
            If returnUnique Then
                returnTable.Rows.Remove(dRow__2)
            Else
                For Each dupRec As DataRow In dTable.Select(colName & " = '" & dRow__2.Item(colName) & "'")
                    returnTable.ImportRow(dupRec)
                    'Dim dupRow As DataRow = returnTable.NewRow
                    'dupRow.ItemArray = dupRec.ItemArray.Clone()

                    'returnTable.Rows.Add(dupRec)
                Next
            End If
        Next

        'Datatable which contains unique records will be return as output.
        Return returnTable
    End Function

    Function Update_Oracle_Proceed() As Boolean
        Dim updateOracle As Boolean = True
        Dim tblImportResults As DataTable = dst.Tables("TATWMRLR")
        tblTATWMRLR_dups = FindDuplicateRows(tblImportResults, "ORDR_CUST_PO")

        If tblTATWMRLR_dups.Rows.Count > 0 Then
            Using fr As New ASFMSGBF
                'fr.Text = "Click OK to proced with update."
                'fr.Controls("cmdOK").Text = "Proceed with update."
                fr.Show_grd(tblTATWMRLR_dups, Me, "The following duplicate records have been identified")
                If fr.user_option = -1 Then
                    updateOracle = False
                End If
            End Using
        End If


        Return updateOracle
    End Function
    Sub Inject_Login_Scripts()
        Dim pageHeadElement As HtmlElement = wb.Document.GetElementsByTagName("head")(0)
        Dim injectedScriptElement As HtmlElement = wb.Document.CreateElement("script")
        Dim injectedFunction As String = "function absValidateLogin() {" &
        " var userValid = $('#txtUser')[0].value != ''; " &
        " var passValid = $('#txtPass')[0].value != ''; " &
        " var shipValid = $('#txtShip')[0].value != ''; " &
        " if (userValid  && passValid && shipValid ){ " &
        "   $('#absLogin').html('Authenticating...'); " &
        "   $('#Login').click(); " &
        " } else { " &
        "   var invalidMsg = ''; " &
        "   invalidMsg += (userValid)?'':'User ID is required. \n'; " &
        "   invalidMsg += (passValid)?'':'Password is required. \n'; " &
        "   invalidMsg += (shipValid)?'':'Ship Point is required.'; " &
        "   alert(invalidMsg); " &
        "   return false; " &
        " } " &
        "} "

        injectedScriptElement.SetAttribute("text", injectedFunction)
        pageHeadElement.AppendChild(injectedScriptElement)
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
        If Table.Value.Contains("<TH") Then
            ' Set the HeadersExist flag  
            HeadersExist = True

            ' Get a match for all the rows in the table  
            Dim Headers As MatchCollection = Regex.Matches(Table.Value, "<TH[^>]*>(.*?)</TH>", RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)


            ' Loop through each header element  
            For Each Header As Match In Headers
                Dim cn As String = Replace(Replace(Header.Groups(1).ToString, " ", "_"), "#", "NO").ToUpper
                If cn = "PO_NO" Then
                    cn = "ORDR_CUST_PO"
                End If
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
                    If InStr(Column.Groups(1).ToString, "<A") > 0 Then
                        Dim cellData As MatchCollection = Regex.Matches(Column.Groups(1).Value, "<A[^>]*>(.*?)</A>", RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)
                        dr(iCurrentColumn) = cellData.Item(0).Groups(1).ToString
                    Else
                        dr(iCurrentColumn) = Column.Groups(1).ToString
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

    Private Sub wb_Navigated(sender As Object, e As WebBrowserNavigatedEventArgs) Handles wb.Navigated
        Dim uriResult As String = e.Url.AbsoluteUri
        Select Case rlCurrentStage
            Case 0
            Case 1
                If InStr(uriResult, "BAD_PWD_OR_USER") > 0 Then
                    loginErrorMessage = "Invalid User ID or Password"
                    Set_Current_Stage(0)
                End If
            Case 2
            Case 3
            Case 4
            Case 5
                Dim s As String = ""
            Case 6
                Dim s As String = ""
            Case 7
            Case 8
            Case 9
            Case 10
        End Select
    End Sub

    Private Sub wb_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs) Handles wb.Navigating

        ASCMAIN1.Progress(rlStageDesc(rlCurrentStage))

        Select Case rlCurrentStage
            Case 0
            Case 1
            Case 2
            Case 3
            Case 4
            Case 5
                Dim s As String = ""
            Case 6
                Dim S As String = ""
            Case 7
            Case 8
            Case 9
            Case 10

        End Select

    End Sub

    Sub Update_Stage_Status(stage As Integer, complete As Boolean)
        If rlStageComplete.ContainsKey(stage) Then
            rlStageComplete.Remove(stage)
            rlStageComplete.Add(stage, complete)
        End If
    End Sub
    Sub Set_Current_Stage(newStage As Integer)
        rlCurrentStage = newStage
        If rlStageDesc.ContainsKey(newStage) Then
            ASCMAIN1.Progress(rlStageDesc(newStage))
        End If
        If newStage = 0 Then
            For i As Integer = 0 To 10
                Update_Stage_Status(i, False)
            Next
        Else
            Update_Stage_Status(newStage - 1, True)
        End If
        If rlStageUrls.ContainsKey(newStage) Then
            Dim newUrl As String = rlStageUrls(newStage)
            Dim forceCleanSession As String = IIf(newStage = 0, "?refresh=" & Guid.NewGuid().ToString(), "")
            If newUrl <> "" Then
                Try
                    If newStage = 0 Then
                        InternetSetOption(IntPtr.Zero, INTERNET_OPTION_END_BROWSER_SESSION, IntPtr.Zero, 0)
                    End If
                    Dim browserURL As String = newUrl & forceCleanSession
                    wb.Navigate(New Uri(browserURL))
                Catch ex As Exception
                    Dim fubar As String = ""
                End Try
            End If

        End If

    End Sub

    Private Sub grdRoutingStatus_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdTATWMRLR.InitializeLayout
        e.Layout.Override.ColumnAutoSizeMode = UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand


    End Sub

    <DllImport("wininet.dll", SetLastError:=True)> _
    Private Shared Function InternetSetOption(hInternet As IntPtr, dwOption As Integer, lpBuffer As IntPtr, lpdwBufferLength As Integer) As Boolean
    End Function

    Private Sub grdRoutingStatus_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdTATWMRLR.InitializeRow
        Dim dupCheckSql As String = "IMPORT_CTL_NO = '" & IMPORT_CTL_NO & "' AND ORDR_CUST_PO = '" & e.Row.Cells("ORDR_CUST_PO").Value & "'"
        If tblTATWMRLR_dups IsNot Nothing AndAlso tblTATWMRLR_dups.Select(dupCheckSql).Length > 0 Then
            'e.Row.Cells("ORDR_CUST_PO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub
End Class

Public Class StyleGenerator
    Dim styleDB As Dictionary(Of String, String)

    Public Sub New()
        styleDB = New Dictionary(Of String, String)()
    End Sub


    Public Function ContainsStyle(ByVal name As String) As Boolean
        Return styleDB.ContainsKey(name)
    End Function


    Public Function SetStyle(ByVal name As String, ByVal value As String) As String
        Dim oldValue As String = ""

        If (Not name.Length > 0) Then
            Throw New ArgumentException("Parameter name cannot be zero-length.")
        End If
        If (Not value.Length > 0) Then
            Throw New ArgumentException("Parameter value cannot be zero-length.")
        End If

        If (styleDB.ContainsKey(name)) Then
            oldValue = styleDB(name)
        End If

        styleDB(name) = value

        Return oldValue
    End Function

    Public Function GetStyle(ByVal name As String) As String
        If (Not name.Length > 0) Then
            Throw New ArgumentException("Parameter name cannot be zero-length.")
        End If

        If (styleDB.ContainsKey(name)) Then
            Return styleDB(name)
        Else
            Return ""
        End If
    End Function

    Public Sub RemoveStyle(ByVal name As String)
        If (styleDB.ContainsKey(name)) Then
            styleDB.Remove(name)
        End If
    End Sub

    Public Function GetStyleString() As String
        If (styleDB.Count > 0) Then
            Dim styleString As New System.Text.StringBuilder("")
            Dim key As String
            For Each key In styleDB.Keys
                styleString.Append(String.Format("{0}:{1};", CType(key, Object), CType(styleDB(key), Object)))
            Next key

            Return styleString.ToString()
        Else
            Return ""
        End If
    End Function

    Public Sub ParseStyleString(ByVal styles As String)
        If (styles.Length) > 0 Then
            Dim stylePairs As String() = styles.Split(New Char() {";"c})
            Dim stylePair As String
            For Each stylePair In stylePairs
                If (stylePairs.Length > 0) Then
                    Dim styleNameValue As String() = stylePair.Split(New Char() {":"c})
                    If (styleNameValue.Length = 2) Then
                        styleDB(styleNameValue(0)) = styleNameValue(1)
                    End If
                End If
            Next stylePair
        End If
    End Sub


    Public Sub Clear()
        styleDB.Clear()
    End Sub
End Class