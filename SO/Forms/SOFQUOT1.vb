Imports System.Drawing
Imports System.Xml
Imports Microsoft.Office.Interop

Public Class SOFQUOT1
    Dim Remote As New REMOTE(Me)
    Private mExcelProcesses() As Process
    Dim ExcelFolder As String = GetExcelFolder()
    Dim rowSOTORDR2_Current As DataRow
    Dim STYLE_CODE_Current As String
    Dim ExtOptions As String() = New String() {".XLS", ".XLSX", ".XLSM"}
#Region "ABS Standard Routines"
    ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Get_PARM("SOTPARM1")
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst
            ASCMAIN1.sql = "SELECT * FROM SOTQUOT2"
            Create_TDA(.Tables.Add, "SOTQUOT2", "**", 0, True, "", 1)
            'Create_TDA(.Tables.Add, "SOTQUOT2", "**", 1, True)
            Fill_Records("SOTQUOT2", "", , ASCMAIN1.sql)

            For Each TABLE_INSERT As String In New String() {"SOTORDR1", "SOTORDR2", "SOTORDR5", "ICTSTYL1"}
                SQLs.Length = 0
                SQLs.AppendLine("SELECT * FROM SOTQUOT1 WHERE TABLE_NAME = :PARM1")
                ASCMAIN1.sql = SQLs.ToString
                Create_TDA(.Tables.Add, TABLE_INSERT, "**", 0, False, "V", 0)
                .Tables(TABLE_INSERT).Columns.Add("COPY_BTN", GetType(System.String))
                Fill_Records(TABLE_INSERT, TABLE_INSERT, , ASCMAIN1.sql)
                For Each rowTABLE As DataRow In dst.Tables(TABLE_INSERT).Select()
                    rowTABLE.Item("COPY_BTN") = "X"
                Next
            Next

            Dim SQLW As String = ""
            If REMOTE.SQLWhere.Length > 0 Then
                SQLW = " Where " & REMOTE.SQLWhere
                SQLW += String.Format(" AND ORDR_DATE >= '{0}'", Format(Now.AddMonths(-1), "dd-MMM-yy"))
            Else
                SQLW = String.Format(" Where ORDR_DATE >= '{0}'", Format(Now.AddMonths(-1), "dd-MMM-yy"))
            End If
            SQLW += " OR ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1_L)"
            ASCMAIN1.sql = "SELECT SOTORDR1.*,  TO_CHAR(ORDR_DATE, 'YYYY') AS YEAR FROM SOTORDR1 " & SQLW
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False)
            .Tables("SOTORDRX").Columns.Add("TORDR", GetType(System.String))
            Fill_Records("SOTORDRX", "", , ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDX1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR2 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDX2", "**", 0, False, "V", 2)
            .Tables("SOTORDX2").Columns.Add("FACTORY_CODE", GetType(System.String))
            .Tables("SOTORDX2").Columns.Add("UPC_CODE", GetType(System.String))

            ASCMAIN1.sql = "SELECT * FROM SOTORDR5 where CUST_ADDR_TYPE = 'ST' AND ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDX5", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1 where STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTYX1", "**", 0, False, "V", 1)
            .Tables("ICTSTYX1").Columns.Add("PORT_CODE", GetType(System.String))

        End With

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        Sort_grdColumns(grdSOTORDR1, "COLUMN_DESC".ToUpper, True)

        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        Sort_grdColumns(grdSOTORDR2, "COLUMN_DESC".ToUpper, True)

        grdSOTORDR5.DataSource = dst.Tables("SOTORDR5")
        Sort_grdColumns(grdSOTORDR5, "COLUMN_DESC".ToUpper, True)

        grdICTSTYL1.DataSource = dst.Tables("ICTSTYl1")
        Sort_grdColumns(grdICTSTYL1, "COLUMN_DESC".ToUpper, True)

        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        Sort_grdColumns(grdSOTORDRX, "ORDR_DATE, ORDR_GROUP_NO, ORDR_NO".ToLower(), False)
        CalculateOrderTotalX()

        RefreshSOTQUOT2()

        Create_Summary(grdSOTORDRX, "TORDR", "Sum", "", "###,##0.00")
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("TORDR").Format = "###,##0.00"
        ASCMAIN1.Add_Value_List(grdSOTORDRX, "ORDR_STATUS", , New String() {":", "L:Laptop", "Q:Quote", "C:Cancelled"})
        'grdSOTORDRX.Parent = tab.Parent

        'tab.Visible = False
        'WebBrowser1.Navigate("C:\Users\Wayne\Dropbox\Regency\Quote Templates From Danny\MASTER REGENCY QUOTE SHEET.xls", False)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Done"
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If
        Call Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)
        Select Case eItemKey
            Case "Cancel"
                Call Mode_Settings(False)
            Case "Done"
                Update_Record()
                Call Mode_Settings(False)
        End Select

    End Sub

    Private Sub ShowMaintGrids(ByVal ShowControl As Boolean)
        If ShowControl Then
            SplitContainer3.Panel2.Show()
        Else
            SplitContainer3.Panel2.Hide()
        End If
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                '.Groups("Allocation").Visible = False
                '.Groups("Image").Visible = False
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            'With grdARTPYMT2.DisplayLayout.Override
            '    If This_Record_Inquiry_Only Then
            '        .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '        .AllowDelete = DefaultableBoolean.False
            '        .AllowUpdate = DefaultableBoolean.False
            '    Else
            '        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            '        .AllowDelete = DefaultableBoolean.True
            '        .AllowUpdate = DefaultableBoolean.True
            '    End If
            'End With
            Clear_Record()
        Else
            Clear_Record()
        End If

        lblMessage.Visible = False
        ShowMaintGrids(False)

        btnImportNewFile.Visible = True
        btnSaveQuote.Visible = False

        lblQUOTE_NO.Visible = False
        txtQUOTE_NO.Visible = False

        lblQUOTE_NAME.Visible = False
        txtQUOTE_NAME.Visible = False
        btnSaveQuote.Visible = False

        grdSOTQUOT2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

    End Sub

    Sub Clear_Record()
        'txtUPC_CODE.Text = ""
        'lblUPC_CODE.Text = "UPC Code"
        txtQUOTE_NO.Text = ""
        txtQUOTE_NAME.Text = ""
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)
        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
        'ClearStyle()
    End Sub

    Sub Delete_Record()
        'Call BeginTrans()
        'Call Delete_Rows("SOTORDR1")
        'Call CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("SOTQUOT2")
        CommitTrans()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record()
    End Sub

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDRX, "SSB", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTQUOT2, "SSB", "Show Filter", "Show GroupBox", "Design This Template")
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

        Select Case e.Tool.Key
            Case "Design This Template"
                Dim QUOTE_NO As String = grd.ActiveRow.Cells("QUOTE_NO").Text
                If QUOTE_NO.Length <> 0 Then
                    OpenExcelToWork(QUOTE_NO)
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                EnforceConstraints(False)
                Fill_Records("ICTSTYL1", Absx1.txtFor("STYLE_CODE").Text, True)
                EnforceConstraints(True)
                If dst.Tables("ICTSTYL1").Rows.Count() = 0 Then
                    MsgBox("Style Not Found In Masterfile", MsgBoxStyle.Critical, "Invalid Style")
                    'ClearStyle()
                    Exit Sub
                End If
        End Select
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

    Private Sub grdClickGrid_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles _
        grdSOTORDR1.ClickCellButton, _
        grdSOTORDR2.ClickCellButton, _
        grdSOTORDR5.ClickCellButton, _
        grdICTSTYL1.ClickCellButton
        Dim CopyText As String = ""
        Dim ClickGrid As Infragistics.Win.UltraWinGrid.UltraGrid = sender
        CopyText = String.Format("<<{0}.{1}>>", ClickGrid.ActiveRow.Cells.Item("TABLE_NAME").Text, ClickGrid.ActiveRow.Cells.Item("COLUMN_NAME").Text)
        ClickGrid.ActiveRow.Cells.Item("COPY_BTN").Value = "O"
        Clipboard.SetText(CopyText)
        lblMessage.Text = CopyText & _
            " has been copied to your clipboard." & _
            vbCrLf & "Paste it into the open excel cell" & _
            vbCrLf & " That you want it to appear in."
        lblMessage.Visible = True
    End Sub

    Private Sub grdClickGrid_MouseLeave(sender As Object, e As System.EventArgs) Handles _
        grdSOTORDR1.MouseLeave, _
        grdSOTORDR2.MouseLeave, _
        grdSOTORDR5.MouseLeave, _
        grdICTSTYL1.MouseLeave
        lblMessage.Visible = False
    End Sub

    Private Sub btnImportNewFile_Click(sender As System.Object, e As System.EventArgs) Handles btnImportNewFile.Click
        OpenFileDialog1.DefaultExt = "xls"
        OpenFileDialog1.ShowDialog()
        Dim BadFileName As Boolean = True
        Dim ThisFileExt As String = ""
        If OpenFileDialog1.FileNames.Length = 1 Then
            Dim FullFileName As String = OpenFileDialog1.FileNames(0)
            Dim SafeFileName As String = OpenFileDialog1.SafeFileName
            If SafeFileName.Length >= 4 Then
                For Each FileExt As String In ExtOptions
                    If SafeFileName.ToUpper.EndsWith(FileExt) Then
                        BadFileName = False
                        ThisFileExt = FileExt
                    End If
                Next
            End If
            If Not BadFileName Then
                Dim QUOTE_NO As String = ASCMAIN1.Next_Control_No("SOFQUOT2.QUOTE_NO")
                Dim NextExcelName As String = QUOTE_NO & ThisFileExt
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Copy File"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine("This Will Copy the Following File to The Following Destination:")
                iMSG.AppendLine("Source: " & FullFileName)
                iMSG.AppendLine(String.Format("Destination: {0}{1}", ExcelFolder, NextExcelName))
                iMSG.AppendLine("Make Sure The File Is Saved And Ready To Copy.")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    If FileIO.FileSystem.FileExists(FullFileName) Then
                        FileIO.FileSystem.CopyFile(FullFileName, String.Format("{0}{1}", ExcelFolder, NextExcelName))
                        Dim newSOTQUOT2 As DataRow = dst.Tables("SOTQUOT2").NewRow
                        newSOTQUOT2.Item("QUOTE_NO") = QUOTE_NO
                        newSOTQUOT2.Item("QUOTE_NAME") = NextExcelName
                        newSOTQUOT2.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        newSOTQUOT2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        newSOTQUOT2.Item("INIT_DATE") = Now()
                        newSOTQUOT2.Item("LAST_DATE") = Now()
                        dst.Tables.Item("SOTQUOT2").Rows.Add(newSOTQUOT2)
                        Call Update_Record_TDA("SOTQUOT2")
                        Dim QuoteFound As Boolean = False
                        Dim QUOTE_NAME As String = ""
                        For Each rowSOTQUOT2 As DataRow In dst.Tables("SOTQUOT2").Select(String.Format("QUOTE_NO = '{0}'", QUOTE_NO))
                            If Not QuoteFound Then
                                QUOTE_NAME = rowSOTQUOT2.Item("QUOTE_NAME")
                                QuoteFound = True
                            End If
                        Next
                        If QuoteFound Then
                            txtQUOTE_NO.Text = QUOTE_NO
                            txtQUOTE_NAME.Text = QUOTE_NAME
                            txtQUOTE_NO.Visible = True
                            txtQUOTE_NAME.Visible = True
                            btnSaveQuote.Visible = True
                        End If
                        'OpenExcelToWork(QUOTE_NO)
                        'ShowMaintGrids(True)
                    End If
                End If
            Else
                MsgBox("Bad File Name (Probably The Extension)", MsgBoxStyle.Critical, "Error Importing")
            End If
        End If
    End Sub

    Private Function GetExcelFolder() As String
        Dim RetVal As String = ASCMAIN1.Folders("Work")
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        If Not IsNothing(rowSOTPARM3) Then
            RetVal = rowSOTPARM3.Item("RO_PARM_EXCEL_DIR").ToString
        End If
        If RetVal.Length > 0 Then
            If RetVal.Substring(RetVal.Length - 1, 1) <> "\" Then
                RetVal = RetVal & "\"
            End If
        End If
        Return RetVal
    End Function

    Private Sub RefreshSOTQUOT2()
        grdSOTQUOT2.DataSource = dst.Tables("SOTQUOT2")
        Sort_grdColumns(grdSOTQUOT2, "QUOTE_NO".ToLower(), False)
    End Sub

    Private Sub grdSOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow
        With e.Row
            Dim ORDR_SOURCE As String = .Cells("ORDR_SOURCE").Value & ""
            Dim ORDR_STATUS As String = .Cells("ORDR_STATUS").Value & ""
            Dim ORDR_GROUP_NO As String = .Cells("ORDR_GROUP_NO").Value & ""
            If ORDR_GROUP_NO.Length = 0 Then
                .Cells("ORDR_GROUP_NO").Value = .Cells("ORDR_BATCH_NO").Value
            End If
            If ORDR_STATUS = "L" Then
                Select Case ORDR_SOURCE
                    Case "L"
                        .Appearance.BackColor = Drawing.Color.Empty
                    Case "Q"
                        .Appearance.BackColor = Drawing.Color.BlanchedAlmond
                    Case Else
                        .Appearance.BackColor = Drawing.Color.Cyan
                End Select
            Else
                If ORDR_STATUS = "Q" Then
                    .Appearance.BackColor = Drawing.Color.BlanchedAlmond
                Else
                    .Appearance.BackColor = Drawing.Color.Cyan
                End If
            End If
            'CalculateOrderTotal(.Cells("ORDR_NO").Value)
            grdSOTORDRX.UpdateData()
        End With
    End Sub

    Private Sub CalculateOrderTotalX()
        For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select()
            ASCMAIN1.sql = String.Format("select sum(nvl(ordr_unit_price,0) * nvl(ordr_qty,0)) from sotordr2 where ordr_no = '{0}'", rowSOTORDRX.Item("ORDR_NO"))
            rowSOTORDRX.Item("TORDR") = Format(Val(ASCDATA1.GetDataValue), "###,##0.00")
        Next
    End Sub

    Private Sub OpenExcelToWork(QUOTE_NO As String)

        Dim QuoteFound As Boolean = False
        Dim QUOTE_NAME As String = ""
        For Each rowSOTQUOT2 As DataRow In dst.Tables("SOTQUOT2").Select(String.Format("QUOTE_NO = '{0}'", QUOTE_NO))
            If Not QuoteFound Then
                QUOTE_NAME = rowSOTQUOT2.Item("QUOTE_NAME")
                QuoteFound = True
            End If
        Next
        If QuoteFound Then
            Dim ThisFileExt As String = GetExtForFile(QUOTE_NO)
            If ThisFileExt.Length <> 0 Then
                If Open_Excel(QUOTE_NO) Then
                    txtQUOTE_NO.Text = QUOTE_NO
                    txtQUOTE_NAME.Text = QUOTE_NAME
                    btnImportNewFile.Visible = False

                    lblQUOTE_NO.Visible = True
                    txtQUOTE_NO.Visible = True

                    lblQUOTE_NAME.Visible = True
                    txtQUOTE_NAME.Visible = True
                    btnSaveQuote.Visible = True
                    lblMessage.Text = "The Selected Quote Has been opened." & _
                        vbCrLf & "When You Are Finished Editing Don't Forget" & _
                        vbCrLf & "To Save It!"
                    lblMessage.Visible = True
                    ShowMaintGrids(True)
                    tabQUOTES.Tabs(1).Selected = True
                End If
            Else
                MsgBox("Bad Or Missing File Extension In File Name " & QUOTE_NO & ThisFileExt, MsgBoxStyle.Critical, "File Name")
            End If
        End If
    End Sub

    Private Sub btnSaveQuote_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveQuote.Click
        Dim QuoteFound As Boolean = False
        Dim QUOTE_NO As String = txtQUOTE_NO.Text
        If txtQUOTE_NO.Text.Length > 0 Then
            For Each rowSOTQUOT2 As DataRow In dst.Tables("SOTQUOT2").Select(String.Format("QUOTE_NO = '{0}'", QUOTE_NO))
                If Not QuoteFound Then
                    rowSOTQUOT2.Item("QUOTE_NAME") = txtQUOTE_NAME.Text
                    rowSOTQUOT2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    rowSOTQUOT2.Item("LAST_DATE") = Now()
                    QuoteFound = True
                End If
            Next
            If QuoteFound Then
                Call Update_Record_TDA("SOTQUOT2")
                btnSaveQuote.Visible = False
                btnImportNewFile.Visible = True

                lblQUOTE_NO.Visible = False
                txtQUOTE_NO.Visible = False

                lblQUOTE_NAME.Visible = False
                txtQUOTE_NAME.Visible = False
                btnSaveQuote.Visible = False
                ShowMaintGrids(False)
                Dim msg As String = "Database Entries Have Been Saved." & _
                    vbCrLf & "Don't Forget To Save Any Open Excel Workbooks" & _
                    vbCrLf & "You May Have Been Working On"
                MsgBox(msg, vbOKOnly, "Save Your Excel Work")
            Else
                MsgBox(String.Format("Quote No{0} Not Found", QUOTE_NO), vbExclamation, "Invalid Quote")
            End If
        End If
    End Sub

    Private Sub ExcelProcessInit()
        Try
            'Get all currently running process Ids for Excel applications
            mExcelProcesses = Process.GetProcessesByName("Excel")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ExcelProcessKill()
        Dim oProcesses() As Process
        Dim bFound As Boolean

        Try
            'Get all currently running process Ids for Excel applications
            oProcesses = Process.GetProcessesByName("Excel")

            If oProcesses.Length > 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    bFound = False

                    For j As Integer = 0 To mExcelProcesses.Length - 1
                        If oProcesses(i).Id = mExcelProcesses(j).Id Then
                            bFound = True
                            Exit For
                        End If
                    Next

                    If Not bFound Then
                        oProcesses(i).Kill()
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Function Open_Excel(ByVal QUOTE_NO As String) As Boolean
        Dim iResult As Boolean = False
        Dim FILE_NAME As String = ""
        Dim FILE_EXT As String = GetExtForFile(QUOTE_NO)
        If FILE_EXT.Length > 0 Then
            Dim FullExcelName As String = String.Format("{0}{1}{2}", ExcelFolder, QUOTE_NO, FILE_EXT)
            Dim FileExists As Boolean = False
            If Not FileIO.FileSystem.FileExists(FullExcelName) Then
                FullExcelName = String.Format("{0}{1}{2}", ExcelFolder, QUOTE_NO, FILE_EXT)
                If FileIO.FileSystem.FileExists(FullExcelName) Then
                    FileExists = True
                End If
            Else
                FileExists = True
            End If
            If Not FileExists Then
                MsgBox("The Selected Template Is Not Found In The Excel Folder!", MsgBoxStyle.Critical, "File Error")
                iResult = False
            Else
                Dim excel As Excel.Application = New Microsoft.Office.Interop.Excel.Application
                Try
                    Dim XWB As Excel.Workbook = excel.Workbooks.Open(FullExcelName)
                    excel.Visible = True
                    iResult = True
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Opening File")
                    iResult = False
                End Try
                excel = Nothing
            End If
        Else
            MsgBox(String.Format("Quote No{0} Not Found", QUOTE_NO), vbExclamation, "Invalid Quote")
        End If
        Return iResult
    End Function

    Private Sub GenerateQuoteStandard()
        Dim FileExists As Boolean = False
        Dim ErrMsg As String = ""
        Dim NewSheetName As String = txtNewSheetName.Text
        Dim TemplateName As String = ""
        Dim TemplateExt As String = "xls"
        Dim XWB As Excel.Workbook
        Dim XWS As Excel.Worksheet
        Dim XWS_NEW As Excel.Worksheet
        If grdSOTORDRX.Selected.Rows.Count <> 1 Or grdSOTQUOT2.Selected.Rows.Count <> 1 Then
            ErrMsg = ErrMsg & vbCrLf & "You Must Select One Order And One Template To Proceed."
        End If
        Dim QUOTE_NO As String = grdSOTQUOT2.ActiveRow.Cells.Item("QUOTE_NO").Text
        Dim ORDR_NO As String = grdSOTORDRX.ActiveRow.Cells.Item("ORDR_NO").Text
        TemplateName = String.Format("{0}{1}.{2}", ExcelFolder, QUOTE_NO, TemplateExt)
        If FileIO.FileSystem.FileExists(TemplateName) Then
            FileExists = True
        Else
            TemplateExt = "xlsx"
            TemplateName = String.Format("{0}{1}.{2}", ExcelFolder, QUOTE_NO, TemplateExt)
            If FileIO.FileSystem.FileExists(TemplateName) Then
                FileExists = True
            Else
                TemplateExt = "xlsm"
                TemplateName = String.Format("{0}{1}.{2}", ExcelFolder, QUOTE_NO, TemplateExt)
                If FileIO.FileSystem.FileExists(TemplateName) Then
                    FileExists = True
                End If
            End If
        End If
        If Not FileExists Then
            ErrMsg = ErrMsg & vbCrLf & "The Selected Template Does Not" & vbCrLf & "Exist at " & ExcelFolder
        End If
        If NewSheetName.Length = 0 Then
            ErrMsg = ErrMsg & vbCrLf & "You Must Specify The Name Of The Sheet To Create."
        Else
            If NewSheetName.EndsWith(".xlsm") Then
                NewSheetName = NewSheetName.Replace(".xlsm", "")
            End If
            If NewSheetName.EndsWith(".xlsx") Then
                NewSheetName = NewSheetName.Replace(".xlsx", "")
            End If
            If NewSheetName.EndsWith(".xls") Then
                NewSheetName = NewSheetName.Replace(".xls", "")
            End If
            NewSheetName = String.Format("{0}{1}.{2}", ExcelFolder, NewSheetName, TemplateExt)
            If FileIO.FileSystem.FileExists(NewSheetName) Then
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Destination Already Exists"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine("The File Name You Selected Already Exists:")
                iMSG.AppendLine(NewSheetName)
                iMSG.AppendLine("Would You Like To Replace It?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    FileIO.FileSystem.DeleteFile(NewSheetName)
                Else
                    ErrMsg = ErrMsg & vbCrLf & "Please Pick A New File Name."
                End If
            End If
        End If
        If ErrMsg.Length <> 0 Then
            MsgBox(ErrMsg, MsgBoxStyle.OkOnly, "Problem Creating Spreadsheet")
            Exit Sub
        End If
        FileIO.FileSystem.CopyFile(TemplateName, NewSheetName)
        Dim excel As Excel.Application = New Microsoft.Office.Interop.Excel.Application
        Try
            XWB = excel.Workbooks.Open(NewSheetName)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Opening File")
            excel = Nothing
            Exit Sub
        End Try
        If XWB.Worksheets.Count > 1 Then
            Dim TotalBooks As Integer = XWB.Worksheets.Count
            For i As Integer = TotalBooks To 2 Step -1
                XWS = XWB.Worksheets(i)
                XWS.Delete()
            Next
        End If
        Dim SheetCount As Integer = 1
        XWS = XWB.Worksheets(1)
        XWS_NEW = XWB.Worksheets(1)
        'Dim xRow As Integer = 1
        'Dim xCol As Integer = 1
        Fill_Records("SOTORDX1", ORDR_NO, True)
        Fill_Records("SOTORDX2", ORDR_NO, True)
        Fill_Records("SOTORDX5", ORDR_NO, True)
        For Each rowSOTORDX2 As DataRow In dst.Tables("SOTORDX2").Select()
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine(String.Format("Select VEND_CODE from ICTSTYL1 where STYLE_CODE = '{0}'", rowSOTORDX2.Item("STYLE_CODE")))
            ASCMAIN1.sql = SQLS.ToString()
            Dim VEND_CODE As String = ASCDATA1.GetDataValue
            Dim FactoryCode As String = GetVendorData(VEND_CODE, "VEND_SUPPLIER_ID")
            rowSOTORDX2.Item("FACTORY_CODE") = FactoryCode

            rowSOTORDX2.Item("UPC_CODE") = GetUPC_CODE(rowSOTORDX2.Item("STYLE_CODE").ToString, rowSOTORDX2.Item("COLOR_CODE").ToString)
        Next
        Dim ExcelReplacements As New List(Of ExcelReplacement)
        BuildReplacements(XWS, ExcelReplacements)
        Dim SORT_ORDER As String = ""
        If chkSortFact.Checked Then
            SORT_ORDER = "FACTORY_CODE"
        End If
        For Each rowSOTORDX2 As DataRow In dst.Tables("SOTORDX2").Select("", SORT_ORDER)
            rowSOTORDR2_Current = rowSOTORDX2
            STYLE_CODE_Current = rowSOTORDX2.Item("STYLE_CODE").ToString
            SheetCount += 1
            XWS.Copy(After:=XWS_NEW)
            XWS_NEW = XWB.Worksheets(SheetCount)
            If chkSortFact.Checked Then
                Try
                    XWS_NEW.Name = MakeSheetName(rowSOTORDX2.Item("STYLE_CODE"), rowSOTORDX2.Item("COLOR_CODE"), rowSOTORDX2.Item("FACTORY_CODE"))
                Catch ex As Exception
                    XWS_NEW.Name = MakeSheetName(rowSOTORDX2.Item("STYLE_CODE"), rowSOTORDX2.Item("COLOR_CODE"), rowSOTORDX2.Item("FACTORY_CODE")) + "_1"
                End Try

            Else
                XWS_NEW.Name = MakeSheetName(rowSOTORDX2.Item("STYLE_CODE"), rowSOTORDX2.Item("COLOR_CODE"))
            End If
            Dim workingon As String = String.Format("Now Creating Sheet {0}, Style {1}", SheetCount, rowSOTORDX2.Item("STYLE_CODE"))
            ASCMAIN1.Progress(workingon)
            For Each EX As ExcelReplacement In ExcelReplacements
                Select Case Trim(EX.OldString)
                    Case Is = "<<ICTSTYL1.IMAGE_NAME>>"
                        Dim PictureFileName As String = GetImageLocation(rowSOTORDX2.Item("STYLE_CODE"), rowSOTORDX2.Item("COLOR_CODE"))
                        If PictureFileName.Length > 0 Then
                            Dim PicRange As Microsoft.Office.Interop.Excel.Range
                            If XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col)).MergeArea.MergeCells Then
                                Dim mRow As Integer = EX.Row + XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col)).MergeArea.Rows.Count - 1
                                Dim mCol As Integer = EX.Col + XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col)).MergeArea.Columns.Count - 1
                                PicRange = XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(mRow, mCol))
                            Else
                                PicRange = XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col))
                            End If
                            InsertPictureInRange(PictureFileName, PicRange, XWS_NEW)
                        Else
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Image Not Found"
                        End If
                    Case "<<ICTSTYL1.DUTY_RATE>>"
                        Dim rowICTDUTY1 As DataRow = LookUp("ICTDUTY1", ReplaceText("<<ICTSTYL1.DUTY_RATE_CODE>>"))
                        If IsNothing(rowICTDUTY1) Then
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = rowICTDUTY1.Item("DUTY_RATE").ToString()
                        End If
                    Case "<<ICTSTYL1.DUTY_RATE_PCT>>"
                        Dim rowICTDUTY1 As DataRow = LookUp("ICTDUTY1", ReplaceText("<<ICTSTYL1.DUTY_RATE_CODE>>"))
                        If IsNothing(rowICTDUTY1) Then
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            If IsNumeric(rowICTDUTY1.Item("DUTY_RATE").ToString()) Then
                                XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = Val(rowICTDUTY1.Item("DUTY_RATE").ToString) / 100
                            Else
                                XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                            End If
                        End If
                    Case "<<ICTSTYL1.FACTORY_CODE>>"
                        Dim VendCode As String = ReplaceText("<<ICTSTYL1.VEND_CODE>>")
                        Dim FactoryCode As String = GetVendorData(VendCode, "VEND_SUPPLIER_ID")
                        If FactoryCode.Length = 0 Then
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = FactoryCode
                        End If
                    Case "<<ICTSTYL1.PORT_CODE>>"
                        Dim VendCode As String = ReplaceText("<<ICTSTYL1.VEND_CODE>>")
                        Dim PortCode As String = GetVendorData(VendCode, "PORT_CODE")
                        If PortCode.Length = 0 Then
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = PortCode
                        End If
                    Case "<<ICTSTYL1.PORT_NAME>>"
                        Dim VendCode As String = ReplaceText("<<ICTSTYL1.VEND_CODE>>")
                        Dim PortCode As String = GetVendorData(VendCode, "PORT_CODE")
                        Dim rowICTPORT1 As DataRow = LookUp("ICTPORT1", PortCode)
                        If IsNothing(rowICTPORT1) Then
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = rowICTPORT1.Item("PORT_NAME").ToString
                        End If
                    Case Else
                        XWS_NEW.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = ReplaceText(EX.OldString)
                End Select
            Next
        Next
        'XWS = XWB.Worksheets(1)
        'XWS.Delete()
        MsgBox("Creation Of Excel From Template Is Complete", vbOKOnly, "Excel Output")
        excel.Visible = True
        excel = Nothing
    End Sub

    Private Sub btnGenerateQuote_Click(sender As System.Object, e As System.EventArgs) Handles btnGenerateQuote.Click
        If chkKirklands.Checked Then
            GenerateQuoteKirklands()
        Else
            GenerateQuoteStandard()
        End If
    End Sub

    Private Function MakeSheetName(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, Optional FACTORY_CODE As String = "") As String

        Dim iResult As String
        If FACTORY_CODE.Length > 0 Then
            iResult = String.Format("{0}-{1}-{2}", FACTORY_CODE, STYLE_CODE, COLOR_CODE)
        Else
            iResult = String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE)
        End If
        iResult = iResult.Replace("/", "")
        iResult = iResult.Replace("&", "")
        iResult = iResult.Replace("*", "")
        iResult = iResult.Replace("%", "")
        iResult = iResult.Replace("$", "")
        Return iResult
    End Function

    Private Function ReplaceText(CurrentText As String) As String
        Dim iResult As String = CurrentText
        Dim FoundField As Boolean = False
        Dim TableName As String = ""
        Dim ColumnName As String = ""
        Dim ReplaceData As String = ""
        Dim ReplaceString As String = ""
        Dim ReplaceStart As Integer = 0
        Dim ReplaceEnd As Integer = 0
        If iResult.Contains("<<") And iResult.Contains(">>") Then
            FoundField = True
            ReplaceStart = iResult.IndexOf("<<")
            ReplaceEnd = iResult.IndexOf(">>")
            TableName = iResult.Substring(ReplaceStart + 2, 8)
            ColumnName = iResult.Substring(ReplaceStart + 11, ReplaceEnd - (ReplaceStart + 11))
            ReplaceString = iResult.Substring(ReplaceStart, (ReplaceEnd + 2) - ReplaceStart)
            ReplaceData = GetDataFromTable(TableName, ColumnName, "Not Found")
            iResult = iResult.Replace(ReplaceString, ReplaceData)
        End If
        Return iResult
    End Function

    Private Function ExtractText(CurrentText As String) As String
        Dim iResult As String = CurrentText
        Dim FoundField As Boolean = False
        Dim TableName As String = ""
        Dim ColumnName As String = ""
        Dim ReplaceData As String = ""
        Dim ReplaceString As String = ""
        Dim ReplaceStart As Integer = 0
        Dim ReplaceEnd As Integer = 0
        If iResult.Contains("<<") And iResult.Contains(">>") Then
            FoundField = True
            ReplaceStart = iResult.IndexOf("<<")
            ReplaceEnd = iResult.IndexOf(">>")
            TableName = iResult.Substring(ReplaceStart + 2, 8)
            ColumnName = iResult.Substring(ReplaceStart + 11, ReplaceEnd - (ReplaceStart + 11))
            ReplaceString = iResult.Substring(ReplaceStart, (ReplaceEnd + 2) - ReplaceStart)
            'ReplaceData = GetDataFromTable(TableName, ColumnName, ReplaceString)
            iResult = ReplaceString
        Else
            iResult = ""
        End If
        Return iResult
    End Function

    Private Function GetDataFromTable(ByVal TableName As String, ByVal ColumnName As String, ByVal DefaultStringifNotFound As String) As String
        Dim iResult As String = DefaultStringifNotFound
        Dim FieldData As String = ""
        Select Case TableName
            Case "SOTORDR1"
                If dst.Tables.Item("SOTORDX1").Rows.Count >= 0 Then
                    Dim rowSOTORDR1 As DataRow = dst.Tables.Item("SOTORDX1").Rows(0)
                    If Not IsNothing(rowSOTORDR1) Then
                        Try
                            FieldData = rowSOTORDR1.Item(ColumnName).ToString
                        Catch ex As Exception
                            'Eat the Error and let FieldData be null.
                        End Try
                    End If
                End If
            Case "SOTORDR2"
                Try
                    FieldData = rowSOTORDR2_Current.Item(ColumnName).ToString
                Catch ex As Exception
                    'Eat the Error and let FieldData be null.
                End Try
            Case "SOTORDR5"
                If dst.Tables.Item("SOTORDX5").Rows.Count >= 0 Then
                    Dim rowSOTORDX5 As DataRow = dst.Tables.Item("SOTORDX5").Rows(0)
                    If Not IsNothing(rowSOTORDX5) Then
                        Try
                            FieldData = rowSOTORDX5.Item(ColumnName).ToString
                        Catch ex As Exception
                            'Eat the Error and let FieldData be null.
                        End Try
                    End If
                End If
            Case "ICTSTYL1"
                Fill_Records("ICTSTYX1", STYLE_CODE_Current, True)
                If dst.Tables.Item("ICTSTYX1").Rows.Count > 0 Then
                    dst.Tables.Item("ICTSTYX1").Rows(0).Item("FACTORY_CODE") = GetVendorData(dst.Tables("ICTSTYX1").Rows(0).Item("VEND_CODE").ToString, "VEND_SUPPLIER_ID")
                    dst.Tables.Item("ICTSTYX1").Rows(0).Item("PORT_CODE") = GetVendorData(dst.Tables("ICTSTYX1").Rows(0).Item("VEND_CODE").ToString, "PORT_CODE")
                End If
                If dst.Tables.Item("ICTSTYX1").Rows.Count >= 0 Then
                    Dim rowICTSTYX1 As DataRow = dst.Tables.Item("ICTSTYX1").Rows(0)
                    If Not IsNothing(rowICTSTYX1) Then
                        Try
                            FieldData = rowICTSTYX1.Item(ColumnName).ToString
                        Catch ex As Exception
                            'Eat the Error and let FieldData be null.
                        End Try
                    End If
                End If
        End Select
        If FieldData.Length > 0 Then
            iResult = FieldData
        End If
        Return iResult
    End Function

    Private Sub BuildReplacements(XWS As Excel.Worksheet, ByRef ExcelReplacements As List(Of ExcelReplacement))
        Dim xRow As Integer = 1
        Dim xCol As Integer = 1
        For R As Integer = 1 To 100
            For C As Integer = 1 To 100
                Dim workingon As String = String.Format("Now Analyzing Row {0}, Col {1}", R, C)
                ASCMAIN1.Progress(workingon)
                Dim ExcelReplacement As New ExcelReplacement
                ExcelReplacement.Row = R
                ExcelReplacement.Col = C
                ExcelReplacement.OldString = XWS.Range(Excel_Cell(R, C), Excel_Cell(R, C)).Text
                If ExcelReplacement.OldString.Length > 0 Then
                    'ExcelReplacement.NewString = ReplaceText(ExcelReplacement.OldString)
                    Dim FoundString As String = ExtractText(ExcelReplacement.OldString)
                    If FoundString.Length <> 0 Then
                        ExcelReplacements.Add(ExcelReplacement)
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub InsertPictureInRange(ByVal PictureFileName As String, _
            ByVal TargetCells As Microsoft.Office.Interop.Excel.Range, _
            ByVal XWS As Microsoft.Office.Interop.Excel.Worksheet)
        ASCMAIN1.Progress("Picture:" & PictureFileName)
        Dim pp As Microsoft.Office.Interop.Excel.Shape

        If TypeName(XWS) <> "Worksheet" Then Exit Sub
        If Dir(PictureFileName) = "" Then Exit Sub

        pp = XWS.Shapes.AddPicture(PictureFileName, _
           Microsoft.Office.Core.MsoTriState.msoFalse, _
           Microsoft.Office.Core.MsoTriState.msoCTrue, TargetCells.Left, TargetCells.Top, TargetCells.Width, TargetCells.Height)

        pp.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize
        pp.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse
        pp = Nothing
    End Sub

    Private Function GetImageLocation(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim RetVal As String = ""
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        Dim RO_PARM_STYLE_IMG_DIR As String = ""
        Dim FileMatch As String
        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
        Dim COLOR_CODE_LONG As String = ""
        If Not IsNothing(rowICTCOLR1) Then
            COLOR_CODE_LONG = rowICTCOLR1.Item("COLOR_CODE_LONG").ToString()
        End If
        Dim WebVal As String = ""
        If Not IsNothing(rowSOTPARM3) Then
            RO_PARM_STYLE_IMG_DIR = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            If RO_PARM_STYLE_IMG_DIR.Length > 0 Then
                FileMatch = Dir(String.Format("{0}\{1}-{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                If FileMatch.Length > 0 Then
                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                Else
                    FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                    If FileMatch.Length > 0 Then
                        RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                    Else
                        FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE_LONG))
                        If FileMatch.Length > 0 Then
                            RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                        Else
                            FileMatch = Dir(String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                            If FileMatch.Length > 0 Then
                                RetVal = String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE)
                            Else
                                FileMatch = Dir(String.Format("{0}\{1}*", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                                If FileMatch.Length > 0 Then
                                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Try
            If WebVal.Length > 0 Then
                Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(WebVal)
                Dim response As System.Net.WebResponse = req.GetResponse()
                Dim stream As IO.Stream = response.GetResponseStream()
                Dim img As System.Drawing.Image = System.Drawing.Image.FromStream(stream)
                stream.Close()
                If System.IO.File.Exists(RetVal) Then
                    System.IO.File.Delete(RetVal)
                    img.Save(RetVal)
                Else
                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE))
                    img.Save(RetVal)
                End If
            End If
        Catch ex As Exception
        End Try
        Return RetVal
    End Function

    Private Sub grdSOTQUOT2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTQUOT2.AfterRowsDeleted
        Update_Record()
    End Sub

    Private Sub grdSOTQUOT2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTQUOT2.BeforeRowsDeleted
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Delete Template"
        Dim iMSG As New System.Text.StringBuilder
        Dim QUOTE_NO As String = grdSOTQUOT2.ActiveRow.Cells.Item("QUOTE_NO").Text
        iMSG.Length = 0
        iMSG.AppendLine("Would You Also Like To Delete The Template File?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            Dim FullFileName As String = String.Format("{0}{1}.xls", ExcelFolder, QUOTE_NO)
            If FileIO.FileSystem.FileExists(FullFileName) Then
                FileIO.FileSystem.DeleteFile(FullFileName)
            Else
                FullFileName = String.Format("{0}{1}.xlsx", ExcelFolder, QUOTE_NO)
                If FileIO.FileSystem.FileExists(FullFileName) Then
                    FileIO.FileSystem.DeleteFile(FullFileName)
                Else
                    FullFileName = String.Format("{0}{1}.xlsm", ExcelFolder, QUOTE_NO)
                    If FileIO.FileSystem.FileExists(FullFileName) Then
                        FileIO.FileSystem.DeleteFile(FullFileName)
                    End If
                End If
            End If
        End If
    End Sub

    Private Function GetVendorData(ByVal VEND_CODE As String, ByVal COLUMN As String) As String
        Dim RetVal As String = ""
        If VEND_CODE.Length > 0 And COLUMN.Length > 0 Then
            ASCMAIN1.sql = String.Format("SELECT {0} FROM APTVEND1 WHERE VEND_CODE = '{1}'", COLUMN, VEND_CODE)
            RetVal = ASCDATA1.GetDataValue
        End If
        Return RetVal
    End Function

    Private Function GetUPC_CODE(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
        SQLS.AppendLine("SELECT NVL(UPC_CODE,'') AS UPC_CODE")
        SQLS.AppendLine("FROM ICTSTYC1")
        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Return ASCDATA1.GetDataValue
    End Function

    Private Sub GenerateQuoteKirklands()
        Dim FileExists As Boolean = False
        Dim ErrMsg As String = ""
        Dim NewSheetName As String = txtNewSheetName.Text
        Dim TemplateName As String = ""
        Dim TemplateExt As String = "xls"
        Dim XWB_SOURCE As Excel.Workbook
        Dim XWS_SOURCE As Excel.Worksheet = Nothing
        If grdSOTORDRX.Selected.Rows.Count <> 1 Or grdSOTQUOT2.Selected.Rows.Count <> 1 Then
            ErrMsg = ErrMsg & vbCrLf & "You Must Select One Order And One Template To Proceed."
        End If
        Dim QUOTE_NO As String = grdSOTQUOT2.ActiveRow.Cells.Item("QUOTE_NO").Text
        Dim ORDR_NO As String = grdSOTORDRX.ActiveRow.Cells.Item("ORDR_NO").Text
        Dim KirkFolder As String = ExcelFolder & Format(Now(), "yyyyMMdd")
        Dim TodayVersion As Integer = 0
        Do While FileIO.FileSystem.DirectoryExists(KirkFolder)
            TodayVersion += 1
            KirkFolder = ExcelFolder & Format(Now(), "yyyyMMdd") & "_" & TodayVersion
        Loop
        FileIO.FileSystem.CreateDirectory(KirkFolder)

        TemplateName = String.Format("{0}{1}.{2}", ExcelFolder, QUOTE_NO, TemplateExt)
        If FileIO.FileSystem.FileExists(TemplateName) Then
            FileExists = True
        Else
            TemplateExt = "xlsx"
            TemplateName = String.Format("{0}{1}.{2}", ExcelFolder, QUOTE_NO, TemplateExt)
            If FileIO.FileSystem.FileExists(TemplateName) Then
                FileExists = True
            Else
                TemplateExt = "xlsm"
                TemplateName = String.Format("{0}{1}.{2}", ExcelFolder, QUOTE_NO, TemplateExt)
                If FileIO.FileSystem.FileExists(TemplateName) Then
                    FileExists = True
                End If
            End If
        End If
        Fill_Records("SOTORDX1", ORDR_NO, True)
        Fill_Records("SOTORDX2", ORDR_NO, True)
        Fill_Records("SOTORDX5", ORDR_NO, True)
        For Each rowSOTORDX2 As DataRow In dst.Tables("SOTORDX2").Select()
            Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
            SQLS.AppendLine(String.Format("Select VEND_CODE from ICTSTYL1 where STYLE_CODE = '{0}'", rowSOTORDX2.Item("STYLE_CODE")))
            ASCMAIN1.sql = SQLS.ToString()
            Dim VEND_CODE As String = ASCDATA1.GetDataValue
            Dim FactoryCode As String = GetVendorData(VEND_CODE, "VEND_SUPPLIER_ID")
            rowSOTORDX2.Item("FACTORY_CODE") = FactoryCode

            rowSOTORDX2.Item("UPC_CODE") = GetUPC_CODE(rowSOTORDX2.Item("STYLE_CODE").ToString, rowSOTORDX2.Item("COLOR_CODE").ToString)
        Next
        Dim excel_source As Excel.Application = New Microsoft.Office.Interop.Excel.Application
        Dim ExcelReplacements As New List(Of ExcelReplacement)
        Try
            XWB_SOURCE = excel_source.Workbooks.Open(TemplateName)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Opening File")
            excel_source = Nothing
            Exit Sub
        End Try
        Dim FoundSheet As Boolean = False
        For i As Integer = 1 To XWB_SOURCE.Worksheets.Count
            XWS_SOURCE = XWB_SOURCE.Worksheets(i)
            If XWS_SOURCE.Name = "Item Info" Then
                FoundSheet = True
                Exit For
            End If
        Next
        If Not FoundSheet Then
            MsgBox("Can Not Find Workbook", MsgBoxStyle.Critical, "Workbook")
        End If

        BuildReplacements(XWS_SOURCE, ExcelReplacements)
        XWB_SOURCE.Saved = True
        XWB_SOURCE.Save()
        XWB_SOURCE.Close()
        excel_source.Quit()
        releaseObject(XWS_SOURCE)
        releaseObject(XWB_SOURCE)
        releaseObject(excel_source)
        Dim SORT_ORDER As String = ""

        If chkSortFact.Checked Then
            SORT_ORDER = "FACTORY_CODE"
        End If
        For Each rowSOTORDX2 As DataRow In dst.Tables("SOTORDX2").Select("", SORT_ORDER)
            Dim XWB_DEST As Excel.Workbook = Nothing
            Dim XWS_DEST As Excel.Worksheet = Nothing
            rowSOTORDR2_Current = rowSOTORDX2
            STYLE_CODE_Current = rowSOTORDX2.Item("STYLE_CODE").ToString
            Dim ThisWorkBookName As String = String.Format("{0}.{1}", MakeSheetName(rowSOTORDX2.Item("STYLE_CODE"), rowSOTORDX2.Item("COLOR_CODE")), TemplateExt)
            Dim ThisWorkBookFullPath As String = String.Format("{0}\{1}", KirkFolder, ThisWorkBookName)
            If System.IO.File.Exists(ThisWorkBookFullPath) Then
                System.IO.File.Delete(ThisWorkBookFullPath)
            End If

            FileIO.FileSystem.CopyFile(TemplateName, ThisWorkBookFullPath)
            Dim excel_DEST As Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Try
                XWB_DEST = excel_DEST.Workbooks.Open(ThisWorkBookFullPath)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Opening File")
                releaseObject(XWS_DEST)
                releaseObject(XWB_DEST)
                releaseObject(excel_DEST)
                Exit Sub
            End Try
            XWS_DEST = XWB_DEST.Worksheets(1)
            'XWS_NEW = XWB.Worksheets(1)
            Dim workingon As String = String.Format("Now Creating Workbook For Style {0}", rowSOTORDX2.Item("STYLE_CODE"))
            ASCMAIN1.Progress(workingon)
            For Each EX As ExcelReplacement In ExcelReplacements
                Select Case Trim(EX.OldString)
                    Case Is = "<<ICTSTYL1.IMAGE_NAME>>"
                        Dim PictureFileName As String = GetImageLocation(rowSOTORDX2.Item("STYLE_CODE"), rowSOTORDX2.Item("COLOR_CODE"))
                        If PictureFileName.Length > 0 Then
                            Dim PicRange As Microsoft.Office.Interop.Excel.Range
                            If XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col)).MergeArea.MergeCells Then
                                Dim mRow As Integer = EX.Row + XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col)).MergeArea.Rows.Count - 1
                                Dim mCol As Integer = EX.Col + XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col)).MergeArea.Columns.Count - 1
                                PicRange = XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(mRow, mCol))
                            Else
                                PicRange = XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col))
                            End If
                            InsertPictureInRange(PictureFileName, PicRange, XWS_DEST)
                        Else
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Image Not Found"
                        End If
                    Case "<<ICTSTYL1.DUTY_RATE>>"
                        Dim rowICTDUTY1 As DataRow = LookUp("ICTDUTY1", ReplaceText("<<ICTSTYL1.DUTY_RATE_CODE>>"))
                        If IsNothing(rowICTDUTY1) Then
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = rowICTDUTY1.Item("DUTY_RATE").ToString()
                        End If
                    Case "<<ICTSTYL1.DUTY_RATE_PCT>>"
                        Dim rowICTDUTY1 As DataRow = LookUp("ICTDUTY1", ReplaceText("<<ICTSTYL1.DUTY_RATE_CODE>>"))
                        If IsNothing(rowICTDUTY1) Then
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            If IsNumeric(rowICTDUTY1.Item("DUTY_RATE").ToString()) Then
                                XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = Val(rowICTDUTY1.Item("DUTY_RATE").ToString) / 100
                            Else
                                XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                            End If
                        End If
                    Case "<<ICTSTYL1.FACTORY_CODE>>"
                        Dim VendCode As String = ReplaceText("<<ICTSTYL1.VEND_CODE>>")
                        Dim FactoryCode As String = GetVendorData(VendCode, "VEND_SUPPLIER_ID")
                        If FactoryCode.Length = 0 Then
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = FactoryCode
                        End If
                    Case "<<ICTSTYL1.PORT_CODE>>"
                        Dim VendCode As String = ReplaceText("<<ICTSTYL1.VEND_CODE>>")
                        Dim PortCode As String = GetVendorData(VendCode, "PORT_CODE")
                        If PortCode.Length = 0 Then
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = PortCode
                        End If
                    Case "<<ICTSTYL1.PORT_NAME>>"
                        Dim VendCode As String = ReplaceText("<<ICTSTYL1.VEND_CODE>>")
                        Dim PortCode As String = GetVendorData(VendCode, "PORT_CODE")
                        Dim rowICTPORT1 As DataRow = LookUp("ICTPORT1", PortCode)
                        If IsNothing(rowICTPORT1) Then
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = "Not Found"
                        Else
                            XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = rowICTPORT1.Item("PORT_NAME").ToString
                        End If
                    Case Else
                        XWS_DEST.Range(Excel_Cell(EX.Row, EX.Col), Excel_Cell(EX.Row, EX.Col)).Value = ReplaceText(EX.OldString)
                End Select
            Next

            XWB_DEST.Saved = True
            XWB_DEST.Save()
            XWB_DEST.Close()
            excel_DEST.Quit()
            releaseObject(XWS_DEST)
            releaseObject(XWB_DEST)
            releaseObject(excel_DEST)
        Next
        MsgBox("Creation Of Excel From Template Is Complete", vbOKOnly, "Excel Output")
        ASCMAIN1.Progress("")
        For Each p As Process In Process.GetProcesses
            If p.ProcessName = "EXCEL" Then
                p.Kill()
            End If
        Next
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
        Catch ex As Exception
            Stop
        Finally
            obj = Nothing
        End Try
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Function GetExtForFile(ByVal FileName As String) As String
        Dim RetVal As String = ""
        If FileName.Length > 0 Then
            Dim IndexOfExtension As Integer = FileName.IndexOf(".")
            If IndexOfExtension > 1 Then
                RetVal = FileName.Substring(IndexOfExtension, Len(FileName) - IndexOfExtension)
            Else
                For Each ExtName As String In ExtOptions
                    If FileIO.FileSystem.FileExists(ExcelFolder & FileName & ExtName) Then
                        RetVal = ExtName
                        Exit For
                    End If
                Next
            End If
            Dim ValidFileExt As Boolean = False
            For Each ExtName As String In ExtOptions
                If RetVal = ExtName Then
                    ValidFileExt = True
                    Exit For
                End If
            Next
            If Not ValidFileExt Then
                RetVal = ""
            End If
        End If
        Return RetVal
    End Function
End Class

Public Class ExcelReplacement
    Public Property Row As Integer
    Public Property Col As Integer
    Public Property OldString As String
    Public Property NewString As String
End Class