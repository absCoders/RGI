Public Class SOFINVE1

    Dim CUST_CODE As String
    Dim XLS_GENERATED As Boolean

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            .Tables("SOTINVH1").Columns.Add("SEL")
            .Tables("SOTINVH1").Columns("SEL").DefaultValue = "0"

        End With

        grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")
        Create_Summary(grdSOTINVH1, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1, "SEL")
        Create_Summary(grdSOTINVH1, New String() _
                       {"INV_SALES", "INV_COGS", "INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT", "INV_TOTAL_AMOUNT_CURR", _
                        "INV_SALES_CURR", "INV_FREIGHT_CURR", "INV_MISC_CHG_CURR", "INV_TOTAL_AMT_CURR", "GST_TAX", "GST_TAX_CURR", "INV_STAX", "INV_STAX_CURR"})


        With grdSOTINVH1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With

        With grdSOTINVH1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                If New String() {"INV_SALES", "INV_COGS", "INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT", "INV_TOTAL_AMOUNT_CURR", _
                                 "INV_SALES_CURR", "INV_FREIGHT_CURR", "INV_MISC_CHG_CURR", "INV_TOTAL_AMT_CURR", _
                                 "GST_TAX", "GST_TAX_CURR", "INV_STAX", "INV_STAX_CURR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                End If

                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If


            Next
        End With

        tabMain.Tabs("XLS").Visible = False

        cbeYPS.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -120) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYPS.SelectedItem = cbeYPS.Items(Val(Mid(ASCMAIN1.CYP, 5, 2)) - 1)
        cbeYPE.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -120) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYPE.SelectedItem = cbeYPE.Items(0)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE", False, True)

                CUST_CODE = Absx1.txtFor("CUST_CODE").Text

                If CUST_CODE <> "LOBLAW" And CUST_CODE <> "SDM" Then
                    EMsg &= vbCr & "Only LOBLAW & SDM are supported in this screen"
                End If

            Case "Generate XLS"

                If dst.Tables("SOTINVH1").Select("SEL='1'").Length = 0 Then
                    EMsg &= vbCr & "No Invoices Selected"
                End If
                If XLS_GENERATED Then
                    EMsg &= vbCr & "Already Generated - Click Done and Re-Load"
                End If

            Case "Done"

                'If MessageBox.Show("Do you want to Cancel changes?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                '    Exit Sub
                'End If

            Case "Update"


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

                'Case "Update"
                '    Update_Record()
                '    Mode_Settings(False)

            Case "Generate XLS"
                Generate_XLS()
                XLS_GENERATED = True

            Case "Save XLS"
                WorkbookView1.GetLock()
                Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("SOFINVE1.XLSX_NO") & ".XLSX"
                WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                Show_Document(FILENAME)
                WorkbookView1.ReleaseLock()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                 
                    .Items("Generate XLS").Visible = ScreenMode
                    .Items("Save XLS").Visible = False ' ScreenMode

                End With

            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = ScreenMode

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

        grdSOTINVH1.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("SOTINVH1").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Clear()

        Clear_All_Filters(grdSOTINVH1)

    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading 1099 information")

        Dim YPS As String = cbeYPS.Value
        Dim YPE As String = cbeYPE.Value

        EnforceConstraints(False)

        ASCMAIN1.sql = "Select SOTINVH1.*, TATSTATE.STATE_NAME " & vbCrLf _
                     & " from SOTINVH1, TATSTATE " & vbCrLf _
                     & " where SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'" & vbCrLf _
                     & "   and SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & YPS & "'" & vbCrLf _
                     & "   and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & YPE & "'" & vbCrLf _
                     & "   and TATSTATE.STATE_CODE (+) = SOTINVH1.CUST_SHIP_TO_STATE"

        ASCMAIN1.sql &= " and SOTINVH1.INV_TYPE = 'I'" & vbCrLf
        ASCMAIN1.sql &= " and nvl(SOTINVH1.INV_SALES_CURR,0) > 0" & vbCrLf
        ASCMAIN1.sql &= " and SOTINVH1.INV_NO_REV is Null and SOTINVH1.INV_NO_REV_BY is Null" & vbCrLf

        Fill_Records("SOTINVH1", String.Empty, True, ASCMAIN1.sql)
        EnforceConstraints(True)

        Sort_grdColumns(grdSOTINVH1, "INV_TYPE,INV_NO")

        Load_Template()

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        'Try
        '    BeginTrans()
        '    Update_Record_TDA("SOTCARRA", "DELETE FROM SOTCARRA where CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
        '    CommitTrans("Update Complete")
        'Catch ex As Exception
        '    Rollback(ex.Message)
        'End Try

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTINVH1, "SSBBBB", "Show Filter", "Show GroupBox", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
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

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                Next
                grdSOTINVH1.Selected.Rows.Clear()
        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

    Sub Load_Template()
        If 1 = 1 Then Exit Sub

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & "\" & CUST_CODE & ".xlsb"
        WorkbookView1.GetLock()
        WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)

        Dim X As SpreadsheetGear.Commands.CommandManager = New MyCommandManager(WorkbookView1.ActiveWorkbookSet)

        'XLS_Validation(True)
        'XLS_STD()
        'XLS_Refresh_Checkbooks()
        'XLS_SubTotals()
        'Set_Month_Headings(SEASON_CODE)
        'XLS_Refresh_Stores()

        WorkbookView1.ReleaseLock()


    End Sub


    Sub Generate_XLS()

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & "\" & CUST_CODE & ".xlsb"

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        ws = wb.Worksheets(1)


        ws.Cells(2, 5).value2 = "New York Accessory Group"
        ws.Cells(3, 5).value2 = "1700326"

        Dim dvw As DataView = dst.Tables("SOTINVH1").DefaultView
        dvw.RowFilter = "SEL = '1'"
        dvw.Sort = "INV_NO"

        Dim DataTable As DataTable = dvw.ToTable
        Dim MaxCols As Integer = DataTable.Columns.Count

        Dim iRx As Integer = 11
        Dim r As Integer = 0 ' since we are using XLS Automation
        Dim c As Integer
 

        For Each row As DataRow In DataTable.Select("", "INV_NO")
            r += 1
            ASCMAIN1.Progress("-", r)
            c = 1

            '    ws.Range(ws.Cells(iRx + r, c), ws.Cells(iRx + r, c + MaxCols - 1)).Value2 = row.ItemArray

            ws.Cells(iRx + r, 4).Value2 = "Warehouse"
            ws.Cells(iRx + r, 5).Value2 = row.Item("CUST_STORE_NO")
            ws.Cells(iRx + r, 6).Value2 = row.Item("ORDR_CUST_PO")
            ws.Cells(iRx + r, 8).Value2 = Format(row.Item("INV_DATE"), "MM/dd/yyyy")
            ws.Cells(iRx + r, 9).Value2 = "'" & row.Item("INV_NO")
            ws.Cells(iRx + r, 10).Value2 = Val(row.Item("INV_SALES_CURR") & "")
            ws.Cells(iRx + r, 11).Value2 = Val(row.Item("GST_TAX_CURR") & "")

        Next
  
        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                'XLS_FILENAME = ASCMAIN1.Next_Control_No("SOFINVE1.XLSX_NO") & ".XLSB"
                Dim DT As String = Format(Now, "yyyyMMddHHmmss")
                XLS_FILENAME = "Inquiry Template_Registre de demande" & "_" & DT & ".XLSB"
                '  XLS_FILENAME = BUDGET_YEAR & "_" & BUDGET_VERSION & "_Budgets"
                ' XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsX"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , 50) ' Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
                ', Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        Show_Document(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        '  Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")

        MsgBox("XLS File has been Generated")



    End Sub
End Class


Public Class MyCommandManager
    Inherits SpreadsheetGear.Commands.CommandManager
    Friend Sub New(workbookSet As SpreadsheetGear.IWorkbookSet)
        MyBase.New(workbookSet)
    End Sub
    Public Overrides Function CreateCommandPaste(range As SpreadsheetGear.IRange) As SpreadsheetGear.Commands.Command
        ' This is what would normally be called...
        ' return new CommandRange.Paste(range);  

        ' Anytime a Paste command is invoked, this will force a "Paste Values"
        Return New SpreadsheetGear.Commands.CommandRange.PasteSpecial(range, SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)
    End Function
End Class

