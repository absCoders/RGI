Imports System.Text
Imports SpreadsheetGear
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports Infragistics.Win.UltraWinGrid

Public Class SORCUSTS
    Dim S As New StringBuilder With {.Length = 0}
    Dim SQL_WHERE_ORDR_RSV As String
    Dim REPORT_NAME As String = "SORCUSTS"
    Dim ICTSTATDSQL As String = ""
    Dim ICTSTAT2SQL As String = ""
    Dim XLS_NO As Integer = 0
    Dim exlExt As String = ".xlsx"
    Dim SQL_REPORT As New StringBuilder With {.Length = 0}
    Dim GRP_IN As String = ""
    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp
    Dim STYLE_CODE As String
    Dim IMG_Error_Reported As Boolean = False

#Region "Report Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

        RWU = "N"
        Get_PARM("ICTPARM1")

        Build_Init_Sel()

        Fill_Records("WEBLINKS")

        grdSOTINVHX.DataSource = dst.Tables.Item("SOTINVHX")

        grdWEBLINKS.DataSource = dst.Tables.Item("WEBLINKS")

        With grdWEBLINKS.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
            .Bands(0).Columns("DATE_ADDED").Format = "MM/dd/yy"
        End With

        With grdSOTINVHX.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
            .Bands(0).Columns("INV_DATE").Format = "MM/dd/yy"
            .Bands(0).Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        Sort_grdColumns(grdSOTINVHX, "INV_DATE".ToLower())

        dteShip_Beg.Value = DateSerial(Now.Year, Now.Month, 1)
        dteShip_End.Value = DateSerial(Now.Year, Now.Month, 1).AddMonths(1).AddDays(-1)

        With UltraExplorerBar1.Groups("Special Functions")
            .Visible = False
        End With

    End Sub

    Protected Overrides Sub Build_Workfile()
        Call ASCMAIN1.Progress("Building Work File")
        'Prepare Parameters for Report
        'VDDVDDVDD
        Dim ODF As String = "dd-MMM-yyyy"
        Dim CUST_CODE As String = txtCUST_CODE.Text & String.Empty
        Dim Ship_Beg As String = Format(dteShip_Beg.DateTime, ODF)
        Dim Ship_End As String = Format(dteShip_End.DateTime, ODF)

        ' Prepare filters from Run-Time Options
        SUBT = ""
        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")
        Dim sql_TABLE_NAMEs_orig As String = sql_TABLE_NAMEs
        Dim sql_JOIN_orig As String = sql_JOIN

        Dim sql_filter2 As String = ""

        '-- Shit you may need here --
        'sql_SELECT_cols, sql_TABLE_NAMEs, sql_WHERE, sql_JOIN, sql_filter, sql_filter2
        FixSqlWhere(sql_WHERE)
        FixSqlGroup(sql_GROUP_BY_cols)

        S = BUILD_SOTCUSTS(False)
        ASCMAIN1.sql = S.ToString
        Create_TDA(dst.Tables.Add, "SOTCUSTS", "**", 0, False)
        With dst.Tables("SOTCUSTS")
            .Columns.Add("SEQ").DataType = GetType(System.Int32)
            .Columns.Add("STYLE_CODE_PLM")
            .Columns.Add("IMAGE", GetType(System.Byte()))
            .Columns.Add("SELECTED")
            .Columns("SELECTED").DefaultValue = "0"
        End With

        '''With dst.Tables("SOTCUSTS").Columns
        '''    For iCOL As Integer = 0 To 4
        '''        .Add("QTY_AVA" & CStr(iCOL), GetType(System.Int64))
        '''        .Add("DTE" & CStr(iCOL), GetType(System.DateTime))
        '''    Next
        '''    .Add("QTY_AVA", GetType(System.Int64), "ISNULL(QTY_AVA0,0)+ISNULL(QTY_AVA1,0)+ISNULL(QTY_AVA2,0)+ISNULL(QTY_AVA3,0)+ISNULL(QTY_AVA4,0)")
        '''    .Add("OPEN_PICK_RSRV", GetType(System.Int64))
        '''    '     .Add("COUNT_COLOR", GetType(System.Int32), String.Format(COUNT_COLOR, 0))
        '''    .Add("SKIP_COLOR")
        '''    .Add("LAST_RCD_DATE")
        '''    .Add("EVER_ORDRED", GetType(System.Int64))
        '''    '.Add("LAST_SHIP_DATE")
        '''End With


        Dim TABLE_TEMP As String = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "select *  from ictstat2 WHERE (STYLE_CODE,COLOR_CODE) IN (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & TABLE_TEMP & ")"
        ICTSTAT2SQL = ASCMAIN1.sql
        Create_TDA(dst.Tables.Add, "ICTSTAT2", "**", 2, False)


        ASCMAIN1.sql = "Select * from (" & vbCrLf _
            & " Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.INIT_DATE, POTSHIP1.WHSE_CODE, POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & ", POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
            & ", POTORDR1.FACTORY_CODE, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & ", POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
            & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_REF_NO, POTSHIP2.CONTAINER_NO" & vbCrLf _
            & ", POTSHIP2.PO_DATE_RECEIVED" & vbCrLf _
            & ", POTSHIP3.PO_QTY_SHP, POTSHIP3.PO_QTY_REC" & vbCrLf _
            & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
            & ", POTORDR2.PO_QTY_ORD, 0 PO_QTY_OPN" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_ETA + NVL(POTSHIP1.PO_SHIP_LANDING_LEAD_DAYS,0) PO_ARRIVAL_DATE" & vbCrLf _
            & ", POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY" & vbCrLf _
            & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
            & "From POTSHIP1, POTSHIP2, POTSHIP3, POTORDR1, POTORDR2, ICTATOP2" & vbCrLf _
            & "Where POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & " And POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & " And POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & " And POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "  And POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & " And POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "  And ICTATOP2.PS_CODE (+) = 'S'" & vbCrLf _
            & " And ICTATOP2.PS_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & " And POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & " And (POTORDR2.STYLE_CODE) IN  (SELECT DISTINCT STYLE_CODE FROM  " & TABLE_TEMP & ")" & vbCrLf _
            & " ) union (" & vbCrLf _
            & "Select  POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE,POTORDR1.INIT_DATE, POTORDR1.WHSE_CODE, POTORDR2.PO_ORDER_NO" & vbCrLf _
            & ", POTORDR1.PO_DATE_SHIP_BY PO_DATE_SHIP_BY_REQ, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
            & ", POTORDR1.FACTORY_CODE, POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & ", Null PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO" & vbCrLf _
            & ", Decode(nvl(POTORDR2.PO_QTY_OPN,0),0,'ClosedPO','OpenPO') PO_SHIP_VESSEL" & vbCrLf _
            & ", POTORDR2.PO_DATE_SHIP_BY, POTORDR2.PO_DATE_ETA" & vbCrLf _
            & ", 10 PO_SHIP_LANDING_LEAD_DAYS" & vbCrLf _
            & ", Null PO_SHIP_REF_NO, Null CONTAINER_NO" & vbCrLf _
            & ", NULL PO_DATE_RECEIVED" & vbCrLf _
            & ", 0 PO_QTY_SHP, 0 PO_QTY_REC" & vbCrLf _
            & ", POTORDR1.VEND_CODE, POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
            & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN" & vbCrLf _
            & ", POTORDR2.PO_DATE_ETA + 10 PO_ARRIVAL_DATE" & vbCrLf _
            & ", POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.LAST_OPER_SHIP_BY" & vbCrLf _
            & ", ICTATOP2.STYLE_ARRIVAL_BUFFER_DAYS, ICTATOP2.STYLE_AT_ONCE_UNTIL, ICTATOP2.STYLE_AT_ONCE_ACTIVE" & vbCrLf _
            & " From POTORDR1, POTORDR2, ICTATOP2" & vbCrLf _
            & "Where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
            & "  And ICTATOP2.PS_CODE (+) = 'P'" & vbCrLf _
            & "   And ICTATOP2.PS_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   And POTORDR2.PO_QTY_OPN <> 0" & vbCrLf _
            & " And (POTORDR2.STYLE_CODE) IN  (SELECT DISTINCT STYLE_CODE FROM  " & TABLE_TEMP & ")" & vbCrLf _
            & ")"

        ICTSTATDSQL = ASCMAIN1.sql
        Create_TDA(dst.Tables.Add, "ICTSTATD", "**", 0, False)


        'Create_TDA(dst.Tables.Add, "SOTCUSTS", "**", 0, False, "VDDVDDVDD")

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        'Sticking a stupid row in the table to keep the standards from being an ass.
        Dim newASTSRPT1 As DataRow = dst.Tables("ASTSRPT1").NewRow
        newASTSRPT1.Item("G1") = "XX"
        dst.Tables("ASTSRPT1").Rows.Add(newASTSRPT1)

        UpdateReportRows()
    End Sub

    Public Overrides Sub Print_Report()
        ASCMAIN1.Progress("Creating Customer Open Report", "")
        Dim XLS_FILENAME1 As String = MakeExcelWorkbook()
        Dim XLS_FILENAME2 As String = ""
        Show_Document(XLS_FILENAME1)


        ASCMAIN1.Progress("", "")
    End Sub

    Private Function MakeExcelWorkbook() As String
        Dim XLS_FILENAME As String = ""

        Dim StyleList As New List(Of String)
        For Each rowSOTCUSTS As DataRow In dst.Tables("SOTCUSTS").Select()
            Dim STYLE_CODE As String = rowSOTCUSTS.Item("STYLE_CODE").ToString & String.Empty
            If Not StyleList.Contains(STYLE_CODE) Then
                StyleList.Add(STYLE_CODE)
            End If
        Next


        If chk1Sheet.Checked Then
            Dim fileName As String = ""
            fileName = Create_Excel()
        Else
            Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
            Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
            worksheet.Name = "Style Info
"
            Create_Excel_WorkSheet(worksheet, StyleList)


            If ASCMAIN1.Folders("Temp").EndsWith("\") Then
                XLS_FILENAME = ASCMAIN1.Folders("Temp") & String.Format("{0}.XLSX", REPORT_NAME)
            Else
                XLS_FILENAME = ASCMAIN1.Folders("Temp") & "\" & String.Format("{0}.XLSX", REPORT_NAME)
            End If
            Dim success As Boolean = False

            ASCMAIN1.Progress("Now Saving Workbook")

            Do Until success
                Try
                    If System.IO.File.Exists(XLS_FILENAME) Then
                        System.IO.File.Delete(XLS_FILENAME)
                    End If
                    workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                    If chkWebLinks.Checked Then
                        SaveLinks(XLS_FILENAME)
                    End If
                    success = True
                Catch ex As Exception

                End Try
            Loop
            Return XLS_FILENAME


        End If

    End Function

    Private Sub SaveLinks(ByVal FILENAME_FULL As String)
        Dim SESSION_NO As String = ASCMAIN1.Next_Control_No(String.Format("{0}.SESSION_NO", REPORT_NAME))
        Dim FILE_NAME As String = String.Format("{0}_{1}.XLSX", REPORT_NAME, SESSION_NO)
        Dim FULLPATH As String = "\\192.168.180.34\g\VDI\ARCHIVE\VAN\Links\" & FILE_NAME
        Dim HASHVALUE As String = ASCMAIN1.Get_Hash(SESSION_NO & String.Format("{0}.XLSX", REPORT_NAME))

        If IsNothing(dst.Tables.Item("WEBLINKS")) Then
            ASCMAIN1.sql = SQL_REPORT.ToString
            Create_TDA(dst.Tables.Add, "WEBLINKS", "**", 0, True)
        End If
        If Not (ASCMAIN1.Running_in_VS) Then
            If System.IO.File.Exists(FULLPATH) Then
                System.IO.File.Delete(FULLPATH)
            End If
            System.IO.File.Copy(FILENAME_FULL, FULLPATH)
        End If

        Dim rowWEBLINKS As DataRow = dst.Tables.Item("WEBLINKS").NewRow
        rowWEBLINKS.Item("HASHVALUE") = HASHVALUE
        rowWEBLINKS.Item("FILE_NAME") = FILE_NAME
        rowWEBLINKS.Item("USER_NAME") = ASCMAIN1.USER_ID
        rowWEBLINKS.Item("CUST_CODE") = txtCUST_CODE.Text
        rowWEBLINKS.Item("STYLE_CODE") = ""
        rowWEBLINKS.Item("IS_PRIVATE") = "0"
        rowWEBLINKS.Item("DATE_ADDED") = Now()
        rowWEBLINKS.Item("FORM_NAME") = REPORT_NAME
        dst.Tables.Item("WEBLINKS").Rows.Add(rowWEBLINKS)
        Update_Record_TDA("WEBLINKS")

        Dim FileNameLocalFull As String = FILENAME_FULL
        Dim FileNameRemote As String = FILE_NAME
        Dim eMsg As Text.StringBuilder = FTP_BLUEHOST(FileNameLocalFull, FileNameRemote)
        If eMsg.Length > 0 Then
            MsgBox(eMsg.ToString, vbCritical, "Error Sending To Remote Server")
        End If
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            Dim BadCol As Boolean = False
            If txtCUST_CODE.Text.Length = 0 Then
                EMsg &= vbCr & "You Must Select A Customer To Run This Report."
            End If
            For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE")
                If rowASTDSQLA.Item("COLUMN_NAME") = "ORDR_GROUP_NO" Or rowASTDSQLA.Item("COLUMN_NAME") = "RSRV_NO" Then
                    BadCol = True
                End If
            Next
            If BadCol Then
                EMsg &= vbCr & "You Can Not Sort By Group Or Reservations"
            End If

            If optSelectBy.Value = "D" Then
                If Not IsDate(dteShip_Beg.Value) Or Not IsDate(dteShip_End.Value) Then
                    EMsg &= vbCr & "Invalid Beginning Or Ending Date Selected"
                Else
                    If dteShip_Beg.Value > dteShip_End.Value Then
                        EMsg &= vbCr & "Beginning Date > Ending Date"
                    End If
                End If
            Else
                Dim RecCnt As Int64 = dst.Tables.Item("SOTINVHX").Select("SEL = '1'").Count
                If RecCnt = 0 Then
                    EMsg &= vbCr & "No POs Selected When Running By PO Option"
                Else
                    GRP_IN = ""
                    For Each rowSOTINVHX As DataRow In dst.Tables.Item("SOTINVHX").Select("SEL = '1'")
                        GRP_IN += String.Format("'{0}',", rowSOTINVHX.Item("ORDR_GROUP_NO").ToString & String.Empty)
                    Next
                    GRP_IN = GRP_IN.Substring(0, GRP_IN.Length - 1)
                End If
            End If

        End If
        If eItemKey = "Done" Then
            Build_Init_Sel()
            With UltraExplorerBar1.Groups("Special Functions")
                .Visible = False
            End With
        End If

        If eItemKey = "Buyer Chart" Then
            If chk1Sheet.Checked Then
                Create_Excel_BuyerChart_DIV()
            Else
                Create_Excel_BuyerChart()
            End If
            'With UltraExplorerBar1.Groups("Special Functions")
            '    .Visible = False
            'End With
            Exit Sub
        End If

        If eItemKey = "Print Full CADs" Then
            Print_Full_CAD_Print(eItemKey)
            'With UltraExplorerBar1.Groups("Special Functions")
            '    .Visible = False
            'End With

            '     UltraExplorerBar1.Groups("Special Functions").Items(eItemKey).Visible = False
        End If
    End Sub

    Public Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
            With UltraExplorerBar1.Groups("Special Functions")
                .Visible = True
            End With
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        If parms.Length > 0 Then
        End If

        EnforceConstraints(False)
        S = BUILD_SOTCUSTS(False)
        Fill_Records("SOTCUSTS",,, S.ToString)

        Fill_Records("ICTSTAT2", "", True, ICTSTAT2SQL)

        Fill_Records("ICTSTATD", "", True, ICTSTATDSQL)

        '''       Get_Availability()

        EnforceConstraints(True)
    End Sub

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWEBLINKS, "BBBB", "View File", "Replace File", "Copy Link", "Extend Expiration")
        Load_Popup_Menu(grdSOTINVHX, "S", "Show Filter")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.Name = "grd" Then
            Exit Sub
        End If

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd Is Nothing OrElse grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "View File"
                If Not IsNothing(grd.ActiveRow) Then
                    Dim FN As String = grd.ActiveRow.Cells.Item("FILE_NAME").Text
                    Show_Document(FN)
                End If
            Case "Replace File"
                Dim openFileDialog1 As New OpenFileDialog()
                openFileDialog1.Filter = "excel files (*.xlsx)|*.xlsx"
                If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    Dim FN_FROM As String = openFileDialog1.FileName

                    Dim FN_TO As String = "G:\VDI\ARCHIVE\VAN\Links\" & grd.ActiveRow.Cells.Item("FILE_NAME").Text

                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Replace file"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("This Action Will Replace The Generated File")
                    iMSG.AppendLine("With The Following File You Selected:")
                    iMSG.AppendLine(FN_FROM)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Is That What You Want?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)

                    If iResult = MsgBoxResult.Yes Then
                        If System.IO.File.Exists(FN_TO) Then
                            System.IO.File.Delete(FN_TO)
                        End If
                        System.IO.File.Copy(FN_FROM, FN_TO)
                    End If
                    MsgBox("You File Has Been Replaced", vbInformation, "Done")
                End If
            Case "Copy Link"
                'Dim FILE_NAME As String = grd.ActiveRow.Cells.Item("FILE_NAME").Text
                'Dim HASH As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
                'Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
                'My.Computer.Clipboard.SetText(FILE_NAME & vbCrLf & LINEPFX)
                'MsgBox("You Link Has Been Copied To Your Clipboard", vbInformation, "Done")
                Dim FILE_NAME As String = grd.ActiveRow.Cells.Item("FILE_NAME").Text
                Dim HASHVALUE As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
                Dim LINEPFX As String = $"https://vandaledocs.azurewebsites.net/Documents/{HASHVALUE}"
                My.Computer.Clipboard.SetText(FILE_NAME & vbCrLf & LINEPFX)
                MsgBox("You Link Has Been Copied To Your Clipboard", vbInformation, "Done")
            Case "Extend Expiration"
                If grd.Selected.Rows.Count = 1 Then
                    Dim grow As UltraWinGrid.UltraGridRow = grd.Selected.Rows(0)
                    Dim HASHVALUE As String = grow.Cells.Item("HASHVALUE").Text.ToString & String.Empty
                    Dim RetVal As Boolean = EXTEND_LINK(HASHVALUE)
                    If RetVal = True Then
                        MsgBox("Your Link Is Extended For 20 Days From Today.", vbOKOnly, "Extend Expiration")
                    Else
                        MsgBox("Could Not Find The Related Link.  Please Inform ABS.", vbOKOnly, "Extend Expiration")
                    End If
                Else
                    If grd.Selected.Rows.Count > 1 Then
                        MsgBox("You Can Only Update One Link At A Time.", vbOKOnly, "Extend Expiration")
                    Else
                        MsgBox("You Select A Row To Update.", vbOKOnly, "Extend Expiration")
                    End If
                End If
        End Select
    End Sub

#End Region

#Region "Form Methods"
    Private Sub optSelectBy_ValueChanged(sender As Object, e As EventArgs) Handles optSelectBy.ValueChanged
        If optSelectBy.Value = "G" Then
            dteShip_Beg.Value = Null
            dteShip_End.Value = Null
            dteShip_Beg.ReadOnly = True
            dteShip_End.ReadOnly = True
            If txtCUST_CODE.Text.Length > 0 Then
                FillShippedData()
            End If
        Else
            dteShip_Beg.Value = DateSerial(Now.Year, Now.Month, 1)
            dteShip_End.Value = DateSerial(Now.Year, Now.Month, 1).AddMonths(1).AddDays(-1)
            dteShip_Beg.ReadOnly = False
            dteShip_End.ReadOnly = False
            If dst.Tables.Contains("SOTINVHX") Then
                dst.Tables.Item("SOTINVHX").Clear()
            End If
        End If
    End Sub

    Private Sub txtCUST_CODE_Leave(sender As Object, e As EventArgs) Handles txtCUST_CODE.Leave
        dst.Tables.Item("SOTINVHX").Clear()
        If txtCUST_CODE.Text.Length > 0 Then
            If isValidCustomer() Then
                If optSelectBy.Value = "G" Then
                    FillShippedData()
                End If
            Else
                MsgBox(String.Format("Invalid Customer: {0}", txtCUST_CODE.Text))
                txtCUST_CODE.Text = ""
            End If
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub Build_Init_Sel()

        S = BUILD_SOTCUSTS(True)
        ASCMAIN1.sql = S.ToString
        If IsNothing(dst.Tables.Item("SOTCUSTS")) Then
            Create_TDA(dst.Tables.Add, "SOTCUSTS", "**", 0, False)
        End If

        SQL_REPORT.Length = 0
        SQL_REPORT.AppendLine("SELECT")
        SQL_REPORT.AppendLine("'0' AS SEL,")
        SQL_REPORT.AppendLine("I1.CUST_CODE,")
        SQL_REPORT.AppendLine("I1.ORDR_CUST_PO,")
        SQL_REPORT.AppendLine("I1.INV_DATE,")
        SQL_REPORT.AppendLine("O1.ORDR_GROUP_NO,")
        SQL_REPORT.AppendLine("SUM(I1.INV_TOTAL_AMOUNT) INV_TOTAL_AMOUNT")
        SQL_REPORT.AppendLine("FROM SOTINVH1 I1, SOTORDR1 O1")
        SQL_REPORT.AppendLine("WHERE I1.ORDR_NO = O1.ORDR_NO")
        SQL_REPORT.AppendLine("AND I1.CUST_CODE = :PARM1")
        SQL_REPORT.AppendLine("AND I1.INV_TYPE = 'I'")
        SQL_REPORT.AppendLine("GROUP BY I1.CUST_CODE,")
        SQL_REPORT.AppendLine("I1.ORDR_CUST_PO,")
        SQL_REPORT.AppendLine("I1.INV_DATE,")
        SQL_REPORT.AppendLine("O1.ORDR_GROUP_NO")
        If IsNothing(dst.Tables.Item("SOTINVHX")) Then
            ASCMAIN1.sql = SQL_REPORT.ToString
            Create_TDA(dst.Tables.Add, "SOTINVHX", "**", 0, False, "V")
        End If

        SQL_REPORT.Length = 0
        SQL_REPORT.AppendLine("SELECT *")
        SQL_REPORT.AppendLine("FROM WEBLINKS")
        SQL_REPORT.AppendLine(String.Format("WHERE FORM_NAME = '{0}'", REPORT_NAME))
        If IsNothing(dst.Tables.Item("WEBLINKS")) Then
            ASCMAIN1.sql = SQL_REPORT.ToString
            Create_TDA(dst.Tables.Add, "WEBLINKS", "**", 0, True)
        End If
    End Sub

    Private Function BUILD_SOTCUSTS(ByVal initTable As Boolean) As StringBuilder
        Dim RetVal As New StringBuilder With {.Length = 0}
        RetVal.Length = 0
        RetVal.AppendLine("SELECT")
        RetVal.AppendLine("JN.CUST_CODE,")
        RetVal.AppendLine("JN.STYLE_CODE,")
        RetVal.AppendLine("JN.COLOR_CODE,")
        RetVal.AppendLine("JN.STYLE_DESC,")
        RetVal.AppendLine("JN.COLOR_DESC,")
        RetVal.AppendLine("JN.FABRIC_CODE,")
        RetVal.AppendLine("JN.SEASON_CODE,")
        RetVal.AppendLine("JN.SUB_BODY_CODE,")
        RetVal.AppendLine("JN.SALES_DIVISION_CODE,")
        RetVal.AppendLine("JN.INNER_PACK_QTY,")
        RetVal.AppendLine("JN.CARTON_PACK_QTY,")
        RetVal.AppendLine("JN.STYLE_CUST_CODE,")
        RetVal.AppendLine("JN.IMAGE_NAME,")
        RetVal.AppendLine("MIN(JN.MIN_PRICE) AS MIN_PRICE,")
        RetVal.AppendLine("MAX(JN.MAX_PRICE) AS MAX_PRICE,")
        If chkAveragePrice.Checked Then
            RetVal.AppendLine("ROUND((SUM(VAL_SHP) / SUM(QTY_SHP)),2) AS AVG_PRICE,")
        Else
            RetVal.AppendLine("0 AS AVG_PRICE,")
        End If
        RetVal.AppendLine("SUM(JN.QTY_SHP) AS QTY_SHP,")
        RetVal.AppendLine("SUM(JN.VAL_SHP) AS VAL_SHP")
        RetVal.AppendLine("FROM")
        RetVal.AppendLine("(")
        RetVal.AppendLine(" SELECT")
        RetVal.AppendLine(" I1.CUST_CODE,")
        RetVal.AppendLine(" I2.STYLE_CODE,")
        RetVal.AppendLine(" I2.COLOR_CODE,")
        RetVal.AppendLine(" S1.STYLE_DESC,")
        RetVal.AppendLine(" C1.COLOR_DESC,")
        RetVal.AppendLine(" S1.FABRIC_CODE,")
        RetVal.AppendLine(" S1.SEASON_CODE,")
        RetVal.AppendLine(" S1.SUB_BODY_CODE,")
        RetVal.AppendLine(" S1.SALES_DIVISION_CODE,")
        RetVal.AppendLine(" S1.INNER_PACK_QTY,")
        RetVal.AppendLine(" S1.CARTON_PACK_QTY,")
        RetVal.AppendLine(" S1.CUST_CODE AS STYLE_CUST_CODE,")
        RetVal.AppendLine(" S1.IMAGE_NAME,")
        RetVal.AppendLine(" MIN(I2.ORDR_UNIT_PRICE) AS MIN_PRICE,")
        RetVal.AppendLine(" MAX(I2.ORDR_UNIT_PRICE) AS MAX_PRICE,")
        RetVal.AppendLine(" SUM(I2.ORDR_QTY_SHIP) AS QTY_SHP,")
        RetVal.AppendLine(" SUM(I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE) AS VAL_SHP")
        RetVal.AppendLine(" FROM SOTINVH1 I1, SOTINVH2 I2, ICTSTYL1 S1, ICTCOLR1 C1")
        RetVal.AppendLine(" WHERE I1.INV_TYPE = I2.INV_TYPE")
        RetVal.AppendLine(" AND I1.INV_NO = I2.INV_NO")
        RetVal.AppendLine(" AND I2.STYLE_CODE = S1.STYLE_CODE")
        RetVal.AppendLine(" AND I2.COLOR_CODE = C1.COLOR_CODE")
        RetVal.AppendLine(" AND I1.INV_TYPE = 'I'")
        If initTable Then
            RetVal.AppendLine(" AND I1.CUST_CODE = 'NOT_ON_FILE'")
        Else
            If optSelectBy.Value = "D" Then
                RetVal.AppendLine(String.Format(" AND I1.CUST_CODE = '{0}'", txtCUST_CODE.Text))
                RetVal.AppendLine(String.Format(" AND INV_DATE >= '{0}'", Format(dteShip_Beg.DateTime, "dd-MMM-yy")))
                RetVal.AppendLine(String.Format(" AND INV_DATE <= '{0}'", Format(dteShip_End.DateTime, "dd-MMM-yy")))
            Else
                RetVal.AppendLine(String.Format(" AND I1.ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1 WHERE ORDR_GROUP_NO IN ({0}))", GRP_IN))
            End If
        End If
        RetVal.AppendLine(" GROUP BY")
        RetVal.AppendLine(" I1.CUST_CODE,")
        RetVal.AppendLine(" I2.STYLE_CODE,")
        RetVal.AppendLine(" I2.COLOR_CODE,")
        RetVal.AppendLine(" S1.STYLE_DESC,")
        RetVal.AppendLine(" C1.COLOR_DESC,")
        RetVal.AppendLine(" S1.FABRIC_CODE,")
        RetVal.AppendLine(" S1.SEASON_CODE,")
        RetVal.AppendLine(" S1.SUB_BODY_CODE,")
        RetVal.AppendLine(" S1.SALES_DIVISION_CODE,")
        RetVal.AppendLine(" S1.INNER_PACK_QTY,")
        RetVal.AppendLine(" S1.CARTON_PACK_QTY,")
        RetVal.AppendLine(" S1.CUST_CODE,")
        RetVal.AppendLine(" S1.IMAGE_NAME")
        RetVal.AppendLine(") JN")
        RetVal.AppendLine(String.Format("WHERE NVL(JN.CUST_CODE,'NULL') = '{0}'", txtCUST_CODE.Text))
        RetVal.AppendLine("AND NVL(QTY_SHP,0) > 0")
        If Absx1.optFor("OPTASN").Value = "S" Then
            RetVal.AppendLine("AND JN.STYLE_CUST_CODE IS NULL")
        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
            RetVal.AppendLine("AND JN.STYLE_CUST_CODE IS NOT NULL")
        End If
        RetVal.AppendLine(sql_WHERE)
        RetVal.AppendLine("GROUP BY")
        RetVal.AppendLine("JN.CUST_CODE,")
        RetVal.AppendLine("JN.STYLE_CODE,")
        RetVal.AppendLine("JN.COLOR_CODE,")
        RetVal.AppendLine("JN.STYLE_DESC,")
        RetVal.AppendLine("JN.COLOR_DESC,")
        RetVal.AppendLine("JN.FABRIC_CODE,")
        RetVal.AppendLine("JN.SEASON_CODE,")
        RetVal.AppendLine("JN.SUB_BODY_CODE,")
        RetVal.AppendLine("JN.SALES_DIVISION_CODE,")
        RetVal.AppendLine("JN.INNER_PACK_QTY,")
        RetVal.AppendLine("JN.CARTON_PACK_QTY,")
        RetVal.AppendLine("JN.STYLE_CUST_CODE,")
        RetVal.AppendLine("JN.IMAGE_NAME")
        If Not IsNothing(sql_GROUP_BY_cols) Then
            If sql_GROUP_BY_cols.Length > 0 Then
                RetVal.AppendLine("ORDER BY")
                RetVal.AppendLine(sql_GROUP_BY_cols)
            End If
        End If
        Return RetVal
    End Function

    Private Sub BuildSpecialWhere()
        Dim SWO As String = ""
        Dim SWR As String = ""
        Dim Filter As String = "SEL = '1'"
        Dim rowCnt As Int64 = 0
        For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select(Filter)
            rowCnt += 1
            Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO") + String.Empty
            If rowCnt = 1 Then
                SWO = String.Format("JN.ORDR_GROUP_NO IN ('{0}'", ORDR_GROUP_NO)
            Else
                SWO = SWO & String.Format(",'{0}'", ORDR_GROUP_NO)
            End If
        Next
        If rowCnt > 0 Then
            SWO = SWO & ")"
        End If

        rowCnt = 0
        For Each rowSOTRSRV1 As DataRow In dst.Tables("SOTRSRV1").Select(Filter)
            rowCnt += 1
            Dim RSRV_NO As String = rowSOTRSRV1.Item("RSRV_NO") + String.Empty
            If rowCnt = 1 Then
                SWR = String.Format("JN.RSRV_NO IN ('{0}'", RSRV_NO)
            Else
                SWR = SWR & String.Format(",'{0}'", RSRV_NO)
            End If
        Next
        If rowCnt > 0 Then
            SWR = SWR & ")"
        End If

        SQL_WHERE_ORDR_RSV = ""
        If SWO.Length > 0 And SWR.Length > 0 Then
            SQL_WHERE_ORDR_RSV = String.Format("AND ({0} OR {1})", SWO, SWR)
        Else
            If SWO.Length > 0 Then
                SQL_WHERE_ORDR_RSV = String.Format("AND {0}", SWO)
            Else
                If SWR.Length > 0 Then
                    SQL_WHERE_ORDR_RSV = String.Format("AND {0}", SWR)
                End If
            End If
        End If
    End Sub

    Private Function EXTEND_LINK(ByVal HASHVALUE As String) As Boolean
        Dim RetVal As Boolean = False
        For Each TABLE_NAME As String In New String() {"ASTATTA2", "ICTQUOH2", "WEBLINKS"}
            If RetVal = False Then
                Dim SQLS As New Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine(String.Format("SELECT COUNT(*) FROM {0} WHERE HASHVALUE = '{1}'", TABLE_NAME, HASHVALUE))
                ASCMAIN1.sql = SQLS.ToString()
                Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                If REC_CNT > 0 Then
                    RetVal = True
                    Dim SQLE As New System.Text.StringBuilder With {.Length = 0}
                    Dim NEW_DATE As String = Format(Now(), "dd-MMM-yyyy")
                    SQLE.AppendLine(String.Format("UPDATE {0} SET NEW_HASH_EXP = '{1}' WHERE HASHVALUE = '{2}'", TABLE_NAME, NEW_DATE, HASHVALUE))
                    ASCMAIN1.sql = SQLE.ToString
                    ASCDATA1.ExecuteSQL()
                End If
            End If
        Next
        Return RetVal
    End Function

    Private Sub FillShippedData()
        Fill_Records("SOTINVHX", txtCUST_CODE.Text)
        Sort_grdColumns(grdSOTINVHX, "INV_DATE".ToLower())
    End Sub

    Private Sub FixSqlWhere(ByRef SQL_FIX As String)
        Dim FIXES As New Dictionary(Of String, String)
        FIXES.Add("ICTSTYL1.CUST_CODE", "JN.CUST_CODE")
        FIXES.Add("ICTSTYL1.STYLE_CODE", "JN.STYLE_CODE")
        FIXES.Add("ICTSTYL1.FABRIC_CODE", "JN.FABRIC_CODE")
        FIXES.Add("ICTSTYL1.SALES_DIVISION_CODE", "JN.SALES_DIVISION_CODE")
        FIXES.Add("ICTSTYL1.SEASON_CODE", "JN.SEASON_CODE")
        FIXES.Add("ICTSTYL1.SUB_BODY_CODE", "JN.SUB_BODY_CODE")
        FIXES.Add("SOTORDR0.ORDR_GROUP_NO", "JN.ORDR_GROUP_NO")
        FIXES.Add("SOTRSRV1.RSRV_NO", "JN.RSRV_NO")
        For Each FX As KeyValuePair(Of String, String) In FIXES
            SQL_FIX = SQL_FIX.Replace(FX.Key, FX.Value)
        Next
    End Sub

    Private Sub FixSqlGroup(ByRef SQL_FIX As String)
        Dim FIXES As New Dictionary(Of String, String)
        FIXES.Add("ICTSTYL1.", "JN.")
        FIXES.Add("SOTORDR0.", "JN.")
        For Each FX As KeyValuePair(Of String, String) In FIXES
            SQL_FIX = SQL_FIX.Replace(FX.Key, FX.Value)
        Next
    End Sub

    Private Function getCostForStyleColor(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As Double
        Dim Retval As Double = 0
        ASCMAIN1.sql = "Select STYLE_COST from (" & vbCrLf _
                            & "Select STYLE_COST from ICTCOSTA " & vbCrLf _
                            & "where (STYLE_CODE, COLOR_CODE) in (" & vbCrLf _
                            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                            & " from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'" _
                            & " and WHSE_QTY_ON_HAND > 0)" & vbCrLf _
                            & " order by OPS_YYYYPP DESC) where ROWNUM < 2"
        Dim STYLE_COST As Decimal = Val(ASCDATA1.GetDataValue)

        If STYLE_COST = 0 Then
            ASCMAIN1.sql = "Select NVL(PO_COST_LANDED,PO_COST) STYLE_COST" & vbCrLf _
                                & " from (" & vbCrLf _
                                & " Select POTSHIP3.PO_SHIPMENT_NO, POTORDR2.PO_ORDER_NO, " & vbCrLf _
                                & " POTORDR2.PO_COST, POTSHIP3.PO_COST_LANDED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
                                & " from POTORDR2,POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
                                & " where POTORDR2.STYLE_CODE = '" & STYLE_CODE & "' and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                                & "   and POTSHIP3.PO_ORDER_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
                                & "   and POTSHIP3.PO_ORDER_LNO (+) = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_LNO (+) = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                                & "   and POTSHIP1.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                                & " order by POTSHIP3.PO_SHIPMENT_NO DESC, POTORDR2.PO_ORDER_NO DESC" & vbCrLf _
                                & ") where ROWNUM <2"
            STYLE_COST = Val(ASCDATA1.GetDataValue)
        End If

        If STYLE_COST <> 0 Then
            Retval = Math.Round(STYLE_COST, 2)
        End If

        Return Retval
    End Function

    Private Function GetCustShipDates(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim RetVal As String = ""
        Dim SQLS As New StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("MIN(S1.INV_DATE) AS MIN_INV_DATE")
        SQLS.AppendLine("FROM SOTINVH1 S1, SOTINVH2 S2")
        SQLS.AppendLine("WHERE S1.INV_NO = S2.INV_NO")
        SQLS.AppendLine("AND S1.INV_TYPE = S2.INV_TYPE")
        SQLS.AppendLine(String.Format("AND S2.STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND S2.COLOR_CODE = '{0}'", COLOR_CODE))
        SQLS.AppendLine(String.Format("AND S1.CUST_CODE = '{0}'", txtCUST_CODE.Text))
        SQLS.AppendLine(String.Format(" AND S1.INV_DATE >= '{0}'", Format(dteShip_Beg.DateTime, "dd-MMM-yy")))
        SQLS.AppendLine(String.Format(" AND S1.INV_DATE <= '{0}'", Format(dteShip_End.DateTime, "dd-MMM-yy")))
        ' new DGJ 2 LINES ABOVE Only consider Shipments in Date Range


        ASCMAIN1.sql = SQLS.ToString()
        Dim MIN_INV_DATE As String = ASCDATA1.GetDataValue

        SQLS.Length = 0
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("MAX(S1.INV_DATE) AS MAX_INV_DATE")
        SQLS.AppendLine("FROM SOTINVH1 S1, SOTINVH2 S2")
        SQLS.AppendLine("WHERE S1.INV_NO = S2.INV_NO")
        SQLS.AppendLine("AND S1.INV_TYPE = S2.INV_TYPE")
        SQLS.AppendLine(String.Format("AND S2.STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND S2.COLOR_CODE = '{0}'", COLOR_CODE))
        SQLS.AppendLine(String.Format("AND S1.CUST_CODE = '{0}'", txtCUST_CODE.Text))
        SQLS.AppendLine(String.Format(" AND S1.INV_DATE >= '{0}'", Format(dteShip_Beg.DateTime, "dd-MMM-yy")))
        SQLS.AppendLine(String.Format(" AND S1.INV_DATE <= '{0}'", Format(dteShip_End.DateTime, "dd-MMM-yy")))
        ' new DGJ 2 LINES ABOVE Only consider Shipments in Date Range

        ASCMAIN1.sql = SQLS.ToString()
        Dim MAX_INV_DATE As String = ASCDATA1.GetDataValue

        If IsDate(MIN_INV_DATE) And IsDate(MAX_INV_DATE) Then
            RetVal = String.Format("{0} - {1}", Format(CDate(MIN_INV_DATE), "MM/dd/yy"), Format(CDate(MAX_INV_DATE), "MM/dd/yy"))
        End If

        Return RetVal
    End Function

    Private Function GetAltColorCode(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal COLOR_DESC_ORIG As String) As String
        Dim RetVal As String = COLOR_DESC_ORIG
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        Dim SIZE_SCALE As String = rowICTSTYL1.Item("SIZE_SCALE") & String.Empty
        Dim MAX_LENGTH As Integer = 60
        Dim I As Integer = InStr(SIZE_SCALE, COLOR_CODE)
        If I <> 0 Then
            Dim S As String = Trim(Mid(SIZE_SCALE, I + 3))
            Dim J As Integer = InStr(Mid(S & "  ", 1, MAX_LENGTH), "  ")
            Dim K As Integer = InStr(Mid(S & vbCrLf, 1, MAX_LENGTH), vbCrLf)
            If J = 0 And K = 0 Then
                J = InStr(Mid(S & " ", 1, MAX_LENGTH), " ")
            End If
            If J = 0 Or J > K Then J = K
            Dim SC As String = ""
            If J <> 0 Then
                SC = Mid(S, 1, J)
                SIZE_SCALE = Mid(SIZE_SCALE, 1, I - 1) & Mid(S, J)
                For C As Integer = 1 To SC.Length - 1
                    If C = 1 Or (C > 1 AndAlso Mid(SC, C + 1, 1) <> " " AndAlso (Mid(SC, C - 1, 1) = " " Or Mid(SC, C - 1, 1) = "/")) Then
                        Mid(SC, C, 1) = Mid(SC, C, 1).ToUpper
                    End If
                Next
                If Trim(SC) <> "" Then
                    If SC.Length > 35 Then
                        RetVal = SC.Substring(0, 34)
                    Else
                        RetVal = SC
                    End If

                End If
            End If
        End If
        If RetVal = COLOR_DESC_ORIG Then
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT NVL(STYLE_COLOR_DESC,'') STYLE_COLOR_DESC")
            SQLS.AppendLine("FROM ICTSTYC1")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = SQLS.ToString()
            Dim COLOR_DESC_MF As String = ASCDATA1.GetDataValue
            If COLOR_DESC_MF.Length > 35 Then
                COLOR_DESC_MF = COLOR_DESC_MF.Substring(0, 35)
            End If
            If COLOR_DESC_MF.Length > 0 Then
                RetVal = COLOR_DESC_MF
            End If
        End If
        Return RetVal
    End Function

    Private Function isValidCustomer() As Boolean
        Dim RetVal As Boolean = False
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", txtCUST_CODE.Text)
        If Not IsNothing(rowARTCUST1) Then
            RetVal = True
        End If
        Return RetVal
    End Function

    Private Sub UpdateReportRows()

    End Sub

#End Region

#Region "Excel Methods"
    Sub Create_Excel_WorkSheet(worksheet As SpreadsheetGear.IWorksheet,
                               ByVal StyleList As List(Of String), Optional sqlWB As String = "")

        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        If (ASCMAIN1.Running_in_VS) Then
            If Not System.IO.Directory.Exists(IMAGE_FOLDER) Then
                Stop 'You Need to Set up Image Folder.
            End If
        End If

        Dim CX As Integer = 0
        Dim RX As Integer = 0

        Dim I As Integer = 0
        I += 4

        Dim COL0 As Integer = 12

        Dim COL As Integer = COL0

        Excel_DefaultColumns(worksheet, COL)

        With worksheet.Cells(I, 0, I, COL)
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        Dim I0 As Integer = 0
        Dim IA As Integer = 0
        Dim RT(11) As String
        Dim ROW0 As Integer = I
        Dim style_count As Integer = 0
        Dim pages As Integer = 0

        For Each STYLE_CODE As String In StyleList
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            ASCMAIN1.Progress("-", STYLE_CODE)
            I += 1
            I0 = I
            COL = COL0

            Excel_StyleHeader(worksheet, COL, I, COL0)

            I += 1

            Dim ImageRows = 0
            Dim ImageRowsBig = 0
            Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""
            Excel_ImageInsert(worksheet, IMAGE_NAME, IMAGE_FOLDER, ImageRows, ImageRowsBig, I)

            CX = 1

            Excel_StyleMasterfile(worksheet, I, CX, rowICTSTYL1, STYLE_CODE)

            Dim CI As Integer = 0
            Excel_ColorDetails(worksheet, STYLE_CODE, I, COL, COL0, CI)

            For iCOL As Integer = 1 To 3
                COL += 1
                Select Case iCOL
                    Case 1
                        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        'Case 2
                        '    If chkShip2.Checked Then
                        '        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        '    Else
                        '        COL -= 1
                        '    End If
                    Case 3
                        If chkAveragePrice.Checked Then
                            worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        Else
                            COL -= 1
                        End If
                End Select

                RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
            Next

            COL += 1

            Dim colsLess As Int16 = 0
            If chkAveragePrice.Checked = False Then
                colsLess += 1
            End If
            If chkShipDates.Checked = False Then
                colsLess += 1
            End If
            If chkStyleStats.Checked Then
                COL = COL - colsLess
                For iCOL As Integer = 5 To 11
                    COL += 1
                    worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                Next
                COL += 0
            End If

            If chkStyleStats.Checked Then
                worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray
            Else
                worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL - colsLess).Interior.Color = SpreadsheetGear.Colors.LightGray
            End If


            With worksheet.Cells(I, COL0 - 1, I + CI - 1, COL - colsLess)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
            End With

            I += ImageRowsBig

            Dim CJ As Integer = ImageRows

            If CJ < 6 Then CJ = 6

            If CI > CJ Then
                I += CI
            Else
                I += CJ
            End If

            style_count += 1

            If (((I - 5) Mod 80) < ((I0 - 5) Mod 80)) Or (style_count >= 5) Or style_count >= 9 Then
                Dim R As SpreadsheetGear.IRange = worksheet.Cells(I0, 0).EntireRow
                worksheet.HPageBreaks.Add(R)
                style_count = 1
                pages += 1
            End If

            If chkStyleStats.Checked Then
                Dim interior As SpreadsheetGear.IInterior
                Dim range As SpreadsheetGear.IRange
                '  I += 1
                COL = COL0
                Dim chkcnt As Int64 = 0
                Dim NEWSTYLE As Boolean = True


                For Each rowICTSTATD As DataRow In dst.Tables("ICTSTATD").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE, PO_DATE_SHIP_BY")
                    If NEWSTYLE = True Then
                        worksheet.Cells(I - 1, COL - 1).Value = "In-Transit Details"
                        I += 1
                        ' Headinds and headingsFOrmat
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Color"
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Factory"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        End With
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Qty Ord"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Qty Shp"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        End With
                        If chkAveragePrice.Checked Or chkShipDates.Checked Then
                        Else
                            worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        End If

                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Rev Ship Dt"
                        If chkAveragePrice.Checked And chkShipDates.Checked Then
                        Else
                            worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        End If
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "ETA"
                        If chkAveragePrice.Checked And chkShipDates.Checked Then
                        Else
                            worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        End If
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Vessel"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 20
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        End With

                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Shp Dt Rev"
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).ColumnWidth = 15
                        range = worksheet.Cells(I - 1, COL - 1, I - 1, COL + 6)
                        interior = range.Interior
                        interior.Color = SpreadsheetGear.Colors.Aquamarine

                        NEWSTYLE = False
                    End If



                    I += 1
                    chkcnt = 1
                    If sql = sql Then
                        ' avoid printing if no records in ICTSTATD
                        ' worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & "***"

                    End If



                    '  worksheet.Cells(I - 1, COL - 2 + chkcnt).Value = Val(rowICTSTATD.Item(1) & String.Empty)

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = Format(Val(rowICTSTATD.Item("COLOR_CODE") & String.Empty), "000")
                        .Font.Size = 14
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                    End With
                    chkcnt += 1

                    '   worksheet.Cells(I - 1, COL - 2 + chkcnt).Value = Val(rowICTSTATD.Item(4) & String.Empty)

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Value = rowICTSTATD.Item("FACTORY_CODE") & String.Empty
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                        .Font.Size = 14
                    End With
                    chkcnt += 1


                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowICTSTATD.Item("PO_QTY_ORD") & String.Empty)
                        .NumberFormat = "#,##0"
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                        .Font.Size = 14
                    End With
                    chkcnt += 1
                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = Val(rowICTSTATD.Item("PO_QTY_SHP") & String.Empty)
                        .NumberFormat = "#,##0"
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If


                        .Font.Size = 14
                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Value = rowICTSTATD.Item("PO_DATE_SHIP_BY") & String.Empty
                        .NumberFormat = "MM/dd/yy"
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                        .Font.Size = 14
                    End With
                    chkcnt += 1
                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .NumberFormat = "MM/dd/yy"
                        .Value = rowICTSTATD.Item("PO_SHIP_ETA") & String.Empty
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If

                        .Font.Size = 14
                    End With
                    chkcnt += 1

                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Left
                        .Value = rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If
                        .Font.Size = 14
                    End With
                    chkcnt += 1
                    With worksheet.Cells(I - 1, COL - 2 + chkcnt)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .NumberFormat = "MM/dd/yy"
                        .Value = rowICTSTATD.Item("LAST_DATE_SHIP_BY") & String.Empty
                        If rowICTSTATD.Item("PO_SHIP_VESSEL") & String.Empty = "OpenPO" Then
                            .Font.Color = SpreadsheetGear.Colors.Green
                        Else
                            .Font.Color = SpreadsheetGear.Colors.Blue
                        End If
                        .Font.Size = 14
                    End With
                    chkcnt += 1

                Next
                'T = ""
                'COL += 1


            End If





            With worksheet.Cells(I0, 0, I + 1 - 1, COL)
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With
        Next

        I += 2
        COL = COL0

        'Trying to get away without totals here :)
        worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"

        Dim GT = ""
        For iCOL As Integer = 1 To 3
            COL += 1
            Select Case iCOL
                Case 1
                    worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                Case 3
                    If chkAveragePrice.Checked Then
                        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                        GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")

                    Else
                        COL -= 2
                    End If

                    '    If chkShip2.Checked Then
                    '        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    '        GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    '    Else
                    '        COL -= 1
                    '    End If
                    'Case 6
                    '    If chkShip2.Checked Then
                    '        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    '        GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    '    Else
                    '        COL -= 1
                    '    End If
            End Select
        Next

        If chkShipDates.Checked Then
            COL += 1
        End If

        If chkStyleStats.Checked Then
            For iCOL As Integer = 1 To 7
                COL += 1
                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL + 4), 2)
                GT &= "+" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "")
            Next
            COL += 0
        End If

        worksheet.Cells(I - 1, COL0 - 1, I - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

        Excel_Header(worksheet)

        Excel_PageSetup(worksheet)
    End Sub

    Private Sub Excel_ColorDetails(ByRef worksheet As IWorksheet,
                                   ByVal STYLE_CODE As String,
                                   ByRef i As Integer,
                                   ByRef COL As Integer,
                                   ByRef COL0 As Integer,
                                   ByRef CI As Integer)
        Dim SZMAX As Integer = 0
        Dim SZTOT As Integer = 0
        Dim T As String = ""
        Dim styleTotal As Int64 = 0
        Dim LAST_COLOR As String = ""
        For Each rowSOTCUSTS As DataRow In dst.Tables("SOTCUSTS").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE")
            CI += 1
            COL = COL0 + 1
            'COL = COL0
            Dim chkcnt As Int64 = 1
            If LAST_COLOR <> rowSOTCUSTS.Item("COLOR_CODE") & String.Empty Then
                worksheet.Cells(i + CI - 1, COL - 2).Value = "'" & rowSOTCUSTS.Item("COLOR_CODE") & String.Empty
                'worksheet.Cells(i + CI - 1, COL - 1).Value = rowSOTCUSTS.Item("COLOR_DESC") & String.Empty
                worksheet.Cells(i + CI - 1, COL - 1).Value = GetAltColorCode(STYLE_CODE, rowSOTCUSTS.Item("COLOR_CODE") & String.Empty, rowSOTCUSTS.Item("COLOR_DESC") & String.Empty)
                LAST_COLOR = rowSOTCUSTS.Item("COLOR_CODE") & String.Empty
            End If
            'worksheet.Cells(i + CI - 1, COL + 1).Value = "ORDR_CUST_PO" 'rowSOTCUSTQ.Item("ORDR_CUST_PO") & String.Empty
            'worksheet.Cells(i + CI - 1, COL + 2).Value = "ORDR_SHIP_DATE" 'rowSOTCUSTQ.Item("ORDR_SHIP_DATE") & String.Empty
            'worksheet.Cells(i + CI - 1, COL + 3).Value = "ORDR_CANCEL_DATE" 'rowSOTCUSTQ.Item("ORDR_CANCEL_DATE") & String.Empty
            worksheet.Cells(i + CI - 1, COL).Value = rowSOTCUSTS.Item("QTY_SHP") & String.Empty
            'chkcnt += 1

            If chkAveragePrice.Checked Then
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTS.Item("AVG_PRICE") & String.Empty
                chkcnt += 1
            End If


            If chkAveragePrice.Checked Then
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTS.Item("VAL_SHP") & String.Empty
                chkcnt += 1
            End If

            If chkShipDates.Checked Then
                'worksheet.Cells(i + CI - 1, COL + chkcnt).Value = "1/1/2018 - 12/31/2018"
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = GetCustShipDates(rowSOTCUSTS.Item("STYLE_CODE") & String.Empty, rowSOTCUSTS.Item("COLOR_CODE") & String.Empty)
                chkcnt += 1
            End If

            'If chkShip2.Checked Then
            '    worksheet.Cells(i + CI - 1, COL + 1).Value = rowSOTCUSTS.Item("QTY_SHP_02") & String.Empty
            '    'chkcnt += 1
            'End If
            'If chkShip3.Checked Then
            '    worksheet.Cells(i + CI - 1, COL + 2).Value = rowSOTCUSTS.Item("QTY_SHP_03") & String.Empty
            '    'chkcnt += 1
            'End If
            'T = ""
            'COL += 1

            If chkStyleStats.Checked Then

                ' ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty & "'"
                '       For Each rowICTSTAT2 As DataRow In ASCDATA1.GetDataTable.Select("")
                For Each rowICTSTAT2 As DataRow In dst.Tables("ICTSTAT2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowSOTCUSTS.Item("COLOR_CODE") & String.Empty & "'")


                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & String.Empty)
                    chkcnt += 1
                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & String.Empty)
                    chkcnt += 1
                    Dim OTS As Integer = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "") - Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & "")
                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = OTS
                    chkcnt += 1

                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty)
                    chkcnt += 1

                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty)
                    chkcnt += 1

                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
                    chkcnt += 1

                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = OTS + Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty) + Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty) - Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
                    chkcnt += 1

                Next
                T = ""
                COL += 1
            End If


        Next

        CI += 2
        COL = COL0

        worksheet.Cells(i - 1, COL - 1, i + CI - 1, COL - 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & "***"
        worksheet.Cells(i + CI - 1, COL - 0).Value = "'" & "Total"
    End Sub

    Private Sub Excel_StyleMasterfile(ByRef worksheet As IWorksheet, ByRef i As Integer, ByRef cx As Integer, ByRef rowICTSTYL1 As DataRow, ByVal STYLE_CODE As String)
        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange

        With worksheet.Cells(i - 1, 3)
            .Value = "'" & STYLE_CODE
            .Font.Color = SpreadsheetGear.Colors.Purple
            .Font.Size = 24
            .Font.Bold = True
        End With

        cx = 3

        worksheet.Cells(i + 2, cx).Value = "Case Qty"

        range = worksheet.Cells(i + 1, 3, i + 2, 4)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.LightGray

        range = worksheet.Cells(i + 1, 3 + 4, i + 2, 4 + 4)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.LightGray

        cx = 5
        worksheet.Cells(i, cx - 2).Value = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
        worksheet.Cells(i + 2, cx).Value = rowICTSTYL1.Item("CARTON_PACK_QTY")
    End Sub

    Private Sub Excel_ImageInsert(ByRef worksheet As IWorksheet,
                                  ByVal iMAGE_NAME As String,
                                  ByVal IMAGE_FOLDER As String,
                                  ByRef ImageRows As Integer,
                                  ByRef ImageRowsBig As Integer,
                                  ByRef i As Integer)
        Dim imageFileStyle As String = IMAGE_FOLDER & "\" & iMAGE_NAME
        If Not System.IO.File.Exists(imageFileStyle) Then
            iMAGE_NAME = ""
        End If

        If iMAGE_NAME <> "" _
                AndAlso My.Computer.FileSystem.FileExists(imageFileStyle) Then

            Dim widthStyle As Double
            Dim heightStyle As Double

            Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
            Try
                widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution / 3
                heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution / 3
            Finally
                imageStyle.Dispose()
            End Try

            Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

            Dim col_adj As Decimal = 0
            If heightStyle > widthStyle Then
                col_adj = 0.3
            Else
                col_adj = 0.05
            End If

            Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(0) + col_adj
            Dim topStyle As Double = windowInfoStyle.RowToPoints(i - 1) + 0.1

            ImageRows = windowInfoStyle.PointsToRow(heightStyle)
            worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
        End If
    End Sub

    Private Sub Excel_StyleHeader(ByRef worksheet As IWorksheet, ByRef COL As Integer, ByRef i As Integer, ByVal COL0 As Integer)
        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange

        worksheet.Cells(i, COL - 1).Value = "" & Chr(13) & Chr(10) & "Color"
        worksheet.Cells(i, COL - 1).Font.Size = 12
        worksheet.Cells(i, COL).Value = "" & Chr(13) & Chr(10) & "Description"
        worksheet.Cells(i, COL).Font.Size = 12

        'COL += 1
        'With worksheet.Cells(i, COL)
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .Value = "PO"
        'End With

        'COL += 1
        'With worksheet.Cells(i, COL)
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .Value = "Ship"
        'End With

        'COL += 1
        'With worksheet.Cells(i, COL)
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .Value = "Cancel"
        'End With

        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            Dim d1TEXT As String = Format(dteShip_Beg.DateTime, "MM/dd/yy")
            Dim d2TEXT As String = Format(dteShip_End.DateTime, "MM/dd/yy")
            If optSelectBy.Value = "D" Then
                .Value = String.Format("{0} to {1}", d1TEXT, d2TEXT) & Chr(13) & Chr(10) & "Shp Units"
            Else
                .Value = "Selected POs" & Chr(13) & Chr(10) & "Shp Units"
            End If
            .Font.Size = 12
        End With

        If chkAveragePrice.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "Price"
                .Font.Size = 12
            End With
        End If

        If chkAveragePrice.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                If optSelectBy.Value = "D" Then
                    .Value = "" & Chr(13) & Chr(10) & "Shp Amt"
                Else
                    .Value = ""
                End If

                .Font.Size = 12
            End With
        End If

        'If chkShip2.Checked Then
        '    COL += 1
        '    With worksheet.Cells(i, COL)
        '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '        Dim d1TEXT As String = Format(dteShip_Beg_2.DateTime, "MM/dd/yy")
        '        Dim d2TEXT As String = Format(dteShip_End_2.DateTime, "MM/dd/yy")
        '        .Value = String.Format("{0} to {1}", d1TEXT, d2TEXT)
        '        .Font.Size = 12
        '    End With
        'End If

        If chkShipDates.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "1st & Last Ship Dates"
                .Font.Size = 12
            End With
        End If

        'If chkShip3.Checked Then
        '    COL += 1
        '    With worksheet.Cells(i, COL)
        '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '        Dim d1TEXT As String = Format(dteShip_Beg_3.DateTime, "MM/dd/yy")
        '        Dim d2TEXT As String = Format(dteShip_End_3.DateTime, "MM/dd/yy")
        '        If chkLastOnFile.Checked Then
        '            d2TEXT = "Last"
        '        End If
        '        .Value = String.Format("{0} to {1}", d1TEXT, d2TEXT)
        '        .Font.Size = 12
        '    End With
        'End If

        range = worksheet.Cells(i, COL0 - 1, i, COL)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.Gold

        If chkStyleStats.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "On Hand"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "In Pick"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "OTS"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "In Transit"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "WIP"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "Open"
            End With
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "Net Pos"
            End With

            range = worksheet.Cells(i, COL - 6, i, COL)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.Aquamarine

        End If




    End Sub

    Private Sub Excel_Header(worksheet As IWorksheet)
        Dim H0 As Integer = 8 + 6

        worksheet.Cells(0, H0).Value = "Prep"
        worksheet.Cells(1, H0).Value = "By"
        worksheet.Cells(2, H0).Value = "XNo"

        worksheet.Cells(0, H0, 2, H0).Interior.Color = SpreadsheetGear.Colors.LightGray


        worksheet.Cells(0, H0 + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Left
        worksheet.Cells(0, H0 + 1).Value = Now
        worksheet.Cells(0, H0 + 1).NumberFormat = "MM/dd/yy"

        worksheet.Cells(1, H0 + 1).Value = ASCMAIN1.USER_ID
        worksheet.Cells(2, H0 + 1).Value = "'" & Mid(XNO, 5)

        With worksheet.Cells(0, H0, 2, H0 + 1)
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Font.Color = SpreadsheetGear.Colors.Black
            .Font.Size = 10
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        Dim H1 As Integer = 11
        Dim HEAD1 As String = ""
        Dim HEAD2 As String = ""
        If optASN.Value = "S" Then
            HEAD1 = "Stock"
        ElseIf optASN.Value = "N" Then
            HEAD1 = "NonStock"
        Else
            HEAD1 = "All Styles"
        End If

        'If chkOpen.Checked Then
        '    HEAD2 = "Open"
        'End If
        'If chkPick.Checked Then
        '    If HEAD2 = "" Then
        '        HEAD2 = "Pick"
        '    Else
        '        HEAD2 = HEAD2 & "," & "Pick"
        '    End If
        'End If
        'If chkReservations.Checked Then
        '    If HEAD2 = "" Then
        '        HEAD2 = "Res"
        '    Else
        '        HEAD2 = HEAD2 & "," & "Res"
        '    End If
        'End If

        worksheet.Cells(0, 2).Value = "Customer Shipped Report with Pictures"
        worksheet.Cells(0, 2).Font.Bold = True
        worksheet.Cells(1, 2).Value = "Customer: " & txtCUST_CODE.Text & "   Styles: " & HEAD1
        worksheet.Cells(1, 2).Font.Bold = True
        If optSelectBy.Value = "D" Then
            worksheet.Cells(2, 2).Value = "Ship Date Range: " & dteShip_Beg.Value & " - " & dteShip_End.Value
            worksheet.Cells(2, 2).Font.Bold = True
        Else
            worksheet.Cells(2, 2).Value = "Report By Selected PO's"
            worksheet.Cells(2, 2).Font.Bold = True
        End If


        worksheet.Cells(0, H1).Value = "Note"
        worksheet.Cells(1, H1).Value = "For"

        worksheet.Cells(0, H1, 2, H1).Interior.Color = SpreadsheetGear.Colors.LightGray

        worksheet.Cells(0, H1 + 1).NumberFormat = "MM/dd/yy"
        worksheet.Cells(0, H1 + 1).Value = "Notes"
        worksheet.Cells(1, H1 + 1).Value = txtCUST_CODE.Text & String.Empty

        With worksheet.Cells(0, H1, 2, H1 + 2)
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Font.Color = SpreadsheetGear.Colors.Black
            .Font.Size = 10
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        With worksheet.Cells(3, 3)
            .Font.Color = SpreadsheetGear.Colors.Purple
            .Font.Size = 20
            .Font.Bold = True
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With
    End Sub

    Private Sub Excel_PageSetup(ByRef worksheet As IWorksheet)
        With worksheet.PageSetup
            .TopMargin = 0.25
            .LeftMargin = 0.25
            .RightMargin = 0.25
            .BottomMargin = 0.25
            .FitToPagesWide = 1
            .FitToPagesTall = Nothing
            .PrintTitleRows = "A1:S5"
            .CenterFooter = "&P"
        End With
    End Sub

    Private Sub Excel_DefaultColumns(ByRef worksheet As IWorksheet, ByRef COL As Int64)
        worksheet.Cells("A1:Z1").EntireColumn.Font.Size = 16

        Dim CWC() As String = Split("A,B, C,D,E,F,G,H,I,J,K,L, M", ",")
        Dim CWS() As String = Split("1,1,40,6,6,6,6,6,6,6,6,6,20", ",")
        CWS(2) = 45
        For CWCi As Integer = 0 To CWC.Length - 1
            worksheet.Cells(Trim(CWC(CWCi)) & "1").EntireColumn.ColumnWidth = Val(CWS(CWCi))
        Next

        worksheet.Cells(0, 0).EntireColumn.Hidden = True
        worksheet.Cells(0, 1).EntireColumn.Hidden = True

        Dim _COL As Int64 = 1
        ''PO Column
        'COL += 1
        'With worksheet.Cells(_COL, COL)
        '    .ColumnWidth = 20
        '    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        'End With

        ''Ship Date Column
        'COL += 1
        '_COL += 1
        'With worksheet.Cells(_COL, COL)
        '    .ColumnWidth = 15
        '    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        '    .EntireColumn.NumberFormat = "MM/dd/yy"
        'End With

        ''Cancel Date Column
        'COL += 1
        '_COL += 1
        'With worksheet.Cells(_COL, COL)
        '    .ColumnWidth = 15
        '    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '    .EntireColumn.NumberFormat = "MM/dd/yy"
        'End With

        'Ship 1 Column
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 25
            .EntireColumn.NumberFormat = "#,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With

        If chkAveragePrice.Checked Then
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 12
                .EntireColumn.NumberFormat = "#,##0.00"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 12
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

        End If

        If chkShipDates.Checked Then
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 30
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
        End If

        'Ship 2 Column
        'If chkShip2.Checked Then
        '    COL += 1
        '    _COL += 1
        '    With worksheet.Cells(_COL, COL)
        '        .ColumnWidth = 25
        '        .EntireColumn.NumberFormat = "#,##0"
        '        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '    End With
        'End If

        'Ship 3 Column
        'If chkShip3.Checked Then
        '    COL += 1
        '    _COL += 1
        '    With worksheet.Cells(_COL, COL)
        '        .ColumnWidth = 25
        '        .EntireColumn.NumberFormat = "#,##0"
        '        .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '    End With

        'End If

        If chkStyleStats.Checked Then
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 17
                .EntireColumn.NumberFormat = "#,###,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
        End If


    End Sub
    Private Function Create_Excel(Optional SALES_DIVISION_CODE As String = "") As String
        Dim RetVal As String = ""



        ''  RESEQ()

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim sqlWB As String = ""
        If SALES_DIVISION_CODE <> "" Then
            sqlWB = " and SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
            ASCMAIN1.Progress("Now Creating Workbook for Divison " & SALES_DIVISION_CODE, "")
        Else
            ASCMAIN1.Progress("Now Creating Workbook", "")
        End If
        Dim sql0 As String = ""
        ''  Dim sql0 As String = " and COUNT_COLOR > 0" ' & Val(numMinQty.Value & "")
        ''If chkShowSelectedOnly.Checked Then
        ''    sql0 &= " and SELECTED = '1'"
        ''End If


        ''  CUSTPOSs.Clear()

        Dim CUSTPOi As Integer = 0
        ''dst.Tables("SOTORDRC").Rows.Clear()

        ''For Each row As DataRow In dst.Tables("ICTSTYC1").Select("")
        ''    row.Item("OPEN_PICK_RSRV") = 0
        ''Next

        ''If chkShowPOs.Checked Then
        ''    For Each row As DataRow In dst.Tables("ICTQUOT2").Select("")
        ''        STYLE_CODE = row.Item("STYLE_CODE_PLM")
        ''        Fill_Records("SOTORDRC", New String() {txtQuoteCUST_CODE.Text, STYLE_CODE}, False)
        ''    Next
        ''    For Each row As DataRow In dst.Tables("SOTORDRC").Select("", "ORDR_CANCEL_DATE")
        ''        Dim OPO As String = row.Item("ORDR_TYPE") & vbTab & row.Item("ORDR_CUST_PO") & vbTab & Format(row.Item("ORDR_SHIP_DATE"), "MM/dd/yyyy") & vbTab & Format(row.Item("ORDR_CANCEL_DATE"), "MM/dd/yyyy")
        ''        If Not CUSTPOSs.ContainsKey(OPO) Then
        ''            CUSTPOi += 1
        ''            CUSTPOSs.Add(OPO, CUSTPOi)
        ''        End If
        ''        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
        ''        Dim COLOR_CODE As String = row.Item("COLOR_CODE")
        ''        Dim QTY As Int64 = Val(row.Item("QTY") & "")
        ''        Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
        ''        If rowICTSTYC1 IsNot Nothing Then
        ''            rowICTSTYC1.Item("OPEN_PICK_RSRV") = Val(rowICTSTYC1.Item("OPEN_PICK_RSRV") & "") + QTY
        ''        End If
        ''    Next
        ''End If

        Dim XLS_CREATED As Boolean = False

        If chk1Sheet.Checked Then
            Dim wsi As Integer = 0
            'Dim WJZ As Integer = dst.Tables("ICTQUOT2").Rows.Count

            Dim CODES As String = ""
            ''If opt1Sheet.Value = "S" Then
            ''    CODES = "SUB_BODY_CODE"
            ''ElseIf opt1Sheet.Value = "FS" Then
            ''    CODES = "FABRIC_CODE,SUB_BODY_CODE,STYLE_GROUP_CODE"
            ''ElseIf opt1Sheet.Value = "G" Then
            ''    CODES = "STYLE_GROUP_CODE,FABRIC_CODE,SUB_BODY_CODE"
            ''    ' CODES = "STYLE_GROUP_CODE"
            ''    ' DGJ
            ''ElseIf opt1Sheet.Value = "D" Then
            CODES = "SALES_DIVISION_CODE"

            '' End If

            For Each rowSB As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTCUSTS").Select(Mid(sqlWB & sql0, 6)), Split(CODES, ",")).Select("")
                Dim SHEET_NAME As String = ""
                Dim sqlSB As String = ""
                For Each COLUMN_NAME As String In Split(CODES, ",")
                    Dim CODE_VALUE As String = rowSB.Item(COLUMN_NAME) & ""
                    SHEET_NAME &= "-" & CODE_VALUE
                    If CODE_VALUE = "" Then
                        sqlSB &= " and " & COLUMN_NAME & " IS NULL"
                    Else
                        sqlSB &= " and " & COLUMN_NAME & " = '" & CODE_VALUE & "'"
                    End If
                Next

                If CODES = "SALES_DIVISION_CODE" Then
                    Dim SALES_DIVISION_NAME As String = ""
                    SALES_DIVISION_CODE = Mid(SHEET_NAME, 2)
                    ASCMAIN1.sql = "Select SALES_DIVISION_NAME from SOTSDIV1 where SALES_DIVISION_CODE = :PARM1"
                    Dim rowSOTDIV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", Mid(SHEET_NAME, 2))
                    If rowSOTDIV1 IsNot Nothing Then
                        SALES_DIVISION_NAME = rowSOTDIV1.Item("SALES_DIVISION_NAME")
                    Else
                        SALES_DIVISION_NAME = ""
                    End If
                    SHEET_NAME = "Div-" & Mid(SHEET_NAME, 2) & "-" & SALES_DIVISION_NAME
                Else
                    SHEET_NAME = Mid(SHEET_NAME, 2)
                End If


                If dst.Tables("SOTCUSTS").Select(Mid(sqlWB & sqlSB & sql0, 6)).Length > 0 Then
                    Dim worksheet As SpreadsheetGear.IWorksheet
                    If wsi = 0 Then
                        worksheet = workbook.Worksheets(0)
                    Else
                        worksheet = workbook.Worksheets.Add
                    End If
                    wsi += 1
                    If SHEET_NAME <> "" Then
                        worksheet.Name = SHEET_NAME
                    Else
                        worksheet.Name = "Unknown"
                    End If

                    Dim StyleList As New List(Of String)

                    '        For Each rowICTSTATD As DataRow In dst.Tables("ICTSTATD").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE, PO_DATE_SHIP_BY")

                    For Each rowSOTCUSTS As DataRow In dst.Tables("SOTCUSTS").Select("SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'", "SUB_BODY_CODE,FABRIC_CODE,STYLE_CODE,COLOR_CODE")
                        Dim STYLE_CODE As String = rowSOTCUSTS.Item("STYLE_CODE").ToString & String.Empty
                        If Not StyleList.Contains(STYLE_CODE) Then
                            StyleList.Add(STYLE_CODE)
                        End If
                    Next

                    Create_Excel_WorkSheet(worksheet, StyleList, sqlWB & sqlSB & sql0)
                    XLS_CREATED = True
                End If
            Next
        Else
            ''If dst.Tables("SOTCUSTS").Select(Mid(sqlWB & sql0, 6)).Length > 0 Then
            ''    Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
            ''    worksheet.Name = "Style Info"
            ''    Create_Excel_WorkSheet(worksheet, StyleList, sqlWB & sql0)
            ''    XLS_CREATED = True
            ''End If
        End If

        If XLS_CREATED Then
            Dim XLS_FILENAME As String = ""
            Dim success As Boolean = False

            ASCMAIN1.Progress("Now Saving Workbook")

            Do Until success
                Try
                    XLS_NO += 1
                    ' XLS_FILENAME = Absx1.txtFor("QUOTE_NO").Text
                    XLS_FILENAME = "ShipReport"

                    If SALES_DIVISION_CODE <> "" Then
                        XLS_FILENAME &= "-" & SALES_DIVISION_CODE
                    End If
                    XLS_FILENAME &= "-" & Format(XLS_NO, "000") & exlExt
                    workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                    'workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                    RetVal = XLS_FILENAME
                    success = True
                Catch ex As Exception

                End Try
            Loop

            Show_Document(ASCMAIN1.Folders("Temp") & XLS_FILENAME)
        End If

        ASCMAIN1.Progress("")
        Return RetVal
    End Function

#End Region

    Private Function FTP_BLUEHOST(ByRef FileNameLocalFull As String, ByRef FileNameRemote As String) As StringBuilder
        Dim RetVal As New StringBuilder With {.Length = 0}
        Dim FTPUser As String = "abs@vandalequotes.com"
        Dim FTPPassword As String = "0ff1c3ABS#"
        Dim FTPHost As String = "ftp.tzn.lnr.mybluehost.me"
        Dim FTPRemoteFull As String = $"/public_html/FTP/{FileNameRemote}"

        If Not System.IO.File.Exists(FileNameLocalFull) Then
            RetVal.AppendLine($"FTP File Provided Does Not Exist: {FileNameLocalFull}")
        End If

        If RetVal.Length = 0 Then
            Try
                If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                    Stop
                End If

                Ftp1.User = FTPUser
                Ftp1.Password = FTPPassword
                Ftp1.RemoteHost = FTPHost
                '               Ftp1.Logoff()
                Ftp1.Logon()

                Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                Ftp1.LocalFile = FileNameLocalFull
                Ftp1.RemoteFile = FTPRemoteFull
                'Ftp1.Timeout = 0 'Don't Timeout
                Ftp1.Overwrite = True

                Ftp1.Upload()

                Ftp1.Logoff()
            Catch ex As Exception
                RetVal.AppendLine($"FTP Error: {ex.Message} : {ex.InnerException}")
                'Just bail out for now.  We eventually need some kind of tracking.
            End Try
        End If
        Return RetVal
    End Function


    Sub Get_Availability()

        '''With grdICTQUOT2.DisplayLayout.Bands(1)
        '''    .Columns("QTY_AVA0").Header.Caption = "At Once" ' Format(dte0.Value, "MM/dd")
        '''    .Columns("QTY_AVA1").Header.Caption = Format(dte1.Value, "MM/dd")
        '''    .Columns("QTY_AVA2").Header.Caption = Format(dte2.Value, "MM/dd")
        '''    .Columns("QTY_AVA3").Header.Caption = Format(dte3.Value, "MM/dd")
        '''    .Columns("QTY_AVA4").Header.Caption = "Beyond"

        '''    .Columns("DTE0").Header.Caption = "Dates"
        '''    .Columns("DTE1").Header.Caption = "Dates"
        '''    .Columns("DTE2").Header.Caption = "Dates"
        '''    .Columns("DTE3").Header.Caption = "Dates"
        '''    .Columns("DTE4").Header.Caption = "Dates"

        '''    ' ENABLING THIS CODE MAKES THE ROWHEIGHT OF BAND1 CRAZY

        '''    'grdICTQUOT2.DisplayLayout.Override.RowSizing = UltraWinGrid.RowSizing.Free
        '''    'grdICTQUOT2.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy

        '''    '.Columns("QTY_AVA0").Hidden = False
        '''    '.Columns("QTY_AVA1").Hidden = Not dte1.Visible
        '''    '.Columns("QTY_AVA2").Hidden = Not dte2.Visible
        '''    '.Columns("QTY_AVA3").Hidden = Not dte3.Visible
        '''    '.Columns("QTY_AVA4").Hidden = Not chkBeyond.Checked

        '''    'If Not dte1.Visible Then
        '''    '    .Columns("QTY_AVA1").Width = 1
        '''    'Else
        '''    '    .Columns("QTY_AVA1").Width = 80
        '''    'End If

        '''    'If Not dte2.Visible Then
        '''    '    .Columns("QTY_AVA2").Width = 1
        '''    'Else
        '''    '    .Columns("QTY_AVA2").Width = 80
        '''    'End If

        '''    'If Not dte3.Visible Then
        '''    '    .Columns("QTY_AVA3").Width = 1
        '''    'Else
        '''    '    .Columns("QTY_AVA3").Width = 80
        '''    'End If

        '''    'If Not chkBeyond.Checked Then
        '''    '    .Columns("QTY_AVA4").Width = 1
        '''    'Else
        '''    '    .Columns("QTY_AVA4").Width = 80
        '''    'End If
        '''    ''grdICTQUOT2.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal
        '''    '.Override.MinRowHeight = 1
        '''    '.Override.ResetMinRowHeight()
        '''    '.Override.DefaultRowHeight = 1
        '''    '.Override.ResetDefaultRowHeight()

        '''    '  .Override.DefaultRowHeight = 4


        '''End With


        '''dst.Tables("ICTSTYC1").Columns("QTY_AVA").Expression = "0"
        '''For Each rowICTQUOT2 As DataRow In dst.Tables("ICTQUOT2").Select("")
        '''    Load_Availability(rowICTQUOT2)
        '''Next

        '''Dim MinGrpOpt As Int64 = 0
        '''If chkALLOSTDT.Checked Then
        '''    MinGrpOpt = cboStartPeriod.SelectedIndex
        '''End If

        '''Dim ColVisible(4) As Boolean
        '''If MinGrpOpt < 1 Then
        '''    ColVisible(0) = True
        '''End If
        '''If MinGrpOpt < 2 Then
        '''    ColVisible(1) = (tkb1.Value <= 2)
        '''End If
        '''If MinGrpOpt < 3 Then
        '''    ColVisible(2) = (tkb1.Value <= 1)
        '''End If
        '''If MinGrpOpt < 4 Then
        '''    ColVisible(3) = (tkb1.Value <= 0)
        '''End If
        '''If MinGrpOpt < 5 Then
        '''    ColVisible(4) = chkBeyond.Checked
        '''End If

        '''Dim EX As String = ""
        '''For I As Integer = 0 To 4
        '''    If ColVisible(I) Then
        '''        EX &= "+ISNULL(QTY_AVA" & CStr(I) & ",0)"
        '''    End If
        '''Next
        '''dst.Tables("ICTSTYC1").Columns("QTY_AVA").Expression = Mid(EX, 2)

        '''refresh_required = False
        '''cmdGetAvailability.Appearance.ForeColor = Color.Empty

    End Sub

    Private Function Create_Excel_BuyerChart() As String
        Dim RetVal As String = ""
        Me.Cursor = Cursors.WaitCursor

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        worksheet.Name = "Buyer Spreadsheet"
        ASCMAIN1.Progress("Now Creating Buyer Workbook", "")
        'Make Headers
        worksheet.Cells("A1").EntireColumn.ColumnWidth = 0
        worksheet.Cells("B1").EntireColumn.ColumnWidth = 0
        worksheet.Cells("C1").EntireColumn.ColumnWidth = 0
        worksheet.Cells("D1").EntireColumn.ColumnWidth = 0
        If chkShowFactoryBC.Checked Then
            worksheet.Cells("E1").EntireColumn.ColumnWidth = 13.17
        Else
            worksheet.Cells("E1").EntireColumn.ColumnWidth = 0
        End If
        worksheet.Cells("F1").EntireColumn.ColumnWidth = 20.33
        worksheet.Cells("G1").EntireColumn.ColumnWidth = 17.83
        worksheet.Cells("H1").EntireColumn.ColumnWidth = 27.33
        worksheet.Cells("I1").EntireColumn.ColumnWidth = 17.33
        worksheet.Cells("J1").EntireColumn.ColumnWidth = 19.83
        worksheet.Cells("K1").EntireColumn.ColumnWidth = 29.83
        If chkShowCountry.Checked Then
            worksheet.Cells("L1").EntireColumn.ColumnWidth = 14.83
        Else
            worksheet.Cells("L1").EntireColumn.ColumnWidth = 0
        End If
        worksheet.Cells("M1").EntireColumn.ColumnWidth = 15.83
        worksheet.Cells("N1").EntireColumn.ColumnWidth = 0
        worksheet.Cells("O1").EntireColumn.ColumnWidth = 12
        worksheet.Cells("P1").EntireColumn.ColumnWidth = 13
        If chkShowMSRP.Checked Then
            worksheet.Cells("Q1").EntireColumn.ColumnWidth = 13
        Else
            worksheet.Cells("Q1").EntireColumn.ColumnWidth = 0
        End If
        worksheet.Cells("R4").EntireColumn.ColumnWidth = 12.83
        worksheet.Cells("S1").EntireColumn.ColumnWidth = 15.83
        worksheet.Cells("T1").EntireColumn.ColumnWidth = 29.83
        worksheet.Cells("U1").EntireColumn.ColumnWidth = 15.83

        worksheet.Cells("V1: AB1").EntireColumn.ColumnWidth = 12

        worksheet.Cells("E1: J1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells("K1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
        worksheet.Cells("L1: M1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells("O1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
        worksheet.Cells("P1: Q1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells("J1").EntireColumn.WrapText = True
        worksheet.Cells("K1").EntireColumn.WrapText = True
        worksheet.Cells("S1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells("T1").EntireColumn.WrapText = True
        worksheet.Cells("U1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center

        worksheet.Cells("V1: AB1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right

        worksheet.Cells("A1").RowHeight = 12
        worksheet.Cells("A2").RowHeight = 48.75
        worksheet.Cells("A3").RowHeight = 12
        worksheet.Cells("A4").RowHeight = 56.5
        worksheet.Cells("A4").Value = "Department Number"
        worksheet.Cells("B4").Value = "Season"
        worksheet.Cells("C4").Value = "Class"
        worksheet.Cells("D4").Value = "Category Number"
        If chkShowFactoryBC.Checked Then
            worksheet.Cells("E4").Value = "Factory"
        End If
        worksheet.Cells("F4").Value = "Brand"
        worksheet.Cells("G4").Value = "Size Ratio"
        worksheet.Cells("H4").Value = "Photo"
        worksheet.Cells("I4").Value = "Style Code"
        worksheet.Cells("J4").Value = "Product Description"
        worksheet.Cells("K4").Value = "Color"
        If chkShowCountry.Checked Then
            worksheet.Cells("L4").Value = "Country"
        Else
            worksheet.Cells("L4").Value = ""
        End If
        worksheet.Cells("M4").Value = "ShpDt Range"
        worksheet.Cells("N4").Value = "TKM"
        worksheet.Cells("O4").Value = "Shipped"
        worksheet.Cells("P4").Value = "Vandale Cost"
        If chkShowMSRP.Checked Then
            worksheet.Cells("Q4").Value = "MSRP"
        Else
            worksheet.Cells("Q4").Value = ""
        End If
        worksheet.Cells("R4").Value = ""
        worksheet.Cells("S4").Value = "FOB date"
        worksheet.Cells("T4").Value = "Factory Name"
        worksheet.Cells("U4").Value = "Last Rcvd"

        If chkStyleStats.Checked Then
            worksheet.Cells("V4").Value = "On Hand"
            worksheet.Cells("W4").Value = "In Pick"
            worksheet.Cells("X4").Value = "OTS"
            worksheet.Cells("Y4").Value = "In Transit"
            worksheet.Cells("Z4").Value = "WIP"
            worksheet.Cells("AA4").Value = "Open"
            worksheet.Cells("AB4").Value = "Net Pos"
        End If




        worksheet.Cells("F2").Value = "Buyer Chart"
        With worksheet.Cells("E2:R2")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Font.Bold = True
            .Font.Size = 18
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Colors.LightCyan
        End With
        With worksheet.Cells("A4:D4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Colors.LightGray
        End With
        With worksheet.Cells("E4:N4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
        End With
        With worksheet.Cells("O4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Colors.Yellow
        End With
        With worksheet.Cells("P4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
        End With
        With worksheet.Cells("Q4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
        End With
        With worksheet.Cells("R4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Colors.Yellow
        End With
        With worksheet.Cells("S4:T4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
        End With

        With worksheet.Cells("U4")
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
        End With

        If chkStyleStats.Checked Then
            With worksheet.Cells("V4: AB4")
                .VerticalAlignment = SpreadsheetGear.VAlign.Center
                .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                .Font.Bold = True
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                .Interior.Color = SpreadsheetGear.Colors.Aquamarine
                '.EntireColumn.FormatConditions,A
            End With

        End If


        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

        Dim SORTSOTCUSTS As String = "SUB_BODY_CODE,FABRIC_CODE,STYLE_CODE,COLOR_CODE"
        '''If chkSortStyle.Checked Then
        '''    SORTSOTCUSTQ = "STYLE_CODE,COLOR_CODE"
        '''End If


        '  Dim QTYAVAILFILTER As String = "QTY_AVA <> 0"
        Dim CURR_SALES_DIVISION_CODE As String = ""
        Dim QTYAVAILFILTER As String = ""

        Dim curRow As Int64 = 5
        For Each rowSB As DataRow In dst.Tables.Item("SOTCUSTS").Select("", SORTSOTCUSTS)
            STYLE_CODE = rowSB.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowSB.Item("COLOR_CODE").ToString & String.Empty
            Dim SALES_DIVISION_CODE As String = rowSB.Item("SALES_DIVISION_CODE").ToString & String.Empty
            If CURR_SALES_DIVISION_CODE <> CURR_SALES_DIVISION_CODE Then
                ' NEW SHEET
            End If
            Dim sql As New System.Text.StringBuilder With {.Length = 0}
            sql.AppendLine("Select")
            sql.AppendLine("ST1.FACTORY_CODE,")
            sql.AppendLine("CN1.COUNTRY_NAME,")
            sql.AppendLine("SD1.SALES_DIVISION_NAME,")
            sql.AppendLine("ST1.STYLE_RETAIL")
            sql.AppendLine("FROM ICTSTYL1 ST1, SOTSDIV1 SD1, TATCNTRY CN1")
            sql.AppendLine("WHERE ST1.SALES_DIVISION_CODE = SD1.SALES_DIVISION_CODE")
            sql.AppendLine("And ST1.COUNTRY_CODE = CN1.COUNTRY_CODE (+)")
            sql.AppendLine(String.Format("And STYLE_CODE = '{0}'", STYLE_CODE))
            Dim tblSTYLE As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
            Dim FACTORY_CODE As String = ""
            Dim COUNTRY_NAME As String = ""
            Dim SALES_DIVISION_NAME As String = ""
            Dim STYLE_RETAIL As String = ""
            Dim FACTORY_DESC As String = ""
            Dim VAN_COST As String = ""
            ' ---------

            '''If chkShowCost.Checked Then
            '''    Dim COSTTYPE As String = "FC"
            '''    Dim STYLE_COST As Decimal = 0
            '''    Dim COST_PERIOD As String = ""
            '''    ASCMAIN1.sql = "Select OPS_YYYYPP, STYLE_COST from (" & vbCrLf _
            '''                    & "Select OPS_YYYYPP,STYLE_COST from ICTCOSTA " & vbCrLf _
            '''                    & "where (STYLE_CODE, COLOR_CODE) in (" & vbCrLf _
            '''                    & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            '''                    & " from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'" _
            '''                    & " and WHSE_QTY_ON_HAND > 0)" & vbCrLf _
            '''                    & " order by OPS_YYYYPP DESC) where ROWNUM < 2"

            '''    For Each rowICTCOSTA As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            '''        STYLE_COST = Val(rowICTCOSTA.Item("STYLE_COST") & "")
            '''        COST_PERIOD = rowICTCOSTA.Item("OPS_YYYYPP") & ""
            '''    Next
            '''    ' CHECK FOR MULTIPLE Costs that make it up LC(*), ONE COST MAKES ITS UP LC(TI) TARIFF INC, LC(TNA) TARIFF Not Incl

            '''    If STYLE_COST <> 0 And chkCostCode.Checked Then
            '''        Dim ICTCOSTL_COSTS As Integer = 0
            '''        ASCMAIN1.sql = "Select * From ICTCOSTL Where LOT_QTY_ONHD <> 0 AND OPS_YYYYPP_FIFO = '" & COST_PERIOD & "'AND STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            '''        For Each rowICTCOSTL As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            '''            If ICTCOSTL_COSTS > 0 Then
            '''                COSTTYPE = "FLC(*)"
            '''                Exit For
            '''            End If
            '''            If rowICTCOSTL.Item("TARIFF_FLAG") & "" <> "" Then
            '''                COSTTYPE = "FLC"
            '''            Else
            '''                ' COSTTYPE = "LC(TNA)"
            '''                COSTTYPE = "FLC"
            '''            End If
            '''            ICTCOSTL_COSTS += 1
            '''        Next
            '''    End If

            '''    ' CHANGE PO_COST FIRST TO PO_COST_VCOST FOB A PER GABE 03/05/2025 DGJ
            '''    If STYLE_COST = 0 Then
            '''        ASCMAIN1.sql = "Select NVL(PO_COST_LANDED,PO_COST_VCOST) STYLE_COST, PO_COST_VCOST,PO_COST_LANDED,PO_SHIPMENT_NO" & vbCrLf _
            '''                        & " from (" & vbCrLf _
            '''                        & " Select POTSHIP3.PO_SHIPMENT_NO, POTORDR2.PO_ORDER_NO, " & vbCrLf _
            '''                        & " POTORDR2.PO_COST_VCOST, POTSHIP3.PO_COST_LANDED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
            '''                        & " from POTORDR2,POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
            '''                        & " where POTORDR2.STYLE_CODE = '" & STYLE_CODE & "' and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            '''                        & "   and POTSHIP3.PO_ORDER_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
            '''                        & "   and POTSHIP3.PO_ORDER_LNO (+) = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            '''                        & "   and POTSHIP2.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            '''                        & "   and POTSHIP2.PO_SHIPMENT_LNO (+) = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            '''                        & "   and POTSHIP1.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            '''                        & " order by POTSHIP3.PO_SHIPMENT_NO DESC, POTORDR2.PO_ORDER_NO DESC" & vbCrLf _
            '''                        & ") where ROWNUM <2"
            '''        '  STYLE_COST = Val(ASCDATA1.GetDataValue)
            '''        For Each rowPOTSHIP3 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            '''            STYLE_COST = Val(rowPOTSHIP3.Item("STYLE_COST") & "")
            '''            If chkCostCode.Checked Then
            '''                If STYLE_COST = Val(rowPOTSHIP3.Item("PO_COST_VCOST") & "") Then
            '''                    COSTTYPE = "FOB"
            '''                Else
            '''                    Dim PO_SHIPMENT_NO As String = rowPOTSHIP3.Item("PO_SHIPMENT_NO") & ""
            '''                    If PO_SHIPMENT_NO <> "" Then
            '''                        ASCMAIN1.sql = "Select SUM(LANDING_COST_AMT) From POTSHIP5 Where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' AND COST_CATGY_CODE = 'TARIFF'"
            '''                        Dim TARIFF_AMT As Integer = Val(ASCDATA1.GetDataValue)
            '''                        If TARIFF_AMT <> 0 Then
            '''                            COSTTYPE = "PC(TI)"
            '''                        Else
            '''                            COSTTYPE = "PC(TNA)"
            '''                        End If

            '''                    End If
            '''                End If
            '''            End If
            '''        Next

            '''    End If

            '''    If STYLE_COST = 0 Then
            '''        '   STYLE_COST = Val(row.Item("STYLE_COST") & "")
            '''        COSTTYPE = "SC"
            '''        COSTTYPE = ""
            '''    End If
            '''    STYLE_COST = Format$(STYLE_COST, "$#,##0.00")

            '''    If chkCostCode.Checked = True Then
            '''        COSTTYPE = " - " & COSTTYPE
            '''    Else
            '''        COSTTYPE = ""
            '''    End If
            '''    VAN_COST = STYLE_COST & COSTTYPE
            '''End If


            If tblSTYLE.Rows.Count = 1 Then
                FACTORY_CODE = tblSTYLE.Rows(0).Item("FACTORY_CODE").ToString & String.Empty
                Dim rowICTFACT1 As DataRow = clsASCBASE1.LookUp("ICTFACT1", FACTORY_CODE)
                If rowICTFACT1 Is Nothing Then
                    FACTORY_DESC = ""
                Else
                    FACTORY_DESC = FACTORY_CODE & "-" & rowICTFACT1.Item("FACTORY_DESC") & ""
                End If

                COUNTRY_NAME = tblSTYLE.Rows(0).Item("COUNTRY_NAME").ToString & String.Empty
                SALES_DIVISION_NAME = tblSTYLE.Rows(0).Item("SALES_DIVISION_NAME").ToString & String.Empty
                If chkShowMSRP.Checked Then
                    If IsNumeric(tblSTYLE.Rows(0).Item("STYLE_RETAIL").ToString & String.Empty) Then
                        If Val(tblSTYLE.Rows(0).Item("STYLE_RETAIL").ToString & String.Empty) > 0 Then
                            STYLE_RETAIL = Format(Val(tblSTYLE.Rows(0).Item("STYLE_RETAIL").ToString & String.Empty), "###,##0.00")
                        End If
                    End If
                End If
            End If
            Dim STYLE_COLOR_DESC As String = rowSB.Item("COLOR_DESC").ToString & String.Empty
            '  Dim fltrSOTCUSTQ As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            '     Dim rowSOTCUSTQ As DataRow = dst.Tables.Item("SOTCUSTQ").Select(fltrSOTCUSTQ).FirstOrDefault
            Dim STYLE_DESC As String = rowSB.Item("STYLE_DESC").ToString & String.Empty
            Dim SIZE_SCALE As String = GET_ONLY_SIZE_SCALE(STYLE_CODE)
            Dim IMAGE_NAME As String = rowSB.Item("IMAGE_NAME") & ""
            Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME
            Dim HasImage As Boolean = False
            Dim imageStyle As System.Drawing.Image = Nothing
            If My.Computer.FileSystem.FileExists(imageFileStyle) Then
                imageStyle = System.Drawing.Image.FromFile(imageFileStyle)
                HasImage = True
            End If
            worksheet.Cells("A" & curRow.ToString & ":" & "R" & curRow.ToString).Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            worksheet.Cells("A" & curRow.ToString).RowHeight = 100.5
            worksheet.Cells("E" & curRow.ToString).Value = FACTORY_CODE
            worksheet.Cells("F" & curRow.ToString).Value = SALES_DIVISION_NAME
            worksheet.Cells("G" & curRow.ToString).Value = SIZE_SCALE
            If HasImage Then
                Dim leftStyle As Integer = windowInfoStyle.ColumnToPoints(7)
                Dim topStyle As Integer = windowInfoStyle.RowToPoints(curRow - 1) + 0.1
                Dim WidthStyle As Integer = 100
                Dim HeightStyle As Integer = 99
                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle + 20, topStyle + 1, WidthStyle, HeightStyle)
            End If
            worksheet.Cells("I" & curRow.ToString).Value = STYLE_CODE
            worksheet.Cells("J" & curRow.ToString).Value = STYLE_DESC
            worksheet.Cells("K" & curRow.ToString).Value = COLOR_CODE & " - " & STYLE_COLOR_DESC
            worksheet.Cells("L" & curRow.ToString).Value = COUNTRY_NAME
            Dim TOT_AVAIL As Int64 = 0
            Dim DATES As New System.Text.StringBuilder With {.Length = 0}
            Dim FOBDATES As New System.Text.StringBuilder With {.Length = 0}
            Dim DATES_STRING As String = GetCustShipDates(STYLE_CODE, COLOR_CODE)
            If Val(rowSB.Item("QTY_SHP").ToString & String.Empty) <> 0 Then
                TOT_AVAIL = Val(rowSB.Item("QTY_SHP").ToString & String.Empty)
            End If
            Dim FOBDATES_STRING As String = ""

            With worksheet.Cells("M" & curRow.ToString)
                .Value = DATES_STRING
                .Font.Color = SpreadsheetGear.Colors.Red
            End With
            With worksheet.Cells("O" & curRow.ToString)
                .Value = TOT_AVAIL
                .NumberFormat = "###,##0"
            End With
            With worksheet.Cells("P" & curRow.ToString)
                .Value = VAN_COST 'Get Vandale Cost Here
                '  .NumberFormat = "$###,##0.00"
                .Font.Color = SpreadsheetGear.Colors.Red

            End With
            With worksheet.Cells("Q" & curRow.ToString)
                .Value = STYLE_RETAIL
                .NumberFormat = "$###,##0.00"
            End With
            With worksheet.Cells("R" & curRow.ToString)
                ' .Value = 3.3 'Get TKMAX OFFER Here
                .NumberFormat = "$###,##0.00"
                .Interior.Color = SpreadsheetGear.Colors.Yellow
                .Font.Color = SpreadsheetGear.Colors.Red
                .VerticalAlignment = SpreadsheetGear.VAlign.Center
            End With
            With worksheet.Cells("S" & curRow.ToString)
                .Value = FOBDATES_STRING
                .Font.Color = SpreadsheetGear.Colors.Red
            End With
            worksheet.Cells("T" & curRow.ToString).Value = FACTORY_DESC

            ''If chkShowLastRcd.Checked Then
            ''    If IsDate(rowSB.Item("LAST_RCD_DATE").ToString & String.Empty) Then
            ''        Dim LAST_SHIPPED As Date = CDate(rowSB.Item("LAST_RCD_DATE").ToString & String.Empty)
            ''        worksheet.Cells("U" & curRow.ToString).Value = Format(LAST_SHIPPED, "MM/dd/yy")
            ''    Else
            ''        Dim LAST_SHIPPED As String = rowSB.Item("LAST_RCD_DATE").ToString & String.Empty
            ''        worksheet.Cells("U" & curRow.ToString).Value = LAST_SHIPPED
            ''    End If
            ''Else
            ''    worksheet.Cells("U" & curRow.ToString).Value = ""
            ''End If
            worksheet.Cells("U" & curRow.ToString).Value = ""

            If chkStyleStats.Checked Then

                For Each rowICTSTAT2 As DataRow In dst.Tables("ICTSTAT2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                    worksheet.Cells("V" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & String.Empty)
                    worksheet.Cells("V" & curRow.ToString).NumberFormat = "#,###,##0"
                    worksheet.Cells("W" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & String.Empty)
                    worksheet.Cells("W" & curRow.ToString).NumberFormat = "#,###,##0"

                    Dim OTS As Integer = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "") - Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & "")
                    worksheet.Cells("X" & curRow.ToString).Value = OTS
                    worksheet.Cells("X" & curRow.ToString).NumberFormat = "#,###,##0"

                    worksheet.Cells("Y" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty)
                    worksheet.Cells("Y" & curRow.ToString).NumberFormat = "#,###,##0"

                    worksheet.Cells("Z" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty)
                    worksheet.Cells("Z" & curRow.ToString).NumberFormat = "#,###,##0"

                    worksheet.Cells("AA" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
                    worksheet.Cells("AA" & curRow.ToString).NumberFormat = "#,###,##0"

                    worksheet.Cells("AB" & curRow.ToString).Value = OTS + Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty) + Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty) - Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
                    worksheet.Cells("AB" & curRow.ToString).NumberFormat = "#,###,##0"
                Next

            End If
            curRow += 1
        Next

        'Show Workbook
        Dim XLS_FILENAME As String = "5000"
        Dim success As Boolean = False
        ' Dim RPT_PREFIX As String = Absx1.txtFor("QUOTE_NO").Text
        Dim RPT_PREFIX As String = "BuyerChart"
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = RPT_PREFIX & "_" & Format(XLS_NO, "000") & exlExt
                workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                RetVal = XLS_FILENAME
                success = True
            Catch ex As Exception
                If XLS_NO > 5000 Then
                    success = True
                End If
            End Try
        Loop
        If XLS_FILENAME = "5000" Then
            MsgBox("Reports In Temp Folder Exceeded", vbCritical, "Log Out Of ABS And Get Back In")
        Else
            Show_Document(ASCMAIN1.Folders("Temp") & XLS_FILENAME)
        End If

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default

        Return RetVal
    End Function
    Private Function Create_Excel_BuyerChart_DIV() As String
        Dim RetVal As String = ""
        Me.Cursor = Cursors.WaitCursor

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)

        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

        Dim SORTSOTCUSTS As String = "SUB_BODY_CODE,FABRIC_CODE,STYLE_CODE,COLOR_CODE"
        ''If chkSortStyle.Checked Then
        ''    SORTSOTCUSTQ = "STYLE_CODE,COLOR_CODE"
        ''End If


        Dim QTYAVAILFILTER As String = "QTY_AVA <> 0"
        Dim CURR_SALES_DIVISION_CODE As String = ""
        QTYAVAILFILTER = ""
        Dim curRow As Int64 = 5
        For Each rowSB As DataRow In dst.Tables.Item("SOTCUSTS").Select(QTYAVAILFILTER, "SALES_DIVISION_CODE, SUB_BODY_CODE, FABRIC_CODE, STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowSB.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowSB.Item("COLOR_CODE").ToString & String.Empty
            Dim SALES_DIVISION_CODE As String = rowSB.Item("SALES_DIVISION_CODE").ToString & String.Empty
            Dim SALES_DIVISION_NAME As String = ""
            If CURR_SALES_DIVISION_CODE <> SALES_DIVISION_CODE Then

                ' NEW SHEET
                Dim SHEET_NAME As String = SALES_DIVISION_CODE
                SALES_DIVISION_CODE = SHEET_NAME
                ASCMAIN1.sql = "Select SALES_DIVISION_NAME from SOTSDIV1 where SALES_DIVISION_CODE = :PARM1"
                Dim rowSOTDIV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SALES_DIVISION_CODE)
                If rowSOTDIV1 IsNot Nothing Then
                    SALES_DIVISION_NAME = rowSOTDIV1.Item("SALES_DIVISION_NAME")
                Else
                    SALES_DIVISION_NAME = ""
                End If
                SHEET_NAME = "Div-" & SHEET_NAME & "-" & SALES_DIVISION_NAME
                SHEET_NAME = Replace(SHEET_NAME, "/", " ")
                SHEET_NAME = Replace(SHEET_NAME, ".", "")
                SHEET_NAME = Replace(SHEET_NAME, ",", "")
                SHEET_NAME = Replace(SHEET_NAME, "&", "")

                If SHEET_NAME.Length > 31 Then
                    SHEET_NAME = SHEET_NAME.ToString.Substring(0, 30)
                    ' SHEET_NAME = SUBSTR(SHEET_NAME, 0, 31)
                End If
                If CURR_SALES_DIVISION_CODE = "" Then
                Else
                    worksheet = workbook.Worksheets.Add
                End If
                If SHEET_NAME <> "" Then
                    worksheet.Name = SHEET_NAME
                Else
                    worksheet.Name = "Unknown"
                End If
                CURR_SALES_DIVISION_CODE = SALES_DIVISION_CODE


                ASCMAIN1.Progress("Now Creating Buyer Workbook", "")
                'Make Headers
                worksheet.Cells("A1").EntireColumn.ColumnWidth = 0
                worksheet.Cells("B1").EntireColumn.ColumnWidth = 0
                worksheet.Cells("C1").EntireColumn.ColumnWidth = 0
                worksheet.Cells("D1").EntireColumn.ColumnWidth = 0
                If chkShowFactoryBC.Checked Then
                    worksheet.Cells("E1").EntireColumn.ColumnWidth = 13.17
                Else
                    worksheet.Cells("E1").EntireColumn.ColumnWidth = 0
                End If
                worksheet.Cells("F1").EntireColumn.ColumnWidth = 20.33
                worksheet.Cells("G1").EntireColumn.ColumnWidth = 17.83
                worksheet.Cells("H1").EntireColumn.ColumnWidth = 27.33
                worksheet.Cells("I1").EntireColumn.ColumnWidth = 17.33
                worksheet.Cells("J1").EntireColumn.ColumnWidth = 19.83
                worksheet.Cells("K1").EntireColumn.ColumnWidth = 29.83
                If chkShowCountry.Checked Then
                    worksheet.Cells("L1").EntireColumn.ColumnWidth = 14.83
                Else
                    worksheet.Cells("L1").EntireColumn.ColumnWidth = 0
                End If
                worksheet.Cells("M1").EntireColumn.ColumnWidth = 15.83
                worksheet.Cells("N1").EntireColumn.ColumnWidth = 0
                worksheet.Cells("O1").EntireColumn.ColumnWidth = 12
                worksheet.Cells("P1").EntireColumn.ColumnWidth = 13
                If chkShowMSRP.Checked Then
                    worksheet.Cells("Q1").EntireColumn.ColumnWidth = 13
                Else
                    worksheet.Cells("Q1").EntireColumn.ColumnWidth = 0
                End If
                worksheet.Cells("R4").EntireColumn.ColumnWidth = 12.83
                worksheet.Cells("S1").EntireColumn.ColumnWidth = 15.83
                worksheet.Cells("T1").EntireColumn.ColumnWidth = 29.83
                worksheet.Cells("U1").EntireColumn.ColumnWidth = 15.83
                worksheet.Cells("V1: AB1").EntireColumn.ColumnWidth = 12

                worksheet.Cells("E1: J1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                worksheet.Cells("K1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
                worksheet.Cells("L1: M1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                worksheet.Cells("O1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                worksheet.Cells("P1: Q1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                worksheet.Cells("J1").EntireColumn.WrapText = True
                worksheet.Cells("K1").EntireColumn.WrapText = True
                worksheet.Cells("S1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                worksheet.Cells("T1").EntireColumn.WrapText = True
                worksheet.Cells("U1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center

                worksheet.Cells("V1: AB1").EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right

                worksheet.Cells("A1").RowHeight = 12
                worksheet.Cells("A2").RowHeight = 48.75
                worksheet.Cells("A3").RowHeight = 12
                worksheet.Cells("A4").RowHeight = 56.5
                worksheet.Cells("A4").Value = "Department Number"
                worksheet.Cells("B4").Value = "Season"
                worksheet.Cells("C4").Value = "Class"
                worksheet.Cells("D4").Value = "Category Number"
                If chkShowFactoryBC.Checked Then
                    worksheet.Cells("E4").Value = "Factory"
                End If
                worksheet.Cells("F4").Value = "Brand"
                worksheet.Cells("G4").Value = "Size Ratio"
                worksheet.Cells("H4").Value = "Photo"
                worksheet.Cells("I4").Value = "Style Code"
                worksheet.Cells("J4").Value = "Product Description"
                worksheet.Cells("K4").Value = "Color"
                If chkShowCountry.Checked Then
                    worksheet.Cells("L4").Value = "Country"
                Else
                    worksheet.Cells("L4").Value = ""
                End If
                worksheet.Cells("M4").Value = "ShpDt Range"
                worksheet.Cells("N4").Value = "TKM"
                worksheet.Cells("O4").Value = "Shipped"
                worksheet.Cells("P4").Value = "Vandale Cost"
                If chkShowMSRP.Checked Then
                    worksheet.Cells("Q4").Value = "MSRP"
                Else
                    worksheet.Cells("Q4").Value = ""
                End If
                worksheet.Cells("R4").Value = ""
                worksheet.Cells("S4").Value = "FOB date"
                worksheet.Cells("T4").Value = "Factory Name"
                worksheet.Cells("U4").Value = "Last Rcvd"


                If chkStyleStats.Checked Then
                    worksheet.Cells("V4").Value = "On Hand"
                    worksheet.Cells("W4").Value = "In Pick"
                    worksheet.Cells("X4").Value = "OTS"
                    worksheet.Cells("Y4").Value = "In Transit"
                    worksheet.Cells("Z4").Value = "WIP"
                    worksheet.Cells("AA4").Value = "Open"
                    worksheet.Cells("AB4").Value = "Net Pos"
                End If

                worksheet.Cells("F2").Value = "Buyer Chart"
                With worksheet.Cells("E2:R2")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Left
                    .Font.Bold = True
                    .Font.Size = 18
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Colors.LightCyan
                End With
                With worksheet.Cells("A4:D4")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Colors.LightGray
                End With
                With worksheet.Cells("E4:N4")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
                End With
                With worksheet.Cells("O4")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Colors.Yellow
                End With
                With worksheet.Cells("P4")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
                End With
                With worksheet.Cells("Q4")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
                End With
                With worksheet.Cells("R4")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Colors.Yellow
                End With
                With worksheet.Cells("S4:T4")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
                End With

                With worksheet.Cells("U4")
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .Font.Bold = True
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
                    .Interior.Color = SpreadsheetGear.Color.FromArgb(252, 213, 179)
                End With

                If chkStyleStats.Checked Then
                    With worksheet.Cells("V4: AB4")
                        .VerticalAlignment = SpreadsheetGear.VAlign.Center
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        .Font.Bold = True
                        .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                        .Interior.Color = SpreadsheetGear.Colors.Aquamarine
                        '.EntireColumn.FormatConditions,A
                    End With

                End If


                ''Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
                ''Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo
                curRow = 5
            End If

            Dim sql As New System.Text.StringBuilder With {.Length = 0}
            sql.AppendLine("Select")
            sql.AppendLine("ST1.FACTORY_CODE,")
            sql.AppendLine("CN1.COUNTRY_NAME,")
            sql.AppendLine("SD1.SALES_DIVISION_NAME,")
            sql.AppendLine("ST1.STYLE_RETAIL")
            sql.AppendLine("FROM ICTSTYL1 ST1, SOTSDIV1 SD1, TATCNTRY CN1")
            sql.AppendLine("WHERE ST1.SALES_DIVISION_CODE = SD1.SALES_DIVISION_CODE")
            sql.AppendLine("And ST1.COUNTRY_CODE = CN1.COUNTRY_CODE (+)")
            sql.AppendLine(String.Format("And STYLE_CODE = '{0}'", STYLE_CODE))
            Dim tblSTYLE As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
            Dim FACTORY_CODE As String = ""
            Dim COUNTRY_NAME As String = ""
            '  SALES_DIVISION_NAME As String = ""
            Dim STYLE_RETAIL As String = ""
            Dim FACTORY_DESC As String = ""
            Dim VAN_COST As String = ""
            ' ---------

            ''If chkShowCost.Checked Then
            ''    Dim COSTTYPE As String = "FC"
            ''    Dim TPERC As String = ""
            ''    Dim STYLE_COST As Decimal = 0
            ''    Dim COST_PERIOD As String = ""
            ''    ASCMAIN1.sql = "Select OPS_YYYYPP, STYLE_COST from (" & vbCrLf _
            ''                & "Select OPS_YYYYPP,STYLE_COST from ICTCOSTA " & vbCrLf _
            ''                & "where (STYLE_CODE, COLOR_CODE) in (" & vbCrLf _
            ''                & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            ''                & " from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'" _
            ''                & " and WHSE_QTY_ON_HAND > 0)" & vbCrLf _
            ''                & " order by OPS_YYYYPP DESC) where ROWNUM < 2"

            ''    For Each rowICTCOSTA As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            ''        STYLE_COST = Val(rowICTCOSTA.Item("STYLE_COST") & "")
            ''        COST_PERIOD = rowICTCOSTA.Item("OPS_YYYYPP") & ""
            ''    Next
            ''    ' CHECK FOR MULTIPLE Costs that make it up LC(*), ONE COST MAKES ITS UP LC(TI) TARIFF INC, LC(TNA) TARIFF Not Incl

            ''    If STYLE_COST <> 0 And chkCostCode.Checked Then
            ''        Dim ICTCOSTL_COSTS As Integer = 0
            ''        ASCMAIN1.sql = "Select * From ICTCOSTL Where LOT_QTY_ONHD <> 0 AND OPS_YYYYPP_FIFO = '" & COST_PERIOD & "'AND STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            ''        For Each rowICTCOSTL As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            ''            If ICTCOSTL_COSTS > 0 Then
            ''                COSTTYPE = "FLC(*)"
            ''                Exit For
            ''            End If
            ''            If rowICTCOSTL.Item("TARIFF_FLAG") & "" <> "" Then
            ''                COSTTYPE = "FLC"
            ''                If Len(rowICTCOSTL.Item("TARIFF_FLAG") & "") = 10 Then
            ''                    ' TPERC = "T% " & Mid(rowICTCOSTL.Item("TARIFF_FLAG"), 9, 2)
            ''                End If
            ''            Else
            ''                ' COSTTYPE = "LC(TNA)"
            ''                COSTTYPE = "FLC"
            ''            End If
            ''            ICTCOSTL_COSTS += 1
            ''        Next
            ''    End If

            ''    ' CHANGE PO_COST FIRST TO PO_COST_VCOST FOB A PER GABE 03/05/2025 DGJ
            ''    If STYLE_COST = 0 Then
            ''        ASCMAIN1.sql = "Select NVL(PO_COST_LANDED,PO_COST_VCOST) STYLE_COST, PO_COST_VCOST,PO_COST_LANDED,PO_SHIPMENT_NO" & vbCrLf _
            ''                    & " from (" & vbCrLf _
            ''                    & " Select POTSHIP3.PO_SHIPMENT_NO, POTORDR2.PO_ORDER_NO, " & vbCrLf _
            ''                    & " POTORDR2.PO_COST_VCOST, POTSHIP3.PO_COST_LANDED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
            ''                    & " from POTORDR2,POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
            ''                    & " where POTORDR2.STYLE_CODE = '" & STYLE_CODE & "' and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            ''                    & "   and POTSHIP3.PO_ORDER_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
            ''                    & "   and POTSHIP3.PO_ORDER_LNO (+) = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            ''                    & "   and POTSHIP2.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            ''                    & "   and POTSHIP2.PO_SHIPMENT_LNO (+) = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            ''                    & "   and POTSHIP1.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            ''                    & " order by POTSHIP3.PO_SHIPMENT_NO DESC, POTORDR2.PO_ORDER_NO DESC" & vbCrLf _
            ''                    & ") where ROWNUM <2"
            ''        '  STYLE_COST = Val(ASCDATA1.GetDataValue)
            ''        For Each rowPOTSHIP3 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            ''            STYLE_COST = Val(rowPOTSHIP3.Item("STYLE_COST") & "")
            ''            If chkCostCode.Checked Then
            ''                If STYLE_COST = Val(rowPOTSHIP3.Item("PO_COST_VCOST") & "") Then
            ''                    COSTTYPE = "FOB"
            ''                Else
            ''                    Dim PO_SHIPMENT_NO As String = rowPOTSHIP3.Item("PO_SHIPMENT_NO") & ""
            ''                    If PO_SHIPMENT_NO <> "" Then
            ''                        ASCMAIN1.sql = "Select SUM(LANDING_COST_AMT) From POTSHIP5 Where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' AND COST_CATGY_CODE = 'TARIFF'"
            ''                        Dim TARIFF_AMT As Integer = Val(ASCDATA1.GetDataValue)
            ''                        If TARIFF_AMT <> 0 Then
            ''                            COSTTYPE = "PC(TI)"
            ''                        Else
            ''                            COSTTYPE = "PC(TNA)"
            ''                        End If

            ''                    End If
            ''                End If
            ''            End If
            ''        Next

            ''    End If

            ''    If STYLE_COST = 0 Then
            ''        '   STYLE_COST = Val(row.Item("STYLE_COST") & "")
            ''        COSTTYPE = "SC"
            ''        COSTTYPE = ""
            ''    End If
            ''    STYLE_COST = Format$(STYLE_COST, "$#,##0.00")

            ''    If chkCostCode.Checked = True Then
            ''        COSTTYPE = " - " & COSTTYPE & " " & TPERC
            ''    Else
            ''        COSTTYPE = ""
            ''    End If
            ''    VAN_COST = STYLE_COST & COSTTYPE
            ''End If


            If tblSTYLE.Rows.Count = 1 Then
                FACTORY_CODE = tblSTYLE.Rows(0).Item("FACTORY_CODE").ToString & String.Empty
                Dim rowICTFACT1 As DataRow = clsASCBASE1.LookUp("ICTFACT1", FACTORY_CODE)
                If rowICTFACT1 Is Nothing Then
                    FACTORY_DESC = ""
                Else
                    '     FACTORY_DESC = rowICTFACT1.Item("FACTORY_DESC") & ""
                    FACTORY_DESC = FACTORY_CODE & " " & rowICTFACT1.Item("FACTORY_DESC") & ""

                End If

                COUNTRY_NAME = tblSTYLE.Rows(0).Item("COUNTRY_NAME").ToString & String.Empty
                SALES_DIVISION_NAME = tblSTYLE.Rows(0).Item("SALES_DIVISION_NAME").ToString & String.Empty
                If chkShowMSRP.Checked Then
                    If IsNumeric(tblSTYLE.Rows(0).Item("STYLE_RETAIL").ToString & String.Empty) Then
                        If Val(tblSTYLE.Rows(0).Item("STYLE_RETAIL").ToString & String.Empty) > 0 Then
                            STYLE_RETAIL = Format(Val(tblSTYLE.Rows(0).Item("STYLE_RETAIL").ToString & String.Empty), "###,##0.00")
                        End If
                    End If
                End If
            End If
            Dim STYLE_COLOR_DESC As String = rowSB.Item("COLOR_DESC").ToString & String.Empty
            '   Dim fltrICTQUOT2 As String = String.Format("STYLE_CODE_PLM = '{0}'", STYLE_CODE)
            '   Dim rowICTQUOT2 As DataRow = dst.Tables.Item("ICTQUOT2").Select(fltrICTQUOT2).FirstOrDefault

            Dim STYLE_DESC As String = rowSB.Item("STYLE_DESC").ToString & String.Empty
            Dim SIZE_SCALE As String = GET_ONLY_SIZE_SCALE(STYLE_CODE)

            Dim IMAGE_NAME As String = rowSB.Item("IMAGE_NAME") & ""
            Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME
            Dim HasImage As Boolean = False
            Dim imageStyle As System.Drawing.Image = Nothing
            If My.Computer.FileSystem.FileExists(imageFileStyle) Then
                imageStyle = System.Drawing.Image.FromFile(imageFileStyle)
                HasImage = True
            End If
            worksheet.Cells("A" & curRow.ToString & ":" & "R" & curRow.ToString).Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            worksheet.Cells("A" & curRow.ToString).RowHeight = 100.5
            worksheet.Cells("E" & curRow.ToString).Value = FACTORY_CODE
            worksheet.Cells("F" & curRow.ToString).Value = SALES_DIVISION_NAME
            worksheet.Cells("G" & curRow.ToString).Value = SIZE_SCALE
            If HasImage Then
                Dim leftStyle As Integer = windowInfoStyle.ColumnToPoints(7)
                Dim topStyle As Integer = windowInfoStyle.RowToPoints(curRow - 1) + 0.1
                Dim WidthStyle As Integer = 100
                Dim HeightStyle As Integer = 99
                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle + 20, topStyle + 1, WidthStyle, HeightStyle)
            End If
            worksheet.Cells("I" & curRow.ToString).Value = STYLE_CODE
            worksheet.Cells("J" & curRow.ToString).Value = STYLE_DESC
            worksheet.Cells("K" & curRow.ToString).Value = COLOR_CODE & " - " & STYLE_COLOR_DESC
            worksheet.Cells("L" & curRow.ToString).Value = COUNTRY_NAME
            Dim TOT_AVAIL As Int64 = 0
            Dim DATES As New System.Text.StringBuilder With {.Length = 0}
            Dim FOBDATES As New System.Text.StringBuilder With {.Length = 0}

            Dim DATES_STRING As String = GetCustShipDates(STYLE_CODE, COLOR_CODE)
            '   Dim DATES_STRING As String = CDate(rowSB.Item("ORDR_SHIP_DATE").ToString & String.Empty)
            If Val(rowSB.Item("QTY_SHP").ToString & String.Empty) <> 0 Then
                TOT_AVAIL = Val(rowSB.Item("QTY_SHP").ToString & String.Empty)
            End If
            Dim FOBDATES_STRING As String = ""
            With worksheet.Cells("M" & curRow.ToString)
                .Value = DATES_STRING
                .Font.Color = SpreadsheetGear.Colors.Red
            End With
            With worksheet.Cells("O" & curRow.ToString)
                .Value = TOT_AVAIL
                .NumberFormat = "###,##0"
            End With
            With worksheet.Cells("P" & curRow.ToString)
                .Value = VAN_COST 'Get Vandale Cost Here
                '  .NumberFormat = "$###,##0.00"
                .Font.Color = SpreadsheetGear.Colors.Red

            End With
            With worksheet.Cells("Q" & curRow.ToString)
                .Value = STYLE_RETAIL
                .NumberFormat = "$###,##0.00"
            End With
            With worksheet.Cells("R" & curRow.ToString)
                ' .Value = 3.3 'Get TKMAX OFFER Here
                .NumberFormat = "$###,##0.00"
                .Interior.Color = SpreadsheetGear.Colors.Yellow
                .Font.Color = SpreadsheetGear.Colors.Red
                .VerticalAlignment = SpreadsheetGear.VAlign.Center
            End With
            With worksheet.Cells("S" & curRow.ToString)
                .Value = FOBDATES_STRING
                .Font.Color = SpreadsheetGear.Colors.Red
            End With
            worksheet.Cells("T" & curRow.ToString).Value = FACTORY_DESC

            ''If chkShowLastRcd.Checked Then
            ''    If IsDate(rowSB.Item("LAST_RCD_DATE").ToString & String.Empty) Then
            ''        Dim LAST_SHIPPED As Date = CDate(rowSB.Item("LAST_RCD_DATE").ToString & String.Empty)
            ''        worksheet.Cells("U" & curRow.ToString).Value = Format(LAST_SHIPPED, "MM/dd/yy")
            ''    Else
            ''        Dim LAST_SHIPPED As String = rowSB.Item("LAST_RCD_DATE").ToString & String.Empty
            ''        worksheet.Cells("U" & curRow.ToString).Value = LAST_SHIPPED
            ''    End If
            ''Else
            ''    worksheet.Cells("U" & curRow.ToString).Value = ""
            ''End If

            worksheet.Cells("U" & curRow.ToString).Value = ""
            If chkStyleStats.Checked Then

                For Each rowICTSTAT2 As DataRow In dst.Tables("ICTSTAT2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                    worksheet.Cells("V" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & String.Empty)
                    worksheet.Cells("V" & curRow.ToString).NumberFormat = "#,###,##0"
                    worksheet.Cells("W" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & String.Empty)
                    worksheet.Cells("W" & curRow.ToString).NumberFormat = "#,###,##0"

                    Dim OTS As Integer = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "") - Val(rowICTSTAT2.Item("WHSE_QTY_PICK") & "")
                    worksheet.Cells("X" & curRow.ToString).Value = OTS
                    worksheet.Cells("X" & curRow.ToString).NumberFormat = "#,###,##0"

                    worksheet.Cells("Y" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty)
                    worksheet.Cells("Y" & curRow.ToString).NumberFormat = "#,###,##0"

                    worksheet.Cells("Z" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty)
                    worksheet.Cells("Z" & curRow.ToString).NumberFormat = "#,###,##0"

                    worksheet.Cells("AA" & curRow.ToString).Value = Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
                    worksheet.Cells("AA" & curRow.ToString).NumberFormat = "#,###,##0"

                    worksheet.Cells("AB" & curRow.ToString).Value = OTS + Val(rowICTSTAT2.Item("WHSE_QTY_TRAN") & String.Empty) + Val(rowICTSTAT2.Item("WHSE_QTY_ON_ORDER") & String.Empty) - Val(rowICTSTAT2.Item("WHSE_QTY_OPEN") & String.Empty)
                    worksheet.Cells("AB" & curRow.ToString).NumberFormat = "#,###,##0"
                Next

            End If



            curRow += 1
        Next

        'Show Workbook
        Dim XLS_FILENAME As String = "5000"
        Dim success As Boolean = False
        Dim RPT_PREFIX As String = "BuyerChart"
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = RPT_PREFIX & "_" & Format(XLS_NO, "000") & exlExt
                workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                RetVal = XLS_FILENAME
                success = True
            Catch ex As Exception
                If XLS_NO > 5000 Then
                    success = True
                End If
            End Try
        Loop
        If XLS_FILENAME = "5000" Then
            MsgBox("Reports In Temp Folder Exceeded", vbCritical, "Log Out Of ABS And Get Back In")
        Else
            Show_Document(ASCMAIN1.Folders("Temp") & XLS_FILENAME)
        End If

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default

        Return RetVal
    End Function
    Private Function GET_ONLY_SIZE_SCALE(ByVal STYLE_CODE As String) As String
        Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)
        'If STYLE_CODE = "VCO51279" Then
        '    Stop
        'End If
        Dim SIZEs As String = ""
        Dim QTYs As String = ""
        Dim SIZEs_And_QTYs As String = ""
        If rowICTSTYLS IsNot Nothing Then
            If rowICTSTYLS.Item("SIZE_01") & "" <> "" Then
                For iSZ As Integer = 1 To 24
                    If rowICTSTYLS.Item("SIZE_" & Format(iSZ, "00")) & "" = "" Then
                        Exit For
                    Else
                        SIZEs &= "-" & rowICTSTYLS.Item("SIZE_" & Format(iSZ, "00")) & ""
                        QTYs &= "/" & CStr(Val(rowICTSTYLS.Item("QTY_" & Format(iSZ, "00")) & ""))
                    End If
                Next
                SIZEs = Mid(SIZEs, 2) ' just the sizes
                If Not QTYs.StartsWith("/0") Then
                    SIZEs_And_QTYs = SIZEs & " = " & Mid(QTYs, 2)
                Else
                    SIZEs_And_QTYs = SIZEs
                End If
            End If
        End If
        Return SIZEs_And_QTYs
    End Function
    Sub RESEQ()

        Dim SEQ As Integer = 0
        Dim OLDSTYLE As String = ""


        Dim SORTSOTCUSTS As String = "SUB_BODY_CODE,FABRIC_CODE,STYLE_CODE,COLOR_CODE"
        ''If chkSortStyle.Checked Then
        ''    SORTSOTCUSTQ = "STYLE_CODE,COLOR_CODE"
        ''End If

        Dim sqlWB As String = ""
        If chk1Sheet.Checked Then
            sqlWB = ",SALES_DIVISION_CODE," & SORTSOTCUSTS
        Else
            sqlWB = "," & SORTSOTCUSTS

        End If

        For Each row As DataRow In dst.Tables("SOTCUSTS").Select("", Mid(sqlWB, 2))
            If OLDSTYLE = "" Or OLDSTYLE <> row.Item("STYLE_CODE") Then
                SEQ += 10
                row.Item("SEQ") = SEQ
                row.Item("STYLE_CODE_PLM") = row.Item("STYLE_CODE")
                row.Item("SELECTED") = "1"
                OLDSTYLE = row.Item("STYLE_CODE")
            End If
        Next
    End Sub

    Sub Print_Full_CAD_Print(eItemKey As String, Optional STYLE_CODE As String = "")
        Dim ListPDFSheets As New List(Of String)
        Dim MISSING_IMAGES As New List(Of String)


        RESEQ()

        Dim EXCUDE_FUTURE As String = ""

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        If Not FOLDER_NAME.EndsWith("\") Then FOLDER_NAME &= "\"
        FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")

        For Each row As DataRow In dst.Tables("SOTCUSTS").Select("SELECTED='1'")

            Dim STYLE_CODE_PLM As String = row.Item("STYLE_CODE")
            'If STYLE_CODE_PLM = "500498AVR" And ASCMAIN1.Running_in_VS Then Stop
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "dgj")) Then
                'Stop
                FOLDER_NAME = "S:\VAN\images\"
                'FOLDER_NAME = "\\192.168.180.32\g\VAN\images\"
            End If

            If Not My.Computer.FileSystem.FileExists(FOLDER_NAME & row.Item("IMAGE_NAME")) Then
                row.Item("SELECTED") = "0"
                MISSING_IMAGES.Add(STYLE_CODE_PLM)
            End If

        Next

        Dim RPT As String = ""
        ''If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        ''    RPT = "ICRQUOT2"
        ''End If

        Dim ColVisible(4) As Boolean
        RPT = "SORFCADY"

        If eItemKey = "email" Then
            ''Dim tempFileName As String = rowICTQUOT1.Item("QUOTE_NO")

            ''Dim REPORT_NO As String = Generate_Report(RPT, "Quote Sheet", "", "", "PDF", tempFileName, False)
            ''' Dim FILENAME As String = REPORT_FILENAMES(REPORT_NO)
            ''Print_Report_End(, True)
            ''email_Quote(tempFileName)
        Else

            For Each row As DataRow In dst.Tables("SOTCUSTS").Select("SELECTED='1'")
                row.Item("SELECTED") = "2"
            Next

            Dim REPORT_INDEX As Integer = 0
            Dim PDF_FN As String = ""
            Dim PDF_LINKS As String = ""
            Dim SUB_BODY_DESC As String = ""
            Dim SALES_DIVISION_NAME As String = ""
            Dim FABRIC_DESC As String = ""
            Dim DESCHASH As String = ""

            Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/"
            'Dim LINEPFX_NEW As String = "https://docs.vandalequotes.com/"
            Dim LINEPFX_NEW As String = "https://vandaledocs.azurewebsites.net/Documents/"

            Dim SESSION_NO As String = ASCMAIN1.Next_Control_No("SOTCUSTS.SESSION_NO")
            Dim FILE_NO As Integer = 0

            Do While dst.Tables("SOTCUSTS").Select("SELECTED='2'").Length <> 0

                Print_Report_Begin()

                ''    Dim STYLE_count As Integer = 0
                ''    Dim SRT As String = "SEQ"
                ''    Select Case opt1Sheet.Value
                ''        Case "S"
                ''            SRT = "SUB_BODY_CODE, STYLE_CODE_PLM"
                ''        Case "FS"
                ''            SRT = "FABRIC_CODE, SUB_BODY_CODE, STYLE_CODE_PLM"
                ''        Case "G"
                ''            SRT = "STYLE_GROUP_CODE, STYLE_CODE_PLM"
                ''        Case "D"
                ''            SRT = "SALES_DIVISION_CODE"
                ''    End Select
                ''    For Each row As DataRow In dst.Tables("ICTQUOT2").Select(sqlw, SRT)
                ''        STYLE_count += 1
                ''        row.Item("SELECTED") = "1"
                ''        SetRowImage(row)
                ''    Next
                For Each row As DataRow In dst.Tables("SOTCUSTS").Select()
                    row.Item("IMAGE") = Null
                Next
                Dim STYLE_count As Integer = 0
                For Each row As DataRow In dst.Tables("SOTCUSTS").Select("SELECTED='2'", "SEQ")
                    STYLE_count += 1
                    row.Item("SELECTED") = "1"
                    SetRowImage(row)
                    If STYLE_count >= 50 Then Exit For
                Next
                Application.DoEvents()

                CR_params.Add("IMAGES_FOLDER", FOLDER_NAME)

                CR_params.Add("TXTSTYLE_CODE", "")

                Dim tempFileName As String = ""
                Do
                    REPORT_INDEX += 1
                    tempFileName = "SORCUSTS" & "-" & Format(REPORT_INDEX, "000")
                Loop While My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")

                Dim REPORT_NO As String = Generate_Report(RPT, "Open Order", "", "", "PDF", tempFileName, False)

                Dim tempNotMade As Boolean = Not System.IO.File.Exists(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")

                If Not tempNotMade Then
                    'Show_Document(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
                    ListPDFSheets.Add(ASCMAIN1.Folders("Temp") & tempFileName & ".PDF")
                    Print_Report_End(, True)
                End If

                For Each row As DataRow In dst.Tables("SOTCUSTS").Select("SELECTED='1'")
                    row.Item("SELECTED") = "3"
                    row.Item("IMAGE") = DBNull.Value
                Next
            Loop

        End If

        For Each row As DataRow In dst.Tables("SOTCUSTS").Select("")
            row.Item("IMAGE") = Nothing
        Next

        For Each PDF As String In ListPDFSheets
            Show_Document(PDF)
        Next

        If MISSING_IMAGES.Count > 0 Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Missing Images"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("The Following Styles Did Not Have")
            iMSG.AppendLine("Set-up In The Style Masterfile:")
            For Each MI As String In MISSING_IMAGES
                iMSG.AppendLine("-> " & MI)
            Next
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        End If
    End Sub

    Private Sub SetRowImage(row As DataRow)
        Dim STYLE_CODE As String
        If row.Table.TableName = "SOTCUSTS" Then
            STYLE_CODE = row.Item("STYLE_CODE_PLM") & ""
        Else
            STYLE_CODE = row.Item("STYLE_CODE") & ""
        End If

        Dim IMAGE_NAME As String = row.Item("IMAGE_NAME") & ""

        If IMAGE_NAME = "" Then IMAGE_NAME = STYLE_CODE

        ''If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
        ''    IMAGE_NAME = STYLE_CODE & "-" & COLOR_CODE
        ''End If

        'Dim imgba() As Byte = Nothing
        Dim imgb As System.Drawing.Bitmap = Nothing
        If IMAGE_NAME <> "" Then
            Dim ex_err As Exception = Nothing
            Dim IMAGE_FILE_USED As String = ""
            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                FOLDER_NAME = Replace(FOLDER_NAME, "G:", "R:")
                ''If chkLowRes.Checked Then
                ''Dim FILE_NAME_LOW_RES As String = String.Format("{0}{1}{2}", FOLDER_NAME, "_lowres\", IMAGE_NAME)
                ''If System.IO.File.Exists(FILE_NAME_LOW_RES) Then
                ''    FOLDER_NAME = FOLDER_NAME & "_lowres"
                ''    IMAGE_FILE_USED = FILE_NAME_LOW_RES
                ''Else
                ''    IMAGE_FILE_USED = String.Format("{0}{1}{2}", FOLDER_NAME, "\", IMAGE_NAME)
                ''End If
                ''End If
            End If
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "dgj")) Then
                'Stop
                FOLDER_NAME = "S:\VAN\images\"
                'FOLDER_NAME = "\\192.168.180.32\g\VAN\images\"
            End If


            Dim img As System.Drawing.Bitmap = Nothing

            Dim image_file_found As Boolean = True

            If IMAGE_NAME = "\.jpg" Then
                image_file_found = False
                Exit Sub
            End If

            If Not FOLDER_NAME.EndsWith("\") Then FOLDER_NAME &= "\"
            Dim IMAGE_FILENAME As String = FOLDER_NAME & IMAGE_NAME
            Try
                If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then

                ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".PNG") Then
                    IMAGE_FILE_USED &= ".PNG"
                ElseIf My.Computer.FileSystem.FileExists(IMAGE_FILENAME & ".JPG") Then
                    IMAGE_FILE_USED &= ".JPG"
                Else
                    image_file_found = False
                    img = Nothing
                End If
            Catch ex As Exception
                image_file_found = False
                img = Nothing
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    ex_err = ex
                End If
            End Try

            Dim fs As System.IO.FileStream = New System.IO.FileStream(IMAGE_FILENAME, System.IO.FileMode.Open)
            Dim newBMP As System.Drawing.Bitmap = New System.Drawing.Bitmap(System.Drawing.Image.FromStream(fs))
            Dim scaleFactor As Double = 1  ' 1 (trkScaleImage.Value / 100)
            Dim newBMP2 As System.Drawing.Bitmap = New System.Drawing.Bitmap(newBMP, newBMP.Width * scaleFactor, newBMP.Height * scaleFactor)
            Application.DoEvents()
            Try
                'newBMP.MakeTransparent(System.Drawing.Color.White)
                Dim converter As New System.Drawing.ImageConverter
                'row.Item("IMAGE") = converter.ConvertTo(newBMP, GetType(Byte()))
                row.Item("IMAGE") = converter.ConvertTo(newBMP2, GetType(Byte()))
                newBMP.Dispose()
                newBMP2.Dispose()
            Catch ex As Exception
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    ex_err = ex
                End If
            End Try
            fs.Close()
            Application.DoEvents()
            If Not IsNothing(ex_err) Then
                If Not IMG_Error_Reported Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Error Getting Image"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("The Following Error Occured While Attempting To ")
                    iMSG.AppendLine("Get An Image:")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Style: " & STYLE_CODE)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Image: " & IMAGE_NAME)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Error: " & ex_err.Message)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Please Relay This Information To Wayne At ABS.")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
                    IMG_Error_Reported = True
                End If
            End If
            'Dim converter As New ImageConverter
            'row.Item("IMAGE") = converter.ConvertTo(imgb, GetType(Byte()))
            'row.Item("IMAGE") = imgb
            'UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE & "-" & COLOR_CODE
        Else
            'row.Item("IMAGE") = DBNull.Value
            'UltraExplorerBar1.Groups("Style Image").Text = "Style Image"
        End If

    End Sub

End Class