Imports System.Text
Imports SpreadsheetGear

Public Class SORCUSTS
    Dim S As New StringBuilder With {.Length = 0}
    Dim SQL_WHERE_ORDR_RSV As String
    Dim REPORT_NAME As String = "SORCUSTS"
    Dim ICTSTATDSQL As String = ""
    Dim ICTSTAT2SQL As String = ""
    Dim SQL_REPORT As New StringBuilder With {.Length = 0}
    Dim GRP_IN As String = ""
    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

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
        For Each rowSOTCUSTQ As DataRow In dst.Tables("SOTCUSTS").Select()
            Dim STYLE_CODE As String = rowSOTCUSTQ.Item("STYLE_CODE").ToString & String.Empty
            If Not StyleList.Contains(STYLE_CODE) Then
                StyleList.Add(STYLE_CODE)
            End If
        Next

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        worksheet.Name = "Style Info"
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
                               ByVal StyleList As List(Of String))

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
                For iCOL As Integer = 1 To 7
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
        'worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"

        Dim GT = ""
        For iCOL As Integer = 1 To 6
            COL += 1
            Select Case iCOL
                Case 4
                    'worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                    GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    'Case 5
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

        'worksheet.Cells(I - 1, COL0 - 1, I - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

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
End Class