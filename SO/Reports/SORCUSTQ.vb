Imports System.Text
Imports SpreadsheetGear

Public Class SORCUSTQ
    Dim S As New StringBuilder With {.Length = 0}
    Dim SQL_WHERE_ORDR_RSV As String
    Dim REPORT_NAME As String = "SORCUSTQ"
    Dim SQL_REPORT As New StringBuilder With {.Length = 0}

#Region "Report Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RWU = "N"
        Get_PARM("ICTPARM1")

        Build_Init_Sel()

        Fill_Records("WEBLINKS")

        grdSOTORDR0.DataSource = dst.Tables.Item("SOTORDR0")
        grdSOTRSRV1.DataSource = dst.Tables.Item("SOTRSRV1")
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

        With grdSOTORDR0.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
            For Each COLNAME As String In New String() {"SEL"}
                .Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
            For Each COLNAME As String In New String() {"SEL"}
                .Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            Next
            For Each COLNAME As String In New String() {"ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Bands(0).Columns(COLNAME).Format = "MM/dd/yy"
            Next
        End With

        With grdSOTRSRV1.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
            For Each COLNAME As String In New String() {"SEL"}
                .Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
            For Each COLNAME As String In New String() {"SEL"}
                .Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            Next
            For Each COLNAME As String In New String() {"ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Bands(0).Columns(COLNAME).Format = "MM/dd/yy"
            Next
        End With

    End Sub

    Protected Overrides Sub Build_Workfile()
        Call ASCMAIN1.Progress("Building Work File")

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

        S.Length = 0
        S.AppendLine("SELECT 'RESERVATION' AS ORDR_CUST_PO, '0000000000' AS ORDR_GROUP_NO FROM DUAL")
        ASCMAIN1.sql = S.ToString
        Dim TABLE_TEMP As String = ASCMAIN1.Temp_Table

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("JN.CUST_CODE, JN.CUST_NAME,")
        S.AppendLine("JN.ORDR_CUST_PO, JN.ORDR_SHIP_DATE, JN.ORDR_CANCEL_DATE,")
        S.AppendLine("JN.STYLE_CODE, JN.COLOR_CODE, JN.COLOR_DESC, JN.STYLE_DESC, JN.FABRIC_CODE, JN.SEASON_CODE,")
        S.AppendLine("JN.SUB_BODY_CODE, JN.SALES_DIVISION_CODE, JN.INNER_PACK_QTY, JN.CARTON_PACK_QTY, JN.STYLE_CUST_CODE, JN.IMAGE_NAME,")
        S.AppendLine("SUM(JN.RSRV_QTY) RSRV_QTY,")
        S.AppendLine("SUM(JN.RSRV_QTY_OPEN) RSRV_QTY_OPEN,")
        S.AppendLine("SUM(JN.ORDR_QTY) ORDR_QTY,")
        S.AppendLine("SUM(JN.ORDR_QTY_OPEN) ORDR_QTY_OPEN,")
        S.AppendLine("SUM(JN.ORDR_QTY_PICK) ORDR_QTY_PICK,")
        S.AppendLine("SUM(JN.ORDR_QTY_CANC) ORDR_QTY_CANC,")
        S.AppendLine("SUM(JN.ORDR_QTY_SHIP) ORDR_QTY_SHIP")
        S.AppendLine("FROM (")
        S.AppendLine("  SELECT")
        S.AppendLine("  O1.ORDR_GROUP_NO,")
        S.AppendLine("  '0000000000' AS RSRV_NO,")
        S.AppendLine("  A1.CUST_CODE, A1.CUST_NAME,")
        S.AppendLine("  O1.ORDR_CUST_PO, O1.ORDR_SHIP_DATE, O1.ORDR_CANCEL_DATE,")
        S.AppendLine("  I1.STYLE_CODE, O2.COLOR_CODE, C1.COLOR_DESC, I1.STYLE_DESC, I1.FABRIC_CODE, I1.SEASON_CODE, I1.SUB_BODY_CODE,")
        S.AppendLine("  I1.SALES_DIVISION_CODE, I1.INNER_PACK_QTY, I1.CARTON_PACK_QTY, I1.CUST_CODE STYLE_CUST_CODE, I1.IMAGE_NAME,")
        S.AppendLine("  0 RSRV_QTY,")
        S.AppendLine("  0 RSRV_QTY_OPEN,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY,0)) ORDR_QTY,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_PICK,0)) ORDR_QTY_PICK,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_CANC,0)) ORDR_QTY_CANC,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP")
        S.AppendLine("  FROM ARTCUST1 A1, SOTORDR1 O1, SOTORDR2 O2, ICTSTYL1 I1, ICTCOLR1 C1")
        S.AppendLine("  WHERE A1.CUST_CODE = O1.CUST_CODE")
        S.AppendLine(String.Format("  AND NVL(O1.CUST_CODE,'NULL') = '{0}'", txtCUST_CODE.Text))
        S.AppendLine("  AND O1.ORDR_NO = O2.ORDR_NO")
        S.AppendLine("  AND O2.STYLE_CODE = I1.STYLE_CODE")
        S.AppendLine("  AND O2.COLOR_CODE = C1.COLOR_CODE")
        S.AppendLine("  AND O1.ORDR_STATUS IN ('P','O')")
        S.AppendLine("  GROUP BY")
        S.AppendLine("  O1.ORDR_GROUP_NO,")
        S.AppendLine("  A1.CUST_CODE, A1.CUST_NAME,")
        S.AppendLine("  O1.ORDR_CUST_PO, O1.ORDR_SHIP_DATE, ORDR_CANCEL_DATE,")
        S.AppendLine("  I1.STYLE_CODE, O2.COLOR_CODE, C1.COLOR_DESC, I1.STYLE_DESC, I1.FABRIC_CODE, I1.SEASON_CODE, I1.SUB_BODY_CODE,")
        S.AppendLine("  I1.SALES_DIVISION_CODE, I1.INNER_PACK_QTY, I1.CARTON_PACK_QTY, I1.CUST_CODE, I1.IMAGE_NAME")
        S.AppendLine("  UNION")
        S.AppendLine("  SELECT")
        S.AppendLine("  '0000000000' AS ORDR_GROUP_NO,")
        S.AppendLine("  R1.RSRV_NO,")
        S.AppendLine("  A1.CUST_CODE, A1.CUST_NAME,")
        S.AppendLine("  'RESERVATION' ORDR_CUST_PO, R1.ORDR_SHIP_DATE, R1.ORDR_CANCEL_DATE,")
        S.AppendLine("  I1.STYLE_CODE, R2.COLOR_CODE, C1.COLOR_DESC, I1.STYLE_DESC, I1.FABRIC_CODE, I1.SEASON_CODE,")
        S.AppendLine("  I1.SUB_BODY_CODE, I1.SALES_DIVISION_CODE, I1.INNER_PACK_QTY, I1.CARTON_PACK_QTY, I1.CUST_CODE STYLE_CUST_CODE, I1.IMAGE_NAME,")
        S.AppendLine("  SUM(NVL(R2.RSRV_QTY,0)) RSRV_QTY,")
        S.AppendLine("  SUM(NVL(R2.RSRV_QTY_OPEN,0)) RSRV_QTY_OPEN,")
        S.AppendLine("  0 ORDR_QTY,")
        S.AppendLine("  0 ORDR_QTY_OPEN,")
        S.AppendLine("  0 ORDR_QTY_PICK,")
        S.AppendLine("  0 ORDR_QTY_CANC,")
        S.AppendLine("  0 ORDR_QTY_SHIP")
        S.AppendLine("  FROM ARTCUST1 A1, SOTRSRV1 R1, SOTRSRV2 R2, ICTSTYL1 I1, ICTCOLR1 C1")
        S.AppendLine("  WHERE A1.CUST_CODE = R1.CUST_CODE")
        S.AppendLine(String.Format("  AND NVL(R1.CUST_CODE,'NULL') = '{0}'", txtCUST_CODE.Text))
        S.AppendLine("  AND R1.RSRV_NO = R2.RSRV_NO")
        S.AppendLine("  AND R2.STYLE_CODE = I1.STYLE_CODE")
        S.AppendLine("  AND R2.COLOR_CODE = C1.COLOR_CODE")
        S.AppendLine("  AND R1.RSRV_STATUS IN ('O')")
        S.AppendLine("  GROUP BY")
        S.AppendLine("  R1.RSRV_NO,")
        S.AppendLine("  A1.CUST_CODE, A1.CUST_NAME,")
        S.AppendLine("  R1.ORDR_SHIP_DATE, R1.ORDR_CANCEL_DATE,")
        S.AppendLine("  I1.STYLE_CODE, R2.COLOR_CODE, C1.COLOR_DESC, I1.STYLE_DESC, I1.FABRIC_CODE, I1.SEASON_CODE,")
        S.AppendLine("  I1.SUB_BODY_CODE, I1.SALES_DIVISION_CODE, I1.INNER_PACK_QTY, I1.CARTON_PACK_QTY, I1.CUST_CODE, I1.IMAGE_NAME")
        S.AppendLine(") JN")
        S.AppendLine(String.Format("WHERE NVL(JN.CUST_CODE,'NULL') = '{0}'", txtCUST_CODE.Text))
        S.AppendLine(sql_WHERE)
        S.AppendLine(SQL_WHERE_ORDR_RSV)
        S.AppendLine(sql_filter2)
        If Absx1.optFor("OPTASN").Value = "S" Then
            S.AppendLine("AND JN.STYLE_CUST_CODE IS NULL")
        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
            S.AppendLine("AND JN.STYLE_CUST_CODE IS NOT NULL")
        End If
        S.AppendLine("GROUP BY")
        S.AppendLine("JN.CUST_CODE, JN.CUST_NAME,")
        S.AppendLine("JN.ORDR_CUST_PO, JN.ORDR_SHIP_DATE, JN.ORDR_CANCEL_DATE,")
        S.AppendLine("JN.STYLE_CODE, JN.COLOR_CODE, JN.COLOR_DESC, JN.STYLE_DESC, JN.FABRIC_CODE, JN.SEASON_CODE,")
        S.AppendLine("JN.SUB_BODY_CODE, JN.SALES_DIVISION_CODE, JN.INNER_PACK_QTY, JN.CARTON_PACK_QTY, JN.STYLE_CUST_CODE, JN.IMAGE_NAME")
        S.AppendLine("ORDER BY")
        S.AppendLine(sql_GROUP_BY_cols)
        ASCMAIN1.sql = S.ToString()
        Create_TDA(dst.Tables.Add, "SOTCUSTQ", "**", 0, False)
        'With dst.Tables("ICTQUOTQ").Columns
        '    .Add("LAST_RCD_DATE")
        'End With
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
        For Each rowSOTCUSTQ As DataRow In dst.Tables("SOTCUSTQ").Select()
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
        Dim FULLPATH As String = "\\192.168.180.35\g\VDI\ARCHIVE\VAN\Links\" & FILE_NAME
        Dim HASHVALUE As String = ASCMAIN1.Get_Hash(SESSION_NO & String.Format("{0}.XLSX", REPORT_NAME))

        If IsNothing(dst.Tables.Item("WEBLINKS")) Then
            ASCMAIN1.sql = SQL_REPORT.ToString
            Create_TDA(dst.Tables.Add, "WEBLINKS", "**", 0, True)
        End If

        If System.IO.File.Exists(FULLPATH) Then
            System.IO.File.Delete(FULLPATH)
        End If
        System.IO.File.Copy(FILENAME_FULL, FULLPATH)

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

        End If
        If eItemKey = "Done" Then
            Build_Init_Sel()
        End If

        If EMsg.Length = 0 Then
            BuildSpecialWhere()
            btnLoadOrders.Enabled = False
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
        Fill_Records("SOTCUSTQ")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWEBLINKS, "BBB", "View File", "Replace File", "Copy Link", "Extend Expiration")
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
                Dim FILE_NAME As String = grd.ActiveRow.Cells.Item("FILE_NAME").Text
                Dim HASH As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
                Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
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

#End Region

#Region "Custom Methods"
    Private Sub Build_Init_Sel()
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("'0' AS SEL,")
        S.AppendLine("CUST_CODE,")
        S.AppendLine("RSRV_NO,")
        S.AppendLine("ORDR_SHIP_DATE,")
        S.AppendLine("ORDR_CANCEL_DATE,")
        S.AppendLine("ORDR_CUST_PO")
        S.AppendLine("FROM SOTRSRV1")
        S.AppendLine("WHERE RSRV_STATUS = 'O'")
        S.AppendLine("AND CUST_CODE = :PARM1")
        ASCMAIN1.sql = S.ToString
        Create_TDA(dst.Tables.Add, "SOTRSRV1", "**", 0, False, "V")
        'Create_TDA(dst.Tables.Add, "SOTRSRV1", "**", 0, False)

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("'0' AS SEL,")
        S.AppendLine("CUST_CODE,")
        S.AppendLine("ORDR_GROUP_NO,")
        S.AppendLine("ORDR_SHIP_DATE,")
        S.AppendLine("ORDR_CANCEL_DATE,")
        S.AppendLine("ORDR_CUST_PO,")
        S.AppendLine("ORDR_AMT")
        S.AppendLine("FROM SOTORDR0")
        S.AppendLine("WHERE ORDR_GROUP_NO IN")
        S.AppendLine("(")
        S.AppendLine("  SELECT ORDR_GROUP_NO")
        S.AppendLine("  FROM SOTORDR1")
        S.AppendLine("  WHERE CUST_CODE = :PARM1")
        S.AppendLine("  AND ORDR_STATUS IN ('O','P')")
        S.AppendLine(")")
        ASCMAIN1.sql = S.ToString
        Create_TDA(dst.Tables.Add, "SOTORDR0", "**", 0, False, "V")

        SQL_REPORT.Length = 0
        SQL_REPORT.AppendLine("SELECT *")
        SQL_REPORT.AppendLine("FROM WEBLINKS")
        SQL_REPORT.AppendLine(String.Format("WHERE FORM_NAME = '{0}'", REPORT_NAME))
        If IsNothing(dst.Tables.Item("WEBLINKS")) Then
            ASCMAIN1.sql = SQL_REPORT.ToString
            Create_TDA(dst.Tables.Add, "WEBLINKS", "**", 0, True)
        End If
    End Sub

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

    Private Sub UpdateReportRows()
        For Each rowSOTCUSTQ As DataRow In dst.Tables("SOTCUSTQ").Select()
            'Dim RSRV_QTY As Int64 = Val(rowSOTCUSTQ.Item("RSRV_QTY") & String.Empty)
            'Dim ORDR_QTY As Int64 = Val(rowSOTCUSTQ.Item("ORDR_QTY") & String.Empty)
            'Dim ORDR_QTY_CANC As Int64 = Val(rowSOTCUSTQ.Item("ORDR_QTY_CANC") & String.Empty)
            'Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTCUSTQ.Item("ORDR_QTY_SHIP") & String.Empty)
            'If rowSOTCUSTQ.Item("STYLE_CODE") & String.Empty = "7038IZ" Then Stop
            Dim RSRV_QTY_OPEN As Int64 = Val(rowSOTCUSTQ.Item("RSRV_QTY_OPEN") & String.Empty)
            Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTCUSTQ.Item("ORDR_QTY_OPEN") & String.Empty)
            Dim ORDR_QTY_PICK As Int64 = Val(rowSOTCUSTQ.Item("ORDR_QTY_PICK") & String.Empty)
            Dim LINETOTAL As Int64 = 0
            If chkReservations.Checked Then
                LINETOTAL += RSRV_QTY_OPEN
            End If
            If chkOpen.Checked Then
                LINETOTAL += ORDR_QTY_OPEN
            End If
            If chkPick.Checked Then
                LINETOTAL += ORDR_QTY_PICK
            End If
            If LINETOTAL > 0 Then
                Dim COLOR_CODE As String = rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty
                Dim STYLE_CODE As String = rowSOTCUSTQ.Item("STYLE_CODE") & String.Empty
                Dim COLOR_DESC As String = rowSOTCUSTQ.Item("COLOR_DESC") & String.Empty
                rowSOTCUSTQ.Item("COLOR_DESC") = GetAltColorCode(STYLE_CODE, COLOR_CODE, COLOR_DESC)
            Else
                rowSOTCUSTQ.Delete()
            End If
        Next
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
        Dim RT(6) As String
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

            For iCOL As Integer = 1 To 6
                COL += 1
                Select Case iCOL
                    Case 4
                        If chkReservations.Checked Then
                            worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        Else
                            COL -= 1
                        End If
                    Case 5
                        If chkOpen.Checked Then
                            worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        Else
                            COL -= 1
                        End If
                    Case 6
                        If chkPick.Checked Then
                            worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                        Else
                            COL -= 1
                        End If
                End Select

                RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
            Next

            COL += 1

            worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL - 1).Interior.Color = SpreadsheetGear.Colors.LightGray

            With worksheet.Cells(I, COL0 - 1, I + CI - 1, COL - 1)
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

            With worksheet.Cells(I0, 0, I + 1 - 1, COL)
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With
        Next

        I += 2
        COL = COL0

        worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"

        Dim GT = ""
        For iCOL As Integer = 1 To 6
            COL += 1
            Select Case iCOL
                Case 4
                    If chkReservations.Checked Then
                        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                        GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    Else
                        COL -= 1
                    End If
                Case 5
                    If chkOpen.Checked Then
                        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                        GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    Else
                        COL -= 1
                    End If
                Case 6
                    If chkPick.Checked Then
                        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                        GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                    Else
                        COL -= 1
                    End If
            End Select
        Next

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
        For Each rowSOTCUSTQ As DataRow In dst.Tables("SOTCUSTQ").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE, ORDR_SHIP_DATE")
            CI += 1
            COL = COL0
            Dim chkcnt As Int64 = 4
            If LAST_COLOR <> rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty Then
                worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty
                worksheet.Cells(i + CI - 1, COL).Value = rowSOTCUSTQ.Item("COLOR_DESC") & String.Empty
                LAST_COLOR = rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty
            End If
            worksheet.Cells(i + CI - 1, COL + 1).Value = rowSOTCUSTQ.Item("ORDR_CUST_PO") & String.Empty
            worksheet.Cells(i + CI - 1, COL + 2).Value = rowSOTCUSTQ.Item("ORDR_SHIP_DATE") & String.Empty
            worksheet.Cells(i + CI - 1, COL + 3).Value = rowSOTCUSTQ.Item("ORDR_CANCEL_DATE") & String.Empty
            If chkReservations.Checked Then
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("RSRV_QTY_OPEN") & String.Empty
                chkcnt += 1
            End If
            If chkOpen.Checked Then
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("ORDR_QTY_OPEN") & String.Empty
                chkcnt += 1
            End If
            If chkPick.Checked Then
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("ORDR_QTY_PICK") & String.Empty
                chkcnt += 1
            End If

            T = ""
            COL += 1
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

        worksheet.Cells(i, COL - 1).Value = "Color"
        worksheet.Cells(i, COL).Value = "Description"

        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "PO"
        End With

        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Ship"
        End With

        COL += 1
        With worksheet.Cells(i, COL)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .Value = "Cancel"
        End With

        If chkReservations.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "Reserved"
            End With
        End If

        If chkOpen.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "Open"
            End With
        End If

        If chkPick.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "Pick"
            End With

        End If

        range = worksheet.Cells(i, COL0 - 1, i, COL)
        interior = range.Interior
        interior.Color = SpreadsheetGear.Colors.Gold
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

        worksheet.Cells(0, 2).Value = "Customer Open Order Report with Pictures"
        worksheet.Cells(0, 2).Font.Bold = True

        worksheet.Cells(0, H1).Value = "Note"
        worksheet.Cells(1, H1).Value = "For"

        worksheet.Cells(0, H1, 2, H1).Interior.Color = SpreadsheetGear.Colors.LightGray

        worksheet.Cells(0, H1 + 1).NumberFormat = "MM/dd/yy"
        worksheet.Cells(0, H1 + 1).Value = "Notes"
        worksheet.Cells(1, H1 + 1).Value = "CUST CODE"

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
        'PO Column
        COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 20
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        'Ship Date Column
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .EntireColumn.NumberFormat = "MM/dd/yy"
        End With

        'Cancel Date Column
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .EntireColumn.NumberFormat = "MM/dd/yy"
        End With

        'Reservation Column
        If chkReservations.Checked Then
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 12
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
        End If

        'Open Column
        If chkOpen.Checked Then
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 12
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
        End If

        'Pick Column
        If chkPick.Checked Then
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 12
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

        End If

    End Sub

    Private Sub btnLoadOrders_Click_1(sender As Object, e As EventArgs) Handles btnLoadOrders.Click
        If txtCUST_CODE.Text.Length = 0 Then
            MsgBox("You Must Select A Customer To Load", vbOKOnly, "Select Customer")
        Else
            Fill_Records("SOTRSRV1", txtCUST_CODE.Text)
            Fill_Records("SOTORDR0", txtCUST_CODE.Text)
        End If

    End Sub
#End Region
End Class