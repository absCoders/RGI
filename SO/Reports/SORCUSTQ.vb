Imports System.Text
Imports SpreadsheetGear

Public Class SORCUSTQ
    Dim S As New StringBuilder With {.Length = 0}
    Dim SQL_WHERE_ORDR_RSV As String
    Dim REPORT_NAME As String = "SORCUSTQ"
    Dim SQL_REPORT As New StringBuilder With {.Length = 0}
    Dim ICTSTAT2SQL As String = ""
    Dim ICTSTATDSQL As String = ""
    Dim XLS_NO As Integer = 0
    Dim exlExt As String = ".xlsx"
    Dim CONDITION As String = ""
    Dim REPORT_DATE0 As Date
    Dim REPORT_DATE1 As Date
    '

    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

#Region "Report Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Absx1.dteFor("DTE0").Value = DateAdd("d", -365, Now.Date)
        Absx1.dteFor("DTE1").Value = Now.Date


        '    working_date = DateAdd("d", -2, working_date)
        '  Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)


        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        'FtpS.RuntimeLicense = "31484E46414431535542323032333033313352415331544531414D483134323600000000000000003335384A30543346000059554A4336594E46335047530000"

        chkNewLinks.Checked = True

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

        With UltraExplorerBar1.Groups("Special Functions")
            .Visible = False
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

        REPORT_DATE0 = Absx1.dteFor("DTE0").Value
        REPORT_DATE1 = Absx1.dteFor("DTE1").Value

        CONDITION = " between '" & Format(REPORT_DATE0, "dd-MMM-yyyy") & "' and '" & Format(REPORT_DATE1, "dd-MMM-yyyy") & "'"


        ''S.Length = 0
        ''S.AppendLine("SELECT 'RESERVATION' AS ORDR_CUST_PO, '0000000000' AS ORDR_GROUP_NO FROM DUAL")
        ''ASCMAIN1.sql = S.ToString
        ''Dim TABLE_TEMP As String = ASCMAIN1.Temp_Table

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("JN.CUST_CODE, JN.CUST_NAME,")
        S.AppendLine("JN.ORDR_CUST_PO, JN.ORDR_DATE, JN.ORDR_SHIP_DATE, JN.ORDR_CANCEL_DATE,")
        S.AppendLine("JN.STYLE_CODE, JN.COLOR_CODE, JN.COLOR_DESC, JN.STYLE_DESC, JN.FABRIC_CODE, JN.SEASON_CODE,")
        S.AppendLine("JN.SUB_BODY_CODE, JN.SALES_DIVISION_CODE, JN.INNER_PACK_QTY, JN.CARTON_PACK_QTY, JN.STYLE_CUST_CODE, JN.IMAGE_NAME,")
        S.AppendLine("SUM(JN.RSRV_QTY) RSRV_QTY,")
        S.AppendLine("SUM(JN.RSRV_QTY_OPEN) RSRV_QTY_OPEN,")
        S.AppendLine("MIN(JN.RSRV_MIN_PRICE) AS RSRV_MIN_PRICE,")
        S.AppendLine("MAX(JN.RSRV_MAX_PRICE) AS RSRV_MAX_PRICE,")
        S.AppendLine("SUM(JN.RSRV_VAL_OPEN) AS RSRV_VAL_OPEN,")
        S.AppendLine("SUM(JN.ORDR_QTY) ORDR_QTY,")
        S.AppendLine("SUM(JN.ORDR_QTY_OPEN) ORDR_QTY_OPEN,")
        S.AppendLine("SUM(JN.ORDR_QTY_PICK) ORDR_QTY_PICK,")
        S.AppendLine("SUM(JN.ORDR_QTY_CANC) ORDR_QTY_CANC,")
        S.AppendLine("SUM(JN.ORDR_QTY_SHIP) ORDR_QTY_SHIP,")
        S.AppendLine("MIN(JN.MIN_PRICE) AS MIN_PRICE,")
        S.AppendLine("MAX(JN.MAX_PRICE) AS MAX_PRICE,")
        S.AppendLine("SUM(JN.VAL_PICK) AS VAL_PICK,")
        S.AppendLine("SUM(JN.VAL_OPEN) AS VAL_OPEN")
        S.AppendLine("FROM (")
        S.AppendLine("  SELECT")
        S.AppendLine("  O1.ORDR_GROUP_NO,")
        S.AppendLine("  '0000000000' AS RSRV_NO,")
        S.AppendLine("  A1.CUST_CODE, A1.CUST_NAME,")
        S.AppendLine("  O1.ORDR_CUST_PO, O1.ORDR_DATE, O1.ORDR_SHIP_DATE, O1.ORDR_CANCEL_DATE,")
        S.AppendLine("  I1.STYLE_CODE, O2.COLOR_CODE, C1.COLOR_DESC, I1.STYLE_DESC, I1.FABRIC_CODE, I1.SEASON_CODE, I1.SUB_BODY_CODE,")
        S.AppendLine("  I1.SALES_DIVISION_CODE, I1.INNER_PACK_QTY, I1.CARTON_PACK_QTY, I1.CUST_CODE STYLE_CUST_CODE, I1.IMAGE_NAME,")
        S.AppendLine("  0 RSRV_QTY,")
        S.AppendLine("  0 RSRV_QTY_OPEN,")
        S.AppendLine("  0 RSRV_MIN_PRICE,")
        S.AppendLine("  0 RSRV_MAX_PRICE,")
        S.AppendLine("  0 RSRV_VAL_OPEN,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY,0)) ORDR_QTY,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_PICK,0)) ORDR_QTY_PICK,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_CANC,0)) ORDR_QTY_CANC,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP,")
        S.AppendLine("  MIN(NVL(O2.ORDR_UNIT_PRICE,0)) AS MIN_PRICE,")
        S.AppendLine("  MAX(NVL(O2.ORDR_UNIT_PRICE,0)) AS MAX_PRICE,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_PICK,0) * NVL(O2.ORDR_UNIT_PRICE,0)) AS VAL_PICK,")
        S.AppendLine("  SUM(NVL(O2.ORDR_QTY_OPEN,0) * NVL(O2.ORDR_UNIT_PRICE,0)) AS VAL_OPEN")
        S.AppendLine("  FROM ARTCUST1 A1, SOTORDR1 O1, SOTORDR2 O2, ICTSTYL1 I1, ICTCOLR1 C1")
        S.AppendLine("  WHERE A1.CUST_CODE = O1.CUST_CODE")
        S.AppendLine(String.Format("  AND NVL(O1.CUST_CODE,'NULL') = '{0}'", txtCUST_CODE.Text))
        S.AppendLine("  AND O1.ORDR_NO = O2.ORDR_NO")
        S.AppendLine("  AND O2.STYLE_CODE = I1.STYLE_CODE")
        S.AppendLine("  AND O2.COLOR_CODE = C1.COLOR_CODE")
        S.AppendLine("  AND O1.ORDR_STATUS IN ('P','O')")
        S.AppendLine("  AND O1.ORDR_DATE " & CONDITION)
        S.AppendLine("  GROUP BY")
        S.AppendLine("  O1.ORDR_GROUP_NO,")
        S.AppendLine("  A1.CUST_CODE, A1.CUST_NAME,")
        S.AppendLine("  O1.ORDR_CUST_PO, O1.ORDR_DATE, O1.ORDR_SHIP_DATE, ORDR_CANCEL_DATE,")
        S.AppendLine("  I1.STYLE_CODE, O2.COLOR_CODE, C1.COLOR_DESC, I1.STYLE_DESC, I1.FABRIC_CODE, I1.SEASON_CODE, I1.SUB_BODY_CODE,")
        S.AppendLine("  I1.SALES_DIVISION_CODE, I1.INNER_PACK_QTY, I1.CARTON_PACK_QTY, I1.CUST_CODE, I1.IMAGE_NAME")
        S.AppendLine("  UNION")
        S.AppendLine("  SELECT")
        S.AppendLine("  '0000000000' AS ORDR_GROUP_NO,")
        S.AppendLine("  R1.RSRV_NO,")
        S.AppendLine("  A1.CUST_CODE, A1.CUST_NAME,")
        S.AppendLine("  'RESERVATION' ORDR_CUST_PO, R1.INIT_DATE ORDR_DATE, R1.ORDR_SHIP_DATE, R1.ORDR_CANCEL_DATE,")
        S.AppendLine("  I1.STYLE_CODE, R2.COLOR_CODE, C1.COLOR_DESC, I1.STYLE_DESC, I1.FABRIC_CODE, I1.SEASON_CODE,")
        S.AppendLine("  I1.SUB_BODY_CODE, I1.SALES_DIVISION_CODE, I1.INNER_PACK_QTY, I1.CARTON_PACK_QTY, I1.CUST_CODE STYLE_CUST_CODE, I1.IMAGE_NAME,")
        S.AppendLine("  SUM(NVL(R2.RSRV_QTY,0)) RSRV_QTY,")
        S.AppendLine("  SUM(NVL(R2.RSRV_QTY_OPEN,0)) RSRV_QTY_OPEN,")
        S.AppendLine("  MIN(NVL(R2.ORDR_UNIT_PRICE,0)) RSRV_MIN_PRICE,")
        S.AppendLine("  MAX(NVL(R2.ORDR_UNIT_PRICE,0)) RSRV_MAX_PRICE,")
        S.AppendLine("  SUM(NVL(R2.RSRV_QTY_OPEN,0) * NVL(R2.ORDR_UNIT_PRICE,0)) AS RSRV_VAL_OPEN,")
        S.AppendLine("  0 ORDR_QTY,")
        S.AppendLine("  0 ORDR_QTY_OPEN,")
        S.AppendLine("  0 ORDR_QTY_PICK,")
        S.AppendLine("  0 ORDR_QTY_CANC,")
        S.AppendLine("  0 ORDR_QTY_SHIP,")
        S.AppendLine("  0 MIN_PRICE,")
        S.AppendLine("  0 MAX_PRICE,")
        S.AppendLine("  0 VAL_PICK,")
        S.AppendLine("  0 VAL_OPEN")
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
        S.AppendLine("  R1.INIT_DATE, R1.ORDR_SHIP_DATE, R1.ORDR_CANCEL_DATE,")
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
        S.AppendLine("JN.ORDR_CUST_PO, JN.ORDR_DATE,JN.ORDR_SHIP_DATE, JN.ORDR_CANCEL_DATE,")
        S.AppendLine("JN.STYLE_CODE, JN.COLOR_CODE, JN.COLOR_DESC, JN.STYLE_DESC, JN.FABRIC_CODE, JN.SEASON_CODE,")
        S.AppendLine("JN.SUB_BODY_CODE, JN.SALES_DIVISION_CODE, JN.INNER_PACK_QTY, JN.CARTON_PACK_QTY, JN.STYLE_CUST_CODE, JN.IMAGE_NAME")
        S.AppendLine("ORDER BY")
        S.AppendLine(sql_GROUP_BY_cols)
        ASCMAIN1.sql = S.ToString()
        Create_TDA(dst.Tables.Add, "SOTCUSTQ", "**", 0, False)

        Dim TABLE_TEMP As String = ASCMAIN1.Temp_Table

        sql = "Delete From " & TABLE_TEMP & " WHERE RSRV_QTY_OPEN = 0 AND ORDR_QTY_OPEN = 0 AND ORDR_QTY_PICK = 0"
        ASCDATA1.ExecuteSQL(sql)


        ASCMAIN1.sql = "select *  from ictstat2 WHERE (STYLE_CODE,COLOR_CODE) IN (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & TABLE_TEMP & ")"
        ICTSTAT2SQL = ASCMAIN1.sql
        Create_TDA(dst.Tables.Add, "ICTSTAT2", "**", 2, False)

        '     Fill_Records("ICTSTAT2", "", True, ASCMAIN1.sql)

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
            & " And (POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE) IN  (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & TABLE_TEMP & ")" & vbCrLf _
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
            & " And (POTORDR2.STYLE_CODE,POTORDR2.COLOR_CODE) IN  (SELECT DISTINCT STYLE_CODE,COLOR_CODE FROM  " & TABLE_TEMP & ")" & vbCrLf _
            & ")"

        ICTSTATDSQL = ASCMAIN1.sql
        Create_TDA(dst.Tables.Add, "ICTSTATD", "**", 0, False)

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
        Dim FULLPATH As String = "\\192.168.180.34\g\VDI\ARCHIVE\VAN\Links\" & FILE_NAME
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

        Dim FileNameLocalFull As String = FILENAME_FULL
        Dim FileNameRemote As String = FILE_NAME
        Dim eMsg As Text.StringBuilder = FTP_BLUEHOST(FileNameLocalFull, FileNameRemote)
        If eMsg.Length > 0 Then
            MsgBox(eMsg.ToString, vbCritical, "Error Sending To Remote Server")
        End If
    End Sub
    Private Function Create_Excel_Buyer() As String
        Dim RetVal As String = ""
        Me.Cursor = Cursors.WaitCursor

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        worksheet.Name = "Buyer Spreadsheet"
        ASCMAIN1.Progress("Now Creating Buyer Workbook", "")
        'Make Headers
        'Set Column widths
        Dim CWC() As String = Split("A, B, C, D, E, F,G,H,I,J,K, L,M, N, O", ",")
        Dim CWS() As String = Split("6,20,40,10,40,10,8,5,5,5,5,15,5,10,25", ",")
        Dim CWA() As String = Split("C, L, L, L, L, C,C,C,C,C,L, L,C, C, C", ",")
        For CWCi As Integer = 0 To CWC.Length - 1
            worksheet.Cells(Trim(CWC(CWCi)) & "1").EntireColumn.ColumnWidth = Val(CWS(CWCi))
            If Trim(CWA(CWCi)) = "C" Then
                worksheet.Cells(Trim(CWC(CWCi)) & "1").HorizontalAlignment = SpreadsheetGear.HAlign.Center
            Else
                worksheet.Cells(Trim(CWC(CWCi)) & "1").HorizontalAlignment = SpreadsheetGear.HAlign.Left
            End If
        Next

        With worksheet
            'Paint cell colors, borders and col width
            With .Cells("A1:O2")
                .Interior.Color = SpreadsheetGear.Colors.LightBlue
                .Font.Bold = True
            End With
            With .Cells("A1:C1")
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            End With
            With .Cells("A2:O2")
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            End With
            With .Cells("D1:O1")
                .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            End With

            'Fill In the Captions.
            .Cells("A1").Value = "Categ"
            .Cells("B1").Value = "Style"
            .Cells("C1").Value = "Description"
            .Cells("F1").Value = "Start Ship"
            .Cells("G1").Value = "Consol" & vbCrLf & "cxl"
            .Cells("H1").Value = "LP"
            .Cells("I1").Value = "YJ"
            .Cells("J1").Value = "Marsh"
            .Cells("L1").Value = "Units"
            .Cells("M1").Value = "Cost"
            .Cells("N1").Value = "Comp"
            .Cells("O1").Value = "Comments"
            .Cells("C2").Value = "Group A"
            .Cells("O2").Value = Now().ToShortDateString
            .Cells("O2").HorizontalAlignment = SpreadsheetGear.HAlign.Right

        End With

        'Fill Rows
        'worksheet.Cells("D1").EntireColumn.NumberFormat = SpreadsheetGear.NumberFormatType.Text

        Dim fltrrowSB As String = " RSRV_QTY > 0"

        Dim curRow As Int64 = 3
            For Each rowSB As DataRow In dst.Tables.Item("SOTCUSTQ").Select(fltrrowSB, "STYLE_CODE, COLOR_CODE")
                Dim STYLE_CODE As String = rowSB.Item("STYLE_CODE").ToString & String.Empty
                ''    Dim fltrICTQUOT2 As String = String.Format("STYLE_CODE_PLM = '{0}'", STYLE_CODE)
                ''   Dim rowICTQUOT2 As DataRow = dst.Tables.Item("ICTQUOT2").Select(fltrICTQUOT2).FirstOrDefault
                worksheet.Cells("B" & curRow.ToString).Value = STYLE_CODE
                worksheet.Cells("C" & curRow.ToString).Value = rowSB.Item("STYLE_DESC").ToString & String.Empty
                worksheet.Cells("D" & curRow.ToString).NumberFormat = "@"
                worksheet.Cells("D" & curRow.ToString).Value = (rowSB.Item("COLOR_CODE").ToString & String.Empty).ToString
                worksheet.Cells("E" & curRow.ToString).Value = rowSB.Item("COLOR_DESC").ToString & String.Empty
                worksheet.Cells("A" & curRow.ToString & ":" & "O" & curRow.ToString).Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                Dim START_SHIP As Date = CDate("01/01/1900")
                Dim UNITS As Int64 = 0
                ''For i As Int64 = 0 To 3
                ''    If IsDate(rowSB.Item("DTE" & i).ToString & String.Empty) And Val(rowSB.Item("QTY_AVA" & i).ToString & String.Empty) > 0 Then
                ''        If CDate(rowSB.Item("DTE" & i).ToString & String.Empty) > START_SHIP Then
                ''            START_SHIP = CDate(rowSB.Item("DTE" & i).ToString & String.Empty)
                ''            UNITS += Val(rowSB.Item("QTY_AVA" & i).ToString & String.Empty)
                ''        End If
                ''    End If
                ''Next

                START_SHIP = CDate(rowSB.Item("ORDR_SHIP_DATE").ToString & String.Empty)
            UNITS += Val(rowSB.Item("RSRV_QTY").ToString & String.Empty)
            If START_SHIP <> CDate("01/01/1900") Then
                    'Start Ship (F)
                    worksheet.Cells("F" & curRow.ToString).Value = START_SHIP.ToShortDateString
                    'Units (L)
                    worksheet.Cells("L" & curRow.ToString).Value = UNITS
                End If
                curRow += 1
            Next

            'Show Workbook
            XLS_NO = 1
        Dim XLS_FILENAME As String = "Buyer Sheet" & Format(XLS_NO, "000") & exlExt
        ''workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        ''Show_Document(ASCMAIN1.Folders("Temp") & XLS_FILENAME)

        Dim success As Boolean = False

        Do Until success
            Try
                If System.IO.File.Exists(ASCMAIN1.Folders("Temp") & XLS_FILENAME) Then
                    System.IO.File.Delete(XLS_FILENAME)
                End If
                workbook.SaveAs(ASCMAIN1.Folders("Temp") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                success = True
            Catch ex As Exception

            End Try
        Loop
        Show_Document(ASCMAIN1.Folders("Temp") & XLS_FILENAME)


        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default

        Return RetVal
    End Function

    Private Function getFileExt(ByVal ATTACHMENT_FILENAME As String) As String
        Dim RetVal As String = ""
        Dim dotLoc As Int64 = ATTACHMENT_FILENAME.IndexOf(".", ATTACHMENT_FILENAME.Length - 5)
        If dotLoc > 0 Then
            RetVal = ATTACHMENT_FILENAME.Substring(dotLoc, ATTACHMENT_FILENAME.Length - dotLoc)
        End If
        Return RetVal
    End Function
    Private Function GetCustShipDates(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim RetVal As String = ""
        Dim SQLS As New StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("MIN(O1.ORDR_SHIP_DATE) AS MIN_ORDR_DATE")
        SQLS.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
        SQLS.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
        SQLS.AppendLine("AND O1.ORDR_STATUS IN ('P','O')")
        SQLS.AppendLine(String.Format("AND O2.STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND O2.COLOR_CODE = '{0}'", COLOR_CODE))
        SQLS.AppendLine(String.Format("AND O1.CUST_CODE = '{0}'", txtCUST_CODE.Text))
        'SQLS.AppendLine(String.Format(" AND S1.INV_DATE >= '{0}'", Format(dteShip_Beg.DateTime, "dd-MMM-yy")))
        'SQLS.AppendLine(String.Format(" AND S1.INV_DATE <= '{0}'", Format(dteShip_End.DateTime, "dd-MMM-yy")))
        '' new DGJ 2 LINES ABOVE Only consider Shipments in Date Range


        ASCMAIN1.sql = SQLS.ToString()
        Dim MIN_ORDR_DATE As String = ASCDATA1.GetDataValue

        SQLS.Length = 0
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("MAX(O1.ORDR_CANCEL_DATE) AS MAX_CANC_DATE")
        SQLS.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
        SQLS.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
        SQLS.AppendLine("AND O1.ORDR_STATUS IN ('P','O')")
        SQLS.AppendLine(String.Format("AND O2.STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND O2.COLOR_CODE = '{0}'", COLOR_CODE))
        SQLS.AppendLine(String.Format("AND O1.CUST_CODE = '{0}'", txtCUST_CODE.Text))


        ASCMAIN1.sql = SQLS.ToString()
        Dim MAX_CANC_DATE As String = ASCDATA1.GetDataValue

        If IsDate(MIN_ORDR_DATE) And IsDate(MAX_CANC_DATE) Then
            RetVal = String.Format("{0} - {1}", Format(CDate(MIN_ORDR_DATE), "MM/dd/yy"), Format(CDate(MAX_CANC_DATE), "MM/dd/yy"))
        End If

        Return RetVal
    End Function
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
            With UltraExplorerBar1.Groups("Special Functions")
                .Visible = False
            End With

        End If

        If eItemKey = "Buyer Sheet" Then
            Create_Excel_Buyer()
            'With UltraExplorerBar1.Groups("Special Functions")
            '    .Visible = False
            'End With
            Exit Sub

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
        Fill_Records("SOTCUSTQ")

        Fill_Records("ICTSTAT2", "", True, ICTSTAT2SQL)


        Fill_Records("ICTSTATD", "", True, ICTSTATDSQL)




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
                If chkNewLinks.Checked Then
                    Dim FILE_NAME As String = grd.ActiveRow.Cells.Item("FILE_NAME").Text
                    Dim HASHVALUE As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
                    'Dim LINEPFX As String = $"https://docs.vandalequotes.com/{FILENAME}"
                    Dim LINEPFX As String = $"https://vandaledocs.azurewebsites.net/Documents/{HASHVALUE}"
                    My.Computer.Clipboard.SetText(FILE_NAME & vbCrLf & LINEPFX)
                    MsgBox("You Link Has Been Copied To Your Clipboard", vbInformation, "Done")
                Else
                    Dim FILE_NAME As String = grd.ActiveRow.Cells.Item("FILE_NAME").Text
                    Dim HASH As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
                    Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
                    My.Computer.Clipboard.SetText(FILE_NAME & vbCrLf & LINEPFX)
                    MsgBox("You Link Has Been Copied To Your Clipboard", vbInformation, "Done")
                End If
            'Case "Copy Link"
            '    Dim FILE_NAME As String = grd.ActiveRow.Cells.Item("FILE_NAME").Text
            '    Dim HASH As String = grd.ActiveRow.Cells.Item("HASHVALUE").Text
            '    Dim LINEPFX As String = "http://showroom.vandale.com/api/showroom/" & HASH
            '    My.Computer.Clipboard.SetText(FILE_NAME & vbCrLf & LINEPFX)
            '    MsgBox("You Link Has Been Copied To Your Clipboard", vbInformation, "Done")
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


            If chkPrice.Checked Then
                For iCOL As Integer = 1 To 11
                    COL += 1
                    Select Case iCOL
                        Case 6
                            If chkReservations.Checked Then
                                worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                            Else
                                COL -= 2
                            End If
                        Case 8
                            If chkOpen.Checked Then
                                worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                            Else
                                COL -= 1
                            End If
                        Case 9
                            If chkOpen.Checked Then
                                worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                            Else
                                COL -= 1
                            End If
                        Case 10
                            If chkPick.Checked Then
                                worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                            Else
                                COL -= 1
                            End If

                        Case 11
                            If chkPick.Checked Then
                                worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                            Else
                                COL -= 1
                            End If

                    End Select

                    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                Next
            Else
                For iCOL As Integer = 1 To 7
                    COL += 1
                    Select Case iCOL
                        Case 5
                            If chkReservations.Checked Then
                                worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                            Else
                                COL -= 1
                            End If
                        Case 6
                            If chkOpen.Checked Then
                                worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                            Else
                                COL -= 1
                            End If
                        Case 7
                            If chkPick.Checked Then
                                worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                            Else
                                COL -= 1
                            End If
                    End Select

                    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                Next
            End If

            If chkShipDates.Checked Then
                COL += 1
            End If



            If chkStyleStats.Checked Then
                For iCOL As Integer = 1 To 7
                    COL += 1
                    worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                Next
                COL += 0
            End If


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


            '  Right here Order Details when Applicable
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

                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Rev Ship Dt"
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "ETA"
                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Vessel"
                        With worksheet.Cells(I - 1, COL - 1 + chkcnt)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                        End With

                        chkcnt += 1
                        worksheet.Cells(I - 1, COL - 1 + chkcnt).Value = "Shp Dt Rev"

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
                        .Value = Val(rowICTSTATD.Item("COLOR_CODE") & String.Empty)
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

        If chkStyleStats.Checked Then
            worksheet.Cells(I - 1, COL - 0).Value = "'" & ""
        Else
            worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"
        End If

        Dim GT = ""

        If chkStyleStats.Checked Then
        Else
            If chkPrice.Checked Then
                For iCOL As Integer = 1 To 11
                    COL += 1
                    Select Case iCOL
                        Case 6
                            If chkReservations.Checked Then
                                '  worksheet.Cells(I + ci - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + ci - 1 - 1, COL).Address, "$", "") & ")"
                                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                            Else
                                COL -= 2
                            End If
                        Case 8
                            If chkOpen.Checked Then
                                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                            Else
                                COL -= 1
                            End If
                        Case 9
                            If chkOpen.Checked Then
                                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                            Else
                                COL -= 1
                            End If
                        Case 10
                            If chkPick.Checked Then
                                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                            Else
                                COL -= 1
                            End If

                        Case 11
                            If chkPick.Checked Then
                                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                            Else
                                COL -= 1
                            End If

                    End Select

                    '    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + ci - 1, COL).Address, "$", "")
                Next
            Else
                For iCOL As Integer = 1 To 7
                    COL += 1
                    Select Case iCOL
                        Case 5
                            If chkReservations.Checked Then
                                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                            Else
                                COL -= 1
                            End If
                        Case 6
                            If chkOpen.Checked Then
                                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                            Else
                                COL -= 1
                            End If
                        Case 7
                            If chkPick.Checked Then
                                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)
                                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                            Else
                                COL -= 1
                            End If
                    End Select
                Next
            End If

            worksheet.Cells(I - 1, COL0 - 1, I - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

        End If
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
            Dim chkcnt As Int64 = 5
            If LAST_COLOR <> rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty Then
                worksheet.Cells(i + CI - 1, COL - 1).Value = "'" & rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty
                worksheet.Cells(i + CI - 1, COL).Value = rowSOTCUSTQ.Item("COLOR_DESC") & String.Empty
                LAST_COLOR = rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty
            End If
            worksheet.Cells(i + CI - 1, COL + 1).Value = rowSOTCUSTQ.Item("ORDR_CUST_PO") & String.Empty
            worksheet.Cells(i + CI - 1, COL + 2).Value = rowSOTCUSTQ.Item("ORDR_DATE") & String.Empty
            worksheet.Cells(i + CI - 1, COL + 3).Value = rowSOTCUSTQ.Item("ORDR_SHIP_DATE") & String.Empty
            worksheet.Cells(i + CI - 1, COL + 4).Value = rowSOTCUSTQ.Item("ORDR_CANCEL_DATE") & String.Empty
            If chkPrice.Checked Then
                Dim PRICE_DISPLAY As String = ""
                If rowSOTCUSTQ.Item("MIN_PRICE") = rowSOTCUSTQ.Item("MAX_PRICE") Then
                    PRICE_DISPLAY = rowSOTCUSTQ.Item("MIN_PRICE") & String.Empty
                Else
                    PRICE_DISPLAY = "Max " & rowSOTCUSTQ.Item("MIN_PRICE") & String.Empty & "Max " & rowSOTCUSTQ.Item("MAX_PRICE") & String.Empty
                End If
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = PRICE_DISPLAY
                chkcnt += 1
                '    If chkReservations.Checked Then
                '        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("RSRV_MAX_PRICE") & String.Empty
                '        chkcnt += 1
                '    End If
            End If



            If chkReservations.Checked Then
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("RSRV_QTY_OPEN") & String.Empty
                chkcnt += 1
                If chkPrice.Checked Then
                    Dim PRICE_DISPLAY As String = ""
                    If chkReservations.Checked Then
                        worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("RSRV_MAX_PRICE") & String.Empty
                        chkcnt += 1
                    End If
                End If

            End If
            If chkOpen.Checked Then
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("ORDR_QTY_OPEN") & String.Empty
                chkcnt += 1
                If chkPrice.Checked Then
                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("VAL_OPEN") & String.Empty
                    chkcnt += 1
                End If

            End If
            If chkPick.Checked Then
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("ORDR_QTY_PICK") & String.Empty
                chkcnt += 1
                If chkPrice.Checked Then
                    worksheet.Cells(i + CI - 1, COL + chkcnt).Value = rowSOTCUSTQ.Item("VAL_PICK") & String.Empty
                    chkcnt += 1
                End If

            End If

            If chkShipDates.Checked Then
                'worksheet.Cells(i + CI - 1, COL + chkcnt).Value = "1/1/2018 - 12/31/2018"
                worksheet.Cells(i + CI - 1, COL + chkcnt).Value = GetCustShipDates(rowSOTCUSTQ.Item("STYLE_CODE") & String.Empty, rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty)
                chkcnt += 1
            End If



            If chkStyleStats.Checked Then

                ' ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty & "'"
                '       For Each rowICTSTAT2 As DataRow In ASCDATA1.GetDataTable.Select("")
                For Each rowICTSTAT2 As DataRow In dst.Tables("ICTSTAT2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & rowSOTCUSTQ.Item("COLOR_CODE") & String.Empty & "'")


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
            .Value = "Date"
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

        If chkPrice.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "Price"
            End With

        End If
        If chkReservations.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "Res Qty"
            End With

            If chkPrice.Checked Then
                COL += 1
                With worksheet.Cells(i, COL)
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "Res Price"
                End With
            End If

        End If

        If chkOpen.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "Open Qty"
            End With
            If chkPrice.Checked Then
                COL += 1
                With worksheet.Cells(i, COL)
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "Open Amt"
                End With
            End If

        End If

        If chkPick.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "Pick Qty"
            End With

            If chkPrice.Checked Then
                COL += 1
                With worksheet.Cells(i, COL)
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Value = "Pick Amt"
                End With
            End If

        End If

        If chkShipDates.Checked Then
            COL += 1
            With worksheet.Cells(i, COL)
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .Value = "" & Chr(13) & Chr(10) & "1st Start & Last Cancel"
                .Font.Size = 14
            End With
        End If

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
        worksheet.Cells("A1:AZ1").EntireColumn.Font.Size = 16

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

        'Order Date Column
        COL += 1
        _COL += 1
        With worksheet.Cells(_COL, COL)
            .ColumnWidth = 15
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .EntireColumn.NumberFormat = "MM/dd/yy"
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
        If chkPrice.Checked Then
            ' Order Price
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 12
                .EntireColumn.NumberFormat = "###.00"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

        End If
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
            If chkPrice.Checked Then
                ' Res Price
                COL += 1
                _COL += 1
                With worksheet.Cells(_COL, COL)
                    .ColumnWidth = 15
                    .EntireColumn.NumberFormat = "###.00"
                    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With

            End If


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
            If chkPrice.Checked Then
                ' Open Amt
                COL += 1
                _COL += 1
                With worksheet.Cells(_COL, COL)
                    .ColumnWidth = 17
                    .EntireColumn.NumberFormat = "###,###.00"
                    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With

            End If

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
            If chkPrice.Checked Then
                ' Pick Amt
                COL += 1
                _COL += 1
                With worksheet.Cells(_COL, COL)
                    .ColumnWidth = 17
                    .EntireColumn.NumberFormat = "###,###.00"
                    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With

            End If
        End If

        If chkShipDates.Checked Then
            COL += 1
            _COL += 1
            With worksheet.Cells(_COL, COL)
                .ColumnWidth = 30
                '          .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
        End If

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