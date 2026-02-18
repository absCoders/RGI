Imports SpreadsheetGear

Public Class ARCRGIDD
    'Systemic
    Private _FF As ASFBASE1
    Private S As New Text.StringBuilder With {.Length = 0}
    'Private YYYY As String = "2025"
    Private YYYYBP As String = ""
    Private YYYYEP As String = ""

    'Excel
    Private oWB As SpreadsheetGear.IWorkbook
    Private oSheet As SpreadsheetGear.IWorksheet = Nothing
    Private range As SpreadsheetGear.IRange = Nothing
    Private xls_path As String = ASCMAIN1.Folders("Work")
    Private xls_name As String = ""
    Private xls_file_name As String = ""
    Private XLS_NO As Integer = 0

    Private TRB As Integer = 0 'Starting Row
    Private TRC As Integer = 0 'Current Row
    Private TC As Integer = 0 'Starting / Current Column

    'Tables and Data
    Private tblEXCEL As DataTable = Nothing
    Private tmpSOTINVH1 As String = ""
    Private tblSOTINVH1 As DataTable = Nothing
    Private tmpSOTINVHH As String = ""
    Private tblSOTINVHH As DataTable = Nothing
    Private tmpSOTORDR1 As String = ""
    Private tblSOTORDR1 As DataTable = Nothing
    Private tmpSOTORDRD As String = ""
    Private tblSOTORDRD As DataTable = Nothing
    Private tmpARTCUST1 As String = ""
    Private tblARTCUST1 As DataTable = Nothing
    Private OnlyOrderDetails As Boolean = False
    Private OnlyInvoiceHeaders As Boolean = False

    'Misc
    Private success As Boolean = False
    Public eMsg As String = ""

#Region "Instantiate Class"
    Public Sub New(ByVal F As ASFBASE1)
        _FF = F
        Dim frmASFMSGBF As New ASFMSGBF
        YYYYBP = GetPeriod("Please Select Beginning Period")
        YYYYEP = GetPeriod("Please Select Ending Period")
        Dim OnlyOD As MsgBoxResult = MsgBox("Do You Only Want Order Details Run?", vbYesNo, "Order Details Run")
        If OnlyOD = MsgBoxResult.Yes Then
            OnlyOrderDetails = True
        Else
            Dim OnlyIH As MsgBoxResult = MsgBox("Do You Only Want Invoice Header Run?", vbYesNo, "Invoice Header Run")
            If OnlyIH = MsgBoxResult.Yes Then
                OnlyInvoiceHeaders = True
            End If
        End If
        If YYYYBP.Length = 0 Or YYYYEP.Length = 0 Then
            eMsg = $"Invalid Periods {YYYYBP} | {YYYYEP}"
        End If
        If eMsg.Length = 0 Then
            ASCMAIN1.Progress("Now Loading Excel Data", "")
            _FF.Cursor = Cursors.WaitCursor
            InitializeVariables()
            InitializeTempTables()
            InitializeDataSets()
            InitializeFormatting()
        End If
    End Sub

    Private Function GetPeriod(ByVal Caption As String) As String
        Dim RetVal As String = ""
        Dim S As New Text.StringBuilder With {.Length = 0}
        S.AppendLine("SELECT OPS_YYYYPP, LEGEND")
        S.AppendLine("FROM GLTPARM2")
        S.AppendLine("WHERE SUBSTR(OPS_YYYYPP,0,4) IN ('2023','2024','2025', '2026')")
        S.AppendLine("ORDER BY OPS_YYYYPP")
        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = False
            .PreviouslySelectedCodes0 = ""
            .Caption = Caption
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable
        End With
        Dim F As New ASFCODE1
        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            RetVal = ASCMAIN1.CodeSelector.SelectedRows(0).Item("OPS_YYYYPP") & ""
        End If
        Return RetVal
    End Function

    Private Sub InitializeVariables()
        S.Length = 0
        TRB = 1
        TRC = 1
        TC = 0
    End Sub
    Private Sub InitializeTempTables()
        Dim rowGLTPARM2 As DataRow = _FF.LookUp("GLTPARM2", YYYYBP)
        Dim BEG_DATE As DateTime = rowGLTPARM2.Item("PRD_END_DATE")
        BEG_DATE = BEG_DATE.AddDays(1).AddMonths(-1)
        Dim BEG_DATE_STR As String = Format(BEG_DATE, "dd-MMM-yyyy")

        rowGLTPARM2 = _FF.LookUp("GLTPARM2", YYYYEP)
        Dim END_DATE As DateTime = rowGLTPARM2.Item("PRD_END_DATE")
        Dim END_DATE_STR As String = Format(END_DATE, "dd-MMM-yyyy")

        If OnlyOrderDetails Then
            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("O1.ORDR_NO,")
            S.AppendLine("O1.ORDR_DATE,")
            S.AppendLine("O1.ORDR_DATE_RECD,")
            S.AppendLine("O1.CUST_CODE,")
            S.AppendLine("O1.CUST_NAME,")
            S.AppendLine("(O2.STYLE_CODE || '-' || O2.COLOR_CODE) AS SKU,")
            S.AppendLine("O2.STYLE_DESC,")
            S.AppendLine("S1.STYLE_CLASS_CODE,")
            S.AppendLine($"'{("").PadRight(30)}' AS SUB_STYLE_CLASS_CODE,")
            S.AppendLine("SUM(NVL(O2.ORDR_QTY,0) * NVL(O2.ORDR_UNIT_PRICE,0)) ORDR_AMT,")
            S.AppendLine("SUM(NVL(O2.ORDR_QTY_OPEN,0) * NVL(O2.ORDR_UNIT_PRICE,0)) OPEN_AMT,")
            S.AppendLine("SUM(NVL(O2.ORDR_QTY_CANC,0) * NVL(O2.ORDR_UNIT_PRICE,0)) CANC_AMT,")
            S.AppendLine("SUM(NVL(O2.ORDR_QTY_SHIP,0) * NVL(O2.ORDR_UNIT_PRICE,0)) SHIP_AMT,")
            S.AppendLine("SUM(NVL(O2.ORDR_QTY,0)) ORDR_QTY,")
            S.AppendLine("O1.ORDR_SHIP_DATE,")
            S.AppendLine("O1.ORDR_CANCEL_DATE,")
            S.AppendLine("O1.WHSE_CODE")
            S.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2, ICTSTYL1 S1")
            S.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
            S.AppendLine("AND O2.STYLE_CODE = S1.STYLE_CODE ")
            S.AppendLine($"AND O1.ORDR_DATE >= '{BEG_DATE_STR}'")
            S.AppendLine($"AND O1.ORDR_DATE <= '{END_DATE_STR}'")
            S.AppendLine("GROUP BY")
            S.AppendLine("O1.ORDR_NO,")
            S.AppendLine("O1.ORDR_DATE,")
            S.AppendLine("O1.ORDR_DATE_RECD,")
            S.AppendLine("O1.CUST_CODE,")
            S.AppendLine("O1.CUST_NAME,")
            S.AppendLine("(O2.STYLE_CODE || '-' || O2.COLOR_CODE),")
            S.AppendLine("O2.STYLE_DESC,")
            S.AppendLine("S1.STYLE_CLASS_CODE,")
            S.AppendLine($"'{("").PadRight(30)}',")
            S.AppendLine("O1.ORDR_SHIP_DATE,")
            S.AppendLine("O1.ORDR_CANCEL_DATE,")
            S.AppendLine("O1.WHSE_CODE")
            S.AppendLine("ORDER BY O1.ORDR_DATE")
            ASCMAIN1.sql = S.ToString
            tmpSOTORDRD = ASCMAIN1.Temp_Table
        Else
            If OnlyInvoiceHeaders Then
                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("I1.ORDR_YYYYPP_UPDATED,")
                S.AppendLine("CASE I1.INV_TYPE")
                S.AppendLine("    WHEN 'I' THEN 'INVOICE'")
                S.AppendLine("    WHEN 'C' THEN 'CREDIT'")
                S.AppendLine("    ELSE 'OTHER'")
                S.AppendLine("END AS INV_TYPE_D,")
                S.AppendLine("I1.INV_DATE,")
                S.AppendLine("I1.INIT_DATE,")
                S.AppendLine("I1.INV_NO,")
                S.AppendLine("I1.SREP_CODE,")
                S.AppendLine("C1.CUST_NAME,")
                S.AppendLine("I1.CUST_CODE,")
                S.AppendLine("'TRADE' AS CUST_TYPE,")
                S.AppendLine("I1.WHSE_CODE,")
                S.AppendLine("C1.CUST_STATE,")
                S.AppendLine("C1.CUST_COUNTRY,")
                S.AppendLine("C1.CUST_ZIP_CODE,")
                S.AppendLine("I1.INV_SALES_CURR,")
                S.AppendLine("I1.INV_FREIGHT_CURR,")
                S.AppendLine("I1.INV_MISC_CHG_CURR,")
                S.AppendLine("I1.INV_TOTAL_AMT_CURR,")
                S.AppendLine("SUM(IM.INV_MISC_CHG_CURR) AS INV_MISC_CHG_CURR_T")
                S.AppendLine("FROM SOTINVH1 I1, ARTCUST1 C1, SOTINVHM IM")
                S.AppendLine("WHERE I1.CUST_CODE = C1.CUST_CODE")
                S.AppendLine("AND I1.INV_NO = IM.INV_NO (+)")
                S.AppendLine("AND IM.MISC_CHG_CODE (+) = 'TARIFF'")
                S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED >= {YYYYBP}")
                S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED <= {YYYYEP}")
                S.AppendLine("GROUP BY")
                S.AppendLine("I1.ORDR_YYYYPP_UPDATED, CASE I1.INV_TYPE WHEN 'I' THEN 'INVOICE' WHEN 'C' THEN 'CREDIT' ELSE 'OTHER' END, I1.INV_TYPE, 'I', 'INVOICE',")
                S.AppendLine("'C', 'CREDIT', 'OTHER', I1.INV_DATE, I1.INIT_DATE,")
                S.AppendLine("I1.INV_NO, I1.SREP_CODE, C1.CUST_NAME, I1.CUST_CODE, I1.WHSE_CODE,")
                S.AppendLine("'TRADE', C1.CUST_STATE, C1.CUST_COUNTRY, C1.CUST_ZIP_CODE, I1.INV_SALES_CURR,")
                S.AppendLine("I1.INV_FREIGHT_CURR, I1.INV_MISC_CHG_CURR, I1.INV_TOTAL_AMT_CURR")
                ASCMAIN1.sql = S.ToString
                tmpSOTINVHH = ASCMAIN1.Temp_Table
            Else
                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("I1.ORDR_YYYYPP_UPDATED,")
                S.AppendLine("CASE I1.INV_TYPE")
                S.AppendLine("    WHEN 'I' THEN 'INVOICE'")
                S.AppendLine("    WHEN 'C' THEN 'CREDIT'")
                S.AppendLine("    ELSE 'OTHER'")
                S.AppendLine("END AS INV_TYPE_D,")
                S.AppendLine("I1.INV_DATE,")
                S.AppendLine("I1.INIT_DATE,")
                S.AppendLine("I1.INV_NO,")
                S.AppendLine("I2.INV_LNO,")
                S.AppendLine("O1.ORDR_DATE,")
                S.AppendLine("I1.ORDR_NO,")
                S.AppendLine("I1.SREP_CODE,")
                S.AppendLine("C1.CUST_NAME,")
                S.AppendLine("I1.CUST_CODE,")
                S.AppendLine("'TRADE' AS CUST_TYPE,")
                S.AppendLine("I1.WHSE_CODE,")
                S.AppendLine("(C1.CUST_STATE || ', ' || C1.CUST_COUNTRY) AS STATE_COUNTRY,")
                S.AppendLine("C1.CUST_ZIP_CODE,")
                S.AppendLine("SYSDATE AS FIRST_SALE,")
                S.AppendLine("(I2.STYLE_CODE || '-' || I2.COLOR_CODE) AS SKU,")
                S.AppendLine("S1.STYLE_DESC,")
                S.AppendLine("S1.STYLE_CLASS_CODE,")
                S.AppendLine($"'{("").PadRight(30)}' AS SUB_STYLE_CLASS_CODE,")
                S.AppendLine("9.99 AS LIST_PRICE,")
                S.AppendLine("I2.ORDR_QTY_SHIP,")
                S.AppendLine("I2.ORDR_UNIT_PRICE,")
                S.AppendLine("(I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE) AS ORDR_REVENUE,")
                S.AppendLine("I2.ORDR_UNIT_COST,")
                S.AppendLine("(I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_COST) AS ORDR_COGS,")
                S.AppendLine(".18 AS STD_LANDED_PCT,")
                S.AppendLine("99999.99 AS LANDED_COST,")
                S.AppendLine("99999.99 AS GP,")
                S.AppendLine("99999.99 AS GP_PCT,")
                S.AppendLine("I2.COMM_RATE,")
                S.AppendLine("V1.VEND_SUPPLIER_ID AS VEND_ID,") '
                S.AppendLine("V1.VEND_CODE,")
                S.AppendLine("V1.PORT_CODE,")
                S.AppendLine("PT.PORT_NAME,")
                S.AppendLine("V1.PORT_CODE AS LAST_PORT,")
                S.AppendLine("Y1.COUNTRY_NAME,")
                S.AppendLine("O1.ORDR_SHIP_DATE,")
                S.AppendLine("O1.ORDR_CANCEL_DATE")
                S.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ARTCUST1 C1, ICTSTYL1 S1, APTVEND1 V1, TATCNTRY Y1, SOTORDR1 O1, ICTPORT1 PT")
                S.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
                S.AppendLine("AND I1.INV_NO = I2.INV_NO")
                S.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
                S.AppendLine("AND I1.CUST_CODE = C1.CUST_CODE")
                S.AppendLine("AND I2.STYLE_CODE = S1.STYLE_CODE")
                S.AppendLine("AND S1.VEND_CODE = V1.VEND_CODE (+)")
                S.AppendLine("AND V1.VEND_COUNTRY = Y1.COUNTRY_CODE (+)")
                S.AppendLine("AND V1.PORT_CODE = PT.PORT_CODE (+)")
                S.AppendLine("AND I2.ORDR_QTY_SHIP <> 0")
                S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED >= {YYYYBP}")
                S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED <= {YYYYEP}")
                ASCMAIN1.sql = S.ToString
                tmpSOTINVH1 = ASCMAIN1.Temp_Table

                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("CUST_CODE,")
                S.AppendLine("CUST_NAME,")
                S.AppendLine("CUST_ADDR1,")
                S.AppendLine("CUST_CITY,")
                S.AppendLine("CUST_STATE,")
                S.AppendLine("CUST_ZIP_CODE,")
                S.AppendLine("CUST_COUNTRY,")
                S.AppendLine("SREP_CODE,")
                S.AppendLine("CUST_PRICE_TIER,")
                S.AppendLine("CUST_PRICE_TIER_PVC,")
                S.AppendLine("SYSDATE AS FIRST_ORDR,")
                S.AppendLine("10000000 AS ORDERS,")
                S.AppendLine("10000000 AS OPEN,")
                S.AppendLine("10000000 AS CANCELED,")
                S.AppendLine("10000000 AS SHIPPED")
                S.AppendLine("FROM ARTCUST1")
                S.AppendLine("WHERE CUST_CODE IN (")
                S.AppendLine("  SELECT")
                S.AppendLine("  DISTINCT O1.CUST_CODE")
                S.AppendLine("  FROM SOTORDR1 O1")
                S.AppendLine($"  WHERE O1.ORDR_DATE >= '{BEG_DATE_STR}'")
                S.AppendLine($"  AND O1.ORDR_DATE <= '{END_DATE_STR}'")
                S.AppendLine(")")
                S.AppendLine("ORDER BY CUST_NAME")
                'S.AppendLine($"WHERE CUST_CODE IN (SELECT DISTINCT CUST_CODE FROM {tmpSOTINVH1})")
                'S.AppendLine($"WHERE CUST_CODE IN (SELECT DISTINCT CUST_CODE FROM {tmpSOTORDR1})")
                ASCMAIN1.sql = S.ToString
                tmpARTCUST1 = ASCMAIN1.Temp_Table

                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("O1.ORDR_NO,")
                S.AppendLine("O1.ORDR_DATE,")
                S.AppendLine("O1.ORDR_DATE_RECD,")
                S.AppendLine("O1.CUST_CODE,")
                S.AppendLine("O1.CUST_NAME,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY,0) * NVL(O2.ORDR_UNIT_PRICE,0)) ORDR_AMT,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY_OPEN,0) * NVL(O2.ORDR_UNIT_PRICE,0)) OPEN_AMT,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY_CANC,0) * NVL(O2.ORDR_UNIT_PRICE,0)) CANC_AMT,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY_SHIP,0) * NVL(O2.ORDR_UNIT_PRICE,0)) SHIP_AMT,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY,0)) ORDR_QTY,")
                S.AppendLine("O1.ORDR_SHIP_DATE,")
                S.AppendLine("O1.ORDR_CANCEL_DATE")
                S.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
                S.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
                'S.AppendLine($"AND O1.ORDR_YYYYPP_BOOKED >= '{YYYYBP}'")
                'S.AppendLine($"AND O1.ORDR_YYYYPP_BOOKED <= '{YYYYEP}'")
                S.AppendLine($"AND O1.ORDR_DATE >= '{BEG_DATE_STR}'")
                S.AppendLine($"AND O1.ORDR_DATE <= '{END_DATE_STR}'")
                S.AppendLine("GROUP BY")
                S.AppendLine("O1.ORDR_NO,")
                S.AppendLine("O1.ORDR_DATE,")
                S.AppendLine("O1.ORDR_DATE_RECD,")
                S.AppendLine("O1.CUST_CODE,")
                S.AppendLine("O1.CUST_NAME,")
                S.AppendLine("O1.ORDR_SHIP_DATE,")
                S.AppendLine("O1.ORDR_CANCEL_DATE")
                S.AppendLine("ORDER BY O1.ORDR_DATE")
                ASCMAIN1.sql = S.ToString
                tmpSOTORDR1 = ASCMAIN1.Temp_Table
            End If
        End If
    End Sub
    Private Sub InitializeDataSets()
        If OnlyOrderDetails Then
            S.Length = 0
            S.AppendLine("Select *")
            S.AppendLine($" from {tmpSOTORDRD}")
            S.AppendLine($" ORDER BY ORDR_DATE")
            ASCMAIN1.sql = S.ToString
            tblSOTORDRD = ASCDATA1.GetDataTable
        Else
            If OnlyInvoiceHeaders Then
                S.Length = 0
                S.AppendLine("Select *")
                S.AppendLine($" from {tmpSOTINVHH}")
                ASCMAIN1.sql = S.ToString
                tblSOTINVHH = ASCDATA1.GetDataTable
            Else
                S.Length = 0
                S.AppendLine("Select *")
                S.AppendLine($" from {tmpSOTINVH1}")
                ASCMAIN1.sql = S.ToString
                tblSOTINVH1 = ASCDATA1.GetDataTable

                S.Length = 0
                S.AppendLine("Select *")
                S.AppendLine($" from {tmpARTCUST1}")
                S.AppendLine($" ORDER BY CUST_NAME")
                ASCMAIN1.sql = S.ToString
                tblARTCUST1 = ASCDATA1.GetDataTable

                S.Length = 0
                S.AppendLine("Select *")
                S.AppendLine($" from {tmpSOTORDR1}")
                S.AppendLine($" ORDER BY ORDR_DATE")
                ASCMAIN1.sql = S.ToString
                tblSOTORDR1 = ASCDATA1.GetDataTable
            End If
        End If

        fillExtraData()

    End Sub

    Private Sub fillExtraData()
        'If OnlyOrderDetails Then Exit Sub

        Dim rowGLTPARM2 As DataRow = _FF.LookUp("GLTPARM2", YYYYBP)
        Dim BEG_DATE As DateTime = rowGLTPARM2.Item("PRD_END_DATE")
        BEG_DATE = BEG_DATE.AddDays(1).AddMonths(-1)
        Dim BEG_DATE_STR As String = Format(BEG_DATE, "dd-MMM-yyyy")

        rowGLTPARM2 = _FF.LookUp("GLTPARM2", YYYYEP)
        Dim END_DATE As DateTime = rowGLTPARM2.Item("PRD_END_DATE")
        Dim END_DATE_STR As String = Format(END_DATE, "dd-MMM-yyyy")

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("(S1.STYLE_CODE || '-' || C1.COLOR_CODE) AS SKU,")
        S.AppendLine("MAX(A1.ATTR_DESC) AS ATTR_DESC")
        S.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1, ICTSTYL3 S3, ICTATTR1 A1")
        S.AppendLine("WHERE S1.STYLE_CODE = C1.STYLE_CODE")
        S.AppendLine("AND S1.STYLE_CODE = S3.STYLE_CODE")
        S.AppendLine("AND S3.ATTR_CODE = A1.ATTR_CODE")
        S.AppendLine("AND NVL(A1.ATT_RANK,'0') = '1'")
        S.AppendLine("GROUP BY (S1.STYLE_CODE || '-' || C1.COLOR_CODE)")
        S.AppendLine("ORDER BY (S1.STYLE_CODE || '-' || C1.COLOR_CODE)")
        ASCMAIN1.sql = S.ToString
        Dim tblICTATTRX As DataTable = ASCDATA1.GetDataTable

        If OnlyOrderDetails Then
            For Each rowSOTORDRD As DataRow In tblSOTORDRD.Select()
                Dim SKU As String = rowSOTORDRD.Item("SKU").ToString & String.Empty
                Dim fltr As String = $"SKU = '{SKU}'"
                Dim rowICTATTRX As DataRow = tblICTATTRX.Select(fltr).FirstOrDefault
                If IsNothing(rowICTATTRX) Then
                    rowSOTORDRD.Item("SUB_STYLE_CLASS_CODE") = "UNKNOWN"
                Else
                    rowSOTORDRD.Item("SUB_STYLE_CLASS_CODE") = rowICTATTRX.Item("ATTR_DESC").ToString.ToUpper & String.Empty
                End If
            Next
        Else
            If OnlyInvoiceHeaders Then
                'S.Length = 0
                'S.AppendLine("SELECT")
                'S.AppendLine("O1.CUST_CODE,")
                'S.AppendLine("MIN(O1.ORDR_DATE) AS FIRST_ORDR")
                'S.AppendLine("FROM SOTORDR1 O1")
                'S.AppendLine("GROUP BY")
                'S.AppendLine("CUST_CODE")
                'ASCMAIN1.sql = S.ToString
                'Dim tblCUSTFST As DataTable = ASCDATA1.GetDataTable

                'S.Length = 0
                'S.AppendLine("SELECT")
                'S.AppendLine("I2.INV_NO,")
                'S.AppendLine("I2.INV_LNO,")
                'S.AppendLine("O2.STYLE_CODE,")
                'S.AppendLine("O2.STYLE_PRICE")
                'S.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, SOTPICK2 P2, SOTORDR2 O2")
                'S.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
                'S.AppendLine("AND I1.PICK_NO = P2.PICK_NO")
                'S.AppendLine("AND I2.INV_LNO = P2.PICK_LNO")
                'S.AppendLine("AND P2.ORDR_NO = O2.ORDR_NO")
                'S.AppendLine("AND P2.ORDR_LNO = O2.ORDR_LNO")
                'S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED >= {YYYYBP}")
                'S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED <= {YYYYEP}")
                'ASCMAIN1.sql = S.ToString
                'Dim tblSOTINVHX As DataTable = ASCDATA1.GetDataTable

                'S.Length = 0
                'S.AppendLine("SELECT")
                'S.AppendLine("P2.STYLE_CODE || '-' || P2.COLOR_CODE AS SKU, P1.PORT_CODE_ORIG")
                'S.AppendLine("FROM POTORDR1 P1, POTORDR2 P2")
                'S.AppendLine("WHERE P1.PO_ORDER_NO = P2.PO_ORDER_NO")
                'S.AppendLine("AND P1.PO_STATUS = 'C'")
                'S.AppendLine("AND NVL(P1.PORT_CODE_ORIG,'NULL') <> 'NULL'")
                'S.AppendLine("AND (P2.STYLE_CODE || '-' || P2.COLOR_CODE, P1.PO_ORDER_NO)")
                'S.AppendLine("IN")
                'S.AppendLine("(")
                'S.AppendLine("SELECT P2.STYLE_CODE || '-' || P2.COLOR_CODE AS SKU, MAX(PO_ORDER_NO) AS PO_ORDER_NO")
                'S.AppendLine("FROM POTORDR2 P2")
                'S.AppendLine("GROUP BY P2.STYLE_CODE || '-' || P2.COLOR_CODE")
                'S.AppendLine(")")
                'ASCMAIN1.sql = S.ToString
                'Dim tblSOTINVLP As DataTable = ASCDATA1.GetDataTable

                'S.Length = 0
                'S.AppendLine("SELECT")
                'S.AppendLine("CUST_CODE")
                'S.AppendLine("FROM ECTECOM1")
                'ASCMAIN1.sql = S.ToString
                'Dim tblECTECOM1 As DataTable = ASCDATA1.GetDataTable

                'For Each rowSOTINVHH As DataRow In tblSOTINVHH.Select()
                '    Dim CUST_CODE As String = rowSOTINVHH.Item("CUST_CODE").ToString & String.Empty
                '    Dim SKU As String = rowSOTINVHH.Item("SKU").ToString & String.Empty
                '    Dim INV_NO As String = rowSOTINVHH.Item("INV_NO").ToString & String.Empty
                '    Dim INV_LNO As Int64 = Val(rowSOTINVHH.Item("INV_LNO").ToString & String.Empty)
                '    Dim fltr As String = $"CUST_CODE = '{CUST_CODE}'"
                '    Dim FIRST_ORDR As String = ""
                '    Dim STYLE_PRICE As Double = 0
                '    Dim rowCUSTFST As DataRow = tblCUSTFST.Select(fltr).FirstOrDefault
                '    If Not IsNothing(rowCUSTFST) Then
                '        FIRST_ORDR = rowCUSTFST.Item("FIRST_ORDR").ToString & String.Empty
                '    End If
                '    If IsDate(FIRST_ORDR) Then
                '        rowSOTINVHH.Item("FIRST_SALE") = CDate(FIRST_ORDR)
                '    Else
                '        rowSOTINVHH.Item("FIRST_SALE") = ""
                '    End If

                '    fltr = $"INV_NO = '{INV_NO}' AND INV_LNO = {INV_LNO}"
                '    Dim rowSOTINVHX As DataRow = tblSOTINVHX.Select(fltr).FirstOrDefault
                '    If Not IsNothing(rowSOTINVHX) Then
                '        If IsNumeric(rowSOTINVHX.Item("STYLE_PRICE").ToString & String.Empty) Then
                '            STYLE_PRICE = Val(rowSOTINVHX.Item("STYLE_PRICE").ToString & String.Empty)
                '        End If
                '    End If
                '    rowSOTINVHH.Item("LIST_PRICE") = STYLE_PRICE

                '    fltr = $"SKU = '{SKU}'"
                '    Dim rowSOTINVLP As DataRow = tblSOTINVLP.Select(fltr).FirstOrDefault
                '    If Not IsNothing(rowSOTINVLP) Then
                '        rowSOTINVHH.Item("LAST_PORT") = rowSOTINVLP.Item("PORT_CODE_ORIG").ToString & String.Empty
                '    End If

                '    Dim rowICTATTRX As DataRow = tblICTATTRX.Select(fltr).FirstOrDefault
                '    If IsNothing(rowICTATTRX) Then
                '        rowSOTINVHH.Item("SUB_STYLE_CLASS_CODE") = "UNKNOWN"
                '    Else
                '        rowSOTINVHH.Item("SUB_STYLE_CLASS_CODE") = rowICTATTRX.Item("ATTR_DESC").ToString.ToUpper & String.Empty
                '    End If

                '    fltr = $"CUST_CODE = '{CUST_CODE}'"
                '    Dim rowECTECOM1 As DataRow = tblECTECOM1.Select(fltr).FirstOrDefault
                '    If Not IsNothing(rowECTECOM1) Then
                '        rowSOTINVHH.Item("CUST_TYPE") = "ECOM"
                '    End If
                'Next
            Else
                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("CUST_CODE,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY,0) * NVL(O2.ORDR_UNIT_PRICE,0)) ORDR_AMT,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY_OPEN,0) * NVL(O2.ORDR_UNIT_PRICE,0)) OPEN_AMT,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY_CANC,0) * NVL(O2.ORDR_UNIT_PRICE,0)) CANC_AMT,")
                S.AppendLine("SUM(NVL(O2.ORDR_QTY_SHIP,0) * NVL(O2.ORDR_UNIT_PRICE,0)) SHIP_AMT")
                S.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
                S.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
                'S.AppendLine($"AND O1.ORDR_YYYYPP_BOOKED >= '{YYYYBP}'")
                'S.AppendLine($"AND O1.ORDR_YYYYPP_BOOKED <= '{YYYYEP}'")
                'S.AppendLine($"AND O1.ORDR_DATE >= '01-MAY-2024'")
                'S.AppendLine($"AND O1.ORDR_DATE <= '30-JUN-2024'")
                S.AppendLine($"AND O1.ORDR_DATE >= '{BEG_DATE_STR}'")
                S.AppendLine($"AND O1.ORDR_DATE <= '{END_DATE_STR}'")
                S.AppendLine("GROUP BY")
                S.AppendLine("CUST_CODE")
                ASCMAIN1.sql = S.ToString
                Dim tblCUSTSLS As DataTable = ASCDATA1.GetDataTable

                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("O1.CUST_CODE,")
                S.AppendLine("MIN(O1.ORDR_DATE) AS FIRST_ORDR")
                S.AppendLine("FROM SOTORDR1 O1")
                S.AppendLine("GROUP BY")
                S.AppendLine("CUST_CODE")
                ASCMAIN1.sql = S.ToString
                Dim tblCUSTFST As DataTable = ASCDATA1.GetDataTable

                For Each rowARTCUST1 As DataRow In tblARTCUST1.Select()
                    Dim ORDR_AMT As Double = 0
                    Dim OPEN_AMT As Double = 0
                    Dim CANC_AMT As Double = 0
                    Dim SHIP_AMT As Double = 0
                    Dim FIRST_ORDR As String = ""
                    Dim CUST_CODE As String = rowARTCUST1.Item("CUST_CODE").ToString & String.Empty
                    Dim fltr As String = $"CUST_CODE = '{CUST_CODE}'"
                    Dim rowCUSTSLS As DataRow = tblCUSTSLS.Select(fltr).FirstOrDefault
                    Dim rowCUSTFST As DataRow = tblCUSTFST.Select(fltr).FirstOrDefault
                    If Not IsNothing(rowCUSTSLS) Then
                        ORDR_AMT = Val(rowCUSTSLS.Item("ORDR_AMT").ToString & String.Empty)
                        OPEN_AMT = Val(rowCUSTSLS.Item("OPEN_AMT").ToString & String.Empty)
                        CANC_AMT = Val(rowCUSTSLS.Item("CANC_AMT").ToString & String.Empty)
                        SHIP_AMT = Val(rowCUSTSLS.Item("SHIP_AMT").ToString & String.Empty)
                    End If
                    If Not IsNothing(rowCUSTFST) Then
                        FIRST_ORDR = rowCUSTFST.Item("FIRST_ORDR").ToString & String.Empty
                    End If
                    rowARTCUST1.Item("ORDERS") = ORDR_AMT
                    rowARTCUST1.Item("OPEN") = OPEN_AMT
                    rowARTCUST1.Item("CANCELED") = CANC_AMT
                    rowARTCUST1.Item("SHIPPED") = SHIP_AMT
                    If IsDate(FIRST_ORDR) Then
                        rowARTCUST1.Item("FIRST_ORDR") = CDate(FIRST_ORDR)
                    Else
                        rowARTCUST1.Item("FIRST_ORDR") = ""
                    End If
                Next

                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("I2.INV_NO,")
                S.AppendLine("I2.INV_LNO,")
                S.AppendLine("O2.STYLE_CODE,")
                S.AppendLine("O2.STYLE_PRICE")
                S.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, SOTPICK2 P2, SOTORDR2 O2")
                S.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
                S.AppendLine("AND I1.PICK_NO = P2.PICK_NO")
                S.AppendLine("AND I2.INV_LNO = P2.PICK_LNO")
                S.AppendLine("AND P2.ORDR_NO = O2.ORDR_NO")
                S.AppendLine("AND P2.ORDR_LNO = O2.ORDR_LNO")
                S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED >= {YYYYBP}")
                S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED <= {YYYYEP}")
                ASCMAIN1.sql = S.ToString
                Dim tblSOTINVHX As DataTable = ASCDATA1.GetDataTable

                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("P2.STYLE_CODE || '-' || P2.COLOR_CODE AS SKU, P1.PORT_CODE_ORIG")
                S.AppendLine("FROM POTORDR1 P1, POTORDR2 P2")
                S.AppendLine("WHERE P1.PO_ORDER_NO = P2.PO_ORDER_NO")
                S.AppendLine("AND P1.PO_STATUS = 'C'")
                S.AppendLine("AND NVL(P1.PORT_CODE_ORIG,'NULL') <> 'NULL'")
                S.AppendLine("AND (P2.STYLE_CODE || '-' || P2.COLOR_CODE, P1.PO_ORDER_NO)")
                S.AppendLine("IN")
                S.AppendLine("(")
                S.AppendLine("SELECT P2.STYLE_CODE || '-' || P2.COLOR_CODE AS SKU, MAX(PO_ORDER_NO) AS PO_ORDER_NO")
                S.AppendLine("FROM POTORDR2 P2")
                S.AppendLine("GROUP BY P2.STYLE_CODE || '-' || P2.COLOR_CODE")
                S.AppendLine(")")
                ASCMAIN1.sql = S.ToString
                Dim tblSOTINVLP As DataTable = ASCDATA1.GetDataTable

                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("CUST_CODE")
                S.AppendLine("FROM ECTECOM1")
                ASCMAIN1.sql = S.ToString
                Dim tblECTECOM1 As DataTable = ASCDATA1.GetDataTable

                For Each rowSOTINVH1 As DataRow In tblSOTINVH1.Select()
                    Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE").ToString & String.Empty
                    Dim SKU As String = rowSOTINVH1.Item("SKU").ToString & String.Empty
                    Dim INV_NO As String = rowSOTINVH1.Item("INV_NO").ToString & String.Empty
                    Dim INV_LNO As Int64 = Val(rowSOTINVH1.Item("INV_LNO").ToString & String.Empty)
                    Dim fltr As String = $"CUST_CODE = '{CUST_CODE}'"
                    Dim FIRST_ORDR As String = ""
                    Dim STYLE_PRICE As Double = 0
                    Dim rowCUSTFST As DataRow = tblCUSTFST.Select(fltr).FirstOrDefault
                    If Not IsNothing(rowCUSTFST) Then
                        FIRST_ORDR = rowCUSTFST.Item("FIRST_ORDR").ToString & String.Empty
                    End If
                    If IsDate(FIRST_ORDR) Then
                        rowSOTINVH1.Item("FIRST_SALE") = CDate(FIRST_ORDR)
                    Else
                        rowSOTINVH1.Item("FIRST_SALE") = ""
                    End If

                    fltr = $"INV_NO = '{INV_NO}' AND INV_LNO = {INV_LNO}"
                    Dim rowSOTINVHX As DataRow = tblSOTINVHX.Select(fltr).FirstOrDefault
                    If Not IsNothing(rowSOTINVHX) Then
                        If IsNumeric(rowSOTINVHX.Item("STYLE_PRICE").ToString & String.Empty) Then
                            STYLE_PRICE = Val(rowSOTINVHX.Item("STYLE_PRICE").ToString & String.Empty)
                        End If
                    End If
                    rowSOTINVH1.Item("LIST_PRICE") = STYLE_PRICE

                    fltr = $"SKU = '{SKU}'"
                    Dim rowSOTINVLP As DataRow = tblSOTINVLP.Select(fltr).FirstOrDefault
                    If Not IsNothing(rowSOTINVLP) Then
                        rowSOTINVH1.Item("LAST_PORT") = rowSOTINVLP.Item("PORT_CODE_ORIG").ToString & String.Empty
                    End If

                    Dim rowICTATTRX As DataRow = tblICTATTRX.Select(fltr).FirstOrDefault
                    If IsNothing(rowICTATTRX) Then
                        rowSOTINVH1.Item("SUB_STYLE_CLASS_CODE") = "UNKNOWN"
                    Else
                        rowSOTINVH1.Item("SUB_STYLE_CLASS_CODE") = rowICTATTRX.Item("ATTR_DESC").ToString.ToUpper & String.Empty
                    End If

                    fltr = $"CUST_CODE = '{CUST_CODE}'"
                    Dim rowECTECOM1 As DataRow = tblECTECOM1.Select(fltr).FirstOrDefault
                    If Not IsNothing(rowECTECOM1) Then
                        rowSOTINVH1.Item("CUST_TYPE") = "ECOM"
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub InitializeFormatting()
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("'Customers' AS EXCEL_BOOK,")
        S.AppendLine("0 AS EXCEL_COL")
        S.AppendLine(" from DUAL")
        ASCMAIN1.sql = S.ToString
        tblEXCEL = ASCDATA1.GetDataTable
        With tblEXCEL.Columns
            .Add("DATA_COL", GetType(System.String))
            .Add("DATA_TYPE", GetType(System.String))
            .Add("WIDTH", GetType(System.Int64))
            .Add("HEADING", GetType(System.String))
            .Add("FORMAT", GetType(System.String))
        End With

        tblEXCEL.Rows(0).Delete()
        If OnlyOrderDetails Then
            tblEXCEL.Rows.Add({"Orders", 0, "ORDR_NO", "S", 10, "Order No", ""})
            tblEXCEL.Rows.Add({"Orders", 1, "ORDR_DATE", "D", 10, "Date", ""})
            tblEXCEL.Rows.Add({"Orders", 2, "ORDR_DATE_RECD", "D", 10, "Recd", ""})
            tblEXCEL.Rows.Add({"Orders", 3, "CUST_CODE", "S", 10, "Cust Code", ""})
            tblEXCEL.Rows.Add({"Orders", 4, "CUST_NAME", "S", 40, "Name", ""})
            tblEXCEL.Rows.Add({"Orders", 5, "SKU", "S", 18, "SKU", ""})
            tblEXCEL.Rows.Add({"Orders", 6, "STYLE_DESC", "S", 45, "Description", ""})
            tblEXCEL.Rows.Add({"Orders", 7, "STYLE_CLASS_CODE", "S", 12, "Class Code", ""})
            tblEXCEL.Rows.Add({"Orders", 8, "SUB_STYLE_CLASS_CODE", "S", 20, "Sub Class Code", ""})
            tblEXCEL.Rows.Add({"Orders", 9, "ORDR_AMT", "2", 10, "Amount", ""})
            tblEXCEL.Rows.Add({"Orders", 10, "OPEN_AMT", "2", 10, "Open", ""})
            tblEXCEL.Rows.Add({"Orders", 11, "CANC_AMT", "2", 10, "Cancelled", ""})
            tblEXCEL.Rows.Add({"Orders", 12, "SHIP_AMT", "2", 10, "Shipped", ""})
            tblEXCEL.Rows.Add({"Orders", 13, "ORDR_QTY", "2", 10, "Order Qty", ""})
            tblEXCEL.Rows.Add({"Orders", 14, "ORDR_SHIP_DATE", "D", 10, "Ship Date", ""})
            tblEXCEL.Rows.Add({"Orders", 15, "ORDR_CANCEL_DATE", "D", 10, "Cancel Date", ""})
            tblEXCEL.Rows.Add({"Orders", 16, "WHSE_CODE", "S", 10, "Whs", ""})
        Else
            If OnlyInvoiceHeaders Then
                tblEXCEL.Rows.Add({"Invoices", 0, "ORDR_YYYYPP_UPDATED", "S", 10, "Period", ""})
                tblEXCEL.Rows.Add({"Invoices", 1, "INV_TYPE_D", "S", 10, "Type", ""})
                tblEXCEL.Rows.Add({"Invoices", 2, "INV_DATE", "D", 10, "Invoice Date", ""})
                tblEXCEL.Rows.Add({"Invoices", 3, "INIT_DATE", "D", 10, "Date Created", ""})
                tblEXCEL.Rows.Add({"Invoices", 4, "INV_NO", "S", 14, "Invoice No", ""})
                tblEXCEL.Rows.Add({"Invoices", 5, "SREP_CODE", "S", 10, "Sales Rep", ""})
                tblEXCEL.Rows.Add({"Invoices", 6, "CUST_NAME", "S", 30, "Customer Name", ""})
                tblEXCEL.Rows.Add({"Invoices", 7, "CUST_CODE", "S", 16, "Customer Code", ""})
                tblEXCEL.Rows.Add({"Invoices", 8, "CUST_TYPE", "S", 16, "Customer Type", ""})
                tblEXCEL.Rows.Add({"Invoices", 9, "WHSE_CODE", "S", 16, "Warehouse", ""})
                tblEXCEL.Rows.Add({"Invoices", 10, "CUST_STATE", "S", 16, "State", ""})
                tblEXCEL.Rows.Add({"Invoices", 11, "CUST_COUNTRY", "S", 16, "Country", ""})
                tblEXCEL.Rows.Add({"Invoices", 12, "CUST_ZIP_CODE", "S", 10, "Zip Code", ""})
                tblEXCEL.Rows.Add({"Invoices", 13, "INV_SALES_CURR", "2", 10, "Sales", ""})
                tblEXCEL.Rows.Add({"Invoices", 14, "INV_FREIGHT_CURR", "2", 10, "Freight", ""})
                tblEXCEL.Rows.Add({"Invoices", 15, "INV_MISC_CHG_CURR", "2", 10, "Misc(Inc. Tariff)", ""})
                tblEXCEL.Rows.Add({"Invoices", 16, "INV_TOTAL_AMT_CURR", "2", 10, "Total", ""})
                tblEXCEL.Rows.Add({"Invoices", 17, "INV_MISC_CHG_CURR_T", "2", 10, "Tariff", ""})

                'tblEXCEL.Rows.Add({"Invoices", 6, "ORDR_DATE", "D", 10, "Order Date", ""})
                'tblEXCEL.Rows.Add({"Invoices", 7, "ORDR_NO", "S", 14, "Order No", ""})
                'tblEXCEL.Rows.Add({"Invoices", 15, "FIRST_SALE", "D", 10, "1st Sale", ""})
                'tblEXCEL.Rows.Add({"Invoices", 16, "SKU", "S", 18, "SKU", ""})
                'tblEXCEL.Rows.Add({"Invoices", 17, "STYLE_DESC", "S", 45, "Description", ""})
                'tblEXCEL.Rows.Add({"Invoices", 18, "STYLE_CLASS_CODE", "S", 12, "Class Code", ""})
                'tblEXCEL.Rows.Add({"Invoices", 19, "SUB_STYLE_CLASS_CODE", "S", 20, "Sub Class Code", ""})
                'tblEXCEL.Rows.Add({"Invoices", 20, "LIST_PRICE", "2", 12, "List Price", ""})
                'tblEXCEL.Rows.Add({"Invoices", 21, "ORDR_QTY_SHIP", "0", 10, "Qty Shipped", ""})
                'tblEXCEL.Rows.Add({"Invoices", 22, "ORDR_UNIT_PRICE", "2", 10, "Price", ""})
                'tblEXCEL.Rows.Add({"Invoices", 23, "ORDR_REVENUE", "2", 10, "Revenue", ""})
                'tblEXCEL.Rows.Add({"Invoices", 24, "ORDR_UNIT_COST", "2", 10, "1st Cost", ""})
                'tblEXCEL.Rows.Add({"Invoices", 25, "ORDR_COGS", "2", 10, "COGS", ""})
                'tblEXCEL.Rows.Add({"Invoices", 26, "STD_LANDED_PCT", "P", 12, "Std Land Pct", ""})
                'tblEXCEL.Rows.Add({"Invoices", 27, "LANDED_COST", "2", 12, "Landed Cost", ""})
                'tblEXCEL.Rows.Add({"Invoices", 28, "GP", "2", 12, "GP", ""})
                'tblEXCEL.Rows.Add({"Invoices", 29, "GP_PCT", "P", 12, "GP%", ""})
                'tblEXCEL.Rows.Add({"Invoices", 30, "COMM_RATE", "2", 12, "Comm Rate", ""})
                'tblEXCEL.Rows.Add({"Invoices", 31, "VEND_ID", "S", 10, "Vendor ID", ""})
                'tblEXCEL.Rows.Add({"Invoices", 32, "VEND_CODE", "S", 12, "Vendor", ""})
                'tblEXCEL.Rows.Add({"Invoices", 33, "PORT_CODE", "S", 12, "Port Code", ""})
                'tblEXCEL.Rows.Add({"Invoices", 34, "PORT_NAME", "S", 12, "Port Name", ""})
                'tblEXCEL.Rows.Add({"Invoices", 35, "LAST_PORT", "S", 12, "Last Port", ""})
                'tblEXCEL.Rows.Add({"Invoices", 36, "COUNTRY_NAME", "S", 10, "Country", ""})
                'tblEXCEL.Rows.Add({"Invoices", 37, "ORDR_SHIP_DATE", "D", 10, "Ship Date", ""})
                'tblEXCEL.Rows.Add({"Invoices", 38, "ORDR_CANCEL_DATE", "D", 10, "Cancel Date", ""})


            Else
                tblEXCEL.Rows.Add({"Invoices", 0, "ORDR_YYYYPP_UPDATED", "S", 10, "Period", ""})
                tblEXCEL.Rows.Add({"Invoices", 1, "INV_TYPE_D", "S", 10, "Type", ""})
                tblEXCEL.Rows.Add({"Invoices", 2, "INV_DATE", "D", 10, "Invoice Date", ""})
                tblEXCEL.Rows.Add({"Invoices", 3, "INIT_DATE", "D", 10, "Date Created", ""})
                tblEXCEL.Rows.Add({"Invoices", 4, "INV_NO", "S", 14, "Invoice No", ""})
                tblEXCEL.Rows.Add({"Invoices", 5, "INV_LNO", "0", 5, "LNo", ""})
                tblEXCEL.Rows.Add({"Invoices", 6, "ORDR_DATE", "D", 10, "Order Date", ""})
                tblEXCEL.Rows.Add({"Invoices", 7, "ORDR_NO", "S", 14, "Order No", ""})
                tblEXCEL.Rows.Add({"Invoices", 8, "SREP_CODE", "S", 10, "Sales Rep", ""})
                tblEXCEL.Rows.Add({"Invoices", 9, "CUST_NAME", "S", 30, "Customer Name", ""})
                tblEXCEL.Rows.Add({"Invoices", 10, "CUST_CODE", "S", 16, "Customer Code", ""})
                tblEXCEL.Rows.Add({"Invoices", 11, "CUST_TYPE", "S", 16, "Customer Type", ""})
                tblEXCEL.Rows.Add({"Invoices", 12, "WHSE_CODE", "S", 16, "Warehouse", ""})
                tblEXCEL.Rows.Add({"Invoices", 13, "STATE_COUNTRY", "S", 16, "State / Country", ""})
                tblEXCEL.Rows.Add({"Invoices", 14, "CUST_ZIP_CODE", "S", 10, "Zip Code", ""})
                tblEXCEL.Rows.Add({"Invoices", 15, "FIRST_SALE", "D", 10, "1st Sale", ""})
                tblEXCEL.Rows.Add({"Invoices", 16, "SKU", "S", 18, "SKU", ""})
                tblEXCEL.Rows.Add({"Invoices", 17, "STYLE_DESC", "S", 45, "Description", ""})
                tblEXCEL.Rows.Add({"Invoices", 18, "STYLE_CLASS_CODE", "S", 12, "Class Code", ""})
                tblEXCEL.Rows.Add({"Invoices", 19, "SUB_STYLE_CLASS_CODE", "S", 20, "Sub Class Code", ""})
                tblEXCEL.Rows.Add({"Invoices", 20, "LIST_PRICE", "2", 12, "List Price", ""})
                tblEXCEL.Rows.Add({"Invoices", 21, "ORDR_QTY_SHIP", "0", 10, "Qty Shipped", ""})
                tblEXCEL.Rows.Add({"Invoices", 22, "ORDR_UNIT_PRICE", "2", 10, "Price", ""})
                tblEXCEL.Rows.Add({"Invoices", 23, "ORDR_REVENUE", "2", 10, "Revenue", ""})
                tblEXCEL.Rows.Add({"Invoices", 24, "ORDR_UNIT_COST", "2", 10, "1st Cost", ""})
                tblEXCEL.Rows.Add({"Invoices", 25, "ORDR_COGS", "2", 10, "COGS", ""})
                tblEXCEL.Rows.Add({"Invoices", 26, "STD_LANDED_PCT", "P", 12, "Std Land Pct", ""})
                tblEXCEL.Rows.Add({"Invoices", 27, "LANDED_COST", "2", 12, "Landed Cost", ""})
                tblEXCEL.Rows.Add({"Invoices", 28, "GP", "2", 12, "GP", ""})
                tblEXCEL.Rows.Add({"Invoices", 29, "GP_PCT", "P", 12, "GP%", ""})
                tblEXCEL.Rows.Add({"Invoices", 30, "COMM_RATE", "2", 12, "Comm Rate", ""})
                tblEXCEL.Rows.Add({"Invoices", 31, "VEND_ID", "S", 10, "Vendor ID", ""})
                tblEXCEL.Rows.Add({"Invoices", 32, "VEND_CODE", "S", 12, "Vendor", ""})
                tblEXCEL.Rows.Add({"Invoices", 33, "PORT_CODE", "S", 12, "Port Code", ""})
                tblEXCEL.Rows.Add({"Invoices", 34, "PORT_NAME", "S", 12, "Port Name", ""})
                tblEXCEL.Rows.Add({"Invoices", 35, "LAST_PORT", "S", 12, "Last Port", ""})
                tblEXCEL.Rows.Add({"Invoices", 36, "COUNTRY_NAME", "S", 10, "Country", ""})
                tblEXCEL.Rows.Add({"Invoices", 37, "ORDR_SHIP_DATE", "D", 10, "Ship Date", ""})
                tblEXCEL.Rows.Add({"Invoices", 38, "ORDR_CANCEL_DATE", "D", 10, "Cancel Date", ""})

                tblEXCEL.Rows.Add({"Orders", 0, "ORDR_NO", "S", 10, "Order No", ""})
                tblEXCEL.Rows.Add({"Orders", 1, "ORDR_DATE", "D", 10, "Date", ""})
                tblEXCEL.Rows.Add({"Orders", 2, "ORDR_DATE_RECD", "D", 10, "Recd", ""})
                tblEXCEL.Rows.Add({"Orders", 3, "CUST_CODE", "S", 10, "Cust Code", ""})
                tblEXCEL.Rows.Add({"Orders", 4, "CUST_NAME", "S", 40, "Name", ""})
                tblEXCEL.Rows.Add({"Orders", 5, "ORDR_AMT", "2", 10, "Amount", ""})
                tblEXCEL.Rows.Add({"Orders", 6, "OPEN_AMT", "2", 10, "Open", ""})
                tblEXCEL.Rows.Add({"Orders", 7, "CANC_AMT", "2", 10, "Cancelled", ""})
                tblEXCEL.Rows.Add({"Orders", 8, "SHIP_AMT", "2", 10, "Shipped", ""})
                tblEXCEL.Rows.Add({"Orders", 9, "ORDR_QTY", "2", 10, "Order Qty", ""})
                tblEXCEL.Rows.Add({"Orders", 10, "ORDR_SHIP_DATE", "D", 10, "Ship Date", ""})
                tblEXCEL.Rows.Add({"Orders", 11, "ORDR_CANCEL_DATE", "D", 10, "Cancel Date", ""})

                tblEXCEL.Rows.Add({"Customers", 0, "CUST_CODE", "S", 10, "Cust Code", ""})
                tblEXCEL.Rows.Add({"Customers", 1, "CUST_NAME", "S", 40, "Name", ""})
                tblEXCEL.Rows.Add({"Customers", 2, "CUST_ADDR1", "S", 25, "Street", ""})
                tblEXCEL.Rows.Add({"Customers", 3, "CUST_CITY", "S", 25, "City", ""})
                tblEXCEL.Rows.Add({"Customers", 4, "CUST_STATE", "S", 10, "State", ""})
                tblEXCEL.Rows.Add({"Customers", 5, "CUST_ZIP_CODE", "S", 10, "Zip Code", ""})
                tblEXCEL.Rows.Add({"Customers", 6, "CUST_COUNTRY", "S", 10, "Counrty", ""})
                tblEXCEL.Rows.Add({"Customers", 7, "SREP_CODE", "S", 10, "Sales Rep", ""})
                tblEXCEL.Rows.Add({"Customers", 8, "CUST_PRICE_TIER", "S", 15, "Pricing", ""})
                tblEXCEL.Rows.Add({"Customers", 9, "CUST_PRICE_TIER_PVC", "S", 15, "Pricing PVC", ""})
                tblEXCEL.Rows.Add({"Customers", 10, "FIRST_ORDR", "D", 15, "1st Order", ""})
                tblEXCEL.Rows.Add({"Customers", 11, "ORDERS", "0", 15, "Ordered", ""})
                tblEXCEL.Rows.Add({"Customers", 12, "OPEN", "0", 15, "Open", ""})
                tblEXCEL.Rows.Add({"Customers", 13, "CANCELED", "0", 15, "Cancelled", ""})
                tblEXCEL.Rows.Add({"Customers", 14, "SHIPPED", "0", 15, "Shipped", ""})
            End If
        End If
        tblEXCEL.AcceptChanges()
    End Sub
#End Region

#Region "Custom Methods"
    Public Sub makeExcel()
        If eMsg.Length = 0 Then
            Dim IP As Integer = 0
            _FF.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Beginning Excel Creation", "")

            setxlsInfo()

            If OnlyOrderDetails Then
                'Orders
                oSheet = oWB.Worksheets.Add()
                oSheet.Name = "Orders"
                ASCMAIN1.Progress("-", oSheet.Name)
                _FF.Load_DataTable_into_SGXLS(TRC + 1, TC + 1, tblSOTORDRD, oSheet, Nothing, Nothing, "ORDR_DATE, ORDR_NO, SKU", "")
                setWorkBookHeadings(oSheet)
                setWorkBookTotals(oSheet)
                oSheet.Range(TRC, 0).EntireRow.AutoFilter()
                setWorkbookAltColors(oSheet, tblSOTORDRD)
                setWorkbookFreeze(oSheet)
            Else
                If OnlyInvoiceHeaders Then
                    oSheet = oWB.Worksheets(0)
                    oSheet.Name = "Invoices"
                    ASCMAIN1.Progress("-", oSheet.Name)
                    _FF.Load_DataTable_into_SGXLS(TRC + 1, TC + 1, tblSOTINVHH, oSheet, Nothing, Nothing, "INV_DATE", "")
                    setWorkBookHeadings(oSheet)
                    setWorkBookTotals(oSheet)
                    oSheet.Range(TRC, 0).EntireRow.AutoFilter()
                    setWorkbookAltColors(oSheet, tblSOTINVHH)
                    'setFormulas(oSheet, tblSOTINVHH)
                    setWorkbookFreeze(oSheet)
                Else
                    'Invoices
                    oSheet = oWB.Worksheets(0)
                    oSheet.Name = "Invoices"
                    ASCMAIN1.Progress("-", oSheet.Name)
                    _FF.Load_DataTable_into_SGXLS(TRC + 1, TC + 1, tblSOTINVH1, oSheet, Nothing, Nothing, "INV_DATE", "")
                    setWorkBookHeadings(oSheet)
                    setWorkBookTotals(oSheet)
                    oSheet.Range(TRC, 0).EntireRow.AutoFilter()
                    setWorkbookAltColors(oSheet, tblSOTINVH1)
                    setFormulas(oSheet, tblSOTINVH1)
                    setWorkbookFreeze(oSheet)

                    'Orders
                    oSheet = oWB.Worksheets.Add()
                    oSheet.Name = "Orders"
                    ASCMAIN1.Progress("-", oSheet.Name)
                    _FF.Load_DataTable_into_SGXLS(TRC + 1, TC + 1, tblSOTORDR1, oSheet, Nothing, Nothing, "ORDR_DATE", "")
                    setWorkBookHeadings(oSheet)
                    setWorkBookTotals(oSheet)
                    oSheet.Range(TRC, 0).EntireRow.AutoFilter()
                    setWorkbookAltColors(oSheet, tblSOTORDR1)
                    setWorkbookFreeze(oSheet)

                    'Customers
                    oSheet = oWB.Worksheets.Add()
                    oSheet.Name = "Customers"
                    ASCMAIN1.Progress("-", oSheet.Name)
                    _FF.Load_DataTable_into_SGXLS(TRC + 1, TC + 1, tblARTCUST1, oSheet, Nothing, Nothing, "CUST_NAME", "")
                    setWorkBookHeadings(oSheet)
                    setWorkBookTotals(oSheet)
                    oSheet.Range(TRC, 0).EntireRow.AutoFilter()
                    setWorkbookAltColors(oSheet, tblARTCUST1)
                    setWorkbookFreeze(oSheet)
                End If

            End If

            oWB.Worksheets(0).Select()

            oWB.SaveAs(xls_file_name, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
            _FF.Show_Document(xls_file_name)
            oWB = Nothing

            _FF.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
        Else
            MsgBox(eMsg, vbCritical, "Bad Bad")
            _FF.Cursor = Cursors.Default
        End If

    End Sub

    Private Sub setWorkbookFreeze(ByRef oSheet As SpreadsheetGear.IWorksheet)
        oSheet.Range(TRC + 1, 0).Select()
        oSheet.WindowInfo.FreezePanes = True
        oSheet.Range("A1:A1").Select()
    End Sub

    Private Sub setWorkbookAltColors(ByRef oSheet As SpreadsheetGear.IWorksheet, ByVal tbl As DataTable)
        Dim fltr As String = $"EXCEL_BOOK = '{oSheet.Name}'"
        Dim LCOL As Double = tblEXCEL.Compute("MAX(EXCEL_COL)", fltr)
        For i As Integer = TRB + 1 To TRB + tbl.Rows.Count Step 2
            oSheet.Range(i, 0, i, LCOL).Interior.Color = SpreadsheetGear.Colors.LightBlue
        Next
    End Sub

    Private Sub setWorkBookTotals(ByRef oSheet As SpreadsheetGear.IWorksheet)
        'oSheet.Cells(TRC + tblSOTINVH1.Rows.Count + 1, 0).Value = "Totals"
        'oSheet.Cells(TRC + tblSOTINVH1.Rows.Count + 1, 14).Formula = $"=SUBTOTAL(9,O3:O23)"
    End Sub

    Private Sub setWorkBookHeadings(ByRef oSheet As SpreadsheetGear.IWorksheet)
        Dim fltr As String = $"EXCEL_BOOK = '{oSheet.Name}'"
        For Each rowEXCEL As DataRow In tblEXCEL.Select(fltr, "EXCEL_COL")
            Dim EXCEL_COL As Int64 = Val(rowEXCEL.Item("EXCEL_COL").ToString & String.Empty)
            Dim HEADING As String = rowEXCEL.Item("HEADING").ToString & String.Empty
            Dim DATA_TYPE As String = rowEXCEL.Item("DATA_TYPE").ToString & String.Empty
            Dim WIDTH As Int64 = Val(rowEXCEL.Item("WIDTH").ToString & String.Empty)
            oSheet.Cells(TRB, EXCEL_COL).Value = HEADING
            oSheet.Cells(TRB, EXCEL_COL).EntireColumn.ColumnWidth = WIDTH
            oSheet.Cells(TRB, EXCEL_COL).Font.Bold = True
            oSheet.Cells(TRB, EXCEL_COL).Interior.Color = SpreadsheetGear.Colors.Yellow
            If IsNumeric(DATA_TYPE) Then
                Select Case Val(DATA_TYPE)
                    Case 0
                        oSheet.Cells(TRB, EXCEL_COL).EntireColumn.NumberFormat = "###,##0"
                    Case 1
                        oSheet.Cells(TRB, EXCEL_COL).EntireColumn.NumberFormat = "###,##0.0"
                    Case 2
                        oSheet.Cells(TRB, EXCEL_COL).EntireColumn.NumberFormat = "###,##0.00"
                    Case Else
                        oSheet.Cells(TRB, EXCEL_COL).EntireColumn.NumberFormat = "###,##0.00"
                End Select
            Else
                Select Case DATA_TYPE
                    Case "D"
                        oSheet.Cells(TRB, EXCEL_COL).EntireColumn.NumberFormat = "MM/dd/yy"
                        oSheet.Cells(TRB, EXCEL_COL).EntireColumn.HorizontalAlignment = HAlign.Center
                    Case "P"
                        oSheet.Cells(TRB, EXCEL_COL).EntireColumn.NumberFormat = "###.00%"
                    Case Else
                        oSheet.Cells(TRB, EXCEL_COL).EntireColumn.NumberFormat = "@"
                End Select

            End If
        Next

    End Sub

    Private Sub setxlsInfo()
        Do Until success
            Try
                XLS_NO += 1
                xls_name = ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                xls_file_name = xls_path & "\" & xls_name & ".XLSx"
                If Not My.Computer.FileSystem.FileExists(xls_file_name) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        oWB = SpreadsheetGear.Factory.GetWorkbook()
        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i
    End Sub

    Private Sub setFormulas(ByRef oSheet As SpreadsheetGear.IWorksheet, ByVal tbl As DataTable)
        Dim fltr As String = $"EXCEL_BOOK = '{oSheet.Name}'"
        Dim LCOL As Double = tblEXCEL.Compute("MAX(EXCEL_COL)", fltr)
        Dim COL_ORDR_COGS As KeyValuePair(Of String, Int64) = getCOL_INFO("ORDR_COGS")
        Dim COL_STD_LANDED_PCT As KeyValuePair(Of String, Int64) = getCOL_INFO("STD_LANDED_PCT")
        Dim COL_LANDED_COST As KeyValuePair(Of String, Int64) = getCOL_INFO("LANDED_COST")
        Dim COL_ORDR_REVENUE As KeyValuePair(Of String, Int64) = getCOL_INFO("ORDR_REVENUE")
        Dim COL_GP As KeyValuePair(Of String, Int64) = getCOL_INFO("GP")
        Dim GP_PCT As KeyValuePair(Of String, Int64) = getCOL_INFO("GP_PCT")
        Dim WHSE_CODE As KeyValuePair(Of String, Int64) = getCOL_INFO("WHSE_CODE")
        Dim ORDR_YYYYPP_UPDATED As KeyValuePair(Of String, Int64) = getCOL_INFO("ORDR_YYYYPP_UPDATED")

        For i As Integer = TRB + 1 To TRB + tbl.Rows.Count
            If oSheet.Cells(i, WHSE_CODE.Value).Text = "FE" Then
                oSheet.Cells(i, COL_STD_LANDED_PCT.Value).Value = 0
            Else
                If oSheet.Cells(i, ORDR_YYYYPP_UPDATED.Value).Text.Substring(0, 4) = "2022" Then
                    oSheet.Cells(i, COL_STD_LANDED_PCT.Value).Value = 0.35
                End If
            End If
            'AA(26) = (Y*(1+Z)  'Y=24, Z =25 
            oSheet.Cells(i, COL_LANDED_COST.Value).Formula = $"=({COL_ORDR_COGS.Key}{i + 1}*(1+{COL_STD_LANDED_PCT.Key}{i + 1}))"
            'AB(27) = W - AA   'AA=26, W=24
            oSheet.Cells(i, COL_GP.Value).Formula = $"=({COL_ORDR_REVENUE.Key}{i + 1} - {COL_LANDED_COST.Key}{i + 1})"
            'AC=AB/W
            oSheet.Cells(i, GP_PCT.Value).Formula = $"=({COL_GP.Key}{i + 1} / {COL_ORDR_REVENUE.Key}{i + 1})"
        Next
    End Sub

    Private Function getCOL_INFO(ByVal COL As String) As KeyValuePair(Of String, Long)
        Dim COL_LTR As String = "A"
        Dim COL_NBR As Int64 = 1
        Dim fltr As String = $"EXCEL_BOOK = 'Invoices' AND DATA_COL = '{COL}'"
        Dim rowEXCEL As DataRow = tblEXCEL.Select(fltr).FirstOrDefault
        If Not IsNothing(rowEXCEL) Then
            COL_NBR = Val(rowEXCEL.Item("EXCEL_COL").ToString)
            COL_LTR = getExcelCol(COL_NBR)
        End If
        Return New KeyValuePair(Of String, Long)(COL_LTR, COL_NBR)
    End Function

    Private Function getExcelCol(ByVal COL_NBR As Long) As String
        Dim MX As Int64 = 25
        Dim RetVal As String = ""
        If COL_NBR > MX Then
            RetVal = "A"
            COL_NBR = COL_NBR - (MX)
            RetVal = RetVal & Chr(Asc("A") + COL_NBR - 1)
        Else
            RetVal = RetVal & Chr(Asc("A") + COL_NBR)
        End If
        Return RetVal
    End Function
#End Region


#Region "Space Code"

    ' Load the DataTable into the Sheet
    'Dim TR As Integer = 4
    'Dim TC As Integer = 0
    'Dim sqlIP As String = ""

    'ASCMAIN1.sql = $"SELECT * FROM {}"
    'Dim tbl2 As DataTable = ASCDATA1.GetDataTable

    'oSheet.Range(TRC, 0, TRC, 4).EntireColumn.NumberFormat = "@"
    'oSheet.Range(TRC, 0).EntireRow.NumberFormat = "@"

    'Dim Cx As Integer = -1
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "Store No" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 10
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "Store Name" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 30
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "ASD" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "AE" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "AC" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20

    'oSheet.Range(0, 0).EntireRow.NumberFormat = "@"
    'oSheet.Range(0, 0).EntireRow.Font.Color = SpreadsheetGear.Colors.Red

    'oSheet.Range(1, 0).EntireRow.NumberFormat = "@"
    'oSheet.Range(2, 0).EntireRow.NumberFormat = "@"
    'oSheet.Range(3, 0).EntireRow.NumberFormat = "@"

    'Dim CIP As Integer = Cx

    'oSheet.Cells(0 + 1, CIP).Value = "Customer PO"
    'oSheet.Cells(0 + 2, CIP).Value = "Description"
    'oSheet.Cells(0 + 3, CIP).Value = "Item Code"

    '.Add("DATA_COL", GetType(System.String))
    '.Add("DATA_TYPE", GetType(System.String))
    '.Add("WIDTH", GetType(System.Int64))
    '.Add("HEADING", GetType(System.String))
    '.Add("FORMAT", GetType(System.String))

    'Paint Headings
    'Dim Cx As Integer = -1
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "Store No" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 10
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "Store Name" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 30
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "ASD" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "AE" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20
    'Cx += 1 : oSheet.Cells(TR, Cx).Value = "AC" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20

    'oSheet.Range(0, 0).EntireRow.NumberFormat = "@"
    'oSheet.Range(0, 0).EntireRow.Font.Color = SpreadsheetGear.Colors.Red

    'oSheet.Range(1, 0).EntireRow.NumberFormat = "@"
    'oSheet.Range(2, 0).EntireRow.NumberFormat = "@"
    'oSheet.Range(3, 0).EntireRow.NumberFormat = "@"

    'Dim CIP As Integer = Cx

    'oSheet.Cells(0 + 1, CIP).Value = "Customer PO"
    'oSheet.Cells(0 + 2, CIP).Value = "Description"
    'oSheet.Cells(0 + 3, CIP).Value = "Item Code"


    'For I As Integer = 1 To IP
    '    Cx = CIP + I
    '    oSheet.Cells(TR, Cx).Value = Format(I, "000")
    '    oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 12
    '    oSheet.Cells(TR, Cx).HorizontalAlignment = SpreadsheetGear.HAlign.Center
    'Next

    'Dim IPx As Integer = 0
    'For Each rowSATALLC2 As DataRow In tblSATALLC2.Select("", "ITEM_CODE, ORDR_CUST_PO")
    '    IPx += 1
    '    Dim ITEM_CODE As String = rowSATALLC2.Item("ITEM_CODE")
    '    Dim ITEM_DESC As String = rowSATALLC2.Item("ITEM_DESC") & ""
    '    Dim ORDR_CUST_PO As String = rowSATALLC2.Item("ORDR_CUST_PO") & ""
    '    oSheet.Cells(0, CIP + IPx).Value = "Wave X"
    '    oSheet.Cells(1, CIP + IPx).Value = ORDR_CUST_PO
    '    oSheet.Cells(2, CIP + IPx).Value = ITEM_DESC
    '    oSheet.Cells(2, CIP + IPx).WrapText = True
    '    oSheet.Cells(3, CIP + IPx).Value = ITEM_CODE

    '    oSheet.Cells(TR + tbl2.Rows.Count + 1, CIP + IPx).Formula = $"=SUBTOTAL(9,{Excel_Cell0(TR + 1, CIP + IPx)}:{Excel_Cell0(TR + tbl2.Rows.Count, CIP + IPx)})"
    'Next

    ' Border around Entry Area
    'If IP > 0 Then
    '    With oSheet.Range(0, CIP + 1, 3, CIP + IP)
    '        .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
    '        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
    '        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
    '        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
    '        .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
    '        .Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
    '    End With
    'End If
#End Region
End Class
