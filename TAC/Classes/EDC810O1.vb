Public Class EDC810O1

    Private tblEDT810O1 As DataTable = Nothing
    Private tblEDT810O2 As DataTable = Nothing
    Private tblEDT810O3 As DataTable = Nothing
    Private tblEDT810O4 As DataTable = Nothing
    Private tblEDT810O5 As DataTable = Nothing
    Private tblEDTSYSIH As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing
    Private rowARTPARM1 As DataRow = Nothing
    Private rowEDTPARM1 As DataRow = Nothing

    ' Added 04/20/2018
    Private tblARTSTAX1 As DataTable = Nothing
    Private rowGLTPARM1 As DataRow = Nothing

    Private COMPANY_CODE As String = String.Empty
    Private Const EDI_PROCESS_IND As String = "1"

    Private EDI_OUTBOUND_DOC_NO As String = String.Empty

    Private tblSOTSVIA1 As DataTable = Nothing
    Private tblTATTERM1 As DataTable = Nothing
    Private tblEDTTRPM1 As DataTable = Nothing

    Private Factor810Code As String = String.Empty

    ' Charming is requesting changes to their invoice.
    ' We need to roll up EDT810O2 into a line matching their original PO line – instead of the SLN breakdown.
    ' 6/13/2016
    Private lstSlnRollUpCustomer As New List(Of String)(New String() {"LANEBRY", "CATHERINE"})

    ' Customers having 850 cartons explode to Eaches
    Private lstCartonPackRollUpCustomer As New List(Of String)(New String() {"LOBLAW"})

    Public Sub New(ByRef datasetIn As DataSet)

        InitializeData(datasetIn.Tables("EDTSYSIH"), _
                       datasetIn.Tables("EDT810O1"), _
                       datasetIn.Tables("EDT810O2"), _
                       datasetIn.Tables("EDT810O3"), _
                       datasetIn.Tables("EDT810O4"), _
                       datasetIn.Tables("EDT810O5"))

    End Sub

    ''' <summary>
    ''' Creates the EDI 810 entry for a Shipment
    ''' </summary>
    ''' <param name="tblEDTSYSIHin">Reference to table EDTSYSIH</param>
    ''' <param name="tblEDT810O1in">Reference to table EDT810O1</param>
    ''' <param name="tblEDT810O2in">Reference to table EDT81002</param>
    ''' <param name="tblEDT810O3in">Reference to table EDT81003</param>
    ''' <param name="tblEDT810O5in">Reference to table EDT81005</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef tblEDTSYSIHin As DataTable, _
                   ByRef tblEDT810O1in As DataTable, _
                   ByRef tblEDT810O2in As DataTable, _
                   ByRef tblEDT810O3in As DataTable, _
                   ByRef tblEDT810O4in As DataTable, _
                   ByRef tblEDT810O5in As DataTable)


        InitializeData(tblEDTSYSIHin, tblEDT810O1in, tblEDT810O2in, tblEDT810O3in, tblEDT810O4in, tblEDT810O5in)
    End Sub

    Private Sub InitializeData(ByRef tblEDTSYSIHin As DataTable, _
                   ByRef tblEDT810O1in As DataTable, _
                   ByRef tblEDT810O2in As DataTable, _
                   ByRef tblEDT810O3in As DataTable, _
                   ByRef tblEDT810O4in As DataTable, _
                   ByRef tblEDT810O5in As DataTable)

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            COMPANY_CODE = "VAN"
        ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            COMPANY_CODE = "NYA"
        ElseIf ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            COMPANY_CODE = "RGI"
        Else
            COMPANY_CODE = ASCMAIN1.CLIENT
        End If

        tblEDTSYSIH = tblEDTSYSIHin
        tblEDT810O1 = tblEDT810O1in
        tblEDT810O2 = tblEDT810O2in
        tblEDT810O3 = tblEDT810O3in
        tblEDT810O4 = tblEDT810O4in
        tblEDT810O5 = tblEDT810O5in

        EDI_OUTBOUND_DOC_NO = String.Empty
        tblSOTSVIA1 = ASCDATA1.GetDataTable("SELECT * FROM SOTSVIA1", "SOTSVIA1", String.Empty, Nothing)
        tblTATTERM1 = ASCDATA1.GetDataTable("SELECT * FROM TATTERM1", "TATTERM1", String.Empty, Nothing)
        tblEDTTRPM1 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM1 where EDI_DOC_NO = '810'", "EDTTRPM1", String.Empty, Nothing)
        tblEDTSLSP1 = ASCDATA1.GetDataTable("SELECT * FROM EDTSLSP1")
        tblEDTSLSP1.PrimaryKey = New DataColumn() {tblEDTSLSP1.Columns("CUST_CODE")}

        ' Added 04/20/2018
        tblARTSTAX1 = ASCDATA1.GetDataTable("SELECT * FROM ARTSTAX1", "ARTSTAX1", String.Empty, Nothing)
        rowGLTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM1 WHERE GL_PARM_KEY = 'Z'")

        rowARTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM ARTPARM1 WHERE AR_PARM_KEY = 'Z'")
        rowEDTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM EDTPARM1 WHERE ED_PARM_KEY = 'Z'")

        If rowEDTPARM1 IsNot Nothing AndAlso rowEDTPARM1.Table.Columns.Contains("ED_PARM_FACTOR") Then
            Factor810Code = rowEDTPARM1.Item("ED_PARM_FACTOR") & String.Empty
        End If

        If Not tblEDT810O2.Columns.Contains("EDI_DTL_SEQ_850") Then
            tblEDT810O2.Columns.Add("EDI_DTL_SEQ_850", GetType(System.Int32))
        End If
        If Not tblEDT810O2.Columns.Contains("EDI_DOC_SEQ_NO_850") Then
            tblEDT810O2.Columns.Add("EDI_DOC_SEQ_NO_850", GetType(System.String))
        End If

    End Sub

    Public Sub Create810(ByVal SHIP_BOL_NO As String, ByRef EDI_OUTBOUND_DOC_NO As String, ByVal EDI_BATCH_NO As String, Optional Factor810 As Boolean = False)

        Dim rowEDT810O1 As DataRow = Nothing
        Dim rowEDT810O2 As DataRow = Nothing
        Dim rowEDT810O3 As DataRow = Nothing
        Dim rowEDT810O5 As DataRow = Nothing

        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim rowSOTORDR0 As DataRow = Nothing
        Dim rowSOTORDR1 As DataRow = Nothing
        Dim rowSOTORDR2 As DataRow = Nothing
        Dim rowEDT850T5 As DataRow = Nothing
        Dim rowSOTSVIA1 As DataRow = Nothing
        Dim rowTATTERM1 As DataRow = Nothing
        Dim rowEDTTRPM1 As DataRow = Nothing
        Dim rowSOTINVH1 As DataRow = Nothing
        Dim rowARTOPEN1 As DataRow = Nothing
        Dim rowICTWHSE1 As DataRow = Nothing

        Dim tblSOTINVH2 As DataTable = Nothing
        Dim tblSOTORDR2 As DataTable = Nothing
        Dim tblSOTORDR5 As DataTable = Nothing
        Dim tblSOTCART1 As DataTable = Nothing
        Dim tblSOTSHIP1 As DataTable = Nothing
        Dim tblEDTPPKS1 As DataTable = Nothing
        Dim tblARTCUST2 As DataTable = Nothing

        Dim STYLE_CODE As String = String.Empty
        Dim PICK_NO As String = String.Empty
        Dim INV_NO As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim CUST_CODE As String = String.Empty

        ' Added 04/20/2018
        Dim STAX_CODE As String = String.Empty
        Dim STAX_TYPE As String = String.Empty
        Dim STATE_CODE As String = String.Empty
        Dim STAX_RATE As Decimal = 0
        Dim CURR_CODE As String = String.Empty
        Dim CURR_EXT As String = String.Empty

        Dim EDI_DOC_SEQ_NO As String = String.Empty
        Dim EDI_DTL_SEQ As Int16 = 0
        Dim EDI_SLN_SEQ As Int16 = 0

        Dim EDI_BUYER_ITEM As String = String.Empty
        Dim sql As String = String.Empty
        Dim BillOfLading As String = String.Empty
        Dim methodType As String = String.Empty

        Dim rowEDTSLSP1 As DataRow = Nothing
        Dim rowEDT850T1 As DataRow = Nothing
        Dim numNonZeroDollarInvoices As Int16 = 0
        Dim caseFactor As Int16 = 1

        Dim tblEDT850T6 As DataTable = New DataTable
        Dim rowEDT850T6 As DataRow = Nothing

        Dim tblEDT850T7 As DataTable = New DataTable

        ' defualt to no discount
        ' Use boolen incase 1 as a double = .9999999999999
        Dim useMiscDiscountMultiplier As Boolean = False
        Dim miscDiscountMultiplier As Double = 1

        If Not Factor810 Then
            tblSOTSHIP1 = ASCDATA1.GetDataTable("select * from SOTSHIP1 where SHIP_BOL_NO = :PARM1 and (SHIP_810_BATCH_NO IS NULL OR SHIP_810_BATCH_NO = '')", "SOTSHIP1", "V", New Object() {SHIP_BOL_NO})
        Else
            tblSOTSHIP1 = ASCDATA1.GetDataTable("select * from SOTSHIP1 where SHIP_BOL_NO = :PARM1 and (FACTOR_TRANS_BATCH_LAST IS NULL OR FACTOR_TRANS_BATCH_LAST = '')", "SOTSHIP1", "V", New Object() {SHIP_BOL_NO})
        End If
        If tblSOTSHIP1.Rows.Count = 0 Then
            Exit Sub
        End If

        'If tblSOTSHIP1.Rows(0).Item("BILL_OF_LADING_NO") & String.Empty <> String.Empty Then
        '    tblSOTSHIP1 = ASCDATA1.GetDataTable("select * from SOTSHIP1 where BILL_OF_LADING_NO = :PARM1", "SOTSHIP1", "V", New Object() {tblSOTSHIP1.Rows(0).Item("BILL_OF_LADING_NO") & String.Empty})
        'End If

        If tblSOTSHIP1.Rows(0).Item("BILL_OF_LADING_NO") & String.Empty <> String.Empty Then
            BillOfLading = tblSOTSHIP1.Rows(0).Item("BILL_OF_LADING_NO") & String.Empty
            BillOfLading = BillOfLading.PadLeft(10, "0")
            BillOfLading = StrReverse(StrReverse(BillOfLading).Substring(0, 10))
        Else
            BillOfLading = "9" & SHIP_BOL_NO.Substring(1)
        End If

        Dim ORDR_GROUP_NO As String = tblSOTSHIP1.Rows(0).Item("ORDR_GROUP_NO") & String.Empty
        If ORDR_GROUP_NO.Length = 0 Then
            Exit Sub
        End If

        sql = "Select * From SOTORDR0 Where ORDR_GROUP_NO = :PARM1"
        rowSOTORDR0 = ASCDATA1.GetDataRow(sql, "V", ORDR_GROUP_NO)
        If rowSOTORDR0 Is Nothing Then
            Exit Sub
        End If

        rowICTWHSE1 = ASCDATA1.GetDataRow("SELECT * FROM ICTWHSE1 WHERE WHSE_CODE = :PARM1", "V", New Object() {tblSOTSHIP1.Rows(0).Item("WHSE_CODE") & String.Empty})
        If rowICTWHSE1 Is Nothing Then
            Exit Sub
        End If

        methodType = "A"
        If Not Factor810 AndAlso rowSOTORDR0 IsNot Nothing AndAlso rowSOTORDR0.Item("EDI_DOC_SEQ_NO") & String.Empty <> String.Empty Then
            EDI_DOC_SEQ_NO = rowSOTORDR0.Item("EDI_DOC_SEQ_NO") & String.Empty
            If EDI_DOC_SEQ_NO.Length > 0 Then
                tblEDT850T6 = ASCDATA1.GetDataTable("SELECT * FROM EDT850T6 WHERE EDI_DOC_SEQ_NO = :PARM1", "", "V", New Object() {EDI_DOC_SEQ_NO})
                tblEDT850T6.Columns.Add("EXT_PRICE", GetType(System.Decimal), "EDI_SLN_PRICE * EDI_PO4_QTY")
                If tblEDT850T6.Rows.Count > 0 Then
                    methodType = "B"
                End If
            End If
        End If


        CUST_CODE = rowSOTORDR0.Item("CUST_CODE") & String.Empty
        Dim rowARTCUST1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
        If rowARTCUST1 Is Nothing Then
            Exit Sub
        End If

        Dim CUST_CONS_INV As Boolean = rowARTCUST1.Item("CUST_CONS_INV") & String.Empty = "1"
        Dim SlnRollUpCustomer As Boolean = Not Factor810 AndAlso lstSlnRollUpCustomer.Contains(CUST_CODE)
        Dim CartonPackRollUpCustomer As Boolean = Not Factor810 AndAlso lstCartonPackRollUpCustomer.Contains(CUST_CODE)

        ' Ed,
        ' This is a bit complicated. 
        ' For WALMART   - when we send the 810, we should send it for the full amount of the invoice.  Robin confirmed that this is what she did on the AS400.
        ' For ROSENTHAL – when we send the 810, we will hard code a 2.5% reduction to the total invoice amount.  So an invoice for $100 of product will go to Rosenthal as $97.50.  This is probably to reduce the complication of Rosenthal recording a short payment.
        ' Walter
        ' Walmart's Rosnethal 810 get reduced by 2.5%


        ' Changed in March 2016 to use Discount Rates in the 850 record
        ' Discount will be on the Invoice Header 
        If CUST_CODE = "WALMART" Then
            'miscDiscountMultiplier = 0.025
            useMiscDiscountMultiplier = True
            'tblEDT850T7 = ASCDATA1.GetDataTable("SELECT * FROM EDT850T7 WHERE EDI_DOC_SEQ_NO = :PARM1", "", "V", New Object() {EDI_DOC_SEQ_NO})
            'miscDiscountMultiplier = Math.Round(Val(tblEDT850T7.Compute("SUM(SAH_PERCENT)", "") & String.Empty) / 100, 2)
        End If

        If Factor810 Then
            If Factor810Code.Length = 0 Then
                Exit Sub
            End If
        End If

        rowEDTSLSP1 = tblEDTSLSP1.Rows.Find(IIf(Factor810, Factor810Code, CUST_CODE))
        Dim EDI_ID_810 As String = rowEDTSLSP1.Item("EDI_ID_810") & String.Empty
        Dim EDI_QUAL_810 As String = rowEDTSLSP1.Item("EDI_QUAL_810") & String.Empty

        sql = "EDI_TP_QUAL = '" & EDI_QUAL_810 & "' and EDI_TP_ID = '" & EDI_ID_810 & "' and EDI_DOC_NO = '810'"

        If tblEDTTRPM1.Select(sql).Length = 0 Then
            Exit Sub
        Else
            rowEDTTRPM1 = tblEDTTRPM1.Select(sql)(0)
        End If

        ' Set based on EDTTRPM1
        Dim EDI_OUR_ID As String = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
        Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty

        Dim SHIP_ADDR_TYPE As String = tblSOTSHIP1.Rows(0).Item("SHIP_ADDR_TYPE") & String.Empty
        Dim SHIP_ADDR_CODE As String = tblSOTSHIP1.Rows(0).Item("SHIP_ADDR_CODE") & String.Empty
        Dim mkNumChars As Int16 = 0
        Dim dcNumChars As Int16 = 0

        mkNumChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
        dcNumChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_DC") & String.Empty)

        If rowEDTSLSP1 IsNot Nothing Then
            Select Case SHIP_ADDR_TYPE
                Case "MK"
                    If mkNumChars > 0 AndAlso IsNumeric(SHIP_ADDR_CODE) Then
                        SHIP_ADDR_CODE = SHIP_ADDR_CODE.PadLeft(mkNumChars, "0")
                        SHIP_ADDR_CODE = StrReverse(StrReverse(SHIP_ADDR_CODE).Substring(0, mkNumChars))
                    End If

                Case "DC"
                    If dcNumChars > 0 AndAlso IsNumeric(SHIP_ADDR_CODE) Then
                        SHIP_ADDR_CODE = SHIP_ADDR_CODE.PadLeft(dcNumChars, "0")
                        SHIP_ADDR_CODE = StrReverse(StrReverse(SHIP_ADDR_CODE).Substring(0, dcNumChars))
                    End If
            End Select
        End If


        '************************************************************************************
        ' This is where there should be a check for consolidated invoices
        ' If so, call another sub procedure.
        '************************************************************************************
        'EDT810O2.EDI_ITEM_DESC from SOTORDR2.STYLE_DESC

        ' set a default incase all invoices are $0.00

        EDI_OUTBOUND_DOC_NO = String.Empty '"xxx"
        tblARTCUST2 = ASCDATA1.GetDataTable("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1", "ARTCUST2", "V", New Object() {CUST_CODE})

        For Each rowSOTSHIP1 In tblSOTSHIP1.Rows

            If CUST_CONS_INV Then
                ' Get the Pick Ticket for the Lead Consolidated Invoice
                sql = "Select SOTPICK1.*"
                sql &= " from SOTPICK1, SOTSHIP1, SOTINVH1"
                sql &= " where SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
                sql &= " and SOTPICK1.INV_NO = SOTINVH1.INV_NO AND SOTINVH1.INV_TYPE = 'I'"
                sql &= " and SOTPICK1.SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'"
                sql &= " and SOTINVH1.INV_NO = SOTINVH1.INV_NO_CONS"
            Else
                sql = "SELECT * FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "' AND PICK_STATUS = 'F'"
            End If

            ' Style Code Sub
            For Each rowSOTPICK1 As DataRow In ASCDATA1.GetDataTable(sql).Select("", "PICK_NO")

                PICK_NO = rowSOTPICK1.Item("PICK_NO") & String.Empty
                INV_NO = rowSOTPICK1.Item("INV_NO") & String.Empty
                ORDR_NO = rowSOTPICK1.Item("ORDR_NO") & String.Empty

                Dim tblEDT810O4_DATA As DataTable = Nothing
                If ASCMAIN1.CLIENT = "RGI" Then
                    ASCMAIN1.sql = "select sotcart1.cart_tracking_no, sotpick2.pick_lno, sotcart2.*" _
                            & " from sotcart1, sotcart2, sotpick1, sotpick2" _
                            & " where sotcart1.cart_no = sotcart2.cart_no" _
                            & " and sotcart1.pick_no = sotpick1.pick_no" _
                            & " and sotpick1.pick_no = sotpick2.pick_no" _
                            & " and sotpick2.ordr_no = sotcart2.ordr_no" _
                            & " and sotpick2.ordr_lno = sotcart2.ordr_lno" _
                            & " and sotpick1.inv_no = :PARM1"

                    tblEDT810O4_DATA = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {INV_NO})
                End If

                rowSOTINVH1 = ASCDATA1.GetDataRow("select * from SOTINVH1 where INV_TYPE = :PARM1 AND INV_NO = :PARM2", "VV", New Object() {"I", INV_NO})

                ' Added 04/20/2018
                STAX_RATE = Val(rowSOTINVH1.Item("STAX_RATE") & String.Empty)
                CURR_CODE = rowSOTINVH1.Item("CURR_CODE") & String.Empty
                If CURR_CODE <> rowGLTPARM1.Item("GL_PARM_CURR_CODE") & String.Empty AndAlso CURR_CODE <> String.Empty Then
                    CURR_EXT = "_CURR"
                End If

                STAX_CODE = rowSOTINVH1.Item("STAX_CODE") & String.Empty
                STAX_TYPE = String.Empty
                STATE_CODE = String.Empty

                If tblARTSTAX1.Select("STAX_CODE = '" & STAX_CODE & "'").Length > 0 Then
                    STAX_TYPE = tblARTSTAX1.Select("STAX_CODE = '" & STAX_CODE & "'")(0).Item("STAX_TYPE") & String.Empty
                    STATE_CODE = tblARTSTAX1.Select("STAX_CODE = '" & STAX_CODE & "'")(0).Item("STATE_CODE") & String.Empty
                End If

                If rowSOTINVH1 Is Nothing Then
                    rowARTOPEN1 = Nothing
                Else
                    rowARTOPEN1 = ASCDATA1.GetDataRow("select * from ARTOPEN1 where CUST_CODE = :PARM1 AND INV_NUM = :PARM2", "VV", New Object() {rowSOTINVH1.Item("CUST_CODE"), INV_NO})
                End If

                sql = "Select * from SOTINVH2 where INV_TYPE = :PARM1 AND INV_NO = :PARM2"
                tblSOTINVH2 = ASCDATA1.GetDataTable(sql, "SOTINVH2", "VV", New Object() {"I", INV_NO})

                If Not CUST_CONS_INV Then
                    rowSOTORDR1 = ASCDATA1.GetDataRow("select * from SOTORDR1 where ordr_no = :PARM1", "V", New Object() {ORDR_NO})
                    EDI_DOC_SEQ_NO = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & String.Empty

                    sql = " SELECT SOTORDR2.*, EDT850T2.*"
                    sql &= " FROM SOTORDR2, EDT850T2 "
                    sql &= " WHERE SOTORDR2.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO (+)"
                    sql &= " AND SOTORDR2.EDI_DTL_SEQ = EDT850T2.EDI_DTL_SEQ (+)"
                    sql &= " AND ORDR_NO = :PARM1"
                    tblSOTORDR2 = ASCDATA1.GetDataTable(sql, "SOTORDR2", "V", New Object() {ORDR_NO})

                    tblSOTORDR5 = ASCDATA1.GetDataTable("select * from SOTORDR5 where ordr_no = :PARM1", "SOTORDR5", "V", New Object() {ORDR_NO})
                    tblSOTCART1 = ASCDATA1.GetDataTable("select * from SOTCART1 where PICK_NO = :PARM1", "SOTCART1", "V", New Object() {PICK_NO})
                Else
                    Dim clsSOTINVH1 As New TAC.SOCINVH1

                    If Not clsSOTINVH1.CreateConsolidatedInvoice(INV_NO, rowSOTINVH1, tblSOTINVH2, methodType = "B") Then
                        EDI_OUTBOUND_DOC_NO = String.Empty
                        Exit Sub
                    End If

                    rowSOTORDR1 = ASCDATA1.GetDataRow("Select * from SOTORDR1 where ordr_no = :PARM1", "V", New Object() {ORDR_NO})
                    EDI_DOC_SEQ_NO = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & String.Empty

                    sql = " SELECT SOTORDR2.*, EDT850T2.*"
                    sql &= " FROM SOTORDR2, EDT850T2 "
                    sql &= " WHERE SOTORDR2.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO (+)"
                    sql &= " AND SOTORDR2.EDI_DTL_SEQ = EDT850T2.EDI_DTL_SEQ (+)"
                    sql &= " AND (SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO) IN"
                    sql &= " ("
                    sql &= " SELECT STYLE_CODE, COLOR_CODE, ORDR_NO, MAX(ORDR_LNO) ORDR_LNO"
                    sql &= " FROM SOTORDR2"
                    sql &= " WHERE (STYLE_CODE, COLOR_CODE, ORDR_NO) IN"
                    sql &= " ("
                    sql &= " SELECT SOTORDR2.STYLE_CODE,  SOTORDR2.COLOR_CODE, MAX(SOTORDR2.ORDR_NO) ORDR_NO"
                    sql &= " FROM SOTORDR1, SOTORDR2, SOTINVH1"
                    sql &= " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
                    sql &= " AND SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO"
                    sql &= " AND SOTINVH1.INV_NO_CONS = :PARM1"
                    sql &= " GROUP BY SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
                    sql &= " )"
                    sql &= " GROUP BY STYLE_CODE, SOTORDR2.COLOR_CODE, ORDR_NO"
                    sql &= " )"

                    tblSOTORDR2 = ASCDATA1.GetDataTable(sql, "SOTORDR2", "V", New Object() {INV_NO})
                    tblSOTORDR5 = ASCDATA1.GetDataTable("select * from SOTORDR5 where ordr_no = :PARM1", "SOTORDR5", "V", New Object() {ORDR_NO})

                    sql = " Select SOTCART1.* "
                    sql &= " from SOTCART1, SOTPICK1, SOTINVH1"
                    sql &= " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO"
                    sql &= " and SOTPICK1.INV_NO = SOTINVH1.INV_NO AND SOTINVH1.INV_TYPE = 'I'"
                    sql &= " and SOTINVH1.INV_NO_CONS = :PARM1"
                    tblSOTCART1 = ASCDATA1.GetDataTable(sql, "SOTCART1", "V", New Object() {INV_NO})
                End If

                ' SLN
                If methodType = "B" AndAlso Not CUST_CONS_INV Then

                    Dim invNos As String = String.Empty
                    For Each row As DataRow In ASCDATA1.SelectDistinct(tblSOTINVH2, New String() {"INV_NO"}).Rows
                        invNos &= ", '" & row.Item("INV_NO") & "'"
                    Next

                    invNos = invNos.Substring(1).Trim

                    sql = "Select SOTINVH2.*, SOTINVH1.ORDR_NO, SOTORDR2.EDI_DOC_SEQ_NO, SOTORDR2.EDI_DTL_SEQ, SOTORDR2.STYLE_UOM"
                    sql &= " from SOTINVH1, SOTINVH2, SOTORDR2"
                    sql &= " where SOTINVH1.inv_type = SOTINVH2.inv_type"
                    sql &= " and SOTINVH1.inv_no = SOTINVH2.inv_no"
                    sql &= " and sotordr2.ordr_lno = sotinvh2.inv_lno"
                    sql &= " and sotordr2.ordr_no = sotinvh1.ordr_no"
                    sql &= " and SOTINVH2.Inv_no in (" & invNos & ")"
                    Dim tblDetails As DataTable = ASCDATA1.GetDataTable(sql, "SOTINVH2")
                    tblDetails.Columns.Add("EXTENDED", GetType(System.Double), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY_SHIP,0)")

                    Dim ediData As String = String.Empty
                    For Each row As DataRow In ASCDATA1.SelectDistinct(tblDetails, New String() {"ORDR_NO", "EDI_DOC_SEQ_NO", "EDI_DTL_SEQ"}).Rows
                        Dim ORDR_NOx As String = (row.Item("ORDR_NO") & String.Empty).ToString.Trim
                        Dim EDI_DOC_SEQ_NOx As String = (row.Item("EDI_DOC_SEQ_NO") & String.Empty).ToString.Trim
                        Dim EDI_DTL_SEQx As Int16 = Val(row.Item("EDI_DTL_SEQ") & String.Empty)

                        If ORDR_NOx.Length > 0 AndAlso EDI_DOC_SEQ_NOx.Length > 0 AndAlso EDI_DTL_SEQx > 0 Then
                            ediData &= ", ('" & ORDR_NOx & "', '" & EDI_DOC_SEQ_NOx & "', " & EDI_DTL_SEQx & ")"
                        End If
                    Next

                    If ediData.Length > 0 Then
                        ediData = ediData.Substring(1).Trim
                    End If

                    sql = " SELECT SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ"
                    'sql &= " , SOTCART2.CART_NO"
                    sql &= " , NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM) EDI_PRICE_UOM"
                    sql &= " , ROUND(SUM(SOTCART2.QTY_PACKED) / SUM(EDT850T6.EDI_SLN_QTY), 0) QTY"
                    sql &= " , ROUND(SUM(SOTORDR2.ORDR_QTY_ORIG) / SUM(EDT850T6.EDI_SLN_QTY), 0) QTYO"
                    sql &= " FROM SOTORDR2, EDT850T6, EDT850T2, SOTCART2"
                    sql &= " where EDT850T6.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO AND EDT850T6.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ"
                    sql &= " AND EDT850T6.EDI_SLN_SEQ  = SOTORDR2.EDI_SLN_SEQ"
                    sql &= " AND EDT850T2.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO AND EDT850T2.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ"
                    sql &= " AND SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO AND SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO"
                    sql &= " AND (SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ) IN "
                    sql &= " (" & ediData & ")"
                    'sql &= " GROUP BY SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ, SOTCART2.CART_NO, NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM)"
                    sql &= " GROUP BY SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ, NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM)"
                    Dim tblcartons As DataTable = ASCDATA1.GetDataTable(sql)

                    For Each row As DataRow In tblcartons.Select("", "ORDR_NO, EDI_DOC_SEQ_NO, EDI_DTL_SEQ")
                        Dim ORDR_NOx As String = (row.Item("ORDR_NO") & String.Empty).ToString.Trim
                        Dim EDI_DOC_SEQ_NOx As String = (row.Item("EDI_DOC_SEQ_NO") & String.Empty).ToString.Trim
                        Dim EDI_DTL_SEQx As Int16 = Val(row.Item("EDI_DTL_SEQ") & String.Empty)

                        sql = "ORDR_NO = '" & ORDR_NOx & "' and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NOx & "' and EDI_DTL_SEQ = " & EDI_DTL_SEQx
                        sql = "ORDR_NO = '" & ORDR_NO & "' and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_DTL_SEQ = " & EDI_DTL_SEQ
                        Dim ordrUnitPrice As Double = Val(tblDetails.Compute("SUM(EXTENDED)", sql) & String.Empty) / Val(row.Item("QTY") & String.Empty)

                        ' Set Order Qty Shipped to 0, left over lines will be deleted later
                        For Each rowDetails As DataRow In tblDetails.Select(sql, "EDI_DTL_SEQ")
                            rowDetails.Item("ORDR_QTY_SHIP") = 0
                        Next

                        ' Apply to only one Line item
                        For Each rowDetails As DataRow In tblDetails.Select(sql, "EDI_DTL_SEQ")
                            rowDetails.Item("ORDR_QTY_SHIP") = row.Item("QTY")
                            rowDetails.Item("STYLE_UOM") = row.Item("EDI_PRICE_UOM")
                            rowDetails.Item("ORDR_UNIT_PRICE") = ordrUnitPrice
                            rowDetails.Item("ORDR_UNIT_PRICE_CURR") = ordrUnitPrice
                            Exit For
                        Next
                    Next

                    ' Delete Ship = 0 records.
                    For Each rowDetails As DataRow In tblDetails.Select("ORDR_QTY_SHIP = 0")
                        rowDetails.Delete()
                    Next

                    tblDetails.AcceptChanges()

                    ' replace the values in tblSOTINVH1
                    ' Get a list of column names in common
                    Dim colList As New List(Of String)
                    For Each col As DataColumn In tblDetails.Columns
                        If tblSOTINVH2.Columns.Contains(col.ColumnName) Then
                            colList.Add(col.ColumnName)
                        End If
                    Next

                    tblSOTINVH2.Rows.Clear()
                    tblSOTINVH2.AcceptChanges()

                    For Each row As DataRow In tblDetails.Select("")
                        Dim rowSOTINVH2 As DataRow = tblSOTINVH2.NewRow
                        For Each field As String In colList
                            rowSOTINVH2.Item(field) = row.Item(field)
                        Next
                        tblSOTINVH2.Rows.Add(rowSOTINVH2)
                    Next

                End If

                ' No $0.00 invoices need to get sent over
                If Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty) = 0 Then
                    numNonZeroDollarInvoices += 1
                    Continue For
                End If

                CUST_CODE = rowSOTINVH1.Item("CUST_CODE") & String.Empty

                rowSOTSVIA1 = tblSOTSVIA1.Rows.Find(rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty)
                rowTATTERM1 = tblTATTERM1.Rows.Find(rowSOTINVH1.Item("TERM_CODE") & String.Empty)

                ' Get data where Cases are exploded to Eaches. Need to put back as cases in 856
                sql = " Select EDTPPKS1.*, ICTSTYL1.CARTON_PACK_QTY CARTON_PACK_QTY_STYLE"
                sql &= "  from EDTPPKS1,ICTSTYL1"
                sql &= "  where ICTSTYL1.STYLE_CODE (+) = EDTPPKS1.STYLE_CODE"
                sql &= "  and EDTPPKS1.CUST_CODE = :PARM1"
                tblEDTPPKS1 = ASCDATA1.GetDataTable(sql, "EDTPPKS1", "V", New Object() {CUST_CODE})

                rowEDT810O1 = tblEDT810O1.NewRow

                EDI_OUTBOUND_DOC_NO = Me.CreateEDTSYSIH(EDI_OUR_ID, EDI_TP_ID, "IN", rowEDTTRPM1.Item("EDI_STATUS") & String.Empty)

                rowEDT810O1.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDT810O1.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT810O1.Item("EDI_INVOICE_DATE") = CDate(rowSOTINVH1.Item("INV_DATE") & String.Empty).ToString("MM/dd/yyyy")
                rowEDT810O1.Item("EDI_INVOICE_NUMBER") = rowSOTINVH1.Item("INV_NO")
                rowEDT810O1.Item("EDI_PO_DATE") = CDate(rowSOTORDR1.Item("ORDR_DATE") & String.Empty).ToString("MM/dd/yyyy")
                rowEDT810O1.Item("EDI_PO_NO") = rowSOTINVH1.Item("ORDR_CUST_PO")
                rowEDT810O1.Item("EDI_DEPT_NO") = rowSOTINVH1.Item("ORDR_DEPT")
                rowEDT810O1.Item("EDI_ORDER_NO") = rowSOTINVH1.Item("ORDR_NO")
                rowEDT810O1.Item("EDI_TERM_CODE") = rowSOTINVH1.Item("TERM_CODE")
                rowEDT810O1.Item("EDI_BATCH_NO") = EDI_BATCH_NO

                ' added 01/11/2019
                If rowEDT810O1.Table.Columns.Contains("SEG4_CODE") Then
                    If rowARTCUST1.Item("SEG4_CODE") & String.Empty <> String.Empty Then
                        rowEDT810O1.Item("SEG4_CODE") = rowARTCUST1.Item("SEG4_CODE")
                    Else
                        rowEDT810O1.Item("SEG4_CODE") = rowGLTPARM1.Item("GL_PARM_DEF_SEG4")
                    End If
                End If

                ' Added 04/20/2018
                rowEDT810O1.Item("CURR_CODE") = CURR_CODE

                rowEDT850T1 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                If rowEDT850T1 IsNot Nothing Then
                    rowEDT810O1.Item("EDI_SUPPLIER_NO") = rowEDT850T1.Item("EDI_SUPPLIER_NO")
                    rowEDT810O1.Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE")

                    ' EDT810O1.EDI_PO_RELEASE_NO = EDT850T1.EDI_PO_RELEASE_NO
                    ' Ricks request 06/12/2018
                    If rowEDT810O1.Table.Columns.Contains("EDI_PO_RELEASE_NO") AndAlso rowEDT850T1.Table.Columns.Contains("EDI_PO_RELEASE_NO") Then
                        rowEDT810O1.Item("EDI_PO_RELEASE_NO") = rowEDT850T1.Item("EDI_PO_RELEASE_NO")
                    End If
                Else
                    If tblEDTTRPM1.Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                        rowEDT810O1.Item("EDI_SUPPLIER_NO") = tblEDTTRPM1.Select("CUST_CODE = '" & CUST_CODE & "'")(0).Item("EDI_ACCT_REF_NO") & String.Empty
                    End If
                End If

                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & String.Empty
                If mkNumChars > 0 AndAlso IsNumeric(CUST_STORE_NO) Then
                    CUST_STORE_NO = CUST_STORE_NO.PadLeft(mkNumChars, "0")
                    CUST_STORE_NO = StrReverse(StrReverse(CUST_STORE_NO).Substring(0, mkNumChars))
                End If

                If rowARTOPEN1 IsNot Nothing Then
                    rowEDT810O1.Item("EDI_TERMS_DISC_DATE") = rowARTOPEN1.Item("INV_DISC_DATE")
                End If

                rowEDT810O1.Item("EDI_BILL_TO") = CUST_STORE_NO
                rowEDT810O1.Item("EDI_SHIP_TO") = SHIP_ADDR_CODE
                rowEDT810O1.Item("EDI_MARK_FOR") = CUST_STORE_NO

                rowEDT810O1.Item("EDI_REMIT_TO_NAME") = rowARTPARM1.Item("AR_PARM_REMIT_NAME") & String.Empty
                rowEDT810O1.Item("EDI_REMIT_TO_ID") = (rowARTPARM1.Item("AR_PARM_DUNS_NO") & String.Empty).ToString.Replace("-", "").Replace(" ", "")

                ' Modified 04/20/2018
                rowEDT810O1.Item("EDI_TOTAL_INV_AMT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT" & CURR_EXT)
                ' March 2016 - Walmart Discount will be in the Invoice Header, which is done above.
                ' EDI Total must be Invoice Total Amount and Inv Misc Charge.
                ' As per Maria on 4/8 - use the INV_TOTAL_AMOUNT on the Walmart 810 without the Misc Charge Added back in.
                'If useMiscDiscountMultiplier = True AndAlso Not Factor810 Then
                '    rowEDT810O1.Item("EDI_TOTAL_INV_AMT") = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty) - Val(rowSOTINVH1.Item("INV_MISC_CHG") & String.Empty)
                'End If

                rowEDT810O1.Item("EDI_BL_NO") = BillOfLading ' SHIP_BOL_NO
                rowEDT810O1.Item("EDI_FRT_TERMS") = IIf(rowSOTSHIP1.Item("FRT_TERMS") & String.Empty = "COL", "CC", "PP")
                'ALTER TABLE EDT810O1 ADD EDI_TERMS_DESC VARCHAR2(35);

                If rowSOTSVIA1 IsNot Nothing Then
                    rowEDT810O1.Item("EDI_ROUTING") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    rowEDT810O1.Item("EDI_SCAC_CODE") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty
                    rowEDT810O1.Item("EDI_CARRIER_MODE") = rowSOTSVIA1.Item("CARRIER_MODE") & String.Empty
                Else
                    rowEDT810O1.Item("EDI_ROUTING") = "Unknown Shipper"
                    rowEDT810O1.Item("EDI_SCAC_CODE") = String.Empty
                    rowEDT810O1.Item("EDI_CARRIER_MODE") = "M"
                End If

                If ASCMAIN1.CLIENT = "RGI" Then
                    If CUST_CODE = "316567" OrElse (rowSOTORDR1 IsNot Nothing AndAlso rowSOTORDR1.Item("ECOM_CODE") & String.Empty = "HOMEDEPOT") Then
                        If rowEDT850T1("EDI_SHIPPER") & String.Empty <> String.Empty Then
                            rowEDT810O1.Item("EDI_ROUTING") = rowEDT850T1("EDI_SHIPPER")
                        End If
                    End If
                End If

                If methodType = "B" Then
                    ' Calculated down below.
                    rowEDT810O1.Item("EDI_TOTAL_UNITS") = 0 ' Val(tblSOTINVH2.Compute("SUM(ORDR_QTY_SHIP)", "") & String.Empty)
                Else
                    rowEDT810O1.Item("EDI_TOTAL_UNITS") = Val(tblSOTCART1.Compute("SUM(CART_TOTAL_UNITS)", "") & String.Empty)
                End If

                rowEDT810O1.Item("EDI_WEIGHT") = Val(tblSOTCART1.Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "") & String.Empty)
                rowEDT810O1.Item("EDI_SHIP_REF") = rowSOTSHIP1.Item("SHIP_REF") & String.Empty  'tblSOTCART1.Select("PICK_NO = '" & PICK_NO & "'")(0).Item("CART_TRACKING_NO") & String.Empty

                If rowTATTERM1 IsNot Nothing Then
                    rowEDT810O1.Item("EDI_TERMS_NET_DAYS") = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & String.Empty)
                    rowEDT810O1.Item("EDI_TERMS_DISC_DAYS_DUE") = Val(rowTATTERM1.Item("TERM_DAYS_DISC") & String.Empty)
                    rowEDT810O1.Item("EDI_TERMS_DISC_PCT") = Val(rowTATTERM1.Item("TERM_DISC_PERC") & String.Empty)
                    rowEDT810O1.Item("EDI_TERMS_DESC") = rowTATTERM1.Item("TERM_DESC") & String.Empty

                    'Dim INV_DUE_DATE As Date
                    'Dim DISC_DUE_DATE As Date
                    ' CALC_DUE_DATE(rowTATTERM1, rowSOTINVH1.Item("INV_DATE"), INV_DUE_DATE, DISC_DUE_DATE)

                    Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Nothing, rowSOTINVH1.Item("TERM_CODE") & String.Empty, rowTATTERM1, rowSOTINVH1.Item("INV_DATE"))
                    rowEDT810O1.Item("EDI_TERMS_DUE_DATE") = INV_DUE_DATE
                Else
                    rowEDT810O1.Item("EDI_TERMS_NET_DAYS") = 30 ' DEFAULT
                    rowEDT810O1.Item("EDI_TERMS_DISC_DAYS_DUE") = 30
                    rowEDT810O1.Item("EDI_TERMS_DISC_PCT") = 0
                    rowEDT810O1.Item("EDI_TERMS_DESC") = String.Empty
                    rowEDT810O1.Item("EDI_TERMS_DUE_DATE") = rowSOTINVH1.Item("INV_DATE")
                End If

                rowEDT810O1.Item("INIT_DATE") = DateTime.Now
                rowEDT810O1.Item("INIT_OPER") = ASCMAIN1.USER_ID

                rowEDT810O1.Item("EDI_CARTON_CT") = tblSOTCART1.Rows.Count

                ' Added 04/20/2018
                If tblEDT810O1.Columns.Contains("EDI_STATE_CODE") Then
                    rowEDT810O1.Item("EDI_STATE_CODE") = STATE_CODE
                End If
                If tblEDT810O1.Columns.Contains("EDI_STAX_TYPE") Then
                    rowEDT810O1.Item("EDI_STAX_TYPE") = STAX_TYPE
                End If

                tblEDT810O1.Rows.Add(rowEDT810O1)

                For Each rowSOTINVH2 As DataRow In tblSOTINVH2.Select("", "STYLE_CODE,COLOR_CODE")
                    EDI_BUYER_ITEM = String.Empty
                    EDI_DOC_SEQ_NO = String.Empty
                    EDI_DTL_SEQ = 0
                    EDI_SLN_SEQ = 0
                    rowSOTORDR2 = Nothing

                    caseFactor = 1
                    Dim STYLE_CODEX As String = rowSOTINVH2.Item("STYLE_CODE") & String.Empty
                    Dim rowEDTPPKS1 As DataRow = Nothing
                    If tblEDTPPKS1.Select("CUST_CODE = '" & CUST_CODE & "' and STYLE_CODE = '" & STYLE_CODEX & "'").Length > 0 Then
                        rowEDTPPKS1 = tblEDTPPKS1.Select("CUST_CODE = '" & CUST_CODE & "' and STYLE_CODE = '" & STYLE_CODEX & "'")(0)
                    End If
                    If rowEDTPPKS1 IsNot Nothing Then
                        caseFactor = Val(rowEDTPPKS1.Item("CARTON_PACK_QTY") & "")
                        If caseFactor = 0 Then
                            caseFactor = Val(rowEDTPPKS1.Item("CARTON_PACK_QTY_STYLE") & "")
                        End If
                        If caseFactor = 0 Then
                            caseFactor = 1
                        End If
                    End If

                    rowEDT810O2 = tblEDT810O2.NewRow
                    rowEDT810O2.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT810O2.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    rowEDT810O2.Item("EDI_DOC_LNO") = rowSOTINVH2.Item("INV_LNO")
                    rowEDT810O2.Item("EDI_QTY_INVOICED") = Convert.ToInt32(rowSOTINVH2.Item("ORDR_QTY_SHIP") / caseFactor)
                    If methodType = "B" Then
                        rowEDT810O1.Item("EDI_TOTAL_UNITS") = Val(rowEDT810O1.Item("EDI_TOTAL_UNITS") & String.Empty) + rowEDT810O2.Item("EDI_QTY_INVOICED")
                    End If

                    If caseFactor = 1 Then
                        ' Modified 04/20/2018
                        rowEDT810O2.Item("EDI_UNIT_PRICE") = rowSOTINVH2.Item("ORDR_UNIT_PRICE" & CURR_EXT)
                    Else
                        ' done incase 25.50 ends up loking like 24.49999
                        ' Modified 04/20/2018
                        rowEDT810O2.Item("EDI_UNIT_PRICE") = Math.Round(rowSOTINVH2.Item("ORDR_UNIT_PRICE" & CURR_EXT) * caseFactor, 2)
                    End If

                    STYLE_CODE = rowSOTINVH2.Item("STYLE_CODE") & String.Empty
                    Dim COLOR_CODE As String = rowSOTINVH2.Item("COLOR_CODE") & String.Empty

                    ' Regency Wayfair Set Quantity
                    Dim SET_QTY As Int16 = 1

                    ' Just incase one day there is an item on the invoice that is not on the order
                    ' I know it will Never ever ever happen - Yeah right!!
                    If tblSOTORDR2.Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'").Length > 0 Then
                        rowSOTORDR2 = tblSOTORDR2.Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'")(0)
                        EDI_DOC_SEQ_NO = rowSOTORDR2.Item("EDI_DOC_SEQ_NO") & String.Empty
                        EDI_DTL_SEQ = Val(rowSOTORDR2.Item("EDI_DTL_SEQ") & String.Empty)
                        EDI_SLN_SEQ = Val(rowSOTORDR2.Item("EDI_SLN_SEQ") & String.Empty)

                        ' Regency Wayfair Set Quantity
                        If ASCMAIN1.CLIENT = "RGI" Then
                            SET_QTY = Val(rowSOTORDR2.Item("SET_QTY") & String.Empty)
                        End If

                        If EDI_DOC_SEQ_NO.Length = 0 OrElse EDI_DTL_SEQ = 0 Then
                            If Not Factor810 Then
                                Continue For
                            End If
                        End If
                    End If

                    ' Regency Wayfair Set Quantity
                    If SET_QTY > 1 AndAlso ASCMAIN1.CLIENT = "RGI" Then
                        rowEDT810O2.Item("EDI_UNIT_PRICE") = Val(rowEDT810O2.Item("EDI_UNIT_PRICE") & String.Empty) * SET_QTY
                        rowEDT810O2.Item("EDI_QTY_INVOICED") = Val(rowEDT810O2.Item("EDI_QTY_INVOICED") & String.Empty) / SET_QTY
                    End If

                    ' Added 06/07/2018
                    If ASCMAIN1.CLIENT = "NYA" AndAlso CartonPackRollUpCustomer Then
                        Dim rowEDT850T2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM EDT850T2 WHERE EDI_DOC_SEQ_NO = :PARM1 AND EDI_DTL_SEQ = :PARM1", "VN", New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ})
                        If rowEDT850T2 IsNot Nothing Then
                            If rowSOTORDR2 IsNot Nothing Then
                                Dim EDI_PRICE_UOM As String = rowEDT850T2.Item("EDI_PRICE_UOM") & String.Empty
                                Dim STYLE_UOM As String = rowSOTORDR2.Item("STYLE_UOM") & String.Empty
                                Dim CARTON_PACK_QTY As Int16 = Val(rowSOTORDR2.Item("CARTON_PACK_QTY") & String.Empty)
                                'If (EDI_PRICE_UOM = "CA" OrElse EDI_PRICE_UOM = "CS") AndAlso STYLE_UOM = "EA" AndAlso CARTON_PACK_QTY > 0 Then
                                ' Changed on 6/21/2018 - As per Walter and Maria since Loblaws may send other EDI_PRICE_UOM in the EDT850T2
                                If EDI_PRICE_UOM <> "EA" AndAlso CARTON_PACK_QTY > 0 Then
                                    rowEDT810O2.Item("EDI_UNIT_PRICE") = Val(rowEDT810O2.Item("EDI_UNIT_PRICE") & String.Empty) * CARTON_PACK_QTY
                                    rowEDT810O2.Item("EDI_QTY_INVOICED") = CInt(Val(rowEDT810O2.Item("EDI_QTY_INVOICED") & String.Empty) / CARTON_PACK_QTY)
                                End If
                            End If
                        End If
                    End If

                    If rowSOTORDR2 IsNot Nothing AndAlso (rowSOTORDR2.Item("EDI_PRICE_UOM") & String.Empty).ToString.Trim.Length > 0 Then
                        rowEDT810O2.Item("EDI_UOM") = (rowSOTORDR2.Item("EDI_PRICE_UOM") & String.Empty).ToString.Trim
                    ElseIf rowSOTORDR2 IsNot Nothing AndAlso (rowSOTORDR2.Item("EDI_PO4_UOM") & String.Empty).ToString.Trim.Length > 0 Then
                        rowEDT810O2.Item("EDI_UOM") = (rowSOTORDR2.Item("EDI_PO4_UOM") & String.Empty).ToString.Trim
                    Else
                        rowEDT810O2.Item("EDI_UOM") = "EA"
                    End If

                    rowEDT810O2.Item("EDI_BUYER_STYLE") = (rowSOTORDR2.Item("EDI_STYLE") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_ITEM_UP") = (rowSOTORDR2.Item("EDI_UPC") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_ITEM_EN") = (rowSOTORDR2.Item("EDI_EAN") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_ITEM_GTIN") = (rowSOTORDR2.Item("EDI_GTIN") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_BUYER_ITEM") = (rowSOTORDR2.Item("EDI_SKU") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_ITEM_DESC") = (rowSOTORDR2.Item("STYLE_DESC") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_SELLER_ITEM") = rowSOTORDR2.Item("STYLE_CODE") & String.Empty
                    rowEDT810O2.Item("EDI_PO4_UOM") = (rowSOTORDR2.Item("EDI_PO4_UOM") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_PO4_QTY") = Val(rowSOTORDR2.Item("EDI_PO4_QTY") & String.Empty)
                    If Val(rowSOTORDR2.Item("EDI_PO4_QTY") & String.Empty) > 1 AndAlso rowSOTORDR2.Item("EDI_PO4_INNER") & String.Empty = String.Empty Then
                        rowEDT810O2.Item("EDI_PO4_INNER") = 1
                    ElseIf Val(rowSOTORDR2.Item("EDI_PO4_INNER") & String.Empty) > 0 Then
                        rowEDT810O2.Item("EDI_PO4_INNER") = Val(rowSOTORDR2.Item("EDI_PO4_INNER") & String.Empty)
                    End If
                    'EDT850T2.EDI_PO_LNO  >> EDT810O2.EDI_PO_LNO
                    rowEDT810O2.Item("EDI_PO_LNO") = (rowSOTORDR2.Item("EDI_PO_LNO") & String.Empty).ToString.Trim

                    'EDI_DTL_SEQ_850 - added 6/13/2016
                    rowEDT810O2.Item("EDI_DTL_SEQ_850") = rowSOTORDR2.Item("EDI_DTL_SEQ")
                    rowEDT810O2.Item("EDI_DOC_SEQ_NO_850") = rowSOTORDR2.Item("EDI_DOC_SEQ_NO")

                    ' Added 04/20/2018
                    If tblEDT810O2.Columns.Contains("EDI_LN_STAX_AMT") Then
                        If STAX_RATE > 0 Then
                            rowEDT810O2.Item("EDI_LN_STAX_AMT") = Math.Round(Val(rowEDT810O2.Item("EDI_UNIT_PRICE") & String.Empty) * Val(rowEDT810O2.Item("EDI_QTY_INVOICED") & String.Empty) * (STAX_RATE / 100), 2)
                        End If
                    End If

                    'Instead of adding a new table EDT810O6, Walt agreed we can just hardcode this for cust_code CHARMING
                    'EDT810O2.EDI_BUYER_ITEM = EDT850T6.EDI_SLN_STYLE
                    'EDT810O2.EDI_ITEM_DESC = EDT850T6.EDI_SLN_ITEM_DESC
                    'EDT810O2.EDI_UOM = EDT850T6.EDI_SLN_UOM

                    If CUST_CODE = "CHARMING" Then
                        If tblEDT850T6.Rows.Count > 0 Then
                            rowEDT850T6 = tblEDT850T6.Rows.Find(New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ, EDI_SLN_SEQ})
                            If rowEDT850T6 IsNot Nothing Then
                                rowEDT810O2.Item("EDI_BUYER_ITEM") = rowEDT850T6.Item("EDI_SLN_STYLE")
                                rowEDT810O2.Item("EDI_ITEM_DESC") = rowEDT850T6.Item("EDI_SLN_ITEM_DESC")
                                rowEDT810O2.Item("EDI_UOM") = rowEDT850T6.Item("EDI_SLN_UOM")
                            End If
                        End If
                    End If

                    ' As per Maria and Walter on 10/23/2019
                    If Not (ASCMAIN1.CLIENT = "NYA" AndAlso CUST_CODE = "DOLTREE") Then
                        rowEDT810O2.Item("EDI_UNIT_PRICE") = Math.Round(Val(rowEDT810O2.Item("EDI_UNIT_PRICE") & String.Empty), 2)
                    End If

                    tblEDT810O2.Rows.Add(rowEDT810O2)

                    If ASCMAIN1.CLIENT = "RGI" Then
                        For Each rowEDT810O4_DATA As DataRow In tblEDT810O4_DATA.Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "' and PICK_LNO = " & rowEDT810O2.Item("EDI_DOC_LNO"))
                            Dim rowEDT810O4 As DataRow = tblEDT810O4.NewRow
                            rowEDT810O4.Item("COMPANY_CODE") = rowEDT810O2.Item("COMPANY_CODE")
                            rowEDT810O4.Item("EDI_OUTBOUND_DOC_NO") = rowEDT810O2.Item("EDI_OUTBOUND_DOC_NO")
                            rowEDT810O4.Item("EDI_DOC_LNO") = rowEDT810O2.Item("EDI_DOC_LNO")
                            rowEDT810O4.Item("EDI_CART_NO") = rowEDT810O4_DATA.Item("CART_NO")
                            rowEDT810O4.Item("EDI_CART_TRACKING_NO") = rowEDT810O4_DATA.Item("CART_TRACKING_NO")
                            tblEDT810O4.Rows.Add(rowEDT810O4)
                        Next
                    End If
                Next

                Dim EDI_SAC_LNO As Int16 = 0
                ' Modified 04/20/2018
                Dim charge As Decimal = Val(rowSOTINVH1.Item("INV_FREIGHT" & CURR_EXT) & String.Empty)
                If charge <> 0 Then
                    rowEDT810O3 = tblEDT810O3.NewRow
                    rowEDT810O3.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT810O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    EDI_SAC_LNO += 1
                    rowEDT810O3.Item("EDI_SAC_LNO") = EDI_SAC_LNO
                    rowEDT810O3.Item("EDI_CHG_ALL_IND") = IIf(charge >= 0, "C", "A")
                    rowEDT810O3.Item("EDI_CHG_ALL_CODE") = "D240"
                    rowEDT810O3.Item("EDI_SAC_AMOUNT") = Math.Abs(charge)
                    rowEDT810O3.Item("EDI_SAC_DESC") = "FREIGHT"
                    tblEDT810O3.Rows.Add(rowEDT810O3)
                End If

                ' Added 04/20/2018
                charge = Val(rowSOTINVH1.Item("INV_STAX" & CURR_EXT) & String.Empty)

                If STAX_CODE.Length > 0 AndAlso charge > 0 Then
                    rowEDT810O3 = tblEDT810O3.NewRow
                    rowEDT810O3.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT810O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    EDI_SAC_LNO += 1
                    rowEDT810O3.Item("EDI_SAC_LNO") = EDI_SAC_LNO
                    rowEDT810O3.Item("EDI_CHG_ALL_IND") = IIf(charge >= 0, "C", "A")
                    rowEDT810O3.Item("EDI_SAC_AMOUNT") = Math.Abs(charge)
                    rowEDT810O3.Item("EDI_SAC_DESC") = "TAX"

                    Select Case STAX_TYPE
                        Case "GST", "HST"
                            rowEDT810O3.Item("EDI_CHG_ALL_CODE") = "D360"
                            tblEDT810O3.Rows.Add(rowEDT810O3)

                        Case "QST", "PST"
                            Throw New Exception("Stax Type " & STAX_TYPE & " not supported. Call ABS!")

                        Case Else
                            ' Do nothing at this time. the entry in EDT810O3 will not be created

                    End Select

                End If

                ' At this time do not make entry in 810 for Misc Charges
                ' March 2016 - Make entry in EDT810T3
                ' Modified 04/20/2018
                charge = Val(rowSOTINVH1.Item("INV_MISC_CHG" & CURR_EXT) & String.Empty)
                If useMiscDiscountMultiplier AndAlso Not Factor810 Then
                    For Each rowSOTINVHM As DataRow In ASCDATA1.GetDataTable("SELECT * FROM SOTINVHM WHERE INV_TYPE = :PARM1 AND INV_NO = :PARM2", "SOTINVHM", "VV", {"I", INV_NO}).Rows
                        charge = Val(rowSOTINVHM.Item("INV_MISC_CHG") & String.Empty)

                        If charge <> 0 Then
                            rowEDT810O3 = tblEDT810O3.NewRow
                            rowEDT810O3.Item("COMPANY_CODE") = COMPANY_CODE
                            rowEDT810O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            EDI_SAC_LNO += 1
                            rowEDT810O3.Item("EDI_SAC_LNO") = EDI_SAC_LNO
                            rowEDT810O3.Item("EDI_CHG_ALL_IND") = IIf(charge >= 0, "C", "A")

                            rowEDT810O3.Item("EDI_CHG_ALL_CODE") = rowSOTINVHM.Item("MISC_CHG_NOTE")
                            rowEDT810O3.Item("EDI_SAC_AMOUNT") = Math.Abs(charge)
                            rowEDT810O3.Item("EDI_SAC_DESC") = "MISCCHG"
                            tblEDT810O3.Rows.Add(rowEDT810O3)
                        End If

                    Next
                End If

                ' Ed - When you get to creating EDT810O5 and EDT856O5 rows, I am thinking maybe we should take the rows from EDT850T5 original PO
                ' Factored 810s will use SOTORDR5 records
                ' Also Non-EDI factored invoices will use SOTORDR5 since there will not be any EDT850T5 records
                Dim tblEDT850T5 As DataTable = Nothing
                If Not Factor810 Then
                    tblEDT850T5 = ASCDATA1.GetDataTable("SELECT * FROM EDT850T5 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                End If

                If tblEDT850T5 Is Nothing OrElse tblEDT850T5.Rows.Count = 0 OrElse Factor810 Then
                    Dim hasBTrecords As Boolean = False

                    For Each rowSOTORDR5 In tblSOTORDR5.Select("")

                        If rowSOTORDR5.Item("CUST_ADDR_TYPE") & String.Empty = "BT" Then
                            hasBTrecords = True
                        End If

                        rowEDT810O5 = tblEDT810O5.NewRow
                        rowEDT810O5.Item("COMPANY_CODE") = COMPANY_CODE
                        rowEDT810O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        rowEDT810O5.Item("EDI_ADDR_TYPE") = rowSOTORDR5.Item("CUST_ADDR_TYPE")
                        rowEDT810O5.Item("EDI_NAME") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS3") = String.Empty
                        rowEDT810O5.Item("EDI_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                        rowEDT810O5.Item("EDI_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty

                        rowEDT810O5.Item("EDI_ZIPCODE") = (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                        If (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                            rowEDT810O5.Item("EDI_ZIPCODE") = (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Substring(0, rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                        End If

                        rowEDT810O5.Item("EDI_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty

                        Dim EDI_ADDR_CODE As String = IIf(rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT", rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty, String.Empty)
                        If rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT" Then
                            ' Bill's rule on 5/10/2013
                            If Factor810 AndAlso COMPANY_CODE = "NYA" Then
                                EDI_ADDR_CODE = CUST_CODE
                            ElseIf mkNumChars > 0 AndAlso IsNumeric(EDI_ADDR_CODE) Then
                                EDI_ADDR_CODE = EDI_ADDR_CODE.PadLeft(mkNumChars, "0")
                                EDI_ADDR_CODE = StrReverse(StrReverse(EDI_ADDR_CODE).Substring(0, mkNumChars))
                            End If
                        Else
                            EDI_ADDR_CODE = SHIP_ADDR_CODE
                        End If

                        rowEDT810O5.Item("EDI_ADDR_CODE") = EDI_ADDR_CODE
                        'rowEDT810O5.Item("EDI_ADDR_CODE_QUAL") = IIf(rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT", "91", "92")
                        tblEDT810O5.Rows.Add(rowEDT810O5)
                    Next

                    If Not hasBTrecords Then
                        'Dim rowSOTORDR5 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", CUST_CODE)
                        rowEDT810O5 = tblEDT810O5.NewRow
                        rowEDT810O5.Item("COMPANY_CODE") = COMPANY_CODE
                        rowEDT810O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        rowEDT810O5.Item("EDI_ADDR_TYPE") = "BT"
                        rowEDT810O5.Item("EDI_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS1") = rowARTCUST1.Item("CUST_ADDR1") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS2") = rowARTCUST1.Item("CUST_ADDR2") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS3") = rowARTCUST1.Item("CUST_ADDR3") & String.Empty
                        rowEDT810O5.Item("EDI_CITY") = rowARTCUST1.Item("CUST_CITY") & String.Empty
                        rowEDT810O5.Item("EDI_STATE") = rowARTCUST1.Item("CUST_STATE") & String.Empty

                        rowEDT810O5.Item("EDI_ZIPCODE") = (rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                        rowEDT810O5.Item("EDI_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty

                        Dim EDI_ADDR_CODE = rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty
                        ' Bill's rule on 5/10/2013
                        If Factor810 AndAlso COMPANY_CODE = "NYA" Then
                            EDI_ADDR_CODE = CUST_CODE
                        ElseIf mkNumChars > 0 AndAlso IsNumeric(EDI_ADDR_CODE) Then
                            EDI_ADDR_CODE = EDI_ADDR_CODE.PadLeft(mkNumChars, "0")
                            EDI_ADDR_CODE = StrReverse(StrReverse(EDI_ADDR_CODE).Substring(0, mkNumChars))
                        End If

                        rowEDT810O5.Item("EDI_ADDR_CODE") = EDI_ADDR_CODE
                        'rowEDT810O5.Item("EDI_ADDR_CODE_QUAL") = IIf(rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT", "91", "92")
                        tblEDT810O5.Rows.Add(rowEDT810O5)
                    End If
                Else
                    Dim EDI_ADDR_TYPE As List(Of String) = New List(Of String)
                    Dim hasSTrecords As Boolean = False
                    Dim EDI_ADDR_CODE_ST As String = String.Empty

                    ' FAMILY DOLLAR-TARHEEL work around
                    ' Tarheel sends multiple ST addresses in EDT850T5. Walt had to break it up into individual orders from WH indication at the SLN
                    If tblEDT850T5.Select("EDI_ADDR_TYPE = 'ST'").Length > 1 AndAlso tblSOTORDR5.Select("CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                        ' Need to use CUST_ADDR_CODE to get to the ship to
                        EDI_ADDR_CODE_ST = tblSOTORDR5.Select("CUST_ADDR_TYPE = 'ST'")(0).Item("CUST_ADDR_CODE") & String.Empty
                        If mkNumChars > 0 AndAlso IsNumeric(EDI_ADDR_CODE_ST) Then
                            EDI_ADDR_CODE_ST = EDI_ADDR_CODE_ST.PadLeft(mkNumChars, "0")
                            EDI_ADDR_CODE_ST = StrReverse(StrReverse(EDI_ADDR_CODE_ST).Substring(0, mkNumChars))
                        End If
                    End If

                    For Each rowEDT850T5 In tblEDT850T5.Select("", "EDI_ADDR_TYPE, EDI_ADR_SEQ")

                        If (rowEDT850T5.Item("EDI_ADDR_TYPE") & String.Empty).ToString.Trim.Length = 0 Then
                            Continue For
                        End If

                        If EDI_ADDR_TYPE.Contains(rowEDT850T5.Item("EDI_ADDR_TYPE") & String.Empty) Then
                            Continue For
                        Else
                            EDI_ADDR_TYPE.Add(rowEDT850T5.Item("EDI_ADDR_TYPE") & String.Empty)
                        End If

                        If rowEDT850T5.Item("EDI_ADDR_TYPE") & String.Empty = "ST" Then
                            hasSTrecords = True
                            If EDI_ADDR_CODE_ST.Length > 0 Then
                                ' FAMILY DOLLAR-TARHEEL work around
                                ' Tarheel sends multiple ST addresses in EDT850T5. Walt had to break it up into individual orders from WH indication at the SLN
                                If tblEDT850T5.Select("EDI_ADDR_TYPE = 'ST' AND EDI_ADDR_CODE = '" & EDI_ADDR_CODE_ST & "'").Length > 0 Then
                                    rowEDT850T5 = tblEDT850T5.Select("EDI_ADDR_TYPE = 'ST' AND EDI_ADDR_CODE = '" & EDI_ADDR_CODE_ST & "'")(0)
                                End If
                            End If

                            'If rowEDT850T5.Item("EDI_CUST_NAME_ADR") & String.Empty = String.Empty Then
                            If rowEDT850T5.Item("EDI_CUST_NAME_ADR") & String.Empty = String.Empty OrElse rowEDT850T5.Item("EDI_ADDRESS1") & String.Empty = String.Empty Then
                                If tblSOTORDR5.Select("CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                                    Dim rowSOTORDR5 As DataRow = tblSOTORDR5.Select("CUST_ADDR_TYPE = 'ST'")(0)
                                    rowEDT850T5.Item("EDI_CUST_NAME_ADR") = rowSOTORDR5.Item("CUST_NAME")
                                    rowEDT850T5.Item("EDI_ADDRESS1") = rowSOTORDR5.Item("CUST_ADDR1")
                                    rowEDT850T5.Item("EDI_ADDRESS2") = rowSOTORDR5.Item("CUST_ADDR2")
                                    rowEDT850T5.Item("EDI_CITY") = rowSOTORDR5.Item("CUST_CITY")
                                    rowEDT850T5.Item("EDI_STATE") = rowSOTORDR5.Item("CUST_STATE")
                                    rowEDT850T5.Item("EDI_ZIPCODE") = rowSOTORDR5.Item("CUST_ZIP_CODE")
                                    rowEDT850T5.Item("EDI_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY")
                                End If
                            End If
                        End If

                        rowEDT810O5 = tblEDT810O5.NewRow
                        rowEDT810O5.Item("COMPANY_CODE") = COMPANY_CODE
                        rowEDT810O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        rowEDT810O5.Item("EDI_ADDR_TYPE") = rowEDT850T5.Item("EDI_ADDR_TYPE")
                        rowEDT810O5.Item("EDI_NAME") = rowEDT850T5.Item("EDI_CUST_NAME_ADR") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS1") = rowEDT850T5.Item("EDI_ADDRESS1") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS2") = rowEDT850T5.Item("EDI_ADDRESS2") & String.Empty
                        rowEDT810O5.Item("EDI_ADDRESS3") = String.Empty
                        rowEDT810O5.Item("EDI_CITY") = rowEDT850T5.Item("EDI_CITY") & String.Empty
                        rowEDT810O5.Item("EDI_STATE") = rowEDT850T5.Item("EDI_STATE") & String.Empty

                        rowEDT810O5.Item("EDI_ZIPCODE") = (rowEDT850T5.Item("EDI_ZIPCODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                        If (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                            rowEDT810O5.Item("EDI_ZIPCODE") = (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Substring(0, rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                        End If

                        rowEDT810O5.Item("EDI_COUNTRY") = rowEDT850T5.Item("EDI_COUNTRY") & String.Empty
                        rowEDT810O5.Item("EDI_ADDR_CODE") = rowEDT850T5.Item("EDI_ADDR_CODE") & String.Empty
                        rowEDT810O5.Item("EDI_ADDR_CODE_QUAL") = rowEDT850T5.Item("EDI_ADDR_CODE_QUAL") & String.Empty
                        tblEDT810O5.Rows.Add(rowEDT810O5)
                    Next

                    If Not hasSTrecords Then
                        Dim EDI_ADDR_CODE = rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty
                        Dim rowARTCUST2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_ADDR_TYPE = :PARM2 AND CUST_ADDR_CODE = :PARM3", _
                                                                            "VVV", New Object() {CUST_CODE, SHIP_ADDR_TYPE, EDI_ADDR_CODE})
                        If rowARTCUST2 IsNot Nothing Then
                            rowEDT810O5 = tblEDT810O5.NewRow
                            rowEDT810O5.Item("COMPANY_CODE") = COMPANY_CODE
                            rowEDT810O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            rowEDT810O5.Item("EDI_ADDR_TYPE") = "ST"
                            rowEDT810O5.Item("EDI_NAME") = rowARTCUST2.Item("CUST_NAME") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS1") = rowARTCUST2.Item("CUST_ADDR1") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS2") = rowARTCUST2.Item("CUST_ADDR1") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS3") = String.Empty
                            rowEDT810O5.Item("EDI_CITY") = rowARTCUST2.Item("CUST_CITY") & String.Empty
                            rowEDT810O5.Item("EDI_STATE") = rowARTCUST2.Item("CUST_STATE") & String.Empty

                            rowEDT810O5.Item("EDI_ZIPCODE") = (rowARTCUST2.Item("CUST_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                            If (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                                rowEDT810O5.Item("EDI_ZIPCODE") = (rowARTCUST2.Item("EDI_ZIPCODE") & String.Empty).ToString.Substring(0, rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                            End If
                            rowEDT810O5.Item("EDI_COUNTRY") = rowARTCUST2.Item("CUST_COUNTRY") & String.Empty

                            If Factor810 AndAlso COMPANY_CODE = "NYA" Then
                                EDI_ADDR_CODE = rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty
                            ElseIf rowARTCUST2 IsNot Nothing AndAlso rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty <> String.Empty Then
                                EDI_ADDR_CODE = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty
                            Else
                                If mkNumChars > 0 AndAlso IsNumeric(EDI_ADDR_CODE) Then
                                    EDI_ADDR_CODE = EDI_ADDR_CODE.PadLeft(mkNumChars, "0")
                                    EDI_ADDR_CODE = StrReverse(StrReverse(EDI_ADDR_CODE).Substring(0, mkNumChars))
                                End If
                            End If

                            rowEDT810O5.Item("EDI_ADDR_CODE") = EDI_ADDR_CODE 'rowARTCUST2.Item("EDI_ADDR_CODE") & String.Empty
                            'rowEDT810O5.Item("EDI_ADDR_CODE_QUAL") = String.Empty 'rowARTCUST2.Item("EDI_ADDR_CODE_QUAL") & String.Empty
                            tblEDT810O5.Rows.Add(rowEDT810O5)
                        End If
                    End If
                End If

                ' Added 5/23/2019
                If ASCMAIN1.CLIENT = "NYA" AndAlso Not Factor810 Then
                    If tblEDT810O5.Select("EDI_ADDR_TYPE = 'SF' and COMPANY_CODE = '" & COMPANY_CODE & "' and EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'").Length = 0 Then
                        rowEDT850T5 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T5 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND EDI_ADDR_TYPE = 'SF'")
                        If rowEDT850T5 IsNot Nothing Then
                            rowEDT810O5 = tblEDT810O5.NewRow
                            rowEDT810O5.Item("COMPANY_CODE") = COMPANY_CODE
                            rowEDT810O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            rowEDT810O5.Item("EDI_ADDR_TYPE") = rowEDT850T5.Item("EDI_ADDR_TYPE")
                            rowEDT810O5.Item("EDI_NAME") = rowEDT850T5.Item("EDI_CUST_NAME_ADR") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS1") = rowEDT850T5.Item("EDI_ADDRESS1") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS2") = rowEDT850T5.Item("EDI_ADDRESS2") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS3") = rowEDT850T5.Item("EDI_ADDRESS3") & String.Empty
                            rowEDT810O5.Item("EDI_CITY") = rowEDT850T5.Item("EDI_CITY") & String.Empty
                            rowEDT810O5.Item("EDI_STATE") = rowEDT850T5.Item("EDI_STATE") & String.Empty

                            rowEDT810O5.Item("EDI_ZIPCODE") = (rowEDT850T5.Item("EDI_ZIPCODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                            If (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                                rowEDT810O5.Item("EDI_ZIPCODE") = (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Substring(0, rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                            End If

                            rowEDT810O5.Item("EDI_COUNTRY") = rowEDT850T5.Item("EDI_COUNTRY") & String.Empty
                            rowEDT810O5.Item("EDI_ADDR_CODE") = rowEDT850T5.Item("EDI_ADDR_CODE") & String.Empty
                            rowEDT810O5.Item("EDI_ADDR_CODE_QUAL") = rowEDT850T5.Item("EDI_ADDR_CODE_QUAL") & String.Empty
                            tblEDT810O5.Rows.Add(rowEDT810O5)
                        Else
                            rowEDT810O5 = tblEDT810O5.NewRow
                            rowEDT810O5.Item("COMPANY_CODE") = COMPANY_CODE
                            rowEDT810O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            rowEDT810O5.Item("EDI_ADDR_TYPE") = "SF"
                            rowEDT810O5.Item("EDI_NAME") = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS1") = rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS2") = rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty
                            rowEDT810O5.Item("EDI_ADDRESS3") = rowICTWHSE1.Item("WHSE_ADDR3") & String.Empty
                            rowEDT810O5.Item("EDI_CITY") = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
                            rowEDT810O5.Item("EDI_STATE") = rowICTWHSE1.Item("WHSE_STATE") & String.Empty

                            rowEDT810O5.Item("EDI_ZIPCODE") = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                            If (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                                rowEDT810O5.Item("EDI_ZIPCODE") = (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Substring(0, rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                            End If

                            rowEDT810O5.Item("EDI_COUNTRY") = rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty
                            rowEDT810O5.Item("EDI_ADDR_CODE") = rowICTWHSE1.Item("WHSE_EDI_ID") & String.Empty
                            'rowEDT810O5.Item("EDI_ADDR_CODE_QUAL") = rowEDT850T5.Item("EDI_ADDR_CODE_QUAL") & String.Empty
                            tblEDT810O5.Rows.Add(rowEDT810O5)
                        End If
                    End If
                End If
            Next
        Next

        For Each rowEDT810O5 In tblEDT810O5.Select("EDI_ADDR_TYPE = 'Z7' AND ISNULL(EDI_NAME, '') = ''")
            Dim EDI_ADDR_CODE As String = rowEDT810O5.Item("EDI_ADDR_CODE")
            EDI_ADDR_CODE = EDI_ADDR_CODE.PadLeft(6, "0")
            sql = "CUST_ADDR_TYPE = 'MK'AND CUST_ADDR_CODE = '" & EDI_ADDR_CODE & "'"
            If tblARTCUST2.Select(sql).Length = 1 Then
                Dim rowARTCUST2 As DataRow = tblARTCUST2.Select(sql)(0)
                rowEDT810O5.Item("EDI_NAME") = rowARTCUST2.Item("CUST_NAME") & String.Empty
                rowEDT810O5.Item("EDI_ADDRESS1") = rowARTCUST2.Item("CUST_ADDR1") & String.Empty
                rowEDT810O5.Item("EDI_ADDRESS2") = rowARTCUST2.Item("CUST_ADDR1") & String.Empty
                rowEDT810O5.Item("EDI_ADDRESS3") = String.Empty
                rowEDT810O5.Item("EDI_CITY") = rowARTCUST2.Item("CUST_CITY") & String.Empty
                rowEDT810O5.Item("EDI_STATE") = rowARTCUST2.Item("CUST_STATE") & String.Empty

                rowEDT810O5.Item("EDI_ZIPCODE") = (rowARTCUST2.Item("CUST_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                If (rowEDT810O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                    rowEDT810O5.Item("EDI_ZIPCODE") = (rowARTCUST2.Item("EDI_ZIPCODE") & String.Empty).ToString.Substring(0, rowEDT810O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                End If
                rowEDT810O5.Item("EDI_COUNTRY") = rowARTCUST2.Item("CUST_COUNTRY") & String.Empty
            End If
        Next

        ' Added 06/07/2018
        If ASCMAIN1.CLIENT = "NYA" AndAlso CartonPackRollUpCustomer Then
            For Each rowEDT810O1x As DataRow In tblEDT810O1.Select("")
                Dim EDI_OUTBOUND_DOC_NOx As String = rowEDT810O1x.Item("EDI_OUTBOUND_DOC_NO")
                rowEDT810O1.Item("EDI_TOTAL_UNITS") = Val(tblEDT810O2.Compute("SUM(EDI_QTY_INVOICED)", "EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NOx & "'") & String.Empty)
            Next
        End If


        If numNonZeroDollarInvoices > 0 AndAlso EDI_OUTBOUND_DOC_NO.Length = 0 Then
            EDI_OUTBOUND_DOC_NO = "Zero"
        End If

        ' New code 6/13/2016
        If SlnRollUpCustomer Then
            Dim tbl As DataTable = ASCDATA1.SelectDistinct(tblEDT810O2, New String() {"EDI_OUTBOUND_DOC_NO", "EDI_DTL_SEQ_850"})
            For Each row As DataRow In tbl.Select("EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'", "EDI_DTL_SEQ_850")
                Dim EDI_DTL_SEQ_850 As Int16 = Val(row.Item("EDI_DTL_SEQ_850") & String.Empty)

                If EDI_DTL_SEQ_850 = 0 Then
                    Continue For
                End If

                Dim query As String = "EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "' and EDI_DTL_SEQ_850 = " & EDI_DTL_SEQ_850

                If tblEDT810O2.Select(query).Length = 0 Then
                    Continue For
                End If

                Dim EDI_UNIT_PRICE As Decimal = Val(tblEDT850T6.Compute("SUM(EXT_PRICE)", "EDI_DTL_SEQ = " & EDI_DTL_SEQ_850) & String.Empty)
                Dim EDI_SLN_QTY As Int32 = Val(tblEDT850T6.Compute("SUM(EDI_SLN_QTY)", "EDI_DTL_SEQ = " & EDI_DTL_SEQ_850) & String.Empty)
                Dim EDI_QTY_INVOICED As Int32 = Val(tblEDT810O2.Compute("SUM(EDI_QTY_INVOICED)", query) & String.Empty)

                If EDI_SLN_QTY = 0 Then
                    Exit Sub
                End If

                Dim firstline As Boolean = True
                For Each rowEDT810O2 In tblEDT810O2.Select(query, "EDI_DOC_LNO")
                    If firstline Then
                        rowEDT810O2.Item("EDI_DOC_LNO") = EDI_DTL_SEQ_850
                        rowEDT810O2.Item("EDI_UNIT_PRICE") = EDI_UNIT_PRICE
                        rowEDT810O2.Item("EDI_QTY_INVOICED") = Math.Round(EDI_QTY_INVOICED / EDI_SLN_QTY, 0, MidpointRounding.AwayFromZero)
                        firstline = False
                    Else
                        rowEDT810O2.Delete()
                    End If
                Next
            Next
        End If
    End Sub

    ''' <summary>
    ''' Creates a record for table EDTSYSIH
    ''' </summary>
    ''' <param name="EDI_OUR_ID"></param>
    ''' <param name="EDI_TP_ID"></param>
    ''' <returns>The key field EDI_OUTBOUND_DOC_NO for table EDTSYSIH</returns>
    ''' <remarks></remarks>
    Private Function CreateEDTSYSIH(ByVal EDI_OUR_ID As String, ByVal EDI_TP_ID As String, ByVal ediApplicationID As String, ByVal EDI_STATUS As String) As String

        CreateEDTSYSIH = String.Empty

        Dim ediOutboundDocNo As String = String.Empty
        ' Moved from up above
        ediOutboundDocNo = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")

        Dim rowEDTSYSIH As DataRow = tblEDTSYSIH.NewRow
        rowEDTSYSIH.Item("COMPANY_CODE") = COMPANY_CODE
        rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO") = ediOutboundDocNo
        rowEDTSYSIH.Item("EDI_APPLICATION_ID") = ediApplicationID
        If EDI_STATUS = "P" Then
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = EDI_PROCESS_IND
        Else
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = "T"
        End If
        rowEDTSYSIH.Item("EDI_OUR_ID") = EDI_OUR_ID
        rowEDTSYSIH.Item("EDI_TP_ID") = EDI_TP_ID
        rowEDTSYSIH.Item("INIT_DATE") = DateTime.Now
        rowEDTSYSIH.Item("INIT_OPER") = ASCMAIN1.USER_ID
        tblEDTSYSIH.Rows.Add(rowEDTSYSIH)

        CreateEDTSYSIH = ediOutboundDocNo

    End Function

    Private Sub CALC_DUE_DATE(ByVal rowTATTERM1 As DataRow, ByVal InvoiceDate As Date,
                      ByRef INV_DUE_DATE As Date, ByRef DISC_DUE_DATE As Date)

        Dim INV_BASE_DATE As Date = CDate(InvoiceDate)

        Select Case rowTATTERM1.Item("TERM_DUE_TYPE") & String.Empty
            Case "D"
                INV_DUE_DATE = INV_BASE_DATE.AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DUE") & ""))

            Case "E"
                Dim ADD_MONTHS_BASE As Integer = 1
                Dim TERM_CUTOFF_DAY As Integer = Val(rowTATTERM1.Item("TERM_CUTOFF_DAY") & "")
                Dim BASE_DD As Integer = Val(Format(INV_BASE_DATE, "dd"))
                Dim TERM_DAYS_DUE As Integer = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & "")
                Dim TERM_ADDL_MOS As Integer = Val(rowTATTERM1.Item("TERM_ADDL_MOS") & "")
                Dim INV_BASE_DATEx As String = Format(INV_BASE_DATE, "MM/dd/yyyy")

                Select Case rowTATTERM1.Item("TERM_EOM_TYPE") & String.Empty
                    Case "F"
                        ASCMAIN1.sql = "Select GLTPARM2.* " _
                         & " from GLTPARM2 " _
                         & " where OPS_YYYYPP = " _
                         & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
                         & "  where GLTPARM2.PRD_END_DATE >= '" & Format(INV_BASE_DATE, "dd-MMM-yyyy") & "')"
                        Dim rowGLTPARM2 As DataRow = ASCDATA1.GetDataRow
                        Dim YYYYMM As String = ASCMAIN1.Get_YYYYMM(rowGLTPARM2.Item("OPS_YYYYPP"), 0)
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case "C"
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case "S"
                        If BASE_DD <= TERM_CUTOFF_DAY _
                        And BASE_DD <= TERM_DAYS_DUE Then
                            ADD_MONTHS_BASE = 0
                        End If
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                End Select
                If TERM_ADDL_MOS > 0 Then
                    INV_DUE_DATE = INV_DUE_DATE.AddMonths(TERM_ADDL_MOS)
                End If
        End Select

        If Val(rowTATTERM1.Item("TERM_DISC_PERC") & "") <> 0 Then
            If rowTATTERM1.Item("TERM_DISC_ELIG_DUE") & String.Empty = "1" Then
                DISC_DUE_DATE = INV_DUE_DATE
            Else
                If Val(rowTATTERM1.Item("TERM_DISC_PERC") & "") <> 0 Then
                    DISC_DUE_DATE = DateValue(INV_DUE_DATE & "").AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DISC") & ""))
                End If
            End If
        End If
    End Sub


End Class
